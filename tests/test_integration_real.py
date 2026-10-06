"""真机集成测试：守护进程 + HTTP + MCP stdio 全链路。

需要本机装有 PowerPoint。默认不跑（pyproject 里 addopts 排除了 ``com`` 标记）：

    .venv\\Scripts\\python.exe -m pytest -m com -v

**注意 stderr**：MCP 服务器若把 traceback 写进 stderr 而宿主没抽干，
Windows 匿名管道缓冲区会满，服务端阻塞、客户端永久等待。所以这里一律
把子进程 stderr 重定向到文件——这个测试同时是那条坑的回归测试。
"""

from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import time
from pathlib import Path

import pytest

pytestmark = pytest.mark.com

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pptd.config import host_interpreter, settings as cfg  # noqa: E402


def _read_json_line(stream, timeout_s: float = 30.0) -> dict:
    """读一行 JSON-RPC；超时就报明确错误，而不是永久挂住。"""
    import threading

    box: dict = {}

    def reader() -> None:
        line = stream.readline()
        box["line"] = line

    thread = threading.Thread(target=reader, daemon=True)
    thread.start()
    thread.join(timeout_s)
    if thread.is_alive():
        raise AssertionError(f"{timeout_s:.0f}s 内没等到 MCP 响应（很可能是 stderr 被堵住了）")
    line = (box.get("line") or "").strip()
    assert line, "MCP 服务器关闭了 stdout"
    return json.loads(line)


@pytest.fixture(scope="module")
def daemon():
    """确保守护进程可用。

    只停掉**本次测试自己拉起**的那一个：如果用户本来就在用，
    测试结束不该把他的守护进程掐掉。
    """
    from pptd.daemon import ensure, is_alive, stop

    already_running = is_alive(cfg)
    info = ensure(cfg, timeout=40)
    yield info
    if not already_running:
        stop(cfg, force=True)


def test_daemon_health(daemon: dict) -> None:
    import urllib.request

    with urllib.request.urlopen(f"http://{cfg.host}:{cfg.port}/health", timeout=5) as resp:
        body = json.loads(resp.read().decode("utf-8"))
    assert body["ok"] is True
    assert body["com"]["alive"] is True
    assert body["pid"] == daemon["pid"]


def test_daemon_com_is_visible(daemon: dict) -> None:
    """可见性铁律：agent 操作时用户必须看得见 PowerPoint 窗口。"""
    import httpx

    with httpx.Client(timeout=30) as client:
        resp = client.get(
            f"http://{cfg.host}:{cfg.port}/state",
            headers={"X-PPT-Token": daemon["token"]},
        )
    assert resp.status_code == 200
    assert resp.json()["com"]["visible"] is True
    assert resp.json()["com"]["degraded"] is False


def test_duplicate_daemon_cannot_take_com_ownership(daemon: dict) -> None:
    from pptd.daemon import health, runtime_info
    before = runtime_info(cfg)
    assert before is not None
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "pptd", "serve"],
        cwd=str(ROOT), capture_output=True, text=True, encoding="utf-8", timeout=20,
    )
    assert proc.returncode == 1
    assert "守护进程已运行" in proc.stderr
    assert runtime_info(cfg) == before
    assert health(cfg)["pid"] == before["pid"]


def test_mcp_stdio_roundtrip(daemon: dict, tmp_path: Path) -> None:
    """走 DSH 将要走的那条命令：基础解释器 + PYTHONPATH + `-m pptd mcp`。"""
    exe, extra_env = host_interpreter()
    env = dict(os.environ)
    env.update(extra_env)
    stderr_path = tmp_path / "mcp-stderr.log"

    proc = subprocess.Popen(
        [exe, "-X", "utf8", "-m", "pptd", "mcp"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=open(stderr_path, "w", encoding="utf-8"),
        text=True,
        encoding="utf-8",
        bufsize=1,
        cwd=str(ROOT),
        env=env,
    )
    assert proc.stdin and proc.stdout
    try:
        proc.stdin.write(
            json.dumps(
                {
                    "jsonrpc": "2.0",
                    "id": 1,
                    "method": "initialize",
                    "params": {
                        "protocolVersion": "2025-06-18",
                        "capabilities": {},
                        "clientInfo": {"name": "pytest", "version": "0"},
                    },
                }
            )
            + "\n"
        )
        proc.stdin.flush()
        init = _read_json_line(proc.stdout)["result"]
        assert init["serverInfo"]["name"] == "powerpoint"

        proc.stdin.write(json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}) + "\n")
        proc.stdin.flush()

        proc.stdin.write(json.dumps({"jsonrpc": "2.0", "id": 2, "method": "tools/list"}) + "\n")
        proc.stdin.flush()
        tools = _read_json_line(proc.stdout)["result"]["tools"]
        names = {t["name"] for t in tools}
        assert {"ppt_status", "ppt_open", "ppt_activate", "ppt_close"} <= names

        proc.stdin.write(
            json.dumps(
                {
                    "jsonrpc": "2.0",
                    "id": 3,
                    "method": "tools/call",
                    "params": {"name": "ppt_status", "arguments": {}},
                }
            )
            + "\n"
        )
        proc.stdin.flush()
        called = _read_json_line(proc.stdout)["result"]
        assert called["isError"] is False
        payload = json.loads(called["content"][0]["text"])
        assert payload["com"]["alive"] is True

        # 失败路径：必须拿到我们的错误码，而不是笼统的 "unexpected exception"
        proc.stdin.write(
            json.dumps(
                {
                    "jsonrpc": "2.0",
                    "id": 4,
                    "method": "tools/call",
                    "params": {"name": "ppt_open", "arguments": {"path": "C:\\definitely\\missing.pptx"}},
                }
            )
            + "\n"
        )
        proc.stdin.flush()
        failed = _read_json_line(proc.stdout)["result"]
        assert failed["isError"] is True
        assert "path/missing" in failed["content"][0]["text"]
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:  # pragma: no cover
            proc.kill()
            proc.wait()

    # stderr 必须保持安静：长 traceback 会撑爆宿主的管道
    text = stderr_path.read_text(encoding="utf-8", errors="replace")
    assert "Traceback" not in text, f"stderr 里出现了 traceback，可能堵住宿主管道：\n{text[:800]}"


def _field(obj, *names):
    """客户端 SDK 用 snake_case，服务端/线协议用 camelCase；两者都兼容。"""
    for name in names:
        if hasattr(obj, name):
            return getattr(obj, name)
    raise AttributeError(f"{type(obj).__name__} 上找不到 {names}")


def test_official_client_sdk_roundtrip(daemon: dict) -> None:
    """用官方 MCP 客户端 SDK 走一遍——DSH 的 mcp-client 用的就是这套。

    这比裸 JSON-RPC 更接近真实宿主：协议由 SDK 协商（客户端偏好 2026-07-28，
    服务端会回落到自己支持的版本），工具的 inputSchema 也要能被 SDK 解析。

    **自己造演示**：这条测试要 ``ppt_shot slide=2``，也就是需要一份至少 2 页的演示。
    第一版指望"环境里恰好开着别人的文件"——前面的编辑测试把演示一关，它就挂了。
    和看门狗那条一样：**测试的前提要自己建立，不能依赖环境。**
    """
    import asyncio

    import httpx

    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client
    from pptd.config import read_runtime

    token = str(read_runtime(cfg)["token"])
    origin = f"http://{cfg.host}:{cfg.port}"

    def call(name: str, args: dict) -> dict:
        with httpx.Client(timeout=90) as client:
            resp = client.post(
                f"{origin}/call", json={"name": name, "args": args}, headers={"X-PPT-Token": token}
            )
            return resp.json()

    created = call("ppt_new", {})["result"]
    deck_name = created["presentation"]["name"]
    call("ppt_add_slide", {})
    call("ppt_add_slide", {})
    # 给第 2 页放点内容：空白页导出的 JPEG 只有 3 KB 左右，
    # 断言"图片够大"会变成断言"这一页恰好不是空白"——那不是我们要测的东西。
    call("ppt_add_textbox", {"slide": 2, "text": "MCP 往返测试：内联截图必须真的有内容"})


    try:
        exe, extra_env = host_interpreter()
        env = dict(os.environ)
        env.update(extra_env)
        params = StdioServerParameters(
            command=exe,
            args=["-X", "utf8", "-m", "pptd", "mcp"],
            env=env,
            cwd=str(ROOT),
        )

        async def run() -> None:
            async with stdio_client(params) as (read, write):
                async with ClientSession(read, write) as session:
                    init = await session.initialize()
                    assert _field(init, "server_info", "serverInfo").name == "powerpoint"

                    tools = await session.list_tools()
                    names = {t.name for t in tools.tools}
                    assert {"ppt_status", "ppt_open", "ppt_activate", "ppt_close"} <= names
                    open_tool = next(t for t in tools.tools if t.name == "ppt_open")
                    schema = _field(open_tool, "input_schema", "inputSchema")
                    assert schema["required"] == ["path"]
                    assert open_tool.description

                    ok = await session.call_tool("ppt_status", {})
                    assert _field(ok, "is_error", "isError") is False
                    payload = json.loads(ok.content[0].text)
                    assert payload["com"]["alive"] is True
                    assert payload["com"]["visible"] is True

                    bad = await session.call_tool("ppt_open", {"path": "C:\\definitely\\missing.pptx"})
                    assert _field(bad, "is_error", "isError") is True
                    assert "path/missing" in bad.content[0].text

                    # 最关键的一项：截图必须以**图片内容块**回到模型手里，
                    # 否则"agent 能看见幻灯片"就只是句口号。
                    # 注意 ``ppt_shot`` 有**两个**宽度：``width`` 是归档大图，
                    # ``thumb_width`` 才是内联给模型的那张——第一版断言 640 却传了 width，
                    # 结果缩略图仍是默认的 1024。
                    shot = await session.call_tool(
                        "ppt_shot", {"deck": deck_name, "slide": 2, "width": 640, "thumb_width": 512}
                    )
                    assert _field(shot, "is_error", "isError") is False
                    kinds = [getattr(c, "type", None) for c in shot.content]
                    assert "image" in kinds, f"没有内联图片内容块，实际是 {kinds}"
                    image = next(c for c in shot.content if getattr(c, "type", None) == "image")
                    assert _field(image, "mime_type", "mimeType") in ("image/jpeg", "image/png")
                    raw = base64.b64decode(_field(image, "data"))
                    # 判"图有没有坏"要看**能不能解码、尺寸对不对**，不是看字节数：
                    # 空白页的 JPEG 也就 3 KB，字节数只能测出"这页恰好有内容"。
                    from io import BytesIO

                    from PIL import Image

                    with Image.open(BytesIO(raw)) as decoded:
                        assert decoded.width == 512, f"内联缩略图宽度应为 512，实际 {decoded.width}"
                        assert decoded.height > 100
                    assert len(raw) > 2000, f"内联图片只有 {len(raw)} 字节，疑似截断"

        asyncio.run(run())
    finally:
        call("ppt_close", {"deck": deck_name})


def test_mcp_streamable_http_roundtrip(daemon: dict) -> None:
    """MCP over HTTP（streamable-http）—— 这是 ChatGPT 侧要接的那个面。

    守护进程在后台线程里起这个服务，所以先等它监听上来。
    """
    import asyncio
    import socket
    import time as _time

    from mcp import ClientSession
    from mcp.client.streamable_http import create_mcp_http_client, streamable_http_client

    port = int(daemon.get("mcp_port") or (cfg.port + 1))
    deadline = _time.monotonic() + 20
    while _time.monotonic() < deadline:
        with socket.socket() as sock:
            sock.settimeout(0.4)
            if sock.connect_ex((cfg.host, port)) == 0:
                break
        _time.sleep(0.3)
    else:
        raise AssertionError(f"MCP over HTTP 端口 {port} 20 秒内没有监听")

    url = f"http://{cfg.host}:{port}/mcp"
    token = daemon["token"]

    async def run() -> None:
        client = create_mcp_http_client(headers={"X-PPT-Token": token})
        async with streamable_http_client(url, http_client=client) as streams:
            read, write = streams[0], streams[1]
            async with ClientSession(read, write) as session:
                init = await session.initialize()
                assert _field(init, "server_info", "serverInfo").name == "powerpoint"

                tools = await session.list_tools()
                names = {t.name for t in tools.tools}
                assert {"ppt_status", "ppt_open", "ppt_shot", "ppt_set_text"} <= names
                # 两种传输必须给出**同一份** schema，否则模型在不同宿主看到的能力不一致
                open_tool = next(t for t in tools.tools if t.name == "ppt_open")
                assert _field(open_tool, "input_schema", "inputSchema")["required"] == ["path"]

                ok = await session.call_tool("ppt_status", {})
                assert _field(ok, "is_error", "isError") is False
                payload = json.loads(ok.content[0].text)
                assert payload["com"]["alive"] is True

    asyncio.run(run())


def test_mcp_http_requires_token(daemon: dict) -> None:
    """这个端口一旦被隧道到公网，没有令牌就等于任何人都能驱动用户的 PowerPoint。"""
    import urllib.error
    import urllib.request

    port = int(daemon.get("mcp_port") or (cfg.port + 1))
    body = json.dumps({"jsonrpc": "2.0", "id": 1, "method": "tools/list"}).encode("utf-8")
    request = urllib.request.Request(
        f"http://{cfg.host}:{port}/mcp",
        data=body,
        method="POST",
        headers={"Content-Type": "application/json", "Accept": "application/json, text/event-stream"},
    )
    try:
        with urllib.request.urlopen(request, timeout=10) as resp:
            raise AssertionError(f"没有令牌竟然通过了：HTTP {resp.status}")
    except urllib.error.HTTPError as exc:
        assert exc.code == 401, f"期望 401，实际 {exc.code}"


def test_stdio_and_http_expose_same_tools(daemon: dict) -> None:
    """两种传输的工具集必须完全一致——这正是"单一注册表"要保证的事。"""
    import httpx

    with httpx.Client(timeout=20) as client:
        stdio_like = {t["name"] for t in client.get(
            f"http://{cfg.host}:{cfg.port}/tools", headers={"X-PPT-Token": daemon["token"]}
        ).json()["tools"]}

    import pptd.ops  # noqa: F401
    from pptd.registry import all_tools

    assert stdio_like == {t.name for t in all_tools()}


def test_addin_https_serves_taskpane_with_trusted_chain(daemon: dict) -> None:
    """Office 任务窗格的 HTTPS 面：证书能被我们的 CA 验证，页面与 API 同源。

    这一段等价于 WebView 的证书检查——把 CA 装进受信任根之后，WebView 做的就是这件事。
    没生成证书时跳过（这个功能是可选的，不该因为它没启用就让套件变红）。
    """
    import socket
    import ssl
    import time as _time

    import httpx

    from pptd.addin_https import ca_path, cert_status

    status = cert_status(cfg)
    if not status["ready"]:
        pytest.skip("还没生成加载项证书：跑 `pptctl addin-https`")

    port = cfg.addin_bind_port
    deadline = _time.monotonic() + 20
    while _time.monotonic() < deadline:
        with socket.socket() as sock:
            sock.settimeout(0.4)
            if sock.connect_ex((cfg.host, port)) == 0:
                break
        _time.sleep(0.3)
    else:
        raise AssertionError(f"加载项 HTTPS 端口 {port} 没有监听")

    origin = f"https://localhost:{port}"
    token = daemon["token"]
    # 只信任我们这张 CA —— 这正是 WebView 把 CA 装进受信任根之后的行为
    context = ssl.create_default_context(cafile=str(ca_path(cfg)))
    with httpx.Client(verify=context, timeout=20) as client:
        page = client.get(f"{origin}/addin/")
        assert page.status_code == 200, page.text[:200]
        assert "监视台" in page.text
        assert token in page.text, "任务窗格拿不到令牌就用不了"

        manifest = client.get(f"{origin}/addin/manifest.xml")
        assert manifest.status_code == 200 and "<OfficeApp" in manifest.text

        icon = client.get(f"{origin}/addin/icon-32.png")
        assert icon.status_code == 200 and icon.content[:4] == b"\x89PNG"

        # 同源 API：任务窗格靠的就是这个（跨源会被混合内容规则拦掉）
        feed = client.get(f"{origin}/feed.txt", headers={"X-PPT-Token": token})
        assert feed.status_code == 200
        assert feed.text.startswith("PPT-AGENT-FEED v1")


def test_no_orphan_mcp_process(daemon: dict) -> None:
    """杀掉我们拉起的 MCP 进程后不应残留**孤儿**。

    注意不能断言"一个 pptd mcp 进程都没有"——DSH 正常运行时本来就该有一个
    （它的连接管理器会拉起 stdio 子进程）。早期版本就是那么写的，
    等插件真的被加载之后立刻误报。所以要区分：
    - 父进程还在 = 宿主正常管理的子进程，**正常**
    - 父进程没了 = 孤儿，**才是残留**
    """
    from pptd.dsh_bundle import mcp_processes

    before = {int(p.get("ProcessId") or 0) for p in mcp_processes()}

    exe, extra_env = host_interpreter()
    env = dict(os.environ)
    env.update(extra_env)
    proc = subprocess.Popen(
        [exe, "-X", "utf8", "-m", "pptd", "mcp"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        encoding="utf-8",
        cwd=str(ROOT),
        env=env,
    )
    time.sleep(2.0)
    proc.terminate()
    proc.wait(timeout=10)
    time.sleep(1.0)

    after = mcp_processes()
    ours = {int(p.get("ProcessId") or 0) for p in after} - before
    assert not ours, f"我们自己拉起的 MCP 进程没退干净：{ours}"

    orphans = [p for p in after if p.get("orphan")]
    assert not orphans, f"残留了孤儿 MCP 进程：{[p.get('ProcessId') for p in orphans]}"

    # 顺带把"宿主管理的实例"和"孤儿"的区别写进断言里，避免以后再改回去
    for item in after:
        assert "orphan" in item, "mcp_processes() 必须标出 orphan，否则无法区分残留与正常子进程"
