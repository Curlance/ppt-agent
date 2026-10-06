"""守护进程：发现、拉起、就绪等待、停止。

进程模型（这是"PowerPoint 卡死也不拖垮 agent"的关键）
--------------------------------------------------
``pptd serve`` 是一个**独立进程**，只有它碰 COM。它同时提供 HTTP/SSE。
CLI 与 MCP stdio 转发器都是它的客户端：

    DSH ──stdio──> pptd mcp ──HTTP──> pptd serve ──COM(STA)──> PowerPoint
    用户 ──浏览器──> 观察台 ──SSE──────┘
    ChatGPT ──HTTPS──> (用户自建隧道) ──┘

因此 stdio 转发器重启不会丢 COM 会话，观察台地址也稳定。
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

from . import __version__
from .config import (
    Settings,
    clear_runtime,
    ensure_dirs,
    host_interpreter,
    pinned_token,
    read_runtime,
    settings as default_settings,
    write_runtime,
)

__all__ = ["health", "runtime_info", "is_alive", "spawn", "wait_ready", "ensure", "serve", "stop", "observer_url"]

# Windows 进程创建标志
_DETACHED_PROCESS = 0x00000008
_CREATE_NEW_PROCESS_GROUP = 0x00000200
_CREATE_NO_WINDOW = 0x08000000


def health(cfg: Settings | None = None, timeout: float = 1.5) -> dict[str, Any] | None:
    """探测 /health；不通返回 None。"""
    s = cfg or default_settings
    url = f"{s.base_url}/health"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            if resp.status != 200:
                return None
            return json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, OSError, ValueError):
        return None


def runtime_info(cfg: Settings | None = None) -> dict[str, Any] | None:
    """读取 runtime.json **并**证实它没骗人。

    两道验证：

    1. ``/health`` 必须通（文件可能是上次崩溃留下的）；
    2. **文件里的 pid 必须和 ``/health`` 自报的 pid 一致**。

    第 2 条很重要：runtime.json 是"哪个进程、用什么令牌"的唯一来源。
    如果 pid 对不上（文件被改、或两个守护进程抢过同一个端口），
    照它去 ``taskkill`` 就会**杀掉一个无关的进程**。这种情况一律判为不可用。
    """
    s = cfg or default_settings
    data = read_runtime(s)
    if not data:
        return None
    probe = health(s)
    if probe is None:
        return None
    file_pid = data.get("pid")
    live_pid = probe.get("pid")
    if file_pid is not None and live_pid is not None and int(file_pid) != int(live_pid):
        return None
    return data


def is_alive(cfg: Settings | None = None) -> bool:
    return runtime_info(cfg) is not None


def observer_url(cfg: Settings | None = None) -> str | None:
    """带令牌的观察台地址。"""
    info = runtime_info(cfg)
    if not info:
        return None
    s = cfg or default_settings
    return f"{s.base_url}/?token={info.get('token', '')}"


def spawn(cfg: Settings | None = None) -> int:
    """以脱离方式拉起守护进程，返回 pid。

    注意：这里刻意使用 ``host_interpreter()`` 而不是 ``sys.executable``。
    venv 的 python.exe 是再执行的壳，用它启动会多出一个中间进程；宿主杀掉壳时
    真正的守护进程会变成孤儿。用基础解释器 + PYTHONPATH 可以做到单进程。
    """
    s = cfg or default_settings
    ensure_dirs(s)
    log_path = s.log_dir / "daemon.log"
    exe, extra_env = host_interpreter()
    command = [exe, "-X", "utf8", "-m", "pptd", "serve"]
    env = dict(os.environ)
    env.update(extra_env)
    env["PPT_AGENT_HOME"] = str(s.home)
    env["PPT_AGENT_PORT"] = str(s.port)
    env["PPT_AGENT_HOST"] = s.host
    env["PPT_AGENT_MCP_PORT"] = str(s.mcp_bind_port)
    env["PPT_AGENT_ADDIN_PORT"] = str(s.addin_bind_port)
    env["PPT_AGENT_CALL_TIMEOUT"] = str(s.call_timeout)
    env["PPT_AGENT_ATTACH"] = "1" if s.attach_preferred else "0"
    flags = _DETACHED_PROCESS | _CREATE_NEW_PROCESS_GROUP
    if os.name == "nt":
        flags |= _CREATE_NO_WINDOW
    with open(log_path, "ab", buffering=0) as log:
        log.write(f"\n=== {time.strftime('%Y-%m-%d %H:%M:%S')} spawn ({exe}) ===\n".encode())
        proc = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=log,
            stderr=log,
            cwd=str(Path(__file__).resolve().parent.parent),
            env=env,
            creationflags=flags,
            close_fds=True,
        )
    return proc.pid


def wait_ready(cfg: Settings | None = None, timeout: float = 30.0) -> dict[str, Any]:
    """等待守护进程就绪并返回 runtime 信息。"""
    s = cfg or default_settings
    deadline = time.monotonic() + timeout
    last_error = "未知原因"
    while time.monotonic() < deadline:
        info = read_runtime(s)
        if info is not None and health(s) is not None:
            return info
        log_path = s.log_dir / "daemon.log"
        if log_path.exists():
            try:
                tail = log_path.read_text(encoding="utf-8", errors="replace").strip().splitlines()[-6:]
                last_error = " | ".join(tail) if tail else last_error
            except OSError:
                pass
        time.sleep(0.25)
    raise TimeoutError(f"守护进程 {timeout:.0f}s 内未就绪。日志尾部：{last_error}")


def _kill_pid(pid: int) -> None:
    """强杀一个进程。COM 卡死时这是唯一可靠的收场方式。"""
    if not pid:
        return
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, check=False)
        else:  # pragma: no cover - 本机为 Windows
            import signal

            os.kill(pid, signal.SIGKILL)
    except OSError:
        pass


def ensure(cfg: Settings | None = None, *, auto_spawn: bool = True, timeout: float = 30.0) -> dict[str, Any]:
    """确保守护进程可用：活着就直接返回，否则拉起再等。

    还有一种中间状态要处理：**端口上有人，但 runtime.json 对不上**。
    这时我们既拿不到令牌、也不敢照文件里的 pid 去杀（可能杀错）。
    但能确认的是——占着这个端口的是我们的程序（端口是我们选的），
    所以用 ``/health`` 自报的 pid 把它收回，再起一个干净的。
    """
    s = cfg or default_settings
    info = runtime_info(s)
    if info is not None:
        return info

    probe = health(s)
    if probe is not None and probe.get("pid"):
        # 端口被一个"我们认领不了"的实例占着：先收回，否则新进程绑不上端口
        _kill_pid(int(probe["pid"]))
        deadline = time.monotonic() + 8.0
        while time.monotonic() < deadline and health(s, timeout=0.5) is not None:
            time.sleep(0.2)
        clear_runtime(s)

    if not auto_spawn:
        raise RuntimeError("守护进程未运行（pptctl serve 可手动启动）")
    spawn(s)
    return wait_ready(s, timeout=timeout)


def serve(cfg: Settings | None = None) -> int:
    """只允许一个守护进程拥有本用户会话中的 PowerPoint COM 通道。"""
    if os.name != "nt":
        return _serve(cfg)
    import ctypes
    from ctypes import wintypes

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateMutexW.argtypes = (ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR)
    kernel.CreateMutexW.restype = wintypes.HANDLE
    kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel.CloseHandle.restype = wintypes.BOOL
    handle = kernel.CreateMutexW(None, False, "Local\\ppt-agent-com-owner")
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
        kernel.CloseHandle(handle)
        print("ppt-agent 守护进程已运行；本次启动未接管 COM，也未修改运行时记录。", file=sys.stderr)
        return 1
    try:
        return _serve(cfg)
    finally:
        kernel.CloseHandle(handle)


def _serve(cfg: Settings | None = None) -> int:
    """在**当前进程**里跑守护进程（阻塞）。

    同时起两个 HTTP 服务：

    - ``port``：观察台 + 通用 API + OpenAPI（给人看、给 ChatGPT Actions 用）；
    - ``mcp_bind_port``：MCP over HTTP（streamable-http），给 DSH / ChatGPT 的
      MCP 连接器用。**独立端口**是刻意的，理由见 :class:`~pptd.config.Settings`。
    """
    import uvicorn

    from .config import new_token
    from .engine import Engine
    from .events import EventBus
    from .http_api import create_app
    from . import ops  # noqa: F401  导入即注册工具

    s = cfg or default_settings
    ensure_dirs(s)
    # 默认一次一换；设了 PPT_AGENT_TOKEN 就固定用它（书签/外部工具友好）
    token = pinned_token() or new_token()
    bus = EventBus()
    engine = Engine.create(s, bus=bus, token=token)
    write_runtime(
        {
            "pid": os.getpid(),
            "host": s.host,
            "port": s.port,
            "mcp_port": s.mcp_bind_port,
            "mcp_url": s.mcp_url,
            "token": token,
            "version": __version__,
            "started_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "python": sys.executable,
            "log": str(s.log_dir / "daemon.log"),
        },
        s,
    )
    bus.publish("log", level="info", message=f"ppt-agent {__version__} 守护进程启动，端口 {s.port}")

    info = engine.start()
    bus.publish(
        "log",
        level="info",
        message=(
            f"PowerPoint {info.get('version')} 就绪"
            f"（{'接管已运行实例' if info.get('attached') else '由 pptd 启动'}），"
            f"观察台：http://{s.host}:{s.port}/"
        ),
    )

    app = create_app(engine)
    config = uvicorn.Config(app, host=s.host, port=s.port, log_level="warning", access_log=False)
    server = uvicorn.Server(config)
    app.state.uvicorn_server = server

    # MCP over HTTP：独立端口 + 独立线程，避免子应用 lifespan 与路径重定向的坑
    mcp_server = _start_mcp_http(s, engine, token, bus)
    # Office 加载项任务窗格：必须 HTTPS，且必须与 API **同源**
    addin_server = _start_addin_https(s, app, bus)

    try:
        server.run()
    finally:
        for extra in (mcp_server, addin_server):
            try:
                if extra is not None:
                    extra.should_exit = True
            except Exception:  # noqa: BLE001 - 退出路径不该再抛
                pass
        engine.stop()
        clear_runtime(s)
    return 0


def _start_mcp_http(
    s: Settings, engine: Any, token: str, bus: Any
) -> Any:
    """在后台线程起 MCP over HTTP 服务。失败只记日志，不影响主服务。"""
    import threading

    import uvicorn

    from .mcp_server import DirectCaller, build_http_app

    try:
        mcp_app = build_http_app(DirectCaller(engine), token, host=s.host)
        config = uvicorn.Config(
            mcp_app, host=s.host, port=s.mcp_bind_port, log_level="warning", access_log=False
        )
        server = uvicorn.Server(config)
        thread = threading.Thread(target=server.run, name="pptd-mcp-http", daemon=True)
        thread.start()
        bus.publish(
            "log",
            level="info",
            message=f"MCP over HTTP 已启动：{s.mcp_url}（需要令牌）",
        )
        return server
    except Exception as exc:  # noqa: BLE001 - MCP 端口起不来不该拖垮观察台
        bus.publish("log", level="error", message=f"MCP over HTTP 启动失败：{type(exc).__name__}: {exc}")
        return None


def _start_addin_https(s: Settings, app: Any, bus: Any) -> Any:
    """在后台线程起 Office 加载项任务窗格的 HTTPS 服务。

    **只在证书已生成时才启动**——没证书就当这个功能不存在，其它一切照常。
    用独立的 HTTPS 端口是必须的：任务窗格页面与它要调的 API 必须同源，
    否则 HTTPS 页面去 fetch HTTP 会被混合内容拦掉。
    """
    import threading

    import uvicorn

    from .addin_https import cert_status, leaf_cert_path, leaf_key_path

    try:
        status = cert_status(s)
    except Exception as exc:  # noqa: BLE001 - 证书检查失败不该拖垮主服务
        bus.publish("log", level="error", message=f"加载项证书检查失败：{exc}")
        return None

    if not status["ready"]:
        bus.publish(
            "log",
            level="info",
            message="Office 加载项未启用（还没生成证书）：跑 `pptctl addin-https` 即可",
        )
        return None

    try:
        config = uvicorn.Config(
            app,
            host=s.host,
            port=s.addin_bind_port,
            log_level="warning",
            access_log=False,
            ssl_certfile=str(leaf_cert_path(s)),
            ssl_keyfile=str(leaf_key_path(s)),
        )
        server = uvicorn.Server(config)
        threading.Thread(target=server.run, name="pptd-addin-https", daemon=True).start()
        trusted = "已信任" if status["trusted"] else "**CA 尚未装进受信任根**"
        bus.publish(
            "log",
            level="info",
            message=f"Office 加载项任务窗格已启动：{s.addin_url}（{trusted}）",
        )
        return server
    except Exception as exc:  # noqa: BLE001
        bus.publish("log", level="error", message=f"加载项 HTTPS 启动失败：{type(exc).__name__}: {exc}")
        return None


def stop(cfg: Settings | None = None, *, force: bool = False, timeout: float = 8.0) -> dict[str, Any]:
    """停止守护进程。``force`` 用强杀——COM 卡死时唯一可靠的收场方式。

    **只杀 ``/health`` 自报的那个 pid。** 文件里的 pid 只用来判断"能不能发优雅停止"，
    绝不直接拿去 ``taskkill``——文件可能陈旧或被改，照它杀会误伤无关进程。
    """
    s = cfg or default_settings
    file_info = read_runtime(s)
    probe = health(s, timeout=2.0)
    live_pid = int(probe["pid"]) if probe and probe.get("pid") else 0

    if probe is None and not file_info:
        return {"stopped": False, "reason": "没有运行中的守护进程"}

    graceful = False
    if not force and probe is not None and file_info and int(file_info.get("pid", 0)) == live_pid:
        try:
            request = urllib.request.Request(f"{s.base_url}/shutdown", method="POST")
            request.add_header("X-PPT-Token", str(file_info.get("token", "")))
            with urllib.request.urlopen(request, timeout=3) as resp:
                graceful = resp.status == 200
        except (urllib.error.URLError, OSError):
            graceful = False

    deadline = time.monotonic() + (timeout if not force else 2.0)
    while time.monotonic() < deadline and health(s, timeout=0.5) is not None:
        time.sleep(0.2)

    alive = health(s, timeout=0.5) is not None
    if alive and live_pid:
        _kill_pid(live_pid)
        time.sleep(0.4)
        alive = health(s, timeout=0.5) is not None

    if not alive:
        clear_runtime(s)
    return {
        "stopped": not alive,
        "graceful": graceful,
        "pid": live_pid,
        "record_matched": bool(file_info and int(file_info.get("pid", 0)) == live_pid),
    }
