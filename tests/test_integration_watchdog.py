"""真机集成测试：守护进程被强杀后，能否干净恢复，且**不伤用户的 PowerPoint**。

这是整个架构最关键的承诺——"PowerPoint 卡死也不拖垮 agent"。它由三件事支撑：

1. COM 跑在**独立进程**里（卡死可强杀，agent 不受影响）；
2. 每请求有超时，卡住时把会话标记为 degraded（而不是永久装死）；
3. 强杀后留下的**陈旧 runtime.json 必须被证伪**，`ensure()` 能重新拉起。

单元测试只能验到第 2 条（用假 COM + 假超时）。第 1、3 条以及"用户的演示还在不在"
必须在真机上验——这正是本文件存在的理由。

    .venv\\Scripts\\python.exe -m pytest -m com -v tests/test_integration_watchdog.py
"""

from __future__ import annotations

import json
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

import pytest

pytestmark = pytest.mark.com

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pptd.config import read_runtime, settings as cfg  # noqa: E402
from pptd.daemon import ensure, health, is_alive, runtime_info, stop  # noqa: E402


def _wait_until(predicate, timeout: float = 20.0, interval: float = 0.3) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return True
        time.sleep(interval)
    return False


def _call(name: str, args: dict | None = None, timeout: float = 60.0) -> dict:
    """每次都重新读令牌——守护进程重启后令牌会变。"""
    info = read_runtime(cfg)
    assert info is not None, "runtime.json 不见了"
    body = json.dumps({"name": name, "args": args or {}}).encode("utf-8")
    request = urllib.request.Request(
        f"http://{cfg.host}:{cfg.port}/call",
        data=body,
        method="POST",
        headers={"Content-Type": "application/json", "X-PPT-Token": info["token"]},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        return json.loads(exc.read().decode("utf-8"))


def _kill(pid: int) -> None:
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, check=False)


@pytest.fixture(scope="module")
def running_daemon():
    info = ensure(cfg, timeout=40)
    yield info
    ensure(cfg, timeout=40)  # 无论测试怎么折腾，结束时保证有守护进程在跑


def test_force_kill_recovers_without_touching_powerpoint(running_daemon: dict) -> None:
    """强杀守护进程 → 陈旧记录被证伪 → 重新拉起 → **正在编辑的演示原封不动**。

    「原封不动」需要一个参照物。这里**自己造一份**演示（``ppt_new``，不保存、不碰用户文件），
    而不是指望环境里恰好开着别人的文件——第一版就是这么依赖环境的，
    前面的测试把演示关掉之后它就挂了。**测试的前提要自己建立。**
    """
    info = runtime_info(cfg)
    assert info is not None, "前置条件：守护进程必须在跑"
    old_pid = int(info["pid"])

    # 造一份属于这次测试的演示，并在里面留下"用户正在做的事"。
    # 注意 ``ppt_new`` 给的是**0 页**的空演示，所以页码要用 ``add_slide`` 的返回值，
    # 不能想当然写死（第一版写死 2，结果插出来的其实是第 1 页，写文字就 400 了）。
    created = _call("ppt_new")["result"]
    name = created["presentation"]["name"]
    added = _call("ppt_add_slide", {})["result"]
    slide_no = int(added["slide"])
    _call("ppt_add_textbox", {"slide": slide_no, "text": "看门狗测试：这段文字必须活下来"})
    _call("ppt_save", {"mode": "copy"})  # 留个备份，别让它因为"未保存"而弹窗

    before = _call("ppt_decks")["result"]
    decks_before = [(d["name"], d["slides"]) for d in before["decks"]]
    assert decks_before, "前置条件：PowerPoint 里要有打开的演示，否则测不出'没伤到'"
    assert any(d["name"] == name for d in before["decks"]), f"自己造的演示没出现在列表里：{before}"

    # ---- 强杀（COM 卡死时唯一可靠的收场方式）----
    _kill(old_pid)
    assert _wait_until(lambda: health(cfg, timeout=0.5) is None, 20), "进程没被杀死"

    # ---- 陈旧记录必须被证伪，而不是被当真 ----
    stale = read_runtime(cfg)
    assert stale is not None, "强杀不走优雅退出，runtime.json 会留下（这是预期行为）"
    assert int(stale["pid"]) == old_pid
    assert runtime_info(cfg) is None, "陈旧 runtime.json 必须被 /health 证伪"
    assert is_alive(cfg) is False

    # ---- 重新拉起：必须能与陈旧记录区分开 ----
    fresh = ensure(cfg, timeout=40)
    assert int(fresh["pid"]) != old_pid, "拉起的必须是新进程"
    assert fresh["token"] != stale["token"], "新进程应当签发新令牌"
    probe = health(cfg, timeout=5)
    assert probe is not None and probe["ok"] is True

    # ---- 最关键的一条：PowerPoint 和里面那份演示完全没受影响 ----
    after = _call("ppt_decks")["result"]
    assert [(d["name"], d["slides"]) for d in after["decks"]] == decks_before, (
        "守护进程被强杀后，演示列表发生了变化——独立进程隔离失效"
    )

    # 光比"名字和页数"还不够：内容也得原样留着。
    # 这一段是强杀之前写进去的，强杀之后必须一个字符都不差。
    # 注意要**显式指名**那份演示：守护进程重启后"活动演示"未必还是它。
    page = _call("ppt_read_slide", {"deck": name, "slide": slide_no})["result"]
    texts = " ".join(
        shape.get("text") or "" for shape in page.get("shapes", [])
    )
    assert "这段文字必须活下来" in texts, f"演示内容被破坏了：{texts!r}"

    # 收尾：关掉自己造的那份，别留给用户
    _call("ppt_close", {"deck": name})


def test_daemon_survives_being_killed_mid_flight(running_daemon: dict) -> None:
    """在请求进行中强杀守护进程：调用方拿到的是明确的失败，而不是挂起。"""
    info = runtime_info(cfg)
    assert info is not None
    pid = int(info["pid"])

    import threading

    result: dict = {}

    def fire() -> None:
        try:
            result["response"] = _call("ppt_status", timeout=15)
        except Exception as exc:  # noqa: BLE001 - 连接被切断也在预期内
            result["error"] = f"{type(exc).__name__}: {exc}"

    thread = threading.Thread(target=fire, daemon=True)
    thread.start()
    time.sleep(0.15)
    _kill(pid)
    thread.join(timeout=20)

    assert not thread.is_alive(), "守护进程被杀后，调用方挂住了——这不可接受"
    assert "response" in result or "error" in result
    # 无论哪种结果，进程必须是死的，且能被重新拉起
    assert _wait_until(lambda: health(cfg, timeout=0.5) is None, 20)
    ensure(cfg, timeout=40)


def test_cli_reports_clearly_when_daemon_is_down(running_daemon: dict) -> None:
    """守护进程不在时，CLI 要给一句人话 + 非零退出码，而不是抛 traceback。"""
    stop(cfg, force=True)
    assert _wait_until(lambda: health(cfg, timeout=0.5) is None, 20)

    proc = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "pptd", "status", "--no-spawn"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(ROOT),
        timeout=60,
    )
    assert proc.returncode == 1, f"守护进程不在时应当非零退出，实际 {proc.returncode}"
    assert "守护进程未运行" in (proc.stdout or ""), proc.stdout
    assert "Traceback" not in (proc.stderr or ""), f"不该抛异常栈：\n{proc.stderr}"

    # 恢复现场
    ensure(cfg, timeout=40)


def test_stale_runtime_is_not_mistaken_for_alive(running_daemon: dict) -> None:
    """runtime.json 被改坏（pid 对不上）时，不能信它——否则 stop 会杀错进程。

    这条描述的是一种真实可达的状态：文件被改、或两个守护进程抢过同一个端口。
    正确行为是：认领不了就**收回端口重新起一个干净的**，绝不照文件里的 pid 去杀。
    """
    info = runtime_info(cfg)
    assert info is not None
    live_pid = int(info["pid"])

    from pptd.config import write_runtime

    fake = dict(info)
    fake["pid"] = 999_999_999  # 可能对应着一个**无关的**真实进程
    fake["token"] = "stale-token-should-not-be-used"
    write_runtime(fake, cfg)

    try:
        # /health 仍然通（真进程还在），但文件与实况对不上 → 必须判为不可用
        assert health(cfg, timeout=2) is not None
        assert runtime_info(cfg) is None, "pid 对不上的记录必须被证伪"

        # stop 只能杀 /health 自报的 pid；文件里的 999999999 碰都不能碰
        result = stop(cfg, force=True)
        assert result["pid"] != 999_999_999, "绝不能照陈旧记录去杀进程"
        assert result["record_matched"] is False

        # 收回之后必须能起来一个干净、可鉴权的实例
        fresh = ensure(cfg, timeout=40)
        assert int(fresh["pid"]) not in (999_999_999, live_pid)
        assert _call("ppt_status")["ok"] is True
    finally:
        ensure(cfg, timeout=40)
