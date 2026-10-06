"""stdio 转发层：**令牌轮换**与错误翻译。

起因是一个真实故障：用户在 DSH 里能看见 ``mcp__ppt__*`` 工具，但每个调用都返回
"未知错误"。根因是 :class:`ForwardingCaller` 在**启动时**把令牌写死进
``httpx.Client`` 的默认头里，而 MCP stdio 进程由宿主拉起、会活很久——
中间只要跑过一次 ``pptctl restart``，守护进程就换了新令牌，
这个长驻进程于是永远拿着旧令牌，全部工具失效，非得再重启一次宿主才能恢复。

这里用一个**假的守护进程**（标准库 http.server）把这件事钉死：
令牌换了之后，同一个 caller 必须自己重读、自己恢复。
"""

from __future__ import annotations

import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest

from pptd.config import Settings, write_runtime
from pptd.mcp_server import ForwardingCaller


class FakeDaemon:
    """只认一个令牌的假守护进程；可以随时换令牌。"""

    def __init__(self) -> None:
        self.token = "token-one"
        self.calls = 0
        self.seen_tokens: list[str] = []
        self.mode = "ok"          # ok | 401 | detail | garbage | unreachable
        outer = self

        class Handler(BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"

            def log_message(self, *args: object) -> None:  # 静音
                pass

            def _respond(self, status: int, payload: bytes, content_type: str = "application/json") -> None:
                self.send_response(status)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)

            def do_GET(self) -> None:  # noqa: N802
                outer.calls += 1
                outer.seen_tokens.append(self.headers.get("X-PPT-Token", ""))
                if outer.token and self.headers.get("X-PPT-Token") != outer.token:
                    self._respond(401, json.dumps({"detail": "令牌不正确"}).encode())
                    return
                self._respond(200, b"PNGDATA", "image/png")

            def do_POST(self) -> None:  # noqa: N802
                outer.calls += 1
                outer.seen_tokens.append(self.headers.get("X-PPT-Token", ""))
                self.rfile.read(int(self.headers.get("Content-Length") or 0))
                if outer.mode == "detail":
                    self._respond(500, json.dumps({"detail": "工具内部炸了"}, ensure_ascii=False).encode())
                    return
                if outer.mode == "garbage":
                    self._respond(200, b"<html>not json</html>", "text/html")
                    return
                if outer.token and self.headers.get("X-PPT-Token") != outer.token:
                    self._respond(401, json.dumps({"detail": "令牌不正确"}).encode())
                    return
                body = json.dumps({"ok": True, "tool": "ppt_status", "result": {"seen": "yes"}}).encode()
                self._respond(200, body)

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.port = self.server.server_address[1]
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def stop(self) -> None:
        self.server.shutdown()
        self.server.server_close()


@pytest.fixture()
def fake_daemon():
    daemon = FakeDaemon()
    yield daemon
    daemon.stop()


@pytest.fixture()
def runtime_settings(tmp_path: Path, fake_daemon: FakeDaemon) -> Settings:
    settings = Settings(
        host="127.0.0.1",
        port=fake_daemon.port,
        home=tmp_path / "home",
        call_timeout=5.0,
        attach_preferred=True,
    )
    settings.home.mkdir(parents=True, exist_ok=True)
    write_runtime({"pid": 1, "token": "token-one", "port": fake_daemon.port}, settings)
    return settings


# --------------------------------------------------------------------------
# 令牌轮换：这条就是那个真实故障
# --------------------------------------------------------------------------


def test_caller_recovers_after_token_rotation(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    """守护进程重启换了令牌，**同一个长驻 caller** 必须自己跟上。"""
    caller = ForwardingCaller(runtime_settings, {"token": "token-one"})
    assert caller.call("ppt_status", {})["ok"] is True

    # 模拟 pptctl restart：令牌变了（runtime.json 与守护进程同时换）
    fake_daemon.token = "token-two"
    write_runtime({"pid": 2, "token": "token-two", "port": fake_daemon.port}, runtime_settings)

    result = caller.call("ppt_status", {})
    assert result["ok"] is True, f"令牌轮换后调用失败：{result}"
    assert "token-two" in fake_daemon.seen_tokens


def test_caller_survives_many_rotations(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    caller = ForwardingCaller(runtime_settings, {"token": "token-one"})
    for index in range(2, 6):
        token = f"token-{index}"
        fake_daemon.token = token
        write_runtime({"pid": index, "token": token, "port": fake_daemon.port}, runtime_settings)
        assert caller.call("ppt_status", {})["ok"] is True, f"第 {index} 次轮换失败"


def test_caller_reads_token_fresh_every_call(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    """不是"失败才刷新"，而是每次调用都现读——更简单也更难写错。"""
    caller = ForwardingCaller(runtime_settings, {})
    write_runtime({"pid": 1, "token": "later-token", "port": fake_daemon.port}, runtime_settings)
    fake_daemon.token = "later-token"
    assert caller.call("ppt_status", {})["ok"] is True
    assert fake_daemon.seen_tokens[-1] == "later-token"


def test_call_uses_no_stale_token_when_started_empty(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    """启动时 runtime.json 恰好读不到，也不该永久瘫痪。"""
    caller = ForwardingCaller(runtime_settings, {})
    fake_daemon.token = "token-one"
    assert caller.call("ppt_status", {})["ok"] is True


# --------------------------------------------------------------------------
# 错误翻译：不许再出现"未知错误"
# --------------------------------------------------------------------------


def test_unauthorized_is_reported_with_a_reason(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    """两次都 401 时要给出**人能看懂**的原因，而不是让上层变成"未知错误"。"""
    write_runtime({"pid": 1, "token": "wrong-token", "port": fake_daemon.port}, runtime_settings)
    fake_daemon.token = "the-real-token"
    caller = ForwardingCaller(runtime_settings, {"token": "wrong-token"})

    result = caller.call("ppt_status", {})
    assert result["ok"] is False
    assert result["error"]["code"] == "auth/unauthorized"
    assert "令牌" in result["error"]["message"]
    assert "pptctl" in result["error"]["message"], "要给出下一步该跑什么"


def test_fastapi_detail_is_surfaced_not_swallowed(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    """FastAPI 的错误体是 ``{"detail": ...}``，必须把它翻出来。"""
    fake_daemon.mode = "detail"
    caller = ForwardingCaller(runtime_settings, {"token": "token-one"})
    result = caller.call("ppt_status", {})
    assert result["ok"] is False
    assert "工具内部炸了" in result["error"]["message"]
    assert result["error"]["code"].startswith("api/http-")


def test_non_json_response_is_reported(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    fake_daemon.mode = "garbage"
    caller = ForwardingCaller(runtime_settings, {"token": "token-one"})
    result = caller.call("ppt_status", {})
    assert result["ok"] is False
    assert result["error"]["code"] in {"daemon/bad-response", "api/http-200"}


def test_unreachable_daemon_is_reported(runtime_settings: Settings) -> None:
    """守护进程没起来时要说"连不上"，而不是抛异常。"""
    caller = ForwardingCaller(runtime_settings, {"token": "x"})
    caller.base_url = "http://127.0.0.1:1"  # 几乎肯定没人监听
    result = caller.call("ppt_status", {})
    assert result["ok"] is False
    assert result["error"]["code"] == "daemon/unreachable"


def test_ensure_hook_tries_to_start_the_daemon(runtime_settings: Settings) -> None:
    """连不上时应当顺手把守护进程拉起来，再把失败说清楚。"""
    attempts: list[int] = []

    def hook() -> None:
        attempts.append(1)

    caller = ForwardingCaller(runtime_settings, {"token": "x"}, ensure_hook=hook)
    caller.base_url = "http://127.0.0.1:1"
    result = caller.call("ppt_status", {})
    assert attempts, "应当尝试拉起守护进程"
    assert result["ok"] is False


# --------------------------------------------------------------------------
# 取图也要带对令牌
# --------------------------------------------------------------------------


def test_fetch_image_uses_the_current_token(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    caller = ForwardingCaller(runtime_settings, {"token": "stale"})
    write_runtime({"pid": 1, "token": "token-one", "port": fake_daemon.port}, runtime_settings)
    fake_daemon.token = "token-one"
    assert caller.fetch_image("/shots/a.png") == b"PNGDATA"
    assert fake_daemon.seen_tokens[-1] == "token-one"


def test_fetch_image_returns_none_on_401(runtime_settings: Settings, fake_daemon: FakeDaemon) -> None:
    fake_daemon.token = "the-real-token"
    write_runtime({"pid": 1, "token": "wrong", "port": fake_daemon.port}, runtime_settings)
    caller = ForwardingCaller(runtime_settings, {"token": "wrong"})
    assert caller.fetch_image("/shots/a.png") is None


# --------------------------------------------------------------------------
# 固定令牌（可选）
# --------------------------------------------------------------------------


def test_pinned_token_from_env(monkeypatch) -> None:
    from pptd.config import pinned_token

    monkeypatch.delenv("PPT_AGENT_TOKEN", raising=False)
    assert pinned_token() == ""
    monkeypatch.setenv("PPT_AGENT_TOKEN", "  fixed-token  ")
    assert pinned_token() == "fixed-token"


def test_lost_response_never_replays_write(runtime_settings, monkeypatch):
    import httpx

    attempts, restarts = [], []
    class BrokenClient:
        def request(self, *args, **kwargs):
            attempts.append(1)
            raise httpx.ReadTimeout("已发送，但没有收到响应")

    caller = ForwardingCaller(runtime_settings, {"token": "t"}, ensure_hook=lambda: restarts.append(1))
    monkeypatch.setattr(caller, "_client_for", lambda token: BrokenClient())
    result = caller.call("ppt_add_slide", {})
    assert result["error"]["code"] == "daemon/result-unknown"
    assert attempts == [1]
    assert restarts == []
