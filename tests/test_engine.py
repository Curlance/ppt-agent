"""引擎：把每次调用变成"可见的一步操作"，并给出中文人话摘要。"""

from __future__ import annotations

from typing import Any

import pytest

from pptd.com import ComInfo
from pptd.engine import Engine, describe_result
from pptd.events import EventBus


class FakeDeck:
    """够用的假 Presentation。"""

    def __init__(self, name: str = "demo.pptx", slides: int = 3) -> None:
        self.Name = name
        self.Slides = _Slides(slides)
        self.Saved = True
        self.ReadOnly = False
        self.FullName = f"D:\\decks\\{name}"
        self.Windows = _Slides(1)


class _Slides:
    def __init__(self, count: int) -> None:
        self.Count = count

    def __call__(self, index: int) -> Any:  # pragma: no cover
        raise IndexError(index)


class DeckPowerPoint:
    """带一个演示的假 Application，用于跑通 ppt_open / ppt_status。"""

    def __init__(self) -> None:
        self.Version = "16.0"
        self.Visible = False
        self.DisplayAlerts = 2
        self.Presentations = _Presentations([FakeDeck()])
        self.ActivePresentation = self.Presentations.items[0]

    def Activate(self) -> None:
        pass


class _Presentations:
    def __init__(self, items: list[Any]) -> None:
        self.items = items

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> Any:
        return self.items[index - 1]


class DeadSession:
    """COM 起不来的会话，用来验证降级行为。"""

    def __init__(self) -> None:
        self.info_value = ComInfo(alive=False, last_error="测试：COM 未启动")

    def info(self) -> ComInfo:
        return self.info_value

    def run(self, *args: Any, **kwargs: Any) -> Any:
        raise AssertionError("needs_com=False 的工具不该触发 COM 调用")

    def stop(self) -> None:
        pass


@pytest.fixture()
def deck_engine(settings, bus: EventBus) -> Engine:
    import pptd.ops  # noqa: F401

    from pptd.com import ComSession

    pp = DeckPowerPoint()
    session = ComSession(app_factory=lambda: pp)
    session.start(timeout=5)
    engine = Engine.create(settings, bus=bus, token="t", session=session)
    yield engine
    session.stop(timeout=2)


def test_status_ok_and_human_readable(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_status", {})
    assert result["ok"] is True
    assert "PowerPoint" in result["human"]
    assert result["result"]["com"]["alive"] is True
    # 假实例里有 1 个演示
    assert len(result["result"]["presentations"]) == 1


def test_open_missing_file_fails_cleanly(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_open", {"path": "C:\\definitely\\missing.pptx"})
    assert result["ok"] is False
    assert result["error"]["code"] == "path/missing"


def test_open_relative_path_rejected(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_open", {"path": "relative.pptx"})
    assert result["ok"] is False
    assert result["error"]["code"] == "path/relative"


def test_unknown_tool(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_nope", {})
    assert result["ok"] is False
    assert result["error"]["code"] == "tool/unknown"


def test_events_are_published(settings, bus: EventBus) -> None:
    import pptd.ops  # noqa: F401

    from pptd.com import ComSession

    session = ComSession(app_factory=lambda: DeckPowerPoint())
    session.start(timeout=5)
    engine = Engine.create(settings, bus=bus, token="t", session=session)
    try:
        engine.invoke("ppt_status", {})
    finally:
        session.stop(timeout=2)

    types = [e.type for e in bus.recent(20)]
    assert "op.start" in types and "op.end" in types
    start = next(e for e in bus.recent(20) if e.type == "op.start")
    # 中文人话摘要必须出现在事件里——观察台就是靠它显示"agent 在做什么"
    assert start.data["summary"]
    assert start.data["tool"] == "ppt_status"
    end = next(e for e in bus.recent(20) if e.type == "op.end")
    assert end.data["ok"] is True
    assert end.data["human"]


def test_failure_emits_failed_event(deck_engine: Engine, bus: EventBus) -> None:
    deck_engine.invoke("ppt_open", {"path": "C:\\missing.pptx"})
    end = next(e for e in bus.recent(20) if e.type == "op.end")
    assert end.data["ok"] is False
    assert end.data["error"]["code"] == "path/missing"
    assert "失败" in end.data["human"]


def test_dead_com_still_reports_status(settings, bus: EventBus) -> None:
    """COM 挂掉时 ppt_status 也要能如实汇报——否则用户没法诊断。"""
    import pptd.ops  # noqa: F401

    engine = Engine.create(settings, bus=bus, token="t", session=DeadSession())
    result = engine.invoke("ppt_status", {})
    assert result["ok"] is True
    assert result["result"]["com_error"] == "测试：COM 未启动"
    assert "不可用" in result["human"]


def test_dead_com_blocks_com_tools(settings, bus: EventBus) -> None:
    import pptd.ops  # noqa: F401

    engine = Engine.create(settings, bus=bus, token="t", session=DeadSession())
    result = engine.invoke("ppt_activate", {})
    assert result["ok"] is False
    assert result["error"]["code"] == "com/unavailable"


def test_describe_result_known_tools(settings, bus: EventBus) -> None:
    from pptd.registry import get_tool

    import pptd.ops  # noqa: F401

    tool = get_tool("ppt_open")
    text = describe_result(tool, {"opened": True, "presentation": {"name": "a.pptx", "slides": 8}})
    assert "a.pptx" in text and "8 页" in text

    text2 = describe_result(tool, {"already_open": True, "presentation": {"name": "a.pptx"}})
    assert "已经" in text2 or "本来就开着" in text2


def test_session_state_shape(deck_engine: Engine) -> None:
    state = deck_engine.session_state()
    assert state["tools"] >= 4
    assert "com" in state and "started_at" in state


def test_concurrent_steps_keep_event_order_and_pace(engine, bus, monkeypatch):
    import threading
    import time
    import pptd.engine as module
    from pptd.registry import Tool

    starts = []
    operation = Tool("test_write", "测试写入", "测试", "write", (), lambda app, args, ctx: starts.append(time.monotonic()), needs_com=False)
    monkeypatch.setattr(module, "get_tool", lambda name: operation)
    engine.state["pace_ms"] = 80
    threads = [threading.Thread(target=lambda: engine.invoke("test_write", {})) for _ in range(2)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(2)
    assert len(starts) == 2
    assert starts[1] - starts[0] >= 0.075
    events = [e.type for e in bus.recent(10)]
    assert events == ["op.start", "op.end", "op.start", "op.end"]


def test_expired_engine_queue_never_writes_later(engine, monkeypatch):
    from dataclasses import replace
    import threading
    import pptd.engine as module
    from pptd.registry import Tool

    entered, writes = threading.Event(), []
    def write(app, args, ctx):
        writes.append(1)
        entered.set()
    operation = Tool("test_write", "测试写入", "测试", "write", (), write, needs_com=False)
    monkeypatch.setattr(module, "get_tool", lambda name: operation)
    engine.settings = replace(engine.settings, call_timeout=0.05)
    engine.state["pace_ms"] = 150
    thread = threading.Thread(target=lambda: engine.invoke("test_write", {}))
    thread.start()
    assert entered.wait(1)
    result = engine.invoke("test_write", {})
    thread.join(1)
    assert result["error"]["code"] == "operation/queue-timeout"
    assert writes == [1]
