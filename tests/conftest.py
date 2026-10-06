"""测试公共装置：假 COM 会话、假 PowerPoint 实例、临时配置。

这些 fixture 让绝大部分逻辑（注册表、校验、事件、引擎、HTTP 面）能脱离
真实 PowerPoint 跑测试；只有标了 ``@pytest.mark.com`` 的用例才碰真机。
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pptd.com import ComInfo, ComSession  # noqa: E402
from pptd.config import Settings  # noqa: E402
from pptd.engine import Engine  # noqa: E402
from pptd.events import EventBus  # noqa: E402


class FakePowerPoint:
    """够用的假 Application：只有被测试用到的属性。"""

    def __init__(self, version: str = "16.0", presentations: int = 0) -> None:
        self.Version = version
        self.Visible = False          # 初始不可见，用来验证"可见性铁律"确实生效
        self.DisplayAlerts = 2
        self.Presentations = _FakeCollection(presentations)
        self.activated = 0

    def Activate(self) -> None:
        self.activated += 1


class _FakeCollection:
    def __init__(self, count: int) -> None:
        self.Count = count

    def __call__(self, index: int) -> Any:  # pragma: no cover - 仅占位
        raise IndexError(index)


@pytest.fixture()
def settings(tmp_path: Path) -> Settings:
    """指向临时目录的配置，避免污染真实的 %LOCALAPPDATA%。"""
    return Settings(
        host="127.0.0.1",
        port=8791,
        home=tmp_path,
        call_timeout=5.0,
        attach_preferred=True,
    )


@pytest.fixture()
def fake_app() -> FakePowerPoint:
    return FakePowerPoint()


@pytest.fixture()
def session(fake_app: FakePowerPoint) -> ComSession:
    """真 STA 线程 + 真队列，但用假 Application。"""
    s = ComSession(visible=True, alerts_off=True, call_timeout=5.0, app_factory=lambda: fake_app)
    s.start(timeout=5.0)
    yield s
    s.stop(timeout=2.0)


@pytest.fixture()
def bus() -> EventBus:
    return EventBus()


@pytest.fixture()
def engine(settings: Settings, session: ComSession, bus: EventBus) -> Engine:
    import pptd.ops  # noqa: F401  导入即注册工具

    eng = Engine.create(settings, bus=bus, token="test-token", session=session)
    # 测试里把跟速归零：默认 400ms 会让每个写操作用例都变慢
    eng.state["pace_ms"] = 0
    return eng


@pytest.fixture()
def com_info() -> ComInfo:
    return ComInfo(alive=True, attached=True, version="16.0", visible=True)
