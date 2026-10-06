"""COM 会话：STA 串行化、可见性铁律、超时与瞬态重试。

这些用假 Application 跑，所以不需要装 PowerPoint，但真实覆盖了线程与队列逻辑。
"""

from __future__ import annotations

import threading
import time

import pytest

from pptd.com import ComSession
from pptd.errors import ComTimeout, ComUnavailable


def test_start_sets_visibility_iron_rule(session: ComSession, fake_app) -> None:
    """拿到实例后必须立刻让窗口可见——Dispatch 默认是不可见的。"""
    info = session.info()
    assert info.alive is True
    assert info.version == "16.0"
    assert fake_app.Visible is True
    assert info.visible is True


def test_run_injects_app(session: ComSession) -> None:
    assert session.run(lambda app: app.Version) == "16.0"


def test_calls_are_serialised(session: ComSession) -> None:
    """两个线程同时调用，结果不能交错（PowerPoint 是单线程的）。"""
    order: list[str] = []
    lock = threading.Lock()

    def slow(app) -> str:
        with lock:
            order.append("in")
        time.sleep(0.05)
        with lock:
            order.append("out")
        return "done"

    results: list[str] = []

    def worker() -> None:
        results.append(session.run(slow))

    threads = [threading.Thread(target=worker) for _ in range(3)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=5)

    assert results == ["done", "done", "done"]
    assert order == ["in", "out"] * 3, "调用发生了交错，STA 串行化被破坏"


def test_timeout_marks_degraded(fake_app) -> None:
    s = ComSession(call_timeout=0.2, app_factory=lambda: fake_app)
    s.start(timeout=5)
    with pytest.raises(ComTimeout):
        s.run(lambda app: time.sleep(1.5), timeout=0.2)
    info = s.info()
    assert info.degraded is True
    assert info.timeouts == 1
    assert info.last_error and "模态" in info.last_error
    s.stop(timeout=3)


def test_transient_error_is_retried(fake_app) -> None:
    """被服务器拒绝（RPC_E_CALL_REJECTED）应自动重试，而不是直接失败。"""
    attempts = {"n": 0}

    class Rejected(Exception):
        hresult = -2147418111

    def flaky(app) -> str:
        attempts["n"] += 1
        if attempts["n"] == 1:
            raise Rejected("调用被拒绝")
        return "ok"

    s = ComSession(app_factory=lambda: fake_app)
    s.start(timeout=5)
    try:
        assert s.run(flaky) == "ok"
        assert attempts["n"] == 2
    finally:
        s.stop(timeout=2)


def test_non_transient_error_propagates(fake_app) -> None:
    s = ComSession(app_factory=lambda: fake_app)
    s.start(timeout=5)

    def boom(app) -> None:
        raise ValueError("不是 COM 的错误，不该重试")

    try:
        with pytest.raises(ValueError):
            s.run(boom)
    finally:
        s.stop(timeout=2)


def test_run_before_start_raises(fake_app) -> None:
    s = ComSession(app_factory=lambda: fake_app)
    with pytest.raises(ComUnavailable):
        s.run(lambda app: app.Version)


def test_stop_marks_not_alive(fake_app) -> None:
    s = ComSession(app_factory=lambda: fake_app)
    s.start(timeout=5)
    assert s.info().alive is True
    s.stop(timeout=3)
    assert s.info().alive is False


def test_app_is_reacquired_when_lost(fake_app) -> None:
    """实例失效（用户关了 PowerPoint）后应自动重取，而不是永久报错。"""
    created = {"n": 0}

    def factory():
        created["n"] += 1
        return fake_app

    s = ComSession(app_factory=factory)
    s.start(timeout=5)
    try:
        s._app = None  # 模拟实例失效
        assert s.run(lambda app: app.Version) == "16.0"
        assert created["n"] >= 2
    finally:
        s.stop(timeout=2)


def test_timed_out_queued_write_never_executes(fake_app):
    s = ComSession(app_factory=lambda: fake_app)
    s.start(timeout=5)
    entered, release = threading.Event(), threading.Event()
    writes = []
    thread = threading.Thread(target=lambda: s.run(lambda app: (entered.set(), release.wait(2)), timeout=3))
    thread.start()
    assert entered.wait(1)
    try:
        with pytest.raises(ComTimeout) as error:
            s.run(lambda app: writes.append("late"), timeout=0.05)
        assert error.value.extra["result_unknown"] is False
        with pytest.raises(ComUnavailable):
            s.run(lambda app: writes.append("new"))
    finally:
        release.set()
        thread.join(3)
        s.stop(timeout=2)
    assert writes == []


def test_partial_write_is_not_replayed(fake_app):
    class Rejected(Exception):
        hresult = -2147418111

    writes = []
    def partial(app):
        writes.append(1)
        raise Rejected("后续属性调用被拒绝")

    s = ComSession(app_factory=lambda: fake_app)
    s.start(timeout=5)
    try:
        with pytest.raises(Rejected):
            s.run(partial, retry_safe=False)
        assert writes == [1]
    finally:
        s.stop(timeout=2)


def test_stop_timeout_cannot_spawn_second_sta(fake_app):
    s = ComSession(app_factory=lambda: fake_app)
    s.start(timeout=5)
    entered, release = threading.Event(), threading.Event()
    thread = threading.Thread(target=lambda: s.run(lambda app: (entered.set(), release.wait(2)), timeout=3))
    thread.start()
    assert entered.wait(1)
    try:
        s.stop(timeout=0.01)
        with pytest.raises(ComUnavailable):
            s.start()
    finally:
        release.set()
        thread.join(3)
        s.stop(timeout=2)


def test_handler_timeout_error_is_not_com_wait_timeout(session):
    def fail(app):
        raise TimeoutError("工具内部异常")
    with pytest.raises(TimeoutError, match="工具内部异常"):
        session.run(fail)
    assert session.info().degraded is False


def test_com_cycles_are_collected_on_sta(fake_app, monkeypatch):
    import gc
    collected = threading.Event()
    owner_threads = []
    class ApartmentObject:
        def __del__(self):
            owner_threads.append(threading.get_ident())
            collected.set()
    was_enabled = gc.isenabled()
    s = ComSession()
    monkeypatch.setattr(s, "_acquire", lambda: setattr(s, "_app", fake_app))
    s.start(timeout=5)
    try:
        def create(app):
            obj = ApartmentObject()
            obj.cycle = obj
            return threading.get_ident()
        sta = s.run(create)
        assert collected.wait(2)
        assert owner_threads == [sta]
        assert gc.isenabled() is False
    finally:
        s.stop(timeout=2)
    assert gc.isenabled() == was_enabled
