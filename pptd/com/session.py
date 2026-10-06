"""单 STA 线程的 PowerPoint 会话。

线程模型
--------
主线程/HTTP 线程通过 :meth:`ComSession.run` 提交任务，任务被投递到一条专用线程；
那条线程先 ``pythoncom.CoInitialize()`` 进入 STA，再 ``Dispatch``/``GetActiveObject``
拿到 ``Application``，此后**只有它**能碰这个对象。提交方通过 ``Future`` 等结果，
可带超时。

``run(fn, *args)`` 中的 ``fn`` 形如 ``def op(app, args, ctx)``，``app`` 由工作线程注入。
"""

from __future__ import annotations

import queue
import gc
import threading
import time
from concurrent.futures import Future
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Any, Callable

from ..errors import ComError, ComTimeout, ComUnavailable

__all__ = ["ComSession", "ComInfo"]

#: 可以安全重试的 COM 瞬态错误（被拒绝 / 稍后重试）。
TRANSIENT_HRESULTS = frozenset(
    {
        -2147418111,  # RPC_E_CALL_REJECTED  调用被拒绝（服务器忙）
        -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
        -2147417851,  # RPC_E_SERVERCALL_RETRYLATER 变体
        -2147418105,  # CO_E_OBJNOTCONNECTED 变体（重取实例）
    }
)

#: PowerPoint 的 PpAlertLevel 常量。
PP_ALERTS_NONE = 1

_WORKER_STOP = object()
_GC_LOCK = threading.Lock()
_GC_USERS = 0
_GC_WAS_ENABLED = False


def _own_com_gc() -> None:
    """动态 COM 包装器的循环引用必须在已初始化 COM 的 STA 上回收。"""
    global _GC_USERS, _GC_WAS_ENABLED
    with _GC_LOCK:
        if _GC_USERS == 0:
            _GC_WAS_ENABLED = gc.isenabled()
            gc.disable()
        _GC_USERS += 1


def _release_com_gc() -> None:
    global _GC_USERS
    with _GC_LOCK:
        _GC_USERS -= 1
        if _GC_USERS == 0 and _GC_WAS_ENABLED:
            gc.enable()


def _detach_tracebacks(exc: BaseException) -> BaseException:
    # Future 不能把持有 COM 局部变量的工作线程 traceback 交给 HTTP/主线程。
    seen = set()
    pending = [exc]
    while pending:
        current = pending.pop()
        if id(current) in seen:
            continue
        seen.add(id(current))
        current.__traceback__ = None
        for linked in (current.__context__, current.__cause__):
            if linked is not None and id(linked) not in seen:
                pending.append(linked)
    return exc


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


@dataclass
class ComInfo:
    """会话对外暴露的状态快照。"""

    alive: bool = False
    attached: bool = False
    version: str | None = None
    visible: bool | None = None
    degraded: bool = False
    acquired_at: str | None = None
    last_error: str | None = None
    last_call_ms: float | None = None
    calls: int = 0
    timeouts: int = 0

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class _Job:
    fn: Callable[..., Any]
    args: tuple[Any, ...]
    kwargs: dict[str, Any]
    retry_safe: bool = True
    future: Future[Any] = field(default_factory=Future)


class ComSession:
    """与 PowerPoint 的唯一通道。"""

    def __init__(
        self,
        *,
        visible: bool = True,
        alerts_off: bool = False,
        attach_preferred: bool = True,
        call_timeout: float = 60.0,
        logger: Callable[[str], None] | None = None,
        app_factory: Callable[[], Any] | None = None,
    ) -> None:
        self.visible = visible
        self.alerts_off = alerts_off
        self.attach_preferred = attach_preferred
        self.call_timeout = call_timeout
        self._log = logger or (lambda _msg: None)
        #: 仅测试使用：替换掉真实的 COM 获取过程，好让 STA/队列/超时逻辑能被单测覆盖。
        self._app_factory = app_factory

        self._queue: queue.Queue[Any] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._app: Any = None
        self._info = ComInfo()
        self._lock = threading.Lock()
        self._stopping = False

    # -- 生命周期 -----------------------------------------------------------

    def start(self, timeout: float = 30.0) -> ComInfo:
        """启动工作线程并获取 Application。失败也能返回状态（不抛）。"""
        if self._thread is not None and self._thread.is_alive():
            if self._stopping:
                raise ComUnavailable("上一条 COM 线程仍在停止，请重启守护进程。")
            return self.info()
        with self._lock:
            self._stopping = False
            self._info = ComInfo()
            self._queue = queue.Queue()
        self._thread = threading.Thread(target=self._worker, name="pptd-com-sta", daemon=True)
        self._thread.start()
        # 工作线程启动后会在队列上等活；这里跑一个空转任务，确认它已拿到实例。
        try:
            self.run(lambda app: app.Version, timeout=timeout)
        except (ComError, ComTimeout):
            pass
        return self.info()

    def stop(self, timeout: float = 5.0) -> None:
        """请求工作线程退出。不主动关闭用户的 PowerPoint。"""
        if self._thread is None:
            return
        with self._lock:
            self._stopping = True
            self._queue.put(_WORKER_STOP)
        self._thread.join(timeout=timeout)
        if not self._thread.is_alive():
            self._thread = None
        with self._lock:
            self._info.alive = False

    def info(self) -> ComInfo:
        with self._lock:
            return ComInfo(**asdict(self._info))

    # -- 调用 ---------------------------------------------------------------

    def run(self, fn: Callable[..., Any], /, *args: Any, timeout: float | None = None, retry_safe: bool = True, **kwargs: Any) -> Any:
        """在 STA 线程上执行 ``fn(app, *args, **kwargs)`` 并等待结果。

        超时抛 :class:`ComTimeout`，同时把会话标记为 degraded——因为任务还卡在
        COM 里，后续调用不可信。
        """
        future: Future[Any] = Future()
        with self._lock:
            if self._thread is None or not self._thread.is_alive() or self._stopping:
                raise ComUnavailable("COM 会话未启动或正在停止")
            if self._info.degraded:
                raise ComUnavailable("COM 会话已因超时降级，请处理 PowerPoint 对话框并运行 pptctl restart。")
            self._queue.put(_Job(fn=fn, args=args, kwargs=kwargs, future=future, retry_safe=retry_safe))
        deadline = self.call_timeout if timeout is None else timeout
        try:
            return future.result(timeout=deadline)
        except TimeoutError as exc:
            # 工具自身抛 TimeoutError 时，Future 已完成；不能误判为等待超时。
            if future.done():
                return future.result()
            cancelled = future.cancel()
            with self._lock:
                self._info.degraded = True
                self._info.timeouts += 1
                self._info.last_error = f"调用超时（{deadline:.0f}s），PowerPoint 可能弹出了模态对话框"
            raise ComTimeout(
                self._info.last_error or "COM 调用超时", timeout_s=deadline,
                result_unknown=not cancelled,
                recovery="尚未执行的操作已取消；正在执行的 COM 调用无法中断，可能仍然完成。请先检查演示，再重启守护进程，勿直接重试写入。",
            ) from exc
        except Exception as exc:  # 透传工具层异常
            raise exc

    # -- 工作线程 -----------------------------------------------------------

    def _worker(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        manages_gc = self._app_factory is None
        if manages_gc:
            _own_com_gc()
        try:
            self._acquire()
            with self._lock:
                self._info.alive = True
            self._log(f"[com] PowerPoint {self._info.version} 就绪（attached={self._info.attached}）")
            while True:
                job = self._queue.get()
                if job is _WORKER_STOP:
                    break
                assert isinstance(job, _Job)
                with self._lock:
                    if not job.future.set_running_or_notify_cancel():
                        continue
                    blocked = self._stopping or self._info.degraded
                if blocked:
                    job.future.set_exception(ComUnavailable("会话停止或降级，排队操作未执行。"))
                    continue
                try:
                    job.future.set_result(self._dispatch(job))
                except BaseException as exc:  # noqa: BLE001 - 必须原样回传给调用方
                    job.future.set_exception(_detach_tracebacks(exc))
                finally:
                    del job
                    if manages_gc:
                        gc.collect()
        except BaseException as exc:  # noqa: BLE001 - 拿到实例就失败了
            with self._lock:
                self._info.last_error = f"{type(exc).__name__}: {exc}"
            self._log(f"[com] 初始化失败：{self._info.last_error}")
        finally:
            self._release()
            with self._lock:
                self._info.alive = False
                while not self._queue.empty():
                    pending = self._queue.get_nowait()
                    if isinstance(pending, _Job) and pending.future.set_running_or_notify_cancel():
                        pending.future.set_exception(ComUnavailable(self._info.last_error or "COM 会话已停止"))
            if manages_gc:
                gc.collect()
            pythoncom.CoUninitialize()
            if manages_gc:
                _release_com_gc()

    def _dispatch(self, job: _Job) -> Any:
        started = time.perf_counter()
        last_exc: BaseException | None = None
        for attempt in range(3):
            try:
                self._ensure_app()
                result = job.fn(self._app, *job.args, **job.kwargs)
                with self._lock:
                    self._info.calls += 1
                    self._info.last_call_ms = round((time.perf_counter() - started) * 1000, 1)
                    if not self._info.degraded:
                        self._info.last_error = None
                return result
            except BaseException as exc:  # noqa: BLE001
                last_exc = exc
                if job.retry_safe and self._is_transient(exc) and attempt < 2:
                    time.sleep(0.25 * (attempt + 1))
                    self._app = None  # 强制重取实例
                    continue
                with self._lock:
                    self._info.last_error = f"{type(exc).__name__}: {exc}"
                raise
        raise last_exc if last_exc else ComError("COM 调用失败")

    @staticmethod
    def _is_transient(exc: BaseException) -> bool:
        hresult = getattr(exc, "hresult", None)
        if hresult is None:
            args = getattr(exc, "args", ())
            hresult = args[0] if args and isinstance(args[0], int) else None
        return hresult in TRANSIENT_HRESULTS

    # -- 实例获取 -----------------------------------------------------------

    def _ensure_app(self) -> None:
        """确认实例还活着；不活就重取。"""
        if self._app is not None:
            try:
                _ = self._app.Version  # 最轻的存活探测
                return
            except Exception:  # noqa: BLE001 - 实例已失效
                self._app = None
        self._acquire()

    def _acquire(self) -> None:
        if self._app_factory is not None:
            self._app = self._app_factory()
            with self._lock:
                self._info.attached = False
                self._info.acquired_at = _now()
            self._apply_visibility()
            return

        import win32com.client

        app: Any = None
        attached = False
        if self.attach_preferred:
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
                attached = True
            except Exception:  # noqa: BLE001 - 没有正在运行的实例
                app = None
        if app is None:
            try:
                app = win32com.client.Dispatch("PowerPoint.Application")
            except Exception as exc:  # noqa: BLE001
                raise ComUnavailable(f"无法启动 PowerPoint：{exc}") from exc
        self._app = app
        with self._lock:
            self._info.attached = attached
            self._info.acquired_at = _now()
        self._apply_visibility()

    def _apply_visibility(self) -> None:
        """可见性铁律：拿到实例立刻让它显示出来，绝不隐身操作。"""
        if self._app is None:
            return
        try:
            if self.visible:
                self._app.Visible = True
        except Exception as exc:  # noqa: BLE001
            self._log(f"[com] 设置 Visible 失败：{exc}")
            if self.visible:
                raise ComUnavailable(f"无法显示 PowerPoint，已停止操作：{exc}") from exc
        if self.alerts_off:
            try:
                self._app.DisplayAlerts = PP_ALERTS_NONE
            except Exception:  # noqa: BLE001 - 部分版本无此属性
                pass
        with self._lock:
            try:
                self._info.version = str(self._app.Version)
            except Exception:  # noqa: BLE001
                pass
            try:
                self._info.visible = bool(self._app.Visible)
            except Exception:  # noqa: BLE001
                self._info.visible = None
        if self.visible and self._info.visible is not True:
            raise ComUnavailable("无法确认 PowerPoint 窗口可见，已停止操作。")

    def _release(self) -> None:
        """只断开引用：不 Quit，用户的 PowerPoint 要留着。"""
        self._app = None
