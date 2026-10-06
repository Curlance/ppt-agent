"""事件总线：把 agent 的每一步操作变成用户看得见的信息流。

这是"全程可视"要求的基础设施。所有面（观察台网页、PowerPoint 任务窗格、
DSH 对话里的工具结果）读的都是同一条流，不各自造数据。

事件类型（约定）
----------------
- ``state``       会话状态变化：是否附着 PowerPoint、当前演示、当前页、运行/暂停。
- ``op.start``    一步操作开始，带**中文人话摘要**与目标（页/形状）。
- ``op.spotlight`` 已把目标选中并滚入视野，供用户眼睛跟随。
- ``op.shot``     该步之后的本页截图（路径 + 可选内联数据）。
- ``op.end``      一步操作结束：成功与否、耗时、错误。
- ``log``         守护进程自身的诊断信息。

线程模型：``publish`` 可能来自 COM 线程，``subscribe`` 一般来自 HTTP 线程，
因此内部用锁保护，订阅者各自持有独立队列。
"""

from __future__ import annotations

import itertools
import queue
import threading
from collections import deque
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Any, Iterator

__all__ = ["Event", "EventBus", "utc_now"]


def utc_now() -> str:
    """ISO 8601 UTC 时间戳（带 Z，便于前端直接解析）。"""
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


@dataclass(frozen=True)
class Event:
    """一条不可变事件。``seq`` 单调递增，便于客户端断线重连后去重。"""

    seq: int
    type: str
    at: str
    data: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class EventBus:
    """进程内发布/订阅，带环形历史，便于新连接补齐上下文。"""

    def __init__(self, capacity: int = 500) -> None:
        self._lock = threading.Lock()
        self._history: deque[Event] = deque(maxlen=capacity)
        self._subscribers: set[queue.Queue[Event]] = set()
        self._counter = itertools.count(1)

    # -- 发布 ---------------------------------------------------------------

    def publish(self, type_: str, **data: Any) -> Event:
        """发布一条事件，并推送给所有订阅者。返回事件本身便于调用方复用。"""
        with self._lock:
            event = Event(seq=next(self._counter), type=type_, at=utc_now(), data=data)
            self._history.append(event)
            subscribers = list(self._subscribers)
        for q in subscribers:
            try:
                q.put_nowait(event)
            except queue.Full:  # 慢消费者不拖垮生产者：丢最旧
                try:
                    q.get_nowait()
                    q.put_nowait(event)
                except queue.Empty:  # pragma: no cover - 竞态兜底
                    pass
        return event

    # -- 订阅 ---------------------------------------------------------------

    def subscribe(self, maxsize: int = 200) -> queue.Queue[Event]:
        """注册一个订阅者，返回其专属队列。调用方负责 ``unsubscribe``。"""
        q: queue.Queue[Event] = queue.Queue(maxsize=maxsize)
        with self._lock:
            self._subscribers.add(q)
        return q

    def unsubscribe(self, q: queue.Queue[Event]) -> None:
        with self._lock:
            self._subscribers.discard(q)

    def stream(self, q: queue.Queue[Event], backlog: int = 0) -> Iterator[Event]:
        """把订阅队列包成生成器，可选先补发最近 ``backlog`` 条历史。"""
        for event in self.recent(backlog):
            yield event
        while True:
            try:
                yield q.get(timeout=30.0)
            except queue.Empty:
                return

    # -- 历史 ---------------------------------------------------------------

    def recent(self, n: int = 50) -> list[Event]:
        """最近 n 条事件（按时间正序）。"""
        if n <= 0:
            return []
        with self._lock:
            items = list(self._history)
        return items[-n:]

    @property
    def subscriber_count(self) -> int:
        with self._lock:
            return len(self._subscribers)
