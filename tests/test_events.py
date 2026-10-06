"""事件总线：观察台的数据源，必须有序、不阻塞、能补齐历史。"""

from __future__ import annotations

from pptd.events import EventBus


def test_publish_and_recent_order() -> None:
    bus = EventBus()
    for i in range(5):
        bus.publish("op.start", id=str(i))
    recent = bus.recent(3)
    assert [e.data["id"] for e in recent] == ["2", "3", "4"]
    assert [e.seq for e in recent] == sorted(e.seq for e in recent)


def test_subscribe_receives_and_unsubscribe_stops() -> None:
    bus = EventBus()
    q = bus.subscribe()
    bus.publish("op.start", tool="ppt_open")
    event = q.get_nowait()
    assert event.type == "op.start"
    assert event.data["tool"] == "ppt_open"

    bus.unsubscribe(q)
    bus.publish("op.end", ok=True)
    assert q.empty()


def test_ring_buffer_is_bounded() -> None:
    bus = EventBus(capacity=10)
    for i in range(50):
        bus.publish("log", n=i)
    assert len(bus.recent(1000)) == 10
    assert bus.recent(1)[0].data["n"] == 49


def test_slow_subscriber_does_not_block_publisher() -> None:
    """慢消费者（比如卡住的浏览器）不能把 agent 拖住——满了就丢最旧。"""
    bus = EventBus()
    q = bus.subscribe(maxsize=3)
    for i in range(10):
        bus.publish("log", n=i)
    assert q.qsize() == 3
    assert q.get_nowait().data["n"] == 7  # 保留的是最新三条


def test_to_dict_shape() -> None:
    bus = EventBus()
    payload = bus.publish("state", active_deck="a.pptx").to_dict()
    assert set(payload) == {"seq", "type", "at", "data"}
    assert payload["type"] == "state"
    assert payload["at"].endswith("Z")
