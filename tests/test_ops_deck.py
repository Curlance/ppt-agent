"""M1 读取与截图工具。

假 COM 在 ``tests/fakes.py``：``Slide.Export`` 会真的写一张 PNG，
所以缩略图生成、命名、裁剪上限、inline 开关都能脱离真机验证。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pptd.com import ComSession
from pptd.engine import Engine, describe_result
from pptd.events import EventBus
from pptd.registry import get_tool

from .fakes import FakePowerPointApp, make_app


@pytest.fixture()
def app(tmp_path: Path) -> FakePowerPointApp:
    return make_app(tmp_path)


@pytest.fixture()
def deck_engine(settings, bus: EventBus, app: FakePowerPointApp) -> Engine:
    import pptd.ops  # noqa: F401

    session = ComSession(app_factory=lambda: app)
    session.start(timeout=5)
    engine = Engine.create(settings, bus=bus, token="t", session=session)
    engine.state["pace_ms"] = 0
    yield engine
    session.stop(timeout=2)


def test_decks_lists_slide_size(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_decks", {})
    assert result["ok"] is True
    data = result["result"]
    assert data["count"] == 1
    assert data["decks"][0]["slide_size"] == {"width_pt": 960.0, "height_pt": 540.0}
    assert data["active_index"] == 1


def test_table_body_is_read_and_can_be_edited(deck_engine, app):
    read = deck_engine.invoke("ppt_read_slide", {"slide": 2})
    table = next(s for s in read["result"]["shapes"] if "table" in s)
    assert table["table"]["cells"][-1]["text"] == "R3C2"
    updated = deck_engine.invoke("ppt_set_table_cell", {"slide": 2, "shape": table["ref"], "row": 3, "column": 2, "text": "更新后的值"})
    assert updated["ok"], updated
    read = deck_engine.invoke("ppt_read_slide", {"slide": 2})
    assert read["result"]["shapes"][2]["table"]["cells"][-1]["text"] == "更新后的值"
    assert deck_engine.invoke("ppt_set_table_cell", {"slide": 2, "shape": "3", "row": 99, "column": 2, "text": "x"})["ok"] is False


def test_group_children_are_addressable_and_bounded(deck_engine, app):
    from .fakes import FakeShape, _Subset

    group = app.ActivePresentation.Slides(4).Shapes(3)
    group.GroupItems = _Subset([FakeShape("子形状", 17, text="原文"), FakeShape("另一个", 17, text="其他")])
    read = deck_engine.invoke("ppt_read_slide", {"slide": 4})
    child = read["result"]["shapes"][2]["children"][0]
    assert child["ref"] == "3/1"
    assert child["text"] == "原文"
    changed = deck_engine.invoke("ppt_set_text", {"slide": 4, "shape": child["ref"], "text": "修改组合内文字"})
    assert changed["ok"], changed
    assert group.GroupItems(1).TextFrame.TextRange.Text == "修改组合内文字"
    limited = deck_engine.invoke("ppt_read_slide", {"slide": 4, "max_shapes": 4})
    assert limited["result"]["shapes"][2]["children_truncated"] is True


def test_table_read_truncation_is_explicit(deck_engine):
    result = deck_engine.invoke("ppt_read_slide", {"slide": 2, "max_table_cells": 2})
    table = result["result"]["shapes"][2]["table"]
    assert len(table["cells"]) == 2
    assert table["truncated"] is True


def test_slides_lists_titles_and_truncates(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_slides", {"limit": 2})
    assert result["ok"] is True
    data = result["result"]
    assert data["total"] == 5
    assert data["returned"] == 2
    assert data["truncated"] is True
    assert data["slides"][0]["title"] == "第 1 页标题"
    assert data["slides"][0]["has_notes"] is True


def test_read_slide_reports_shapes_text_geometry(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_read_slide", {"slide": 2})
    assert result["ok"] is True
    data = result["result"]
    assert data["slide"] == 2
    assert data["layout"] == "标题和内容"      # 走 CustomLayout，不是 Layout
    assert data["notes"] == "备注 2"
    assert data["shape_count"] == 3
    types = {s["type"] for s in data["shapes"]}
    assert "TextBox" in types and "Table" in types
    title = next(s for s in data["shapes"] if s.get("text") == "第 2 页标题")
    assert title["geometry"]["left"] == 60.0
    table = next(s for s in data["shapes"] if s["type"] == "Table")
    assert table["table"]["rows"] == 3 and table["table"]["header"] == ["表头1", "表头2"]


def test_read_slide_chart_and_group(deck_engine: Engine) -> None:
    chart = deck_engine.invoke("ppt_read_slide", {"slide": 3})["result"]
    assert any(s.get("chart", {}).get("chart_type") == 51 for s in chart["shapes"])
    group = deck_engine.invoke("ppt_read_slide", {"slide": 4})["result"]
    assert any(s.get("group_items") == 4 for s in group["shapes"])


def test_read_slide_can_drop_text_and_geometry(deck_engine: Engine) -> None:
    data = deck_engine.invoke(
        "ppt_read_slide", {"slide": 1, "include_text": False, "include_geometry": False}
    )["result"]
    assert all("text" not in s for s in data["shapes"])
    assert all("geometry" not in s for s in data["shapes"])


def test_read_slide_out_of_range(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_read_slide", {"slide": 99})
    assert result["ok"] is False
    assert result["error"]["code"] == "slide/out-of-range"
    assert result["error"]["details"]["total"] == 5


def test_goto_changes_view(deck_engine: Engine, app: FakePowerPointApp) -> None:
    result = deck_engine.invoke("ppt_goto", {"slide": 4})
    assert result["ok"] is True
    assert result["result"]["current_slide"] == 4
    assert app.ActivePresentation.Windows(1).View.Slide.SlideIndex == 4


def test_unknown_deck_lists_candidates(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_slides", {"deck": "不存在.pptx"})
    assert result["ok"] is False
    assert result["error"]["code"] == "deck/not-found"
    assert "deck1.pptx" in result["error"]["details"]["open_decks"][0]


def test_deck_can_be_selected_by_index(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_slides", {"deck": "1"})
    assert result["ok"] is True
    assert result["result"]["deck"]["name"] == "deck1.pptx"


def test_shot_writes_image_and_thumb(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_shot", {"slide": 2, "width": 640, "thumb_width": 320})
    assert result["ok"] is True
    shot = result["result"]["shots"][0]
    assert shot["slide"] == 2
    image = Path(shot["image"]["path"])
    thumb = Path(shot["thumb"]["path"])
    assert image.is_file() and thumb.is_file()
    # 缩略图必须比高清图小——这是"给人看大图、给模型看小图"的前提
    assert thumb.stat().st_size < image.stat().st_size
    assert shot["image"]["width"] == 640
    assert shot["thumb"]["width"] == 320
    assert shot["image"]["url"].startswith("/shots/")
    assert shot["thumb"]["url"].endswith(".thumb.jpg")


def test_shot_can_skip_thumb(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_shot", {"slide": 1, "inline": False})
    assert result["ok"] is True
    assert "thumb" not in result["result"]["shots"][0]


def test_shot_clamps_count(deck_engine: Engine) -> None:
    # 合法请求在文档末尾缩短；超过 schema 的上限必须明确拒绝。
    rejected = deck_engine.invoke("ppt_shot", {"slide": 4, "count": 99})
    assert rejected["error"]["code"] == "args/range"
    result = deck_engine.invoke("ppt_shot", {"slide": 4, "count": 10})
    assert result["ok"] is True
    assert result["result"]["count"] == 2  # 第 4、5 页


def test_shot_jpg_format(deck_engine: Engine) -> None:
    result = deck_engine.invoke("ppt_shot", {"slide": 1, "format": "jpg"})
    assert result["ok"] is True
    assert Path(result["result"]["shots"][0]["image"]["path"]).suffix == ".jpg"


def test_engine_emits_op_shot_event(deck_engine: Engine, bus: EventBus) -> None:
    deck_engine.invoke("ppt_shot", {"slide": 2})
    shots = [e for e in bus.recent(30) if e.type == "op.shot"]
    assert shots, "截图没有广播 op.shot 事件，观察台就挂不上画面"
    assert shots[0].data["shots"][0]["slide"] == 2
    # 顺序：op.shot 必须紧跟它那一步的 op.end
    types = [e.type for e in bus.recent(30)]
    position = types.index("op.shot")
    assert types[position - 1] == "op.end"


def test_capture_live_frame_reads_current_slide(deck_engine: Engine, app: FakePowerPointApp, settings) -> None:
    """实时取景读的是 PowerPoint 当前状态——用户自己翻页也跟得上。"""
    from pptd.ops.deck import capture_live_frame

    target = settings.shot_dir / "live.jpg"
    settings.shot_dir.mkdir(parents=True, exist_ok=True)
    app.ActivePresentation.Windows(1).View.GotoSlide(3)
    index = deck_engine.session.run(capture_live_frame, target, width=480)
    assert index == 3
    assert target.is_file()
    # 覆盖同一个文件，不留档
    app.ActivePresentation.Windows(1).View.GotoSlide(5)
    assert deck_engine.session.run(capture_live_frame, target, width=480) == 5
    assert len(list(settings.shot_dir.glob("live*"))) == 1


def test_capture_live_frame_distinguishes_failure_reasons(settings, bus: EventBus) -> None:
    """拍不到时要说清是"没有演示"还是"演示没页"——两者的排查方向完全不同。"""
    from pptd.ops.deck import LIVE_NO_PRESENTATION, LIVE_NO_SLIDES, capture_live_frame

    import pptd.ops  # noqa: F401

    target = settings.shot_dir / "live.jpg"
    settings.shot_dir.mkdir(parents=True, exist_ok=True)

    empty_app = make_app(settings.home, decks=0)
    assert empty_app.ActivePresentation is None
    assert capture_live_frame(empty_app, target, width=480) == LIVE_NO_PRESENTATION

    deckless = make_app(settings.home, slides=0)
    assert capture_live_frame(deckless, target, width=480) == LIVE_NO_SLIDES


def test_feed_text_derives_deck_and_slide(deck_engine: Engine, bus: EventBus) -> None:
    """面板要显示"现在在哪份演示的第几页"，而这些信息藏在事件流的工具结果里。"""
    from pptd.http_api import _derive_from_events

    assert _derive_from_events(bus.recent(20)) == ("", None)

    deck_engine.invoke("ppt_status", {})
    deck, slide = _derive_from_events(bus.recent(20))
    assert deck == "deck1.pptx"
    assert slide == 1


def test_human_summaries_for_m1_tools() -> None:
    import pptd.ops  # noqa: F401

    assert "2 个" in describe_result(get_tool("ppt_decks"), {"decks": [{"name": "a"}, {"name": "b"}]})
    assert "共 5 页" in describe_result(
        get_tool("ppt_slides"), {"deck": {"name": "a"}, "total": 5, "returned": 2}
    )
    assert "第 3 页" in describe_result(get_tool("ppt_read_slide"), {"slide": 3, "returned": 8})
    assert "第 4 页" in describe_result(get_tool("ppt_goto"), {"current_slide": 4})
    assert "共 2 页" in describe_result(get_tool("ppt_shot"), {"shots": [{"slide": 1}, {"slide": 2}]})
