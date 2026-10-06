"""核心承诺：在真实 PowerPoint 的临时演示上验证跟页、清空、表格和组合。"""
from dataclasses import replace

import pytest

from pptd.config import settings as defaults
from pptd.engine import Engine
from pptd.http_api import create_app
from pptd.ops.deck import _goto_slide
from pptd.ops.common import resolve_deck

pytestmark = pytest.mark.com


@pytest.fixture(scope="module")
def live_engine(tmp_path_factory):
    import pptd.ops  # noqa: F401
    tmp_path = tmp_path_factory.mktemp("live-core")
    engine = Engine.create(replace(defaults, home=tmp_path), token="integration")
    assert engine.start()["alive"]
    engine.state["pace_ms"] = 0
    previous = engine.session.run(lambda app: (
        getattr(app.ActivePresentation, "Name", None) if app.Presentations.Count else None,
        int(app.ActiveWindow.View.Slide.SlideIndex) if app.Presentations.Count and app.ActivePresentation.Slides.Count else None,
    ))
    created = engine.invoke("ppt_new", {})
    assert created["ok"], created
    deck = created["result"]["presentation"]["name"]
    (tmp_path / "test-deck-name.txt").write_text(deck, encoding="utf-8")
    try:
        for _ in range(2):
            added = engine.invoke("ppt_add_slide", {"deck": deck, "layout": "7"})
            assert added["ok"], added
        yield engine, deck
    finally:
        engine.invoke("ppt_close", {"deck": deck, "save": False})
        if previous[0] and previous[1]:
            engine.session.run(lambda app: _goto_slide(app, resolve_deck(app, previous[0]), previous[1]))
        engine.stop()


def test_real_edit_auto_follows_and_clears(live_engine):
    engine, deck = live_engine
    added = engine.invoke("ppt_add_textbox", {"deck": deck, "slide": 2, "text": "原始文字", "left": 0, "top": 0, "color": "#102030"})
    assert added["ok"], added
    shape = added["result"]["shape"]
    engine.invoke("ppt_goto", {"deck": deck, "slide": 1})
    changed = engine.invoke("ppt_set_text", {"deck": deck, "slide": 2, "shape": shape, "text": ""})
    assert changed["ok"], changed
    assert engine.session_state()["active_slide"] == 2
    read = engine.invoke("ppt_read_slide", {"deck": deck, "slide": 2})
    target = next(s for s in read["result"]["shapes"] if s["name"] == shape)
    assert target.get("text", "") == ""
    assert target["geometry"]["left"] == 0
    assert target["geometry"]["top"] == 0
    assert engine.invoke("ppt_set_notes", {"deck": deck, "slide": 2, "text": ""})["ok"]
    assert engine.invoke("ppt_read_slide", {"deck": deck, "slide": 2})["result"]["notes"] == ""


def test_real_table_and_group_read_edit(live_engine):
    engine, deck = live_engine
    def create(app):
        slide = resolve_deck(app, deck).Slides(2)
        table = slide.Shapes.AddTable(2, 2, 60, 100, 400, 100)
        table.Name = "核心测试表格"
        table.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "原始单元格"
        first = slide.Shapes.AddTextbox(1, 60, 240, 160, 40)
        first.TextFrame.TextRange.Text = "组合原文"
        second = slide.Shapes.AddTextbox(1, 260, 240, 160, 40)
        second.TextFrame.TextRange.Text = "其他文字"
        slide.Shapes.Range([first.Name, second.Name]).Group()
    engine.session.run(create, retry_safe=False)
    read = engine.invoke("ppt_read_slide", {"deck": deck, "slide": 2})
    assert read["ok"], read
    table = next(s for s in read["result"]["shapes"] if "table" in s)
    assert table["table"]["cells"][-1]["text"] == "原始单元格"
    result = engine.invoke("ppt_set_table_cell", {"deck": deck, "slide": 2, "shape": table["ref"], "row": 2, "column": 2, "text": "已更新", "color": "#2468AC", "bold": True})
    assert result["ok"], result
    group = next(s for s in read["result"]["shapes"] if s.get("children"))
    child = next(s for s in group["children"] if s.get("text") == "组合原文")
    result = engine.invoke("ppt_set_text", {"deck": deck, "slide": 2, "shape": child["ref"], "text": "组合已更新"})
    assert result["ok"], result
    read = engine.invoke("ppt_read_slide", {"deck": deck, "slide": 2})
    assert next(s for s in read["result"]["shapes"] if "table" in s)["table"]["cells"][-1]["text"] == "已更新"
    assert "组合已更新" in str(read["result"]["shapes"])


def test_real_unsaved_deck_requires_path(live_engine):
    engine, deck = live_engine
    result = engine.invoke("ppt_save", {"deck": deck})
    assert result["error"]["code"] == "save/path-required"
    result = engine.invoke("ppt_save", {"deck": deck, "mode": "copy"})
    assert result["ok"], result


def test_live_view_updates_panel_after_manual_navigation(live_engine):
    from fastapi.testclient import TestClient
    engine, deck = live_engine
    engine.invoke("ppt_goto", {"deck": deck, "slide": 1})
    engine.session.run(lambda app: _goto_slide(app, resolve_deck(app, deck), 2))
    with TestClient(create_app(engine)) as client:
        shot = client.get("/live", headers={"X-PPT-Token": "integration"})
        assert shot.status_code == 200, shot.text
        assert shot.headers["x-ppt-slide"] == "2"
        state = client.get("/state", headers={"X-PPT-Token": "integration"}).json()
        assert state["active_slide"] == 2
        assert "第 2 页" in client.get("/strip.txt", headers={"X-PPT-Token": "integration"}).text
