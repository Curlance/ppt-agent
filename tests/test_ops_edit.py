"""M2 编辑工具：改文字/位置/增删页与形状、插图、备注、撤销、跟速。

假 COM 会**真的维护撤销栈**，所以 ``ppt_undo`` 的"撤到底会失败"这一边界也被覆盖。
"""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from pptd.com import ComSession
from pptd.engine import Engine, describe_result
from pptd.events import EventBus
from pptd.registry import get_tool

from .fakes import FakePowerPointApp, make_app, make_png


@pytest.fixture()
def app(tmp_path: Path) -> FakePowerPointApp:
    return make_app(tmp_path)


@pytest.fixture()
def edit_engine(settings, bus: EventBus, app: FakePowerPointApp) -> Engine:
    import pptd.ops  # noqa: F401

    session = ComSession(app_factory=lambda: app)
    session.start(timeout=5)
    engine = Engine.create(settings, bus=bus, token="t", session=session)
    engine.state["pace_ms"] = 0
    yield engine
    session.stop(timeout=2)


def _shape_count(app: FakePowerPointApp, slide: int = 1) -> int:
    return app.ActivePresentation.Slides(slide).Shapes.Count


def test_write_follows_target_page_and_can_clear(edit_engine, app):
    result = edit_engine.invoke("ppt_set_text", {"slide": 3, "shape": "2", "text": ""})
    assert result["ok"], result
    assert app.ActivePresentation.Slides(3).Shapes(2).TextFrame.TextRange.Text == ""
    assert app.ActivePresentation.Windows(1).View.Slide.SlideIndex == 3
    assert edit_engine.invoke("ppt_set_notes", {"slide": 3, "text": ""})["ok"]


def test_zero_origin_is_honored(edit_engine, app):
    result = edit_engine.invoke("ppt_add_textbox", {"slide": 2, "text": "", "left": 0, "top": 0})
    assert result["ok"], result
    shape = app.ActivePresentation.Slides(2).Shapes(result["result"]["shape_index"])
    assert (shape.Left, shape.Top) == (0, 0)


def test_bad_color_does_not_modify_text(edit_engine, app):
    shape = app.ActivePresentation.Slides(1).Shapes(2)
    before = shape.TextFrame.TextRange.Text
    result = edit_engine.invoke("ppt_set_text", {"slide": 1, "shape": "2", "text": "changed", "color": "invalid"})
    assert result["ok"] is False
    assert result["error"]["code"] == "args/color"
    assert shape.TextFrame.TextRange.Text == before


def test_geometry_write_failure_is_reported(edit_engine, app):
    from .fakes import FakeShape

    class RejectPosition(FakeShape):
        def __setattr__(self, name, value):
            if name == "Left" and getattr(self, "reject", False):
                raise RuntimeError("位置写入被拒绝")
            super().__setattr__(name, value)
    shape = RejectPosition("locked")
    app.ActivePresentation.Slides(1).Shapes._add(shape)
    shape.reject = True
    result = edit_engine.invoke("ppt_set_geometry", {"slide": 1, "shape": "locked", "left": 90})
    assert result["ok"] is False
    assert "位置写入被拒绝" in result["error"]["message"]


def test_failed_navigation_prevents_write(edit_engine, app):
    def fail(index):
        raise RuntimeError("无法跳页")
    app.ActivePresentation.Windows(1).View.GotoSlide = fail
    shape = app.ActivePresentation.Slides(2).Shapes(2)
    before = shape.TextFrame.TextRange.Text
    result = edit_engine.invoke("ppt_set_text", {"slide": 2, "shape": "2", "text": "changed"})
    assert result["error"]["code"] == "visibility/goto-failed"
    assert shape.TextFrame.TextRange.Text == before


# --------------------------------------------------------------------------
# 改文字
# --------------------------------------------------------------------------


def test_set_text_by_name(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_set_text", {"slide": 1, "shape": "正文 1", "text": "换掉的内容"})
    assert result["ok"] is True
    data = result["result"]
    assert data["after"] == "换掉的内容"
    assert data["before"] == "第 1 页正文"
    assert data["chars"] == 5


def test_set_text_by_title_keyword(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_set_text", {"slide": 1, "shape": "title", "text": "新标题"})
    assert result["ok"] is True
    assert result["result"]["shape"] == "标题 1"


def test_set_text_by_index(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_set_text", {"slide": 1, "shape": "2", "text": "按序号改"})
    assert result["ok"] is True
    assert result["result"]["shape_index"] == 2


def test_set_text_applies_formatting(edit_engine: Engine) -> None:
    result = edit_engine.invoke(
        "ppt_set_text",
        {
            "slide": 1,
            "shape": "正文 1",
            "text": "带格式",
            "font_size": 28,
            "bold": True,
            "color": "#4D6BFE",
            "align": "center",
        },
    )
    assert result["ok"] is True
    applied = result["result"]["applied"]
    assert applied["font_size"] == 28.0
    assert applied["bold"] is True
    assert applied["color"] == "#4D6BFE"
    assert applied["align"] == "center"


def test_set_text_unknown_shape_lists_candidates(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_set_text", {"slide": 1, "shape": "并不存在", "text": "x"})
    assert result["ok"] is False
    assert result["error"]["code"] == "shape/not-found"
    assert "标题 1" in result["error"]["details"]["available"]


def test_set_text_ambiguous_substring(edit_engine: Engine) -> None:
    edit_engine.invoke("ppt_add_textbox", {"slide": 1, "text": "A"})
    edit_engine.invoke("ppt_add_textbox", {"slide": 1, "text": "B"})
    result = edit_engine.invoke("ppt_set_text", {"slide": 1, "shape": "TextBox", "text": "x"})
    assert result["ok"] is False
    assert result["error"]["code"] == "shape/ambiguous"
    assert len(result["error"]["details"]["matches"]) >= 2


def test_set_text_rejects_picture(edit_engine: Engine, tmp_path: Path, settings) -> None:
    png = make_png(tmp_path / "pic.png")
    added = edit_engine.invoke("ppt_add_picture", {"slide": 1, "path": str(png)})
    name = added["result"]["shape"]
    result = edit_engine.invoke("ppt_set_text", {"slide": 1, "shape": name, "text": "x"})
    assert result["ok"] is False
    assert result["error"]["code"] == "shape/no-text-frame"


def test_set_text_out_of_range_index(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_set_text", {"slide": 1, "shape": "99", "text": "x"})
    assert result["ok"] is False
    assert result["error"]["code"] == "shape/out-of-range"


# --------------------------------------------------------------------------
# 加文本框 / 几何 / 删形状 / 插图
# --------------------------------------------------------------------------


def test_add_textbox(edit_engine: Engine, app: FakePowerPointApp) -> None:
    before = _shape_count(app)
    result = edit_engine.invoke(
        "ppt_add_textbox", {"slide": 2, "text": "新加的字", "left": 100, "top": 200, "font_size": 24}
    )
    assert result["ok"] is True
    data = result["result"]
    assert data["shape_index"] == _shape_count(app, 2)
    assert data["shape"] in {"TextBox 3", "TextBox 4", "TextBox 5"}
    assert _shape_count(app) == before


def test_set_geometry(edit_engine: Engine) -> None:
    result = edit_engine.invoke(
        "ppt_set_geometry", {"slide": 1, "shape": "正文 1", "left": 123.0, "width": 456.0}
    )
    assert result["ok"] is True
    data = result["result"]
    assert data["changed"] == ["left", "width"]
    assert data["after"]["left"] == 123.0
    assert data["after"]["width"] == 456.0
    assert data["before"]["left"] == 60.0


def test_set_geometry_needs_at_least_one_field(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_set_geometry", {"slide": 1, "shape": "正文 1"})
    assert result["ok"] is False
    assert result["error"]["code"] == "args/empty"


def test_delete_shape(edit_engine: Engine, app: FakePowerPointApp) -> None:
    before = _shape_count(app)
    result = edit_engine.invoke("ppt_delete_shape", {"slide": 1, "shape": "正文 1"})
    assert result["ok"] is True
    assert result["result"]["deleted"] is True
    assert _shape_count(app) == before - 1


def test_add_picture(edit_engine: Engine, tmp_path: Path) -> None:
    png = make_png(tmp_path / "logo.png", (128, 64))
    result = edit_engine.invoke(
        "ppt_add_picture", {"slide": 1, "path": str(png), "left": 500, "top": 300, "width": 128, "height": 64}
    )
    assert result["ok"] is True
    data = result["result"]
    assert data["added"] is True
    assert data["geometry"]["left"] == 500.0
    assert data["source"] == str(png)


def test_add_picture_requires_absolute_path(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_add_picture", {"slide": 1, "path": "pic.png"})
    assert result["ok"] is False
    assert result["error"]["code"] == "path/relative"


def test_add_picture_missing_file(edit_engine: Engine, tmp_path: Path) -> None:
    result = edit_engine.invoke("ppt_add_picture", {"slide": 1, "path": str(tmp_path / "nope.png")})
    assert result["ok"] is False
    assert result["error"]["code"] == "path/missing"


def test_add_picture_rejects_unsupported_suffix(edit_engine: Engine, tmp_path: Path) -> None:
    bad = tmp_path / "notes.txt"
    bad.write_text("not an image", encoding="utf-8")
    result = edit_engine.invoke("ppt_add_picture", {"slide": 1, "path": str(bad)})
    assert result["ok"] is False
    assert result["error"]["code"] == "path/unsupported"


# --------------------------------------------------------------------------
# 增删移动页
# --------------------------------------------------------------------------


def test_add_slide_appends_by_default(edit_engine: Engine, app: FakePowerPointApp) -> None:
    result = edit_engine.invoke("ppt_add_slide", {})
    assert result["ok"] is True
    data = result["result"]
    assert data["slide"] == 6
    assert data["total"] == 6
    # 留空版式 = 沿用最后一页的版式
    assert data["layout"] == "标题和内容"


def test_add_slide_at_position_with_named_layout(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_add_slide", {"index": 1, "layout": "空白"})
    assert result["ok"] is True
    assert result["result"]["slide"] == 1
    assert result["result"]["layout"] == "空白"


def test_add_slide_unknown_layout_lists_available(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_add_slide", {"layout": "不存在的版式"})
    assert result["ok"] is False
    assert result["error"]["code"] == "layout/not-found"
    assert "空白" in result["error"]["details"]["available"]


def test_delete_slide_requires_explicit_page(edit_engine: Engine) -> None:
    """删除绝不能默认动用户正在看的那一页。

    这一层是 schema 拦截（连 COM 都不会碰）；``_slide_at(required=True)`` 里
    还有第二道防线，防止将来有人把 schema 的必填去掉。
    """
    result = edit_engine.invoke("ppt_delete_slide", {})
    assert result["ok"] is False
    assert result["error"]["code"] == "args/missing"
    assert "slide" in result["error"]["message"]

    # 第二道防线：直接调内部辅助也照样拒绝
    from pptd.errors import ToolError
    from pptd.ops.edit import _slide_at

    pres = edit_engine.session.run(lambda app: app.ActivePresentation)
    with pytest.raises(ToolError) as exc:
        _slide_at(None, pres, None, required=True)
    assert exc.value.code == "slide/required"


def test_delete_slide(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_delete_slide", {"slide": 2})
    assert result["ok"] is True
    assert result["result"]["total"] == 4
    assert result["result"]["deleted_title"] == "第 2 页标题"


def test_move_slide(edit_engine: Engine, app: FakePowerPointApp) -> None:
    due = app.ActivePresentation.Slides(3).Shapes(1).TextFrame.TextRange.Text
    result = edit_engine.invoke("ppt_move_slide", {"slide": 3, "to": 1})
    assert result["ok"] is True
    assert app.ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Text == due


# --------------------------------------------------------------------------
# 备注 / 选中 / 保存
# --------------------------------------------------------------------------


def test_set_notes(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_set_notes", {"slide": 1, "text": "讲这一页时要强调三件事"})
    assert result["ok"] is True
    assert result["result"]["after"] == "讲这一页时要强调三件事"
    assert result["result"]["before"] == "备注 1"


def test_select_slide_and_shape(edit_engine: Engine, app: FakePowerPointApp) -> None:
    result = edit_engine.invoke("ppt_select", {"slide": 2, "shape": "正文 2"})
    assert result["ok"] is True
    assert result["result"]["current_slide"] == 2
    assert "正文 2" in result["result"]["selected"]
    assert app.ActivePresentation.Slides(2).Shapes(2).selected is True


def test_select_does_not_modify_document(edit_engine: Engine, app: FakePowerPointApp) -> None:
    """选中不该改动文档——给用户的形状画描边等于污染 deck。"""
    before = _shape_count(app)
    edit_engine.invoke("ppt_select", {"slide": 1, "shape": "标题 1"})
    assert _shape_count(app) == before


def test_save(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_save", {})
    assert result["ok"] is True
    assert result["result"]["saved"] is True


def test_save_copy_writes_backup(edit_engine: Engine, settings) -> None:
    result = edit_engine.invoke("ppt_save", {"mode": "copy"})
    assert result["ok"] is True
    copy_path = Path(result["result"]["copy_path"])
    assert copy_path.is_file()
    # 默认落在 ppt-agent 自己的 backups 目录，不污染用户的文件夹
    assert settings.home in copy_path.parents


# --------------------------------------------------------------------------
# 撤销 / 跟速
# --------------------------------------------------------------------------


def test_undo_reverts_last_write(edit_engine: Engine, app: FakePowerPointApp) -> None:
    before = _shape_count(app)
    edit_engine.invoke("ppt_add_textbox", {"slide": 1, "text": "待撤销"})
    assert _shape_count(app) == before + 1

    result = edit_engine.invoke("ppt_undo", {})
    assert result["ok"] is True
    assert result["result"]["undone"] == 1
    assert result["result"]["stopped"] is None
    assert _shape_count(app) == before


def test_undo_reports_when_stack_is_empty(edit_engine: Engine) -> None:
    result = edit_engine.invoke("ppt_undo", {})
    assert result["ok"] is True
    assert result["result"]["undone"] == 0
    assert result["result"]["stopped"], "撤到底必须如实说明，而不是假装成功"


def test_undo_restores_deleted_shape(edit_engine: Engine, app: FakePowerPointApp) -> None:
    before = _shape_count(app)
    edit_engine.invoke("ppt_delete_shape", {"slide": 1, "shape": "正文 1"})
    assert _shape_count(app) == before - 1
    edit_engine.invoke("ppt_undo", {})
    assert _shape_count(app) == before


def test_set_pace_updates_state_and_emits(edit_engine: Engine, bus: EventBus) -> None:
    result = edit_engine.invoke("ppt_set_pace", {"pace_ms": 1500})
    assert result["ok"] is True
    assert result["result"]["pace_ms"] == 1500
    assert result["result"]["label"] == "慢速"
    assert edit_engine.state["pace_ms"] == 1500
    assert any(e.type == "state" and e.data.get("pace_ms") == 1500 for e in bus.recent(10))


def test_pacing_slows_writes_but_not_reads(edit_engine: Engine) -> None:
    edit_engine.state["pace_ms"] = 150

    started = time.perf_counter()
    edit_engine.invoke("ppt_decks", {})
    read_ms = (time.perf_counter() - started) * 1000
    assert read_ms < 120, f"只读操作不该被跟速拖慢，实测 {read_ms:.0f}ms"

    started = time.perf_counter()
    edit_engine.invoke("ppt_add_textbox", {"slide": 1, "text": "跟速"})
    write_ms = (time.perf_counter() - started) * 1000
    assert write_ms >= 140, f"写操作应该停一下让人跟上，实测 {write_ms:.0f}ms"


def test_human_summaries_for_edit_tools() -> None:
    import pptd.ops  # noqa: F401

    assert "已保存" in describe_result(get_tool("ppt_save"), {"saved": True, "presentation": {"name": "a"}})
    assert "备份" in describe_result(get_tool("ppt_save"), {"copied": True, "copy_path": "D:/b.pptx"})
    assert "第 6 页" in describe_result(
        get_tool("ppt_add_slide"), {"slide": 6, "layout": "空白", "total": 6}
    )
    assert "已删除第 2 页" in describe_result(
        get_tool("ppt_delete_slide"), {"deleted": 1, "from": 2, "total": 4}
    )
    assert "第 3 页移到第 1 页" in describe_result(get_tool("ppt_move_slide"), {"from": 3, "to": 1})
    assert "「正文 1」" in describe_result(get_tool("ppt_set_text"), {"shape": "正文 1", "chars": 4})
    assert "已加文本框" in describe_result(get_tool("ppt_add_textbox"), {"shape": "TextBox 3"})
    assert "已插入图片" in describe_result(get_tool("ppt_add_picture"), {"shape": "Picture 3"})
    assert "已撤销 2 步" in describe_result(get_tool("ppt_undo"), {"undone": 2})
    assert "撤到底" in describe_result(get_tool("ppt_undo"), {"undone": 0, "stopped": "撤到底"})
    assert "1500 毫秒" in describe_result(get_tool("ppt_set_pace"), {"pace_ms": 1500, "label": "慢速"})
