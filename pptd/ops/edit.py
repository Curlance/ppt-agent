"""M2 工具集：真正改这份 deck。

三条安全约束
------------
1. **绝不默认改"当前页"**：删除类操作必须显式给页码，避免手滑毁掉用户正在看的那一页。
2. **外部文件必须绝对路径**：PowerPoint 是独立进程，相对路径会按它自己的工作目录解析。
   实测这会让 ``AddPicture`` 抛出与真正原因毫无关系的错。
3. **撤销走 PowerPoint 自己的撤销栈**：``CommandBars.ExecuteMso("Undo")`` 实测有效。
   但要诚实——**撤销粒度由 PowerPoint 决定**，不是严格"一次工具调用一步"，
   而且撤到底会抛 ``com_error``。所以返回值里如实报告实际撤销了几步、为什么停下。

``ppt_select`` 刻意**不修改文档**：只做选中与跳页。给用户的形状画描边会改动人家的
形状属性，一旦忘记还原就等于污染了 deck——选中的控制柄本身在 PowerPoint 里已经足够显眼。
"""

from __future__ import annotations

import time
import uuid
from pathlib import Path
from typing import Any

from ..errors import ToolError
from ..registry import Context, Param, tool
from .common import (
    absolute_path,
    bring_to_front,
    presentation_info,
    resolve_deck,
    resolve_shape,
    round_pt,
    safe,
    shape_names,
)

__all__ = [
    "op_save",
    "op_add_slide",
    "op_delete_slide",
    "op_move_slide",
    "op_set_text",
    "op_add_textbox",
    "op_set_geometry",
    "op_delete_shape",
    "op_add_picture",
    "op_set_notes",
    "op_select",
    "op_undo",
    "op_set_pace",
    "op_set_table_cell",
]

#: 文本框方向：msoTextOrientationHorizontal
MSO_TEXT_HORIZONTAL = 1
#: MsoTriState
MSO_FALSE = 0
MSO_TRUE = -1
#: 段落对齐
PP_ALIGN = {"left": 1, "center": 2, "right": 3, "justify": 4}
#: 图片扩展名
PICTURE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg"}
#: 默认跟速（毫秒）：让人跟得上，又不至于慢到烦
DEFAULT_PACE_MS = 400


# --------------------------------------------------------------------------
# 辅助
# --------------------------------------------------------------------------


def _rgb(value: Any) -> int:
    """把 ``#RRGGBB`` 转成 PowerPoint 要的整数。

    VBA 的 ``RGB(r, g, b)`` 是 ``r | g<<8 | b<<16``（低字节是红），
    所以这里按同样顺序拼，别直接用 0xRRGGBB。
    """
    if isinstance(value, int):
        return value
    text = str(value).strip().lstrip("#")
    if len(text) != 6:
        raise ToolError(f"颜色要写成 #RRGGBB，收到：{value!r}", code="args/color")
    try:
        r, g, b = int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16)
    except ValueError as exc:
        raise ToolError(f"颜色要写成 #RRGGBB，收到：{value!r}", code="args/color") from exc
    return r | (g << 8) | (b << 16)


def _slide_count(pres: Any) -> int:
    return safe(lambda: int(pres.Slides.Count), 0) or 0


def _slide_at(app: Any, pres: Any, wanted: Any, *, required: bool = False, what: str = "页码") -> int:
    total = _slide_count(pres)
    if total == 0:
        raise ToolError("这份演示里没有任何幻灯片。", code="slide/empty")
    if wanted is None or wanted == "":
        if required:
            raise ToolError(
                f"{what}必须显式指定（不会默认动你正在看的那一页）。这份演示共 {total} 页。",
                code="slide/required",
                total=total,
            )
        index = safe(lambda: int(pres.Windows(1).View.Slide.SlideIndex))
        return max(1, min(int(index) if index else 1, total))
    try:
        index = int(wanted)
    except (TypeError, ValueError) as exc:
        raise ToolError(f"{what}必须是整数，收到：{wanted!r}", code="args/type") from exc
    if not 1 <= index <= total:
        raise ToolError(f"{what} {index} 超出范围（共 {total} 页）。", code="slide/out-of-range", total=total)
    return index


def _slide_obj(pres: Any, index: int) -> Any:
    slide = safe(lambda: pres.Slides(index))
    if slide is None:
        raise ToolError(f"取不到第 {index} 页。", code="slide/not-found", slide=index)
    return slide


def _layout_names(pres: Any) -> list[str]:
    layouts = safe(lambda: pres.SlideMaster.CustomLayouts)
    count = safe(lambda: int(layouts.Count), 0) or 0
    names: list[str] = []
    for i in range(1, count + 1):
        names.append(str(safe(lambda i=i: layouts(i).Name) or f"版式{i}"))
    return names


def _resolve_layout(pres: Any, ref: Any, fallback_slide: Any = None) -> tuple[Any, str]:
    """版式可以用序号或名称指定；留空则沿用最后一页的版式。"""
    layouts = safe(lambda: pres.SlideMaster.CustomLayouts)
    count = safe(lambda: int(layouts.Count), 0) or 0
    if count == 0:
        raise ToolError("这份演示没有任何可用版式。", code="layout/none")

    if ref is None or ref == "":
        if fallback_slide is not None:
            layout = safe(lambda: fallback_slide.CustomLayout)
            if layout is not None:
                return layout, str(safe(lambda: layout.Name) or "沿用上一页")
        layout = layouts(1)
        return layout, str(safe(lambda: layout.Name) or "版式1")

    if isinstance(ref, int) or str(ref).strip().isdigit():
        index = int(ref)
        if not 1 <= index <= count:
            raise ToolError(
                f"版式序号 {index} 超出范围（共 {count} 个）。",
                code="layout/out-of-range",
                available=_layout_names(pres),
            )
        layout = layouts(index)
        return layout, str(safe(lambda: layout.Name) or f"版式{index}")

    wanted = str(ref).strip().casefold()
    matches: list[tuple[Any, str]] = []
    for i in range(1, count + 1):
        layout = layouts(i)
        name = str(safe(lambda l=layout: l.Name) or "")
        if name.casefold() == wanted:
            return layout, name
        if wanted in name.casefold():
            matches.append((layout, name))
    if len(matches) == 1:
        return matches[0]
    raise ToolError(
        f"没有找到名为「{ref}」的版式。",
        code="layout/not-found",
        available=_layout_names(pres),
    )


def _index_of(slide: Any, target: Any) -> int:
    """新加的形状：遍历拿到它的序号，好让模型下次能引用。"""
    shapes = safe(lambda: slide.Shapes)
    count = safe(lambda: int(shapes.Count), 0) or 0
    for i in range(1, count + 1):
        candidate = safe(lambda i=i: shapes(i))
        if candidate is target or (safe(lambda: target.Id) is not None and safe(lambda: candidate.Id) == target.Id):
            return i
    return count


def _apply_font(range_obj: Any, *, size: Any, bold: Any, color: Any, align: Any) -> dict[str, Any]:
    """按需设置字号/加粗/颜色/对齐，返回实际生效的值。"""
    applied: dict[str, Any] = {}
    font = range_obj.Font
    if size not in (None, ""):
        font.Size = float(size)
        applied["font_size"] = round_pt(float(font.Size))
        if abs(float(font.Size) - float(size)) > 0.1:
            raise ToolError("字号设置后读回不一致。", code="shape/write-mismatch")
    if bold is not None:
        font.Bold = MSO_TRUE if bold else MSO_FALSE
        applied["bold"] = bool(font.Bold)
        if bool(font.Bold) != bool(bold):
            raise ToolError("加粗设置后读回不一致。", code="shape/write-mismatch")
    if color not in (None, ""):
        font.Color.RGB = _rgb(color)
        if int(font.Color.RGB) != _rgb(color):
            raise ToolError("颜色设置后读回不一致。", code="shape/write-mismatch")
        applied["color"] = color
    if align not in (None, ""):
        code = PP_ALIGN.get(str(align).lower())
        if code is None:
            raise ToolError(
                f"对齐方式只能是 {'/'.join(PP_ALIGN)}，收到：{align!r}", code="args/align"
            )
        range_obj.ParagraphFormat.Alignment = code
        if int(range_obj.ParagraphFormat.Alignment) != code:
            raise ToolError("对齐设置后读回不一致。", code="shape/write-mismatch")
        applied["align"] = str(align).lower()
    return applied


def _notes_shape(slide: Any) -> Any:
    """备注正文占位符。逐个退路找，不同演示的备注页结构不一样。"""
    page = safe(lambda: slide.NotesPage)
    if page is None:
        return None
    shapes = safe(lambda: page.Shapes)
    if shapes is None:
        return None
    candidates = []
    count = safe(lambda: int(shapes.Count), 0) or 0
    for i in range(1, count + 1):
        shape = safe(lambda i=i: shapes(i))
        if shape is None:
            continue
        if safe(lambda s=shape: int(s.PlaceholderFormat.Type)) == 2:
            candidates.append(shape)
    for candidate in candidates:
        if candidate is not None and safe(lambda c=candidate: bool(c.HasTextFrame)):
            return candidate
    return None


def _show_slide(app: Any, pres: Any, index: int, shape: Any = None) -> None:
    from .deck import _goto_slide

    _goto_slide(app, pres, index)
    if shape is not None:
        shape.Select()


def _validate_format(args: dict[str, Any]) -> None:
    # 在改文字/创建形状之前拒绝坏颜色，避免失败调用也修改了演示。
    if args.get("color") not in (None, ""):
        _rgb(args["color"])


def _deck_arg() -> Param:
    return Param(name="deck", type="string", description="演示名、路径或序号；留空用当前活动演示")


def _slide_arg(description: str = "页码，从 1 开始；留空用当前正在看的那页") -> Param:
    return Param(name="slide", type="integer", description=description)


# --------------------------------------------------------------------------
# 工具
# --------------------------------------------------------------------------


@tool(
    "ppt_save",
    title="保存演示",
    summary="保存这份演示；也可以先另存一份备份再改",
    behavior="write",
    returns="保存结果；mode=copy 时给出备份文件路径",
    tags=("deck",),
    params=(
        _deck_arg(),
        Param(
            name="mode",
            type="string",
            description="save=原地保存；copy=另存一份备份（不动当前文件）",
            default="save",
            enum=("save", "copy"),
        ),
        Param(name="path", type="string", description="mode=copy 时的备份路径；留空自动放到 ppt-agent 的 backups 目录"),
    ),
)
def op_save(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    info = presentation_info(pres)
    mode = str(args.get("mode") or "save").lower()

    if mode == "copy":
        raw = args.get("path")
        if raw:
            target = absolute_path(raw, what="备份文件", must_exist=False, suffixes={".pptx", ".ppt", ".pptm"})
        else:
            from .deck import _slug

            stamp = time.strftime("%Y%m%d-%H%M%S")
            stem = _slug(Path(str(info.get("name") or "deck")).stem)
            suffix = Path(str(info.get("name") or "deck.pptx")).suffix.lower()
            if suffix not in {".pptx", ".ppt", ".pptm"}:
                suffix = ".pptx"
            target = ctx.settings.home / "backups" / f"{stem}-{stamp}-{uuid.uuid4().hex[:8]}{suffix}"
        target.parent.mkdir(parents=True, exist_ok=True)
        pres.SaveCopyAs(str(target))
        if not target.exists():
            raise ToolError(f"备份没有生成：{target}", code="save/copy-failed")
        return {
            "copied": True,
            "saved": False,
            "copy_path": str(target),
            "bytes": target.stat().st_size,
            "presentation": info,
        }

    if args.get("path"):
        target = absolute_path(args["path"], what="保存路径", must_exist=False, suffixes={".pptx", ".ppt", ".pptm"})
        pres.SaveAs(str(target))
        return {"saved": True, "saved_as": str(target), "presentation": presentation_info(pres)}

    if safe(lambda: str(pres.Path)) == "":
        raise ToolError("新演示尚未保存，请显式提供 path，避免弹出另存为对话框。", code="save/path-required")
    pres.Save()
    return {"saved": True, "presentation": presentation_info(pres), "note": "已原地保存"}


@tool(
    "ppt_add_slide",
    title="新增一页",
    summary="在指定位置插入一页新幻灯片",
    behavior="write",
    returns="新页的页码、所用版式与当前总页数",
    tags=("slide",),
    params=(
        _deck_arg(),
        Param(name="layout", type="string", description="版式名称或序号；留空沿用最后一页的版式"),
        Param(name="index", type="integer", description="插到第几页；留空表示追加到末尾", minimum=1),
    ),
)
def op_add_slide(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    total = _slide_count(pres)
    index = args.get("index")
    index = total + 1 if index in (None, "") else int(index)
    index = max(1, min(index, total + 1))

    fallback = _slide_obj(pres, total) if total else None
    layout, layout_name = _resolve_layout(pres, args.get("layout"), fallback_slide=fallback)
    slide = pres.Slides.AddSlide(index, layout)
    if slide is None:
        raise ToolError("插入幻灯片失败。", code="slide/add-failed")
    _show_slide(app, pres, index)
    return {
        "added": True,
        "slide": index,
        "layout": layout_name,
        "shape_count": safe(lambda: int(slide.Shapes.Count), 0),
        "total": _slide_count(pres),
        "presentation": presentation_info(pres),
    }


@tool(
    "ppt_delete_slide",
    title="删除一页",
    summary="删除指定的幻灯片（必须显式给页码，不会默认删当前页）",
    behavior="destroy",
    returns="删除了第几页、剩余总页数",
    tags=("slide",),
    params=(
        _deck_arg(),
        Param(name="slide", type="integer", description="要删除的页码，从 1 开始（必填）", required=True),
        Param(name="count", type="integer", description="连续删除几页", default=1, minimum=1, maximum=50),
    ),
)
def op_delete_slide(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    first = _slide_at(app, pres, args.get("slide"), required=True)
    wanted = max(1, int(args.get("count") or 1))
    total = _slide_count(pres)
    actual = min(wanted, total - first + 1)
    title = ""
    _show_slide(app, pres, first)
    for _ in range(actual):
        slide = _slide_obj(pres, first)
        if not title:
            from .common import slide_title

            title = slide_title(slide)
        slide.Delete()
    remaining = _slide_count(pres)
    if remaining:
        _show_slide(app, pres, min(first, remaining))
    return {
        "deleted": actual,
        "from": first,
        "deleted_title": title,
        "total": _slide_count(pres),
        "presentation": presentation_info(pres),
    }


@tool(
    "ppt_move_slide",
    title="移动一页",
    summary="把某一页挪到另一个位置",
    behavior="write",
    returns="移动后的页码",
    tags=("slide",),
    params=(
        _deck_arg(),
        Param(name="slide", type="integer", description="要移动的页码（必填）", required=True),
        Param(name="to", type="integer", description="移动到第几页（必填）", required=True, minimum=1),
    ),
)
def op_move_slide(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    total = _slide_count(pres)
    source = _slide_at(app, pres, args.get("slide"), required=True)
    target = max(1, min(int(args["to"]), total))
    _show_slide(app, pres, source)
    _slide_obj(pres, source).MoveTo(target)
    _show_slide(app, pres, target)
    return {"moved": True, "from": source, "to": target, "total": total}


@tool(
    "ppt_set_text",
    title="改写形状里的文字",
    summary="把某一页某个形状里的文字替换掉（可同时设字号、加粗、颜色、对齐）",
    behavior="write",
    returns="被改形状的名称与序号，以及实际生效的格式",
    tags=("shape", "text"),
    params=(
        _deck_arg(),
        _slide_arg(),
        Param(name="shape", type="string", description="形状名、序号或组合 ref 路径（如 3/2；可用 title 指标题占位符）；先从 ppt_read_slide 里看", required=True),
        Param(name="text", type="string", description="新的文字内容；空字符串清空文字", required=True, allow_empty=True),
        Param(name="font_size", type="number", description="字号（磅）", minimum=1, maximum=4000),
        Param(name="bold", type="boolean", description="是否加粗"),
        Param(name="color", type="string", description="文字颜色，形如 #4D6BFE"),
        Param(name="align", type="string", description="段落对齐", enum=("left", "center", "right", "justify")),
    ),
)
def op_set_text(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _slide_at(app, pres, args.get("slide"))
    slide = _slide_obj(pres, index)
    shape, shape_index, name = resolve_shape(slide, args.get("shape"))

    if not safe(lambda: bool(shape.HasTextFrame)):
        raise ToolError(
            f"形状「{name}」不能放文字（例如图片）。",
            code="shape/no-text-frame",
            available=shape_names(slide),
        )
    frame = safe(lambda: shape.TextFrame)
    if frame is None:
        raise ToolError(f"形状「{name}」没有文本框。", code="shape/no-text-frame")

    text = str(args["text"])
    _validate_format(args)
    _show_slide(app, pres, index, shape)
    before = safe(lambda: str(frame.TextRange.Text))
    frame.TextRange.Text = text
    applied = _apply_font(
        frame.TextRange,
        size=args.get("font_size"),
        bold=args.get("bold"),
        color=args.get("color"),
        align=args.get("align"),
    )
    after = safe(lambda: str(frame.TextRange.Text))
    if after is None or after.replace("\r\n", "\n").replace("\r", "\n") != text.replace("\r\n", "\n").replace("\r", "\n"):
        raise ToolError(
            f"写入后读回的文本不一致（期望 {text!r}，得到 {after!r}）。",
            code="shape/write-mismatch",
        )
    return {
        "updated": True,
        "slide": index,
        "shape": name,
        "shape_index": shape_index,
        "before": before,
        "after": after,
        "chars": len(text),
        "applied": applied,
    }


@tool(
    "ppt_add_textbox",
    title="加一个文本框",
    summary="在某一页加一个文本框并写入文字",
    behavior="write",
    returns="新形状的名称与序号，后续可用它来改文字或位置",
    tags=("shape", "text"),
    params=(
        _deck_arg(),
        _slide_arg(),
        Param(name="text", type="string", description="文本框里的文字", required=True, allow_empty=True),
        Param(name="left", type="number", description="左边距（磅）", default=60.0),
        Param(name="top", type="number", description="上边距（磅）", default=60.0),
        Param(name="width", type="number", description="宽度（磅）", default=400.0, minimum=0.01),
        Param(name="height", type="number", description="高度（磅）", default=80.0, minimum=0.01),
        Param(name="font_size", type="number", description="字号（磅）", minimum=1, maximum=4000),
        Param(name="bold", type="boolean", description="是否加粗"),
        Param(name="color", type="string", description="文字颜色，形如 #4D6BFE"),
        Param(name="align", type="string", description="段落对齐", enum=("left", "center", "right", "justify")),
    ),
)
def op_add_textbox(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _slide_at(app, pres, args.get("slide"))
    slide = _slide_obj(pres, index)
    _validate_format(args)
    _show_slide(app, pres, index)

    textbox = slide.Shapes.AddTextbox(
        MSO_TEXT_HORIZONTAL,
        float(args["left"] if args.get("left") is not None else 60.0),
        float(args["top"] if args.get("top") is not None else 60.0),
        float(args.get("width") or 400.0),
        float(args.get("height") or 80.0),
    )
    if textbox is None:
        raise ToolError("添加文本框失败。", code="shape/add-failed")
    textbox.TextFrame.TextRange.Text = str(args["text"])
    applied = _apply_font(
        textbox.TextFrame.TextRange,
        size=args.get("font_size"),
        bold=args.get("bold"),
        color=args.get("color"),
        align=args.get("align"),
    )
    name = str(safe(lambda: textbox.Name) or "")
    textbox.Select()
    return {
        "added": True,
        "slide": index,
        "shape": name,
        "shape_index": _index_of(slide, textbox),
        "text": str(args["text"]),
        "applied": applied,
        "shape_count": safe(lambda: int(slide.Shapes.Count), 0),
    }


@tool(
    "ppt_set_geometry",
    title="移动或缩放形状",
    summary="改某个形状的位置与尺寸（只改你给的项）",
    behavior="write",
    returns="改动前后的位置尺寸",
    tags=("shape",),
    params=(
        _deck_arg(),
        _slide_arg(),
        Param(name="shape", type="string", description="形状名、序号或组合 ref 路径（如 3/2；可用 title 指标题占位符）", required=True),
        Param(name="left", type="number", description="左边距（磅）"),
        Param(name="top", type="number", description="上边距（磅）"),
        Param(name="width", type="number", description="宽度（磅）", minimum=0.01),
        Param(name="height", type="number", description="高度（磅）", minimum=0.01),
        Param(name="rotation", type="number", description="旋转角度"),
    ),
)
def op_set_geometry(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _slide_at(app, pres, args.get("slide"))
    slide = _slide_obj(pres, index)
    shape, shape_index, name = resolve_shape(slide, args.get("shape"))
    _show_slide(app, pres, index, shape)

    before = {
        "left": round_pt(safe(lambda: shape.Left)),
        "top": round_pt(safe(lambda: shape.Top)),
        "width": round_pt(safe(lambda: shape.Width)),
        "height": round_pt(safe(lambda: shape.Height)),
        "rotation": round_pt(safe(lambda: shape.Rotation)),
    }
    changed: list[str] = []
    for key in ("left", "top", "width", "height", "rotation"):
        value = args.get(key)
        if value in (None, ""):
            continue
        attr = key.capitalize()
        setattr(shape, attr, float(value))
        actual = float(getattr(shape, attr))
        expected = float(value) % 360 if key == "rotation" else float(value)
        if abs(actual - expected) > 0.1:
            raise ToolError(f"{key} 设置后读回不一致（期望 {expected}，得到 {actual}）。", code="shape/write-mismatch")
        changed.append(key)
    if not changed:
        raise ToolError("至少要给 left/top/width/height/rotation 中的一个。", code="args/empty")

    after = {
        "left": round_pt(safe(lambda: shape.Left)),
        "top": round_pt(safe(lambda: shape.Top)),
        "width": round_pt(safe(lambda: shape.Width)),
        "height": round_pt(safe(lambda: shape.Height)),
        "rotation": round_pt(safe(lambda: shape.Rotation)),
    }
    return {
        "updated": True,
        "slide": index,
        "shape": name,
        "shape_index": shape_index,
        "changed": changed,
        "before": before,
        "after": after,
    }


@tool(
    "ppt_delete_shape",
    title="删除形状",
    summary="删掉某一页里的某个形状",
    behavior="destroy",
    returns="被删形状的名称与剩余形状数",
    tags=("shape",),
    params=(
        _deck_arg(),
        _slide_arg(),
        Param(name="shape", type="string", description="形状名或序号（必填，不会默认删某个形状）", required=True),
    ),
)
def op_delete_shape(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _slide_at(app, pres, args.get("slide"))
    slide = _slide_obj(pres, index)
    shape, shape_index, name = resolve_shape(slide, args.get("shape"))
    shape_type = safe(lambda: int(shape.Type))
    _show_slide(app, pres, index, shape)
    shape.Delete()
    return {
        "deleted": True,
        "slide": index,
        "shape": name,
        "shape_index": shape_index,
        "shape_type": shape_type,
        "shape_count": safe(lambda: int(slide.Shapes.Count), 0),
    }


@tool(
    "ppt_add_picture",
    title="插入图片",
    summary="把一张本地图片插到某一页（必须给绝对路径）",
    behavior="write",
    returns="新图片形状的名称与序号",
    tags=("shape", "media"),
    params=(
        _deck_arg(),
        _slide_arg(),
        Param(name="path", type="string", description="图片的绝对路径", required=True),
        Param(name="left", type="number", description="左边距（磅）", default=60.0),
        Param(name="top", type="number", description="上边距（磅）", default=60.0),
        Param(name="width", type="number", description="宽度（磅）；留空按原图尺寸", minimum=0.01),
        Param(name="height", type="number", description="高度（磅）；留空按原图尺寸", minimum=0.01),
    ),
)
def op_add_picture(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _slide_at(app, pres, args.get("slide"))
    slide = _slide_obj(pres, index)
    picture = absolute_path(args.get("path"), what="图片", suffixes=PICTURE_SUFFIXES)

    left = float(args["left"] if args.get("left") is not None else 60.0)
    top = float(args["top"] if args.get("top") is not None else 60.0)
    width = args.get("width")
    height = args.get("height")
    _show_slide(app, pres, index)
    try:
        if width not in (None, "") or height not in (None, ""):
            shape = slide.Shapes.AddPicture(
                str(picture), MSO_FALSE, MSO_TRUE, left, top,
                float(width) if width not in (None, "") else -1,
                float(height) if height not in (None, "") else -1,
            )
        else:
            shape = slide.Shapes.AddPicture(str(picture), MSO_FALSE, MSO_TRUE, left, top)
    except Exception as exc:  # noqa: BLE001 - 图片格式不被支持等
        raise ToolError(
            f"插入图片失败：{type(exc).__name__}: {exc}。请确认路径是绝对路径、文件是 PowerPoint 支持的格式。",
            code="picture/add-failed",
            path=str(picture),
        ) from exc
    if shape is None:
        raise ToolError("插入图片失败。", code="picture/add-failed")

    name = str(safe(lambda: shape.Name) or "")
    shape.Select()
    return {
        "added": True,
        "slide": index,
        "shape": name,
        "shape_index": _index_of(slide, shape),
        "source": str(picture),
        "geometry": {
            "left": round_pt(safe(lambda: shape.Left)),
            "top": round_pt(safe(lambda: shape.Top)),
            "width": round_pt(safe(lambda: shape.Width)),
            "height": round_pt(safe(lambda: shape.Height)),
        },
        "shape_count": safe(lambda: int(slide.Shapes.Count), 0),
    }


@tool(
    "ppt_set_notes",
    title="写演讲者备注",
    summary="替换某一页的备注文字",
    behavior="write",
    returns="写入后的备注内容",
    tags=("slide", "notes"),
    params=(
        _deck_arg(),
        _slide_arg(),
        Param(name="text", type="string", description="备注内容；传空字符串表示清空", required=True, allow_empty=True),
    ),
)
def op_set_notes(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _slide_at(app, pres, args.get("slide"))
    slide = _slide_obj(pres, index)
    shape = _notes_shape(slide)
    if shape is None:
        raise ToolError("这一页的备注页里找不到可写文字的占位符。", code="notes/no-placeholder")
    _show_slide(app, pres, index)
    before = safe(lambda: str(shape.TextFrame.TextRange.Text))
    shape.TextFrame.TextRange.Text = str(args["text"])
    after = safe(lambda: str(shape.TextFrame.TextRange.Text))
    return {"updated": True, "slide": index, "before": before, "after": after}


@tool(
    "ppt_set_table_cell",
    title="编辑表格单元格",
    summary="修改真实 PowerPoint 表格中指定单元格的文字与格式",
    behavior="write",
    returns="目标页、表格、行列和修改前后的文字",
    tags=("shape", "text", "table"),
    params=(
        _deck_arg(), _slide_arg(),
        Param(name="shape", type="string", description="表格形状名称或 ref（支持组合序号路径，如 3/2）", required=True),
        Param(name="row", type="integer", description="行号，从 1 开始", required=True, minimum=1),
        Param(name="column", type="integer", description="列号，从 1 开始", required=True, minimum=1),
        Param(name="text", type="string", description="单元格文字；空字符串清空", required=True, allow_empty=True),
        Param(name="font_size", type="number", description="字号（磅）", minimum=1, maximum=4000),
        Param(name="bold", type="boolean", description="是否加粗"),
        Param(name="color", type="string", description="文字颜色，如 #4D6BFE"),
        Param(name="align", type="string", description="段落对齐", enum=("left", "center", "right", "justify")),
    ),
)
def op_set_table_cell(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _slide_at(app, pres, args.get("slide"))
    shape, shape_index, name = resolve_shape(_slide_obj(pres, index), args["shape"])
    if not safe(lambda: bool(shape.HasTable)):
        raise ToolError(f"形状「{name}」不是表格。", code="table/not-table")
    table = shape.Table
    row, column = int(args["row"]), int(args["column"])
    if not 1 <= row <= int(table.Rows.Count) or not 1 <= column <= int(table.Columns.Count):
        raise ToolError("单元格行列超出表格范围。", code="table/out-of-range")
    _validate_format(args)
    _show_slide(app, pres, index, shape)
    target = table.Cell(row, column).Shape.TextFrame.TextRange
    before = str(target.Text)
    target.Text = args["text"]
    applied = _apply_font(target, size=args.get("font_size"), bold=args.get("bold"), color=args.get("color"), align=args.get("align"))
    after = str(target.Text)
    if after.replace("\r\n", "\n").replace("\r", "\n") != args["text"].replace("\r\n", "\n").replace("\r", "\n"):
        raise ToolError("表格单元格写入后读回不一致。", code="table/write-mismatch")
    return {"updated": True, "slide": index, "shape": name, "shape_index": shape_index,
        "row": row, "column": column, "before": before, "after": after, "applied": applied}


@tool(
    "ppt_select",
    title="选中并跳页",
    summary="跳到某一页并选中某个形状，让用户看得见 agent 正在动哪里",
    behavior="idempotent",
    returns="当前页码与选中的形状",
    tags=("visibility", "slide"),
    params=(
        _deck_arg(),
        _slide_arg("页码；留空表示不跳页"),
        Param(name="shape", type="string", description="要选中的形状名或序号；留空只跳页不选中"),
    ),
)
def op_select(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    from .deck import PP_VIEW_NORMAL, _goto_slide

    pres = resolve_deck(app, args.get("deck"))
    current = safe(lambda: int(pres.Windows(1).View.Slide.SlideIndex)) or 1
    index = _slide_at(app, pres, args.get("slide")) if args.get("slide") not in (None, "") else current
    actual = _goto_slide(app, pres, index)

    selected: str | None = None
    if args.get("shape") not in (None, ""):
        slide = _slide_obj(pres, index)
        shape, shape_index, name = resolve_shape(slide, args.get("shape"))
        # 只做选中，不改文档——给用户的形状画描边等于污染 deck
        shape.Select()
        selected = f"{name}（第 {shape_index} 个形状）"
        bring_to_front(app, pres)
    return {
        "slide": index,
        "current_slide": actual if actual is not None else index,
        "selected": selected,
        "view_normal": safe(lambda: int(pres.Windows(1).ViewType)) == PP_VIEW_NORMAL,
    }


@tool(
    "ppt_undo",
    title="撤销",
    summary="撤销上一步操作（用的是 PowerPoint 自己的撤销栈）",
    behavior="write",
    returns="实际撤销了几步、为什么停下",
    tags=("slide", "shape"),
    params=(
        _deck_arg(),
        Param(name="steps", type="integer", description="最多撤销几步", default=1, minimum=1, maximum=20),
    ),
)
def op_undo(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    steps = max(1, min(int(args.get("steps") or 1), 20))
    _show_slide(app, pres, _slide_at(app, pres, None))
    restored = 0
    stopped: str | None = None
    for _ in range(steps):
        try:
            app.CommandBars.ExecuteMso("Undo")
        except Exception as exc:  # noqa: BLE001 - 撤到底时 PowerPoint 会抛 com_error
            stopped = f"撤销栈已到底或当前不可用（{type(exc).__name__}）"
            break
        restored += 1
    return {
        "undone": restored,
        "requested": steps,
        "stopped": stopped,
        "total_slides": _slide_count(pres),
        "presentation": presentation_info(pres),
        "caveat": (
            "撤销粒度由 PowerPoint 决定，不保证与一次工具调用一一对应——"
            "实测一次 Undo 可能回退得比上一步更多。"
        ),
        "safety": "要动一批破坏性操作前，先 ppt_save(mode='copy') 留一份备份。",
    }


@tool(
    "ppt_set_pace",
    title="设置跟速",
    summary="设置每一步之间的停顿，好让人跟得上 agent 的操作",
    behavior="idempotent",
    needs_com=False,
    returns="生效后的停顿毫秒数",
    tags=("visibility",),
    params=(
        Param(
            name="pace_ms",
            type="integer",
            description="写操作之间停顿的毫秒数；0=极速，400≈正常，1500=慢速看得清",
            default=DEFAULT_PACE_MS,
            minimum=0,
            maximum=10000,
        ),
    ),
)
def op_set_pace(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    value = max(0, min(int(args.get("pace_ms") or 0), 10000))
    ctx.state["pace_ms"] = value
    ctx.emit("state", pace_ms=value)
    label = "极速" if value == 0 else ("慢速" if value >= 1000 else "正常")
    return {"pace_ms": value, "label": label}
