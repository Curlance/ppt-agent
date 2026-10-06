"""M1 工具集：读懂演示，并把画面拿给人看和给模型看。

关键设计（由真机实测得出）
--------------------------
``Slide.Export`` 导出的 1280×720 PNG 有 **1.5 MB**，直接塞给模型既慢又贵。
所以每次截图都产出两份：

- **高清 PNG**（默认 1280 宽）→ 观察台与任务窗格用，浏览器从本地 HTTP 取；
- **压缩 JPEG 缩略图**（默认 1024 宽、质量 82）→ 内联给模型看。

这样"人看得清"和"模型看得见"同时成立，且不必让模型吞 1.5 MB。
"""

from __future__ import annotations

import time
import uuid
from pathlib import Path
from typing import Any

from ..config import ensure_dirs
from ..errors import ToolError
from ..registry import Context, Param, tool
from .common import (
    bring_to_front,
    list_presentations,
    presentation_info,
    resolve_deck,
    round_pt,
    safe,
    shape_type_name,
    slide_title,
)

__all__ = [
    "op_decks",
    "op_slides",
    "op_read_slide",
    "op_goto",
    "op_shot",
]

#: PowerPoint 的 PpViewType 常量：切页前要确保窗口在普通视图。
PP_VIEW_NORMAL = 9
#: 单次截图最多导出多少页（避免一次生成几十 MB 图片）。
MAX_SHOTS_PER_CALL = 10
#: 截图目录保留的文件数上限。
SHOT_KEEP = 400


# --------------------------------------------------------------------------
# 内部辅助
# --------------------------------------------------------------------------


def _clamp(value: int, low: int, high: int) -> int:
    return max(low, min(high, value))


def _page_size(pres: Any) -> tuple[float, float]:
    setup = safe(lambda: pres.PageSetup)
    width = float(safe(lambda: setup.SlideWidth, 960.0) or 960.0)
    height = float(safe(lambda: setup.SlideHeight, 540.0) or 540.0)
    return width, height


def _slide_count(pres: Any) -> int:
    return safe(lambda: int(pres.Slides.Count), 0) or 0


def _current_slide_index(app: Any, pres: Any) -> int:
    """优先取这份演示自己窗口的当前页，退化到全局活动窗口。"""
    index = safe(lambda: int(pres.Windows(1).View.Slide.SlideIndex))
    return int(index) if index else 1


def _resolve_slide(app: Any, pres: Any, wanted: Any) -> int:
    """把 ``slide`` 参数解析成合法页码（留空 = 当前页）。"""
    total = _slide_count(pres)
    if total == 0:
        raise ToolError("这份演示里没有任何幻灯片。", code="slide/empty")
    if wanted is None or wanted == "":
        return _clamp(_current_slide_index(app, pres), 1, total)
    try:
        index = int(wanted)
    except (TypeError, ValueError) as exc:
        raise ToolError(f"页码必须是整数，收到：{wanted!r}", code="args/type", parameter="slide") from exc
    if not 1 <= index <= total:
        raise ToolError(
            f"页码 {index} 超出范围（这份演示共 {total} 页）。",
            code="slide/out-of-range",
            total=total,
        )
    return index


def _custom_layout_name(slide: Any) -> str | None:
    # 实测 Slide.Layout.Name 会返回 None，CustomLayout.Name 才是可用的那个。
    name = safe(lambda: str(slide.CustomLayout.Name))
    return name or None


def _notes_text(slide: Any) -> str:
    """备注页里第一段非空文字。逐个形状找，比固定取 Shapes(2) 可靠。"""
    page = safe(lambda: slide.NotesPage)
    if page is None:
        return ""
    shapes = safe(lambda: page.Shapes)
    count = safe(lambda: int(shapes.Count), 0) or 0
    for i in range(1, count + 1):
        shape = safe(lambda i=i: shapes(i))
        if shape is None:
            continue
        if safe(lambda s=shape: int(s.PlaceholderFormat.Type)) != 2:
            continue
        text = safe(lambda s=shape: str(s.TextFrame.TextRange.Text))
        if text and text.strip():
            return text.strip()
    return ""


def _shape_info(shape: Any, *, include_text: bool, include_geometry: bool, ref: str = "", budget: list[int] | None = None, depth: int = 0, max_table_cells: int = 200) -> dict[str, Any]:
    raw_type = safe(lambda: int(shape.Type))
    info: dict[str, Any] = {
        "name": safe(lambda: str(shape.Name)),
        "type": shape_type_name(raw_type),
        "z": safe(lambda: int(shape.ZOrderPosition)),
    }
    if ref:
        info["ref"] = ref
    if budget is not None:
        budget[0] -= 1
    if include_geometry:
        info["geometry"] = {
            "left": round_pt(safe(lambda: shape.Left)),
            "top": round_pt(safe(lambda: shape.Top)),
            "width": round_pt(safe(lambda: shape.Width)),
            "height": round_pt(safe(lambda: shape.Height)),
            "rotation": round_pt(safe(lambda: shape.Rotation)),
        }
    if include_text and safe(lambda: bool(shape.HasTextFrame)) and safe(lambda: bool(shape.TextFrame.HasText)):
        text = safe(lambda: str(shape.TextFrame.TextRange.Text)) or ""
        info["text"] = text.replace("\r", "\n").strip()
        font = safe(lambda: shape.TextFrame.TextRange.Font)
        if font is not None:
            info["font"] = {
                "name": safe(lambda: str(font.Name)),
                "size": round_pt(safe(lambda: float(font.Size))),
                "bold": safe(lambda: bool(font.Bold)),
            }
    if safe(lambda: bool(shape.HasTable)):
        table = safe(lambda: shape.Table)
        rows = safe(lambda: int(table.Rows.Count), 0) or 0
        cols = safe(lambda: int(table.Columns.Count), 0) or 0
        info["table"] = {"rows": rows, "cols": cols, "header": _table_header(table, rows, cols)}
        if include_text:
            cells = []
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    if len(cells) >= max_table_cells:
                        break
                    text = safe(lambda r=r, c=c: str(table.Cell(r, c).Shape.TextFrame.TextRange.Text))
                    cells.append({"row": r, "column": c, "text": text})
                if len(cells) >= max_table_cells:
                    break
            info["table"].update(cells=cells, truncated=rows * cols > len(cells))
    if safe(lambda: bool(shape.HasChart)):
        chart = safe(lambda: shape.Chart)
        info["chart"] = {
            "chart_type": safe(lambda: int(chart.ChartType)),
            "has_title": safe(lambda: bool(chart.HasTitle)),
            "title": safe(lambda: str(chart.ChartTitle.Text)),
        }
    if raw_type == 6:  # msoGroup
        info["group_items"] = safe(lambda: int(shape.GroupItems.Count))
        children = []
        if depth < 8:
            for i in range(1, (info["group_items"] or 0) + 1):
                if budget is not None and budget[0] <= 0:
                    break
                child = safe(lambda i=i: shape.GroupItems(i))
                if child is not None:
                    children.append(_shape_info(child, include_text=include_text, include_geometry=include_geometry,
                        ref=f"{ref}/{i}", budget=budget, depth=depth + 1, max_table_cells=max_table_cells))
        info["children"] = children
        info["children_truncated"] = len(children) < (info["group_items"] or 0)
    if raw_type == 14:  # msoPlaceholder
        info["placeholder_type"] = safe(lambda: int(shape.PlaceholderFormat.Type))
    return info


def _table_header(table: Any, rows: int, cols: int) -> list[str] | None:
    """只取首行文字当表头——整张表塞进 JSON 会让上下文爆掉。"""
    if rows < 1 or cols < 1:
        return None
    header: list[str] = []
    for c in range(1, min(cols, 12) + 1):
        cell = safe(lambda c=c: table.Cell(1, c))
        text = safe(lambda cell=cell: str(cell.Shape.TextFrame.TextRange.Text)) if cell is not None else None
        header.append((text or "").strip())
    return header


def _slug(text: str) -> str:
    keep = [ch if (ch.isalnum() or ch in "-_") else "-" for ch in text]
    slug = "".join(keep).strip("-")
    return (slug or "deck")[:40]


def _make_thumb(source: Path, width: int) -> dict[str, Any] | None:
    """生成给模型看的压缩 JPEG 缩略图。Pillow 不可用时返回 None（不致命）。"""
    try:
        from PIL import Image
    except Exception:  # noqa: BLE001 - 环境缺 Pillow
        return None
    target = source.with_suffix(".thumb.jpg")
    try:
        with Image.open(source) as image:
            rgb = image.convert("RGB")
            ratio = width / float(rgb.width)
            size = (width, max(1, round(rgb.height * ratio)))
            rgb = rgb.resize(size, Image.LANCZOS)
            rgb.save(target, "JPEG", quality=82, optimize=True)
    except Exception:  # noqa: BLE001 - 生成缩略图失败不该让截图整体失败
        return None
    return {
        "path": str(target),
        "url": f"/shots/{target.name}",
        "width": size[0],
        "height": size[1],
        "bytes": target.stat().st_size if target.exists() else None,
    }


def _prune_shots(directory: Path, keep: int = SHOT_KEEP) -> int:
    """截图目录只保留最新的若干个文件，避免无限增长。"""
    try:
        files = sorted(directory.glob("*"), key=lambda p: p.stat().st_mtime, reverse=True)
    except OSError:
        return 0
    removed = 0
    for stale in files[keep:]:
        try:
            stale.unlink()
            removed += 1
        except OSError:
            pass
    return removed


def _goto_slide(app: Any, pres: Any, index: int) -> int | None:
    """把视图切到指定页，让用户的视线跟得上。返回切换后的页号。"""
    window = safe(lambda: pres.Windows(1))
    if window is None:
        raise ToolError("目标演示没有可见窗口，操作已停止。", code="visibility/no-window")
    try:
        app.Visible = True
        window.Activate()
        if int(window.ViewType) != PP_VIEW_NORMAL:
            window.ViewType = PP_VIEW_NORMAL
        window.View.GotoSlide(index)
        actual = int(window.View.Slide.SlideIndex)
        if actual != index:
            raise RuntimeError(f"当前页仍是 {actual}")
    except Exception as exc:
        raise ToolError(f"无法显示目标第 {index} 页，操作已停止：{exc}", code="visibility/goto-failed") from exc
    bring_to_front(app, pres)
    return actual


#: ``capture_live_frame`` 的两种"拍不到"返回值（用整数而不是 None，是为了区分原因）
LIVE_NO_PRESENTATION = 0
LIVE_NO_SLIDES = -1


def capture_live_frame(app: Any, target: Path, *, width: int = 960, slide: Any = None) -> int:
    """把"此刻这一页"导出成一张固定路径的 JPEG，供实时预览反复覆盖。

    与 ``ppt_shot`` 的区别：

    - ``ppt_shot`` 是**留档**：唯一文件名、同时产高清图与缩略图，进事件流；
    - ``capture_live_frame`` 是**取景**：永远覆盖同一个文件、不产缩略图、不进事件流。

    因此它很便宜、可以每隔一两秒调一次；而且因为它读的是 PowerPoint **当前**状态，
    用户自己在 PowerPoint 里翻页或改内容，观察台也看得见。

    返回导出的页码；拍不到时返回两个哨兵值之一，让调用方能给出**准确**的原因：

    - :data:`LIVE_NO_PRESENTATION` —— 没有打开的演示；
    - :data:`LIVE_NO_SLIDES` —— 演示在，但一页都没有。
    """
    pres = safe(lambda: app.ActivePresentation)
    if pres is None:
        return LIVE_NO_PRESENTATION
    total = _slide_count(pres)
    if total == 0:
        return LIVE_NO_SLIDES

    if slide in (None, ""):
        index = safe(lambda: int(pres.Windows(1).View.Slide.SlideIndex))
        if not index:
            index = safe(lambda: int(app.ActiveWindow.View.Slide.SlideIndex))
        index = int(index) if index else 1
    else:
        index = int(slide)
    index = max(1, min(index, total))

    page_w, page_h = _page_size(pres)
    export_width = _clamp(int(width), 320, 1920)
    height = max(1, round(export_width * page_h / page_w))
    pres.Slides(index).Export(str(target), "JPG", export_width, height)
    return index if target.exists() else LIVE_NO_SLIDES


# --------------------------------------------------------------------------
# 工具
# --------------------------------------------------------------------------


@tool(
    "ppt_decks",
    title="列出打开的演示",
    summary="列出 PowerPoint 里打开着的演示，以及各自停在第几页",
    behavior="read",
    returns="每份演示的名称、路径、页数、页面尺寸，以及哪一份是活动演示",
    tags=("read", "deck"),
    params=(),
)
def op_decks(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    decks = list_presentations(app)
    for index, item in enumerate(decks, start=1):
        pres = safe(lambda i=index: app.Presentations(i))
        width, height = _page_size(pres) if pres is not None else (None, None)
        item["slide_size"] = {"width_pt": round_pt(width), "height_pt": round_pt(height)}
        item["current_slide"] = safe(lambda p=pres: int(p.Windows(1).View.Slide.SlideIndex)) if pres else None

    active = safe(lambda: presentation_info(app.ActivePresentation))
    active_index = next((d["index"] for d in decks if active and d["name"] == active.get("name")), None)
    return {"count": len(decks), "active_index": active_index, "decks": decks}


@tool(
    "ppt_slides",
    title="列出每一页",
    summary="逐页列出标题与形状数量，快速了解这份演示讲了什么",
    behavior="read",
    returns="页码、标题、形状数；可指定范围并对结果做截断",
    tags=("read", "deck"),
    params=(
        Param(name="deck", type="string", description="演示名、路径或序号；留空用当前活动演示"),
        Param(name="from_slide", type="integer", description="起始页，从 1 开始", default=1),
        Param(name="to_slide", type="integer", description="结束页（含）；留空表示到最后一页"),
        Param(name="limit", type="integer", description="本次最多返回多少页", default=60, minimum=1, maximum=300),
    ),
)
def op_slides(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    total = _slide_count(pres)
    if total == 0:
        return {"deck": presentation_info(pres), "total": 0, "returned": 0, "slides": []}

    first = _clamp(int(args.get("from_slide") or 1), 1, total)
    last = int(args.get("to_slide") or total)
    last = _clamp(last, first, total)
    limit = _clamp(int(args.get("limit") or 60), 1, 300)
    truncated = (last - first + 1) > limit
    if truncated:
        last = first + limit - 1

    rows: list[dict[str, Any]] = []
    for index in range(first, last + 1):
        slide = safe(lambda i=index: pres.Slides(i))
        if slide is None:
            continue
        rows.append(
            {
                "slide": index,
                "title": slide_title(slide),
                "shapes": safe(lambda s=slide: int(s.Shapes.Count)),
                "has_notes": bool(_notes_text(slide)),
            }
        )
    return {
        "deck": presentation_info(pres),
        "total": total,
        "range": [first, last],
        "returned": len(rows),
        "truncated": truncated,
        "slides": rows,
    }


@tool(
    "ppt_read_slide",
    title="读取一页的形状",
    summary="读取某一页里所有形状的文字、位置、尺寸与类型",
    behavior="read",
    returns="该页版式、备注与形状明细；包含表格单元格、组合子形状及可用于编辑的 ref 序号路径",
    tags=("read", "slide"),
    params=(
        Param(name="deck", type="string", description="演示名、路径或序号；留空用当前活动演示"),
        Param(name="slide", type="integer", description="页码，从 1 开始；留空用当前正在看的那页"),
        Param(name="include_text", type="boolean", description="是否读取文字内容", default=True),
        Param(name="include_geometry", type="boolean", description="是否读取位置与尺寸", default=True),
        Param(name="max_shapes", type="integer", description="最多返回多少个形状（含组合子形状）", default=80, minimum=1, maximum=400),
        Param(name="max_table_cells", type="integer", description="每个表格最多读取多少个单元格；截断会明确标记", default=200, minimum=1, maximum=2000),
    ),
)
def op_read_slide(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _resolve_slide(app, pres, args.get("slide"))
    slide = pres.Slides(index)

    include_text = bool(args.get("include_text", True))
    include_geometry = bool(args.get("include_geometry", True))
    max_shapes = _clamp(int(args.get("max_shapes") or 80), 1, 400)

    total_shapes = safe(lambda: int(slide.Shapes.Count), 0) or 0
    shapes: list[dict[str, Any]] = []
    budget = [max_shapes]
    for i in range(1, min(total_shapes, max_shapes) + 1):
        if budget[0] <= 0:
            break
        shape = safe(lambda i=i: slide.Shapes(i))
        if shape is not None:
            shapes.append(_shape_info(shape, include_text=include_text, include_geometry=include_geometry,
                ref=str(i), budget=budget, max_table_cells=int(args.get("max_table_cells") or 200)))

    width, height = _page_size(pres)
    return {
        "deck": presentation_info(pres),
        "slide": index,
        "total_slides": _slide_count(pres),
        "layout": _custom_layout_name(slide),
        "slide_size": {"width_pt": round_pt(width), "height_pt": round_pt(height)},
        "notes": _notes_text(slide),
        "shape_count": total_shapes,
        "returned": len(shapes),
        "truncated": total_shapes > len(shapes),
        "shapes": shapes,
    }


@tool(
    "ppt_goto",
    title="跳到某一页",
    summary="把 PowerPoint 切到指定页，让用户看得见你在看哪一页",
    behavior="idempotent",
    returns="切换后的页码",
    tags=("visibility", "slide"),
    params=(
        Param(name="deck", type="string", description="演示名、路径或序号；留空用当前活动演示"),
        Param(name="slide", type="integer", description="目标页码，从 1 开始", required=True),
    ),
)
def op_goto(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    index = _resolve_slide(app, pres, args.get("slide"))
    actual = _goto_slide(app, pres, index)
    return {
        "deck": presentation_info(pres),
        "requested": index,
        "current_slide": actual if actual is not None else index,
    }


@tool(
    "ppt_shot",
    title="截图查看某一页",
    summary="把某一页导出成图片：给人看高清 PNG，给模型看压缩缩略图",
    behavior="read",
    returns="每页的高清图片与缩略图（含路径与可访问 URL）",
    tags=("read", "slide", "vision"),
    params=(
        Param(name="deck", type="string", description="演示名、路径或序号；留空用当前活动演示"),
        Param(name="slide", type="integer", description="起始页码；留空用当前正在看的那页"),
        Param(name="count", type="integer", description="连续导出几页", default=1, minimum=1, maximum=MAX_SHOTS_PER_CALL),
        Param(name="width", type="integer", description="高清图宽度（像素）", default=1280, minimum=320, maximum=3840),
        Param(
            name="format",
            type="string",
            description="高清图格式：png 清晰、jpg 更小",
            default="png",
            enum=("png", "jpg"),
        ),
        Param(name="thumb_width", type="integer", description="给模型看的缩略图宽度（像素）", default=1024, minimum=320, maximum=1600),
        Param(name="inline", type="boolean", description="是否生成给模型内联查看的缩略图", default=True),
        Param(name="goto", type="boolean", description="截图后把视图切到第一张，便于用户跟随", default=True),
    ),
)
def op_shot(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    total = _slide_count(pres)
    if total == 0:
        raise ToolError("这份演示里没有任何幻灯片。", code="slide/empty")

    first = _resolve_slide(app, pres, args.get("slide"))
    count = _clamp(int(args.get("count") or 1), 1, MAX_SHOTS_PER_CALL)
    width = _clamp(int(args.get("width") or 1280), 320, 3840)
    fmt = str(args.get("format") or "png").lower()
    thumb_width = _clamp(int(args.get("thumb_width") or 1024), 320, 1600)
    want_thumb = bool(args.get("inline", True))

    ensure_dirs(ctx.settings)
    page_w, page_h = _page_size(pres)
    height = max(1, round(width * page_h / page_w))
    ext = "jpg" if fmt in {"jpg", "jpeg"} else "png"
    filter_name = "JPG" if ext == "jpg" else "PNG"
    stem = _slug(Path(str(safe(lambda: pres.Name) or "deck")).stem)
    stamp = time.strftime("%H%M%S")

    shots: list[dict[str, Any]] = []
    for offset in range(count):
        index = first + offset
        if index > total:
            break
        slide = safe(lambda i=index: pres.Slides(i))
        if slide is None:
            continue
        target = ctx.settings.shot_dir / f"{stem}-p{index:03d}-{stamp}-{uuid.uuid4().hex[:6]}.{ext}"
        try:
            slide.Export(str(target), filter_name, width, height)
        except Exception as exc:  # noqa: BLE001 - 导出失败要给可诊断的信息
            raise ToolError(
                f"导出第 {index} 页失败：{type(exc).__name__}: {exc}。"
                "若 PowerPoint 正在放映或弹出对话框，先退出放映再试。",
                code="shot/export-failed",
                slide=index,
            ) from exc
        if not target.exists():
            raise ToolError(f"第 {index} 页导出后没有生成文件：{target}", code="shot/missing-file", slide=index)

        image = {
            "path": str(target),
            "url": f"/shots/{target.name}",
            "width": width,
            "height": height,
            "bytes": target.stat().st_size,
        }
        entry: dict[str, Any] = {"slide": index, "title": slide_title(slide), "image": image}
        if want_thumb:
            thumb = _make_thumb(target, thumb_width)
            if thumb is not None:
                entry["thumb"] = thumb
        shots.append(entry)

    if not shots:
        raise ToolError("没有导出任何页面。", code="shot/empty")

    if args.get("goto", True):
        _goto_slide(app, pres, first)

    pruned = _prune_shots(ctx.settings.shot_dir)
    return {
        "deck": presentation_info(pres),
        "total_slides": total,
        "count": len(shots),
        "pruned_files": pruned,
        "shots": shots,
        "hint": "高清图给人和任务窗格看；thumb 是给模型看的压缩版，直接看它即可。",
    }
