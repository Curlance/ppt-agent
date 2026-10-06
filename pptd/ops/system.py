"""M0 工具集：状态、打开、切前台、关闭。

M1 的读取与截图工具在 ``deck.py``。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..errors import ToolError
from ..registry import Context, Param, tool
from .common import (
    active_slide_index,
    bring_to_front,
    find_open_presentation,
    list_presentations,
    presentation_info,
    resolve_deck,
    safe,
    slide_title,
)

__all__ = ["op_status", "op_open", "op_new", "op_activate", "op_close"]


@tool(
    "ppt_status",
    title="查看 PowerPoint 状态",
    summary="查看 PowerPoint 是否在运行、有哪些演示打开着、当前停在第几页",
    behavior="read",
    needs_com=False,  # 即使 COM 起不来，也要能如实报出故障
    returns="会话状态对象：解释器、COM 会话、已打开的演示列表",
    tags=("read", "diagnose"),
    params=(
        Param(
            name="list_slides",
            type="boolean",
            description="是否列出每个演示的每一页标题（大文件会慢）",
            default=False,
        ),
    ),
)
def op_status(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    from .. import __version__
    from ..config import interpreter_report

    session_info = ctx.session.info().to_dict()
    payload: dict[str, Any] = {
        "version": __version__,
        "interpreter": interpreter_report(),
        "com": session_info,
        "presentations": [],
        "active_slide": None,
    }

    if not ctx.session.info().alive:
        payload["com_error"] = session_info.get("last_error") or "COM 会话未启动"
        return payload

    try:
        payload["presentations"] = ctx.session.run(list_presentations, timeout=20)
        payload["active_slide"] = ctx.session.run(active_slide_index, timeout=15)
        if args.get("list_slides"):
            payload["slide_titles"] = ctx.session.run(_slide_titles, timeout=60)
    except Exception as exc:  # noqa: BLE001 - 状态查询永远不该抛
        payload["com_error"] = f"{type(exc).__name__}: {exc}"
    return payload


def _slide_titles(app: Any) -> dict[str, list[str]]:
    """每个已打开演示的逐页首行文字，用于快速了解内容。"""
    out: dict[str, list[str]] = {}
    for item in list_presentations(app):
        index = int(item["index"])
        pres = safe(lambda i=index: app.Presentations(i))
        if pres is None:
            continue
        name = item.get("name") or f"演示{index}"
        titles: list[str] = []
        slides = safe(lambda p=pres: p.Slides)
        total = safe(lambda s=slides: int(s.Count), 0) or 0
        for slide_no in range(1, total + 1):
            slide = safe(lambda s=slides, n=slide_no: s(n))
            titles.append(slide_title(slide) if slide is not None else "")
        out[str(name)] = titles
    return out


@tool(
    "ppt_open",
    title="打开演示",
    summary="打开一个演示文稿并让它在 PowerPoint 里显示出来",
    behavior="write",
    returns="被打开演示的摘要（页数、路径、是否只读）",
    tags=("deck",),
    params=(
        Param(name="path", type="string", description="演示文件路径（.pptx / .ppt）", required=True),
        Param(
            name="read_only",
            type="boolean",
            description="以只读方式打开，避免误改原文件",
            default=False,
        ),
        Param(
            name="bring_to_front",
            type="boolean",
            description="打开后把 PowerPoint 窗口切到最前，让用户看得见",
            default=True,
        ),
    ),
)
def op_open(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    path = Path(str(args["path"])).expanduser()
    if not path.is_absolute():
        raise ToolError(f"需要一个绝对路径，收到：{path}", code="path/relative")
    if not path.exists():
        raise ToolError(f"文件不存在：{path}", code="path/missing")
    if path.suffix.lower() not in {".pptx", ".ppt", ".pptm", ".ppsx", ".ppsm"}:
        raise ToolError(f"不是 PowerPoint 能打开的类型：{path.suffix}", code="path/unsupported")

    existing = find_open_presentation(app, path)
    if existing is not None:
        if args.get("bring_to_front", True):
            bring_to_front(app, existing)
        return {
            "opened": False,
            "already_open": True,
            "reason": "这份演示已经在 PowerPoint 里打开了，直接切到它",
            "presentation": presentation_info(existing),
        }

    pres = app.Presentations.Open(
        str(path),
        bool(args.get("read_only", False)),  # ReadOnly
        False,                               # Untitled
        True,                                # WithWindow —— 可见性铁律，绝不隐身
    )
    if args.get("bring_to_front", True):
        bring_to_front(app, pres)
    return {
        "opened": True,
        "already_open": False,
        "presentation": presentation_info(pres),
    }


@tool(
    "ppt_activate",
    title="切到 PowerPoint 前台",
    summary="把 PowerPoint 窗口带到最前面",
    behavior="idempotent",
    returns="是否成功",
    tags=("visibility",),
    params=(),
)
def op_activate(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    bring_to_front(app)
    return {"activated": True, "visible": safe(lambda: bool(app.Visible))}


@tool(
    "ppt_new",
    title="新建演示",
    summary="新建一份空白演示并显示出来",
    behavior="write",
    returns="新演示的摘要（页数、路径可能为「未保存」）",
    tags=("deck",),
    params=(
        Param(name="template", type="string", description="可选：以某个 .potx / .pptx 模板为底，需要绝对路径"),
        Param(
            name="bring_to_front",
            type="boolean",
            description="新建后把 PowerPoint 切到最前，让用户看得见",
            default=True,
        ),
    ),
)
def op_new(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    template = args.get("template")
    path = None
    if template not in (None, ""):
        from .common import absolute_path
        path = absolute_path(template, what="模板", suffixes={".potx", ".pptx", ".pot", ".ppt", ".thmx"})
    pres = app.Presentations.Add(True)  # WithWindow=True —— 可见性铁律
    if pres is None:
        raise ToolError("新建演示失败。", code="deck/add-failed")

    if path is not None:
        try:
            pres.ApplyTemplate(str(path))
        except Exception:
            pres.Saved = True
            pres.Close()
            raise

    if args.get("bring_to_front", True):
        bring_to_front(app, pres)
    return {
        "created": True,
        "presentation": presentation_info(pres),
        "templated": template not in (None, ""),
        "hint": "新演示是未保存状态；用 ppt_save 落到磁盘，或用 ppt_close 丢弃。",
    }


@tool(
    "ppt_close",
    title="关闭演示",
    summary="关闭一个已打开的演示（默认不保存，避免误改用户文件）",
    behavior="write",
    returns="被关闭演示的名称",
    tags=("deck",),
    params=(
        Param(name="deck", type="string", description="演示名、路径或序号；留空表示当前活动演示"),
        Param(
            name="save",
            type="boolean",
            description="是否先保存再关闭。默认否——未保存的改动会被丢弃",
            default=False,
        ),
    ),
)
def op_close(app: Any, args: dict[str, Any], ctx: Context) -> dict[str, Any]:
    pres = resolve_deck(app, args.get("deck"))
    info = presentation_info(pres)
    if args.get("save", False):
        if safe(lambda: str(pres.Path)) == "":
            raise ToolError("新演示尚未保存，请先用 ppt_save(path=...) 保存。", code="save/path-required")
        pres.Save()
    else:
        # 明确丢弃这份演示的改动，不依靠关闭整个应用的 DisplayAlerts。
        pres.Saved = True
    pres.Close()
    return {"closed": True, "presentation": info, "saved": bool(args.get("save", False))}
