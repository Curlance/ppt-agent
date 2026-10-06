"""ops 之间的公共辅助：演示定位、形状枚举常量、安全的 COM 取值。

这些函数全部在 STA 线程上被调用（由 ComSession 保证），所以内部可以直接摸 COM。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from ..errors import ToolError

__all__ = [
    "MSO_SHAPE_TYPE",
    "safe",
    "round_pt",
    "presentation_info",
    "list_presentations",
    "find_open_presentation",
    "bring_to_front",
    "active_slide_index",
    "resolve_deck",
    "deck_choices",
    "slide_title",
    "shape_type_name",
    "shape_names",
    "resolve_shape",
    "absolute_path",
]

#: MsoShapeType —— 让输出的 JSON 可读，而不是一堆数字。
MSO_SHAPE_TYPE = {
    1: "AutoShape",
    2: "Callout",
    3: "Chart",
    4: "Comment",
    5: "Freeform",
    6: "Group",
    7: "EmbeddedOLEObject",
    8: "FormControl",
    9: "Line",
    10: "LinkedOLEObject",
    11: "LinkedPicture",
    12: "OLEControlObject",
    13: "Picture",
    14: "Placeholder",
    15: "TextEffect",
    16: "Media",
    17: "TextBox",
    18: "ScriptAnchor",
    19: "Table",
    20: "Canvas",
    21: "Diagram",
    22: "Ink",
    23: "InkComment",
    24: "SmartArt",
}


def safe(fn: Callable[[], Any], default: Any = None) -> Any:
    """读 COM 属性时吞掉异常——遍历一页里的形状时，单个属性失败不该整体崩掉。"""
    try:
        return fn()
    except Exception:  # noqa: BLE001 - COM 属性失败原因很杂
        return default


def round_pt(value: Any, digits: int = 1) -> float | None:
    """坐标/尺寸取整到小数点后一位（单位是磅）。"""
    if isinstance(value, (int, float)):
        return round(float(value), digits)
    return None


def shape_type_name(value: Any) -> str | int | None:
    if value is None:
        return None
    return MSO_SHAPE_TYPE.get(int(value), int(value))


# --------------------------------------------------------------------------
# 演示定位
# --------------------------------------------------------------------------


def presentation_info(pres: Any) -> dict[str, Any]:
    """把 Presentation 摘成可序列化摘要。"""
    info: dict[str, Any] = {
        "name": safe(lambda: str(pres.Name)),
        "slides": safe(lambda: int(pres.Slides.Count)),
        "saved": safe(lambda: bool(pres.Saved)),
        "read_only": safe(lambda: bool(pres.ReadOnly)),
        "path": safe(lambda: str(pres.FullName)),
        "windows": safe(lambda: int(pres.Windows.Count)),
    }
    if info["path"] in {"", "None"}:
        info["path"] = None
    return info


def list_presentations(app: Any) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    count = safe(lambda: int(app.Presentations.Count), 0) or 0
    for index in range(1, count + 1):
        pres = safe(lambda i=index: app.Presentations(i))
        if pres is not None:
            info = presentation_info(pres)
            info["index"] = index
            items.append(info)
    return items


def find_open_presentation(app: Any, path: Path) -> Any | None:
    """按规范化路径查找已经打开的演示，避免重复打开同一文件。"""
    count = safe(lambda: int(app.Presentations.Count), 0) or 0
    target = str(path).casefold()
    for index in range(1, count + 1):
        pres = safe(lambda i=index: app.Presentations(i))
        if pres is None:
            continue
        full = safe(lambda p=pres: str(p.FullName))
        if full and str(full).casefold() == target:
            return pres
    return None


def bring_to_front(app: Any, pres: Any | None = None) -> None:
    """把 PowerPoint 窗口带到前台，让用户看得见。"""
    safe(lambda: app.Activate())
    if pres is not None:
        window = safe(lambda: pres.Windows(1))
        if window is not None:
            safe(lambda: window.Activate())


def active_slide_index(app: Any) -> int | None:
    """当前窗口正在看的页号（从 1 开始）。没有窗口时返回 None。"""
    return safe(lambda: int(app.ActiveWindow.View.Slide.SlideIndex))


def deck_choices(app: Any) -> list[str]:
    """给错误信息用的候选项文本。"""
    return [f"{p['index']}. {p['name']}" + (f"（{p['path']}）" if p.get("path") else "")
            for p in list_presentations(app)]


def resolve_deck(app: Any, name: str | None) -> Any:
    """按 序号 / 名称 / 完整路径 / 文件名 定位一个已打开的演示。

    留空表示"当前活动演示"。找不到时给出**候选人话列表**，而不是干巴巴的报错。
    """
    if name is None or str(name).strip() == "":
        pres = safe(lambda: app.ActivePresentation)
        if pres is None:
            raise ToolError("当前没有活动演示。先用 ppt_open 打开一个。", code="deck/none")
        return pres

    wanted = str(name).strip()
    count = safe(lambda: int(app.Presentations.Count), 0) or 0
    if count == 0:
        raise ToolError("PowerPoint 里还没有打开任何演示。", code="deck/none")

    if wanted.isdigit():
        index = int(wanted)
        if 1 <= index <= count:
            return app.Presentations(index)
        raise ToolError(
            f"演示序号 {index} 超出范围（当前 1..{count}）。",
            code="deck/not-found",
            open_decks=deck_choices(app),
        )

    lowered = wanted.casefold()
    # 完整路径只匹配完整路径，不能误改同名的另一份演示。
    candidates = (lowered,) if Path(wanted).is_absolute() else (lowered, Path(wanted).name.casefold())
    for candidate in candidates:
        matches = []
        for index in range(1, count + 1):
            pres = safe(lambda i=index: app.Presentations(i))
            if pres is None:
                continue
            pres_name = str(safe(lambda p=pres: p.Name) or "").casefold()
            full = str(safe(lambda p=pres: p.FullName) or "").casefold()
            choices = {full} if Path(wanted).is_absolute() else {pres_name, full, Path(full).name if full else ""}
            if candidate in choices:
                matches.append(pres)
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raise ToolError(f"「{wanted}」匹配到多个演示，请用完整路径或序号。", code="deck/ambiguous", open_decks=deck_choices(app))

    raise ToolError(
        f"没有找到名为「{wanted}」的已打开演示。",
        code="deck/not-found",
        open_decks=deck_choices(app),
    )


def slide_title(slide: Any) -> str:
    """粗略提取一页的标题：取最靠上的那段非空文字。

    刻意不用 Slide.Shapes.Title —— 实测不少演示没有标题占位符。
    """
    shapes = safe(lambda: slide.Shapes)
    if shapes is None:
        return ""
    best: tuple[float, str] | None = None
    count = safe(lambda: int(shapes.Count), 0) or 0
    for index in range(1, count + 1):
        shape = safe(lambda i=index: shapes(i))
        if shape is None:
            continue
        text = safe(lambda s=shape: str(s.TextFrame.TextRange.Text))
        if not text or not text.strip():
            continue
        top = safe(lambda s=shape: float(s.Top), 1e9)
        if best is None or top < best[0]:
            best = (top, text.strip())
    return best[1] if best else ""


# --------------------------------------------------------------------------
# 形状定位（写操作用）
# --------------------------------------------------------------------------


def shape_names(slide: Any) -> list[str]:
    """这一页所有形状的名字，供错误信息里列出候选项。"""
    shapes = safe(lambda: slide.Shapes)
    count = safe(lambda: int(shapes.Count), 0) or 0
    names: list[str] = []
    for index in range(1, count + 1):
        shape = safe(lambda i=index: shapes(i))
        if shape is not None:
            names.append(str(safe(lambda s=shape: s.Name) or f"#{index}"))
    return names


def resolve_shape(slide: Any, ref: Any) -> tuple[Any, int, str]:
    """把 ``shape`` 参数解析成 (形状对象, 序号, 名字)。

    支持三种写法，模型从 ``ppt_read_slide`` 的输出里直接抄即可：

    - **序号**：从 1 开始的形状序号；
    - **名称**：精确匹配（忽略大小写），失败则退回唯一子串匹配；
    - **``title`` / ``标题``**：标题占位符。

    找不到或有歧义时，错误信息里**列出这一页所有形状名**——让模型能自己改对。
    """
    shapes = safe(lambda: slide.Shapes)
    count = safe(lambda: int(shapes.Count), 0) or 0
    if count == 0:
        raise ToolError("这一页没有任何形状。", code="shape/empty")

    if ref is None or (isinstance(ref, str) and ref.strip() == ""):
        raise ToolError(
            "需要指定形状：可以用形状名或序号。",
            code="shape/required",
            available=shape_names(slide),
        )

    if isinstance(ref, int) or (isinstance(ref, str) and ref.strip().isdigit()):
        index = int(ref)
        if not 1 <= index <= count:
            raise ToolError(
                f"形状序号 {index} 超出范围（这一页共 {count} 个形状）。",
                code="shape/out-of-range",
                total=count,
                available=shape_names(slide),
            )
        shape = shapes(index)
        return shape, index, str(safe(lambda: shape.Name) or f"#{index}")

    # ppt_read_slide 返回组合内形状的稳定序号路径，如 3/2/1。
    if isinstance(ref, str) and "/" in ref and all(part.isdigit() for part in ref.split("/")):
        parts = [int(part) for part in ref.split("/")]
        collection = shapes
        shape = None
        for depth, index in enumerate(parts):
            if not 1 <= index <= int(collection.Count):
                raise ToolError(f"组合形状路径 {ref} 超出范围。", code="shape/out-of-range")
            shape = collection(index)
            if depth < len(parts) - 1:
                if int(shape.Type) != 6:
                    raise ToolError(f"形状路径 {ref} 的中间节点不是组合。", code="shape/not-group")
                collection = shape.GroupItems
        return shape, parts[0], str(shape.Name)

    wanted = str(ref).strip()
    if wanted.lower() in {"title", "标题"}:
        title = safe(lambda: shapes.Title)
        if title is None:
            raise ToolError(
                "这一页没有标题占位符。",
                code="shape/no-title",
                available=shape_names(slide),
            )
        for index in range(1, count + 1):
            candidate = safe(lambda i=index: shapes(i))
            if candidate is title or (safe(lambda: title.Id) is not None and safe(lambda: candidate.Id) == title.Id):
                return title, index, str(safe(lambda: title.Name) or "title")
        return title, 0, str(safe(lambda: title.Name) or "title")

    lowered = wanted.casefold()
    matches: list[tuple[Any, int, str]] = []
    for index in range(1, count + 1):
        shape = safe(lambda i=index: shapes(i))
        if shape is None:
            continue
        name = str(safe(lambda s=shape: s.Name) or "")
        if name.casefold() == lowered:
            return shape, index, name
        if lowered in name.casefold():
            matches.append((shape, index, name))

    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ToolError(
            f"「{wanted}」匹配到多个形状：{'、'.join(m[2] for m in matches)}。请写完整名称或用序号。",
            code="shape/ambiguous",
            matches=[m[2] for m in matches],
        )
    raise ToolError(
        f"这一页没有名为「{wanted}」的形状。",
        code="shape/not-found",
        available=shape_names(slide),
    )


def absolute_path(raw: Any, *, what: str, must_exist: bool = True, suffixes: set[str] | None = None) -> Path:
    """把外部文件路径规范化并校验。

    **必须绝对路径**：PowerPoint 是独立进程，相对路径会按它自己的工作目录解析——
    实测这会让 ``AddPicture`` 报出与真正原因无关的错（"'str' object has no attribute 'Name'"）。
    """
    if raw is None or str(raw).strip() == "":
        raise ToolError(f"需要提供{what}的绝对路径。", code="path/required")
    path = Path(str(raw)).expanduser()
    if not path.is_absolute():
        raise ToolError(f"{what}需要一个绝对路径，收到：{path}", code="path/relative")
    if must_exist and not path.exists():
        raise ToolError(f"{what}不存在：{path}", code="path/missing")
    if suffixes and path.suffix.lower() not in suffixes:
        raise ToolError(
            f"{what}的类型不支持：{path.suffix}（支持 {'、'.join(sorted(suffixes))}）",
            code="path/unsupported",
        )
    return path
