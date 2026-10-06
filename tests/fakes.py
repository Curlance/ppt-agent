"""共享的假 COM 模型：让绝大多数逻辑脱离真机也能测。

刻意做得比"够用"再真一点：

- ``Slide.Export`` **真的写一张 PNG**（尺寸与请求一致），所以缩略图逻辑能被真实验证；
- ``CommandBars.ExecuteMso("Undo")`` **真的维护撤销栈**，撤到底会**抛异常**——
  这是真机上实测到的行为，不模拟它就会漏掉 ``ppt_undo`` 的边界。
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any, Callable

MSO_FALSE = 0
MSO_TRUE = -1


class FakeComError(Exception):
    """对应真机上的 pythoncom.com_error。"""


# --------------------------------------------------------------------------
# 文本与字体
# --------------------------------------------------------------------------


class FakeFont:
    def __init__(self) -> None:
        self.Name = "微软雅黑"
        self.Size = 18.0
        self.Bold = MSO_FALSE
        self.Color = SimpleNamespace(RGB=0)


class FakeTextRange:
    def __init__(self, text: str) -> None:
        self.Text = text
        self.Font = FakeFont()
        self.ParagraphFormat = SimpleNamespace(Alignment=1)


class FakeTextFrame:
    def __init__(self, text: str) -> None:
        self.TextRange = FakeTextRange(text)

    @property
    def HasText(self) -> bool:
        return bool(self.TextRange.Text)


class FakeTable:
    def __init__(self, rows: int, cols: int) -> None:
        self.Rows = SimpleNamespace(Count=rows)
        self.Columns = SimpleNamespace(Count=cols)
        self._cells = {(r, c): SimpleNamespace(Shape=SimpleNamespace(TextFrame=FakeTextFrame(f"表头{c}" if r == 1 else f"R{r}C{c}")))
            for r in range(1, rows + 1) for c in range(1, cols + 1)}

    def Cell(self, r: int, c: int) -> Any:
        return self._cells[(r, c)]


class FakeChart:
    ChartType = 51
    HasTitle = True
    ChartTitle = SimpleNamespace(Text="销量")


class FakeLine:
    def __init__(self) -> None:
        self.Visible = MSO_FALSE
        self.Weight = 1.0
        self.ForeColor = SimpleNamespace(RGB=0)


# --------------------------------------------------------------------------
# 形状
# --------------------------------------------------------------------------


class FakeShape:
    def __init__(
        self,
        name: str,
        type_: int = 1,
        text: str | None = None,
        table: FakeTable | None = None,
        chart: FakeChart | None = None,
        group_items: int | None = None,
        geometry: tuple[float, float, float, float] = (10.0, 20.0, 100.0, 50.0),
        placeholder_type: int | None = None,
    ) -> None:
        self.Name = name
        self.Type = type_
        self.Left, self.Top, self.Width, self.Height = geometry
        self.Rotation = 0.0
        self.HasTextFrame = type_ not in {3, 13, 11}
        self.TextFrame = FakeTextFrame(text or "")
        self.Line = FakeLine()
        self._has_table = table is not None
        self.Table = table
        self._has_chart = chart is not None
        self.Chart = chart
        self.GroupItems = SimpleNamespace(Count=group_items) if group_items is not None else None
        self.PlaceholderFormat = SimpleNamespace(Type=placeholder_type if placeholder_type is not None else 0)
        self._owner: "FakeShapes | None" = None
        self.selected = False

    # 集合里的位置（1 起）
    @property
    def ZOrderPosition(self) -> int:
        if self._owner is None:
            return 1
        return self._owner._items.index(self) + 1

    @property
    def HasTable(self) -> bool:
        return self._has_table

    @property
    def HasChart(self) -> bool:
        return self._has_chart

    # -- 行为 ------------------------------------------------------------

    def Select(self) -> None:
        if self._owner is not None:
            for item in self._owner._items:
                item.selected = False
        self.selected = True

    def Delete(self) -> None:
        assert self._owner is not None
        owner, index = self._owner, self._owner._items.index(self)
        owner._items.pop(index)
        owner._push_undo(lambda: owner._items.insert(index, self))

    def Duplicate(self) -> Any:
        assert self._owner is not None
        clone = FakeShape(self.Name + " 副本", self.Type, self.TextFrame.TextRange.Text, geometry=(self.Left + 10, self.Top + 10, self.Width, self.Height))
        self._owner._add(clone)
        return clone


class FakeShapes:
    def __init__(self, slide: "FakeSlide", items: list[FakeShape] | None = None) -> None:
        self._slide = slide
        self._items: list[FakeShape] = list(items or [])
        for item in self._items:
            item._owner = self
        self._counter = len(self._items)

    # -- 集合接口 --------------------------------------------------------

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> FakeShape:
        if not 1 <= index <= len(self._items):
            raise IndexError(index)
        return self._items[index - 1]

    @property
    def Title(self) -> FakeShape | None:
        for item in self._items:
            if item.PlaceholderFormat.Type == 1:
                return item
        return None

    @property
    def Placeholders(self) -> Any:
        holders = [s for s in self._items if s.PlaceholderFormat.Type]
        return _Subset(holders)

    # -- 添加 ------------------------------------------------------------

    def _push_undo(self, fn: Callable[[], None]) -> None:
        self._slide._push_undo(fn)

    def _add(self, shape: FakeShape) -> FakeShape:
        self._items.append(shape)
        shape._owner = self
        return shape

    def _new_name(self, prefix: str) -> str:
        self._counter += 1
        return f"{prefix} {self._counter}"

    def AddTextbox(self, orientation: int, left: float, top: float, width: float, height: float) -> FakeShape:
        shape = FakeShape(self._new_name("TextBox"), 17, text="", geometry=(left, top, width, height))
        self._add(shape)
        self._push_undo(lambda: self._items.remove(shape))
        return shape

    def AddPicture(self, path: str, link: int, save_with: int, left: float, top: float, width: float | None = None, height: float | None = None) -> FakeShape:
        if not Path(path).is_file():
            raise FakeComError(f"找不到图片：{path}")
        w = width if width is not None else 200.0
        h = height if height is not None else 120.0
        shape = FakeShape(self._new_name("Picture"), 13, geometry=(left, top, w, h))
        self._add(shape)
        self._push_undo(lambda: self._items.remove(shape))
        return shape

    def AddTable(self, rows: int, cols: int, left: float, top: float, width: float, height: float) -> FakeShape:
        shape = FakeShape(self._new_name("Table"), 19, table=FakeTable(rows, cols), geometry=(left, top, width, height))
        self._add(shape)
        self._push_undo(lambda: self._items.remove(shape))
        return shape


class _Subset:
    """形状子集：只支持按序号取，够用来模拟 Placeholders(2)。"""

    def __init__(self, items: list[FakeShape]) -> None:
        self._items = items
        self.Count = len(items)

    def __call__(self, index: int) -> FakeShape:
        if not 1 <= index <= len(self._items):
            raise IndexError(index)
        return self._items[index - 1]


# --------------------------------------------------------------------------
# 幻灯片
# --------------------------------------------------------------------------


class FakeSlide:
    def __init__(self, index: int, shapes: list[FakeShape], notes: str = "", layout: str = "标题和内容") -> None:
        self.index = index
        #: 由 FakeSlides 回填，删除/移动/撤销都要靠它找到宿主
        self._owner_slides: "FakeSlides | None" = None
        self.Shapes = FakeShapes(self, shapes)
        # 实测 Slide.Layout.Name 返回 None，只有 CustomLayout.Name 可用
        self.Layout = SimpleNamespace(Name=None)
        self.CustomLayout = SimpleNamespace(Name=layout)
        self.NotesPage = SimpleNamespace(Shapes=FakeShapes(self, [FakeShape("备注", 14, text=notes, placeholder_type=2)]))

    def _push_undo(self, fn: Callable[[], None]) -> None:
        if self._owner_slides is not None:
            self._owner_slides._push_undo(fn)

    def Export(self, path: str, filter_name: str, width: int, height: int) -> None:
        from PIL import Image

        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        # 用**噪声**而不是纯色：纯色 PNG 只有一两 KB，比 JPEG 缩略图还小，
        # 那样"缩略图比高清图小"这条断言就失去意义了。噪声更接近真实照片页。
        image = Image.effect_noise((width, height), 96).convert("RGB")
        image.save(target, "JPEG" if filter_name == "JPG" else "PNG")

    def Delete(self) -> None:
        assert self._owner_slides is not None
        slides = self._owner_slides
        index = slides._items.index(self)
        slides._items.pop(index)
        slides._push_undo(lambda: slides._items.insert(index, self))

    def MoveTo(self, target: int) -> None:
        assert self._owner_slides is not None
        slides = self._owner_slides
        index = slides._items.index(self)
        slides._items.pop(index)
        slides._items.insert(target - 1, self)


class FakeSlides:
    def __init__(self, pres: "FakePresentation", items: list[FakeSlide]) -> None:
        self._pres = pres
        self._items = items
        for slide in items:
            slide._owner_slides = self

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> FakeSlide:
        if not 1 <= index <= len(self._items):
            raise IndexError(index)
        return self._items[index - 1]

    def _push_undo(self, fn: Callable[[], None]) -> None:
        self._pres._app._push_undo(fn)

    def AddSlide(self, index: int, layout: Any) -> FakeSlide:
        slide = FakeSlide(index, [], layout=getattr(layout, "Name", "空白"))
        slide._owner_slides = self
        self._items.insert(index - 1, slide)
        self._push_undo(lambda: self._items.remove(slide))
        return slide


# --------------------------------------------------------------------------
# 演示与应用程序
# --------------------------------------------------------------------------


class _Windows:
    def __init__(self, window: "FakeWindow") -> None:
        self.Count = 1
        self._window = window

    def __call__(self, index: int) -> "FakeWindow":
        if index != 1:
            raise IndexError(index)
        return self._window


class FakeView:
    def __init__(self) -> None:
        self.Slide = SimpleNamespace(SlideIndex=1)

    def GotoSlide(self, index: int) -> None:
        self.Slide.SlideIndex = index


class FakeWindow:
    def __init__(self) -> None:
        self.ViewType = 9
        self.View = FakeView()

    def Activate(self) -> None:
        pass


class FakeCustomLayouts:
    def __init__(self, names: list[str]) -> None:
        self._names = names

    @property
    def Count(self) -> int:
        return len(self._names)

    def __call__(self, index: int) -> Any:
        if not 1 <= index <= len(self._names):
            raise IndexError(index)
        return SimpleNamespace(Name=self._names[index - 1])


class FakePresentation:
    def __init__(self, name: str, slides: list[FakeSlide], directory: Path, *, layouts: list[str] | None = None) -> None:
        self.Name = name
        self.FullName = str(directory / name)
        self.Saved = True
        self.ReadOnly = False
        self._window = FakeWindow()
        self.Windows = _Windows(self._window)
        self.PageSetup = SimpleNamespace(SlideWidth=960.0, SlideHeight=540.0)
        self.SlideMaster = SimpleNamespace(CustomLayouts=FakeCustomLayouts(layouts or ["标题和内容", "空白"]))
        self._app: FakePowerPointApp | None = None
        self._slides = FakeSlides(self, slides)
        self.Slides = self._slides
        self.closed = False
        self.saved_as: str | None = None
        self.copied_to: str | None = None

    def Save(self) -> None:
        self.Saved = True

    def SaveAs(self, path: str) -> None:
        self.saved_as = path
        self.FullName = path
        self.Saved = True

    def SaveCopyAs(self, path: str) -> None:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        Path(path).write_bytes(b"fake-pptx")
        self.copied_to = path

    def Close(self) -> None:
        self.closed = True


class FakePresentations:
    def __init__(self, items: list[FakePresentation]) -> None:
        self._items = items

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> FakePresentation:
        if not 1 <= index <= len(self._items):
            raise IndexError(index)
        return self._items[index - 1]

    def Open(self, path: str, read_only: bool, untitled: bool, with_window: bool) -> FakePresentation:
        target = Path(path)
        pres = FakePresentation(target.name, _default_slides(3), target.parent)
        pres.FullName = str(target)
        self._items.append(pres)
        return pres


class FakeCommandBars:
    """真的维护一个撤销栈；撤到底时抛异常（与真机行为一致）。"""

    def __init__(self, app: "FakePowerPointApp") -> None:
        self._app = app
        self.executed: list[str] = []

    def ExecuteMso(self, name: str) -> None:
        self.executed.append(name)
        if name != "Undo":
            return
        if not self._app._undo:
            raise FakeComError("没有可撤销的操作")
        self._app._undo.pop()()


class FakePowerPointApp:
    """够用的假 Application。"""

    def __init__(self, directory: Path, decks: int = 1, slides: int = 5) -> None:
        self.Version = "16.0"
        self.Visible = False
        self.DisplayAlerts = 2
        self._undo: list[Callable[[], None]] = []
        self.CommandBars = FakeCommandBars(self)
        self._decks = [
            FakePresentation(f"deck{i}.pptx", _default_slides(slides), directory)
            for i in range(1, decks + 1)
        ]
        for deck in self._decks:
            deck._app = self
        self.Presentations = FakePresentations(self._decks)
        self.ActivePresentation = self._decks[0] if self._decks else None
        self.ActiveWindow = SimpleNamespace(
            View=self._decks[0]._window.View if self._decks else FakeView(),
            Activate=lambda: None,
            Selection=SimpleNamespace(Type=0),
        )

    def Activate(self) -> None:
        pass

    def _push_undo(self, fn: Callable[[], None]) -> None:
        self._undo.append(fn)


def _default_slides(count: int) -> list[FakeSlide]:
    """一份内容可预期的演示：标题、正文，第 2 页带表格，第 3 页带图表，第 4 页带组合。"""
    slides: list[FakeSlide] = []
    for i in range(1, count + 1):
        shapes = [
            FakeShape(f"标题 {i}", 14, text=f"第 {i} 页标题", geometry=(60.0, 40.0, 800.0, 60.0), placeholder_type=1),
            FakeShape(f"正文 {i}", 17, text=f"第 {i} 页正文", geometry=(60.0, 140.0, 800.0, 300.0)),
        ]
        if i == 2:
            shapes.append(FakeShape("表格 1", 19, table=FakeTable(3, 2)))
        if i == 3:
            shapes.append(FakeShape("图表 1", 3, chart=FakeChart()))
        if i == 4:
            shapes.append(FakeShape("组合 1", 6, group_items=4))
        slides.append(FakeSlide(i, shapes, notes=f"备注 {i}"))
    return slides


def make_app(directory: Path, decks: int = 1, slides: int = 5) -> FakePowerPointApp:
    """建一个带 deck 的假 PowerPoint。"""
    return FakePowerPointApp(directory, decks=decks, slides=slides)


def make_png(path: Path, size: tuple[int, int] = (64, 36)) -> Path:
    from PIL import Image

    path.parent.mkdir(parents=True, exist_ok=True)
    Image.new("RGB", size, (200, 90, 40)).save(path)
    return path
