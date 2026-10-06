"""工具实现。

导入本包即完成注册（各模块用 ``@tool`` 装饰器）。新增操作只需在 ``ops/`` 下加模块，
并在下面的 import 列表里登记——三个面会自动带上它。
"""

from __future__ import annotations

from . import common, deck, edit, system  # noqa: F401  （导入即注册）

__all__ = ["common", "deck", "edit", "system"]
