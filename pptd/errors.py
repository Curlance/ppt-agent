"""统一错误类型：让 CLI / HTTP / MCP 三个面给出同样的错误码。"""

from __future__ import annotations

from typing import Any

__all__ = ["PptAgentError", "ToolError", "ComError", "ComTimeout", "ComUnavailable", "to_error"]


class PptAgentError(RuntimeError):
    """所有可预期错误的基类。``code`` 进入三个面的结构化响应。"""

    code = "internal"

    def __init__(self, message: str, *, code: str | None = None, **extra: Any) -> None:
        super().__init__(message)
        self.message = message
        if code:
            self.code = code
        self.extra = extra

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"code": self.code, "message": self.message}
        if self.extra:
            payload["details"] = self.extra
        return payload


class ToolError(PptAgentError):
    """工具层的参数或前置条件错误。"""

    code = "tool/invalid"


class ComError(PptAgentError):
    """PowerPoint COM 调用失败。"""

    code = "com/error"


class ComTimeout(ComError):
    """COM 调用超时——绝大多数情况是 PowerPoint 弹了模态框。"""

    code = "com/timeout"


class ComUnavailable(ComError):
    """拿不到 PowerPoint Application（未安装、被策略禁用、启动失败）。"""

    code = "com/unavailable"


def to_error(exc: BaseException) -> dict[str, Any]:
    """把任意异常规整成 ``{code, message}``。"""
    if isinstance(exc, PptAgentError):
        return exc.to_dict()
    return {"code": "internal", "message": f"{type(exc).__name__}: {exc}"}
