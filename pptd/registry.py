"""唯一事实源：所有工具在这里定义一次，三个面从它生成。

为什么要有注册表
----------------
MCP、HTTP+OpenAPI、CLI 必须给出**完全一致**的工具名、参数 schema 与行为语义，
否则模型在不同面看到的能力不一样，agent 会犯迷糊。因此：

    Tool（name/title/summary/behavior/params/handler）
        ├─ MCP  面：按 schema 合成函数签名 → inputSchema 由 pydantic 生成
        ├─ HTTP 面：按 schema 校验入参，并出现在 /openapi.json
        └─ CLI  面：按 schema 生成 --help 与参数解析

``summary`` 是**中文人话**，会作为 ``op.start`` 事件推给观察台——这正是
"用户能看见 agent 在做什么"的文字层。
"""

from __future__ import annotations

import inspect
import math
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable, Literal, Mapping, Sequence

__all__ = [
    "Param",
    "Tool",
    "Context",
    "registry",
    "tool",
    "get_tool",
    "all_tools",
    "tool_manifest",
]

Behavior = Literal["read", "write", "destroy", "idempotent"]

_JSON_TO_PY: dict[str, type] = {
    "string": str,
    "integer": int,
    "number": float,
    "boolean": bool,
    "object": dict,
    "array": list,
}


@dataclass(frozen=True)
class Param:
    """一个工具参数。``type`` 用 JSON Schema 的类型名。"""

    name: str
    type: str
    description: str
    required: bool = False
    default: Any = None
    enum: tuple[str, ...] | None = None
    items: Mapping[str, Any] | None = None
    minimum: float | None = None
    maximum: float | None = None
    allow_empty: bool = False

    @property
    def py_type(self) -> type:
        return _JSON_TO_PY.get(self.type, str)

    def to_schema(self) -> dict[str, Any]:
        """转成 JSON Schema 片段（**参数顺序即阅读顺序**）。"""
        node: dict[str, Any] = {"type": self.type, "description": self.description}
        if self.type == "string" and self.required and not self.allow_empty:
            node["minLength"] = 1
        if self.enum:
            node["enum"] = list(self.enum)
        if self.items is not None:
            node["items"] = dict(self.items)
        if self.minimum is not None:
            node["minimum"] = self.minimum
        if self.maximum is not None:
            node["maximum"] = self.maximum
        if not self.required and self.default is not None:
            node["default"] = self.default
        return node

    def to_annotation(self) -> Any:
        """给 MCP 面合成签名用：``Annotated[py_type, Field(description=...)]``。"""
        import pydantic
        from typing import Annotated

        kwargs: dict[str, Any] = {"description": self.description}
        if self.minimum is not None:
            kwargs["ge"] = self.minimum
        if self.maximum is not None:
            kwargs["le"] = self.maximum
        if self.type == "string" and self.required and not self.allow_empty:
            kwargs["min_length"] = 1
        py = Literal[self.enum] if self.enum else self.py_type
        if not self.required and self.default is None:
            py = py | None
        return Annotated[py, pydantic.Field(**kwargs)]  # type: ignore[valid-type]

    def to_parameter(self) -> inspect.Parameter:
        import inspect as _inspect

        return _inspect.Parameter(
            self.name,
            _inspect.Parameter.KEYWORD_ONLY,
            default=_inspect.Parameter.empty if self.required else self.default,
            annotation=self.to_annotation(),
        )


@dataclass(frozen=True)
class Tool:
    """一个可被 agent 调用的操作。"""

    name: str
    title: str
    summary: str
    behavior: Behavior
    params: tuple[Param, ...]
    handler: Callable[..., Any]
    returns: str = "工具结果对象"
    tags: tuple[str, ...] = ()
    #: 是否必须持有 PowerPoint 实例。``ppt_status`` 之类要能在 COM 挂掉时如实汇报，故为 False。
    needs_com: bool = True

    # -- schema ----------------------------------------------------------

    def input_schema(self) -> dict[str, Any]:
        properties = {p.name: p.to_schema() for p in self.params}
        required = [p.name for p in self.params if p.required]
        schema: dict[str, Any] = {
            "type": "object",
            "properties": properties,
            "additionalProperties": False,
        }
        if required:
            schema["required"] = required
        return schema

    def required(self) -> list[str]:
        return [p.name for p in self.params if p.required]

    def param(self, name: str) -> Param | None:
        return next((p for p in self.params if p.name == name), None)

    # -- 语义 ------------------------------------------------------------

    @property
    def read_only(self) -> bool:
        return self.behavior == "read"

    @property
    def destructive(self) -> bool:
        return self.behavior == "destroy"

    def to_manifest(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "title": self.title,
            "summary": self.summary,
            "behavior": self.behavior,
            "returns": self.returns,
            "tags": list(self.tags),
            "needs_com": self.needs_com,
            "parameters": [
                {
                    "name": p.name,
                    "type": p.type,
                    "description": p.description,
                    "required": p.required,
                    "default": p.default,
                    "enum": list(p.enum) if p.enum else None,
                }
                for p in self.params
            ],
            "inputSchema": self.input_schema(),
        }


@dataclass
class Context:
    """工具运行时依赖。刻意保持窄：ops 只该看见这些东西。"""

    session: Any                       # ComSession
    bus: Any                           # EventBus
    settings: Any                      # config.Settings
    state: dict[str, Any] = field(default_factory=dict)

    def emit(self, type_: str, **data: Any) -> None:
        self.bus.publish(type_, **data)

    @property
    def timeout(self) -> float:
        return float(self.settings.call_timeout)


class _Registry:
    """工具集合。按注册顺序返回，给模型稳定、可预期的列表顺序。"""

    def __init__(self) -> None:
        self._tools: dict[str, Tool] = {}

    def add(self, t: Tool) -> Tool:
        if t.name in self._tools:
            raise ValueError(f"工具重名：{t.name}")
        self._tools[t.name] = t
        return t

    def get(self, name: str) -> Tool:
        try:
            return self._tools[name]
        except KeyError as exc:
            from .errors import ToolError

            raise ToolError(f"未知工具：{name}", code="tool/unknown", tool=name) from exc

    def all(self) -> list[Tool]:
        return list(self._tools.values())

    def names(self) -> list[str]:
        return list(self._tools)


registry = _Registry()


def tool(
    name: str,
    *,
    title: str,
    summary: str,
    behavior: Behavior,
    params: Sequence[Param] = (),
    returns: str = "工具结果对象",
    tags: Sequence[str] = (),
    needs_com: bool = True,
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """把 ``def op(app, args, ctx)`` 注册成工具。"""

    def decorate(fn: Callable[..., Any]) -> Callable[..., Any]:
        registry.add(
            Tool(
                name=name,
                title=title,
                summary=summary,
                behavior=behavior,
                params=tuple(params),
                handler=fn,
                returns=returns,
                tags=tuple(tags),
                needs_com=needs_com,
            )
        )
        return fn

    return decorate


def get_tool(name: str) -> Tool:
    return registry.get(name)


def all_tools() -> list[Tool]:
    return registry.all()


def tool_manifest() -> list[dict[str, Any]]:
    """给 HTTP ``/tools`` 与 MCP ``tools/list`` 共用。"""
    return [t.to_manifest() for t in registry.all()]


def validate_args(t: Tool, args: Mapping[str, Any] | None) -> dict[str, Any]:
    """按 schema 做最小必要校验并填默认值。

    刻意不引入 pydantic 校验器：参数面很窄，手写校验更容易给出**中文**错误。
    """
    from .errors import ToolError

    if args is not None and not isinstance(args, Mapping):
        raise ToolError("工具参数必须是对象。", code="args/type")
    raw = dict(args or {})
    unknown = sorted(set(raw) - {p.name for p in t.params})
    if unknown:
        raise ToolError(
            f"{t.name} 不接受参数：{', '.join(unknown)}",
            code="args/unknown",
            allowed=[p.name for p in t.params],
        )
    out: dict[str, Any] = {}
    for p in t.params:
        value = raw.get(p.name, None)
        if value is None or (value == "" and not p.allow_empty):
            if p.required:
                raise ToolError(f"{t.name} 缺少必填参数：{p.name}", code="args/missing", parameter=p.name)
            out[p.name] = p.default
            continue
        coerced = _coerce(p, value)
        out[p.name] = coerced
    return out


def _coerce(p: Param, value: Any) -> Any:
    from .errors import ToolError

    try:
        if p.type in {"integer", "number"}:
            if isinstance(value, bool) or not isinstance(value, (str, int, float)):
                raise ValueError("期望数值")
            numeric = float(value)
            if not math.isfinite(numeric):
                raise ValueError("数值必须有限")
            if p.type == "integer" and not numeric.is_integer():
                raise ValueError("期望整数")
            value = int(value) if p.type == "integer" else numeric
        elif p.type == "boolean":
            if isinstance(value, bool):
                pass
            elif isinstance(value, str) and value.strip().lower() in {"1", "true", "yes", "y", "是", "0", "false", "no", "n", "否"}:
                value = value.strip().lower() in {"1", "true", "yes", "y", "是"}
            elif isinstance(value, int) and value in {0, 1}:
                value = bool(value)
            else:
                raise ValueError("期望布尔值")
        elif p.type == "string":
            if not isinstance(value, (str, int, float)) or isinstance(value, bool):
                raise ValueError("期望字符串")
            value = str(value)
        elif p.type == "object" and not isinstance(value, dict):
            raise ValueError("期望对象")
        elif p.type == "array" and not isinstance(value, list):
            raise ValueError("期望数组")
    except (TypeError, ValueError, OverflowError) as exc:
        raise ToolError(
            f"参数 {p.name} 期望 {p.type}，收到 {value!r}",
            code="args/type",
            parameter=p.name,
        ) from exc
    if p.enum and value not in p.enum:
        raise ToolError(
            f"参数 {p.name} 只能是：{' / '.join(p.enum)}",
            code="args/enum",
            parameter=p.name,
        )
    if p.type in {"integer", "number"}:
        if (p.minimum is not None and value < p.minimum) or (p.maximum is not None and value > p.maximum):
            raise ToolError(
                f"参数 {p.name} 超出允许范围（{p.minimum}..{p.maximum}）。",
                code="args/range", parameter=p.name,
            )
    return value


def iter_tools_sorted() -> Iterable[Tool]:
    return sorted(registry.all(), key=lambda t: t.name)
