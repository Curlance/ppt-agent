"""注册表与参数校验：三个面共用的那份 schema 必须准确。"""

from __future__ import annotations

import pytest

from pptd.errors import ToolError
from pptd.registry import Param, all_tools, get_tool, tool_manifest, validate_args


def test_expected_tools_registered() -> None:
    import pptd.ops  # noqa: F401

    names = {t.name for t in all_tools()}
    assert {"ppt_status", "ppt_open", "ppt_activate", "ppt_close"} <= names


def test_schema_reflects_params() -> None:
    import pptd.ops  # noqa: F401

    schema = get_tool("ppt_open").input_schema()
    assert schema["type"] == "object"
    assert schema["required"] == ["path"]
    assert set(schema["properties"]) == {"path", "read_only", "bring_to_front"}
    # 参数描述必须带上——模型就是靠它选参数的
    assert schema["properties"]["path"]["description"]
    assert schema["properties"]["read_only"]["default"] is False
    assert schema["additionalProperties"] is False


def test_manifest_is_json_serialisable() -> None:
    import json

    import pptd.ops  # noqa: F401

    payload = json.dumps(tool_manifest(), ensure_ascii=False)
    assert "ppt_status" in payload


def test_validate_fills_defaults() -> None:
    import pptd.ops  # noqa: F401

    args = validate_args(get_tool("ppt_open"), {"path": "D:/a.pptx"})
    assert args["read_only"] is False
    assert args["bring_to_front"] is True


def test_validate_missing_required() -> None:
    import pptd.ops  # noqa: F401

    with pytest.raises(ToolError) as exc:
        validate_args(get_tool("ppt_open"), {})
    assert exc.value.code == "args/missing"
    assert "path" in exc.value.message


def test_validate_rejects_unknown_arg() -> None:
    import pptd.ops  # noqa: F401

    with pytest.raises(ToolError) as exc:
        validate_args(get_tool("ppt_activate"), {"nope": 1})
    assert exc.value.code == "args/unknown"


def test_validate_coerces_types() -> None:
    import pptd.ops  # noqa: F401

    # 字符串 "1" / "true" 应该被转成正确的类型（CLI 与 HTTP 查询串都靠这个）
    args = validate_args(get_tool("ppt_open"), {"path": "D:/a.pptx", "read_only": "true"})
    assert args["read_only"] is True

    custom = Param(name="n", type="integer", description="数量", default=0)
    assert custom.to_schema()["type"] == "integer"
    assert custom.py_type is int


def test_unknown_tool_raises() -> None:
    import pptd.ops  # noqa: F401

    with pytest.raises(ToolError) as exc:
        get_tool("ppt_does_not_exist")
    assert exc.value.code == "tool/unknown"


def test_annotations_carry_descriptions() -> None:
    """MCP 面靠 to_annotation() 生成 inputSchema，描述必须能传到 pydantic。"""
    import pptd.ops  # noqa: F401

    param = get_tool("ppt_open").param("path")
    assert param is not None
    annotation = param.to_annotation()
    assert "演示文件路径" in str(annotation)
    assert param.to_parameter().name == "path"


@pytest.mark.parametrize("name,args,code", [
    ("ppt_save", {"mode": "invalid"}, "args/enum"),
    ("ppt_set_pace", {"pace_ms": -1}, "args/range"),
    ("ppt_shot", {"count": 11}, "args/range"),
    ("ppt_goto", {"slide": 1.9}, "args/type"),
    ("ppt_open", {"path": "D:/a.pptx", "read_only": "perhaps"}, "args/type"),
    ("ppt_set_geometry", {"shape": "2", "left": float("nan")}, "args/type"),
    ("ppt_set_geometry", {"shape": "2", "width": float("inf")}, "args/type"),
])
def test_validation_enforces_contract(name, args, code):
    with pytest.raises(ToolError) as error:
        validate_args(get_tool(name), args)
    assert error.value.code == code


@pytest.mark.parametrize("name", ["ppt_set_text", "ppt_set_notes", "ppt_add_textbox"])
def test_explicit_empty_text_is_preserved(name):
    args = {"text": "", **({"shape": "2"} if name == "ppt_set_text" else {})}
    assert validate_args(get_tool(name), args)["text"] == ""


def test_body_schema_preserves_constraints():
    from pptd.http_api import _body_model
    from pydantic import ValidationError

    model = _body_model(get_tool("ppt_set_pace"))
    with pytest.raises(ValidationError):
        model(pace_ms=-1)
    with pytest.raises(ValidationError):
        model(pace_ms=0, typo=1)
    schema = _body_model(get_tool("ppt_save")).model_json_schema()
    assert "copy" in str(schema["properties"]["mode"])
