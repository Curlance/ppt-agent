"""接口文档与 agent 技能：生成是否正确、是否漂移、装了能不能用。

其中"生成的配置必须是合法 JSON"这条是**被真实 bug 逼出来的**：
第一版 CONNECT.md 里的 MCP 客户端配置被我手拼进了一个不存在的 `env_extra` 键，
复制到宿主里根本用不了。
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pptd import apidocs, skill


# --------------------------------------------------------------------------
# 工具参考
# --------------------------------------------------------------------------


def test_tools_markdown_covers_every_tool() -> None:
    import pptd.ops  # noqa: F401
    from pptd.registry import all_tools

    text = apidocs.render_tools_markdown()
    for t in all_tools():
        assert f"`{t.name}`" in text, f"{t.name} 没进工具参考"
    assert f"共 **{len(all_tools())} 个工具**" in text
    assert "请勿手改" in text


def test_tools_markdown_documents_parameters() -> None:
    text = apidocs.render_tools_markdown()
    # ppt_open 的必填 path 与默认值都要出现
    assert "| `path` | string | 是 |" in text
    assert "read_only" in text
    # 枚举与范围要写进说明
    assert "png / jpg" in text
    assert "范围" in text


def test_tools_markdown_is_stable() -> None:
    """生成必须确定：跑两次结果一样，否则漂移检查会永远误报。"""
    assert apidocs.render_tools_markdown() == apidocs.render_tools_markdown()


# --------------------------------------------------------------------------
# OpenAPI
# --------------------------------------------------------------------------


def test_openapi_has_all_tool_endpoints() -> None:
    import pptd.ops  # noqa: F401
    from pptd.registry import all_tools

    spec = apidocs.build_openapi()
    paths = spec["paths"]
    for t in all_tools():
        assert f"/tools/{t.name}" in paths, f"{t.name} 没进 OpenAPI"
        assert paths[f"/tools/{t.name}"]["post"]["operationId"] == t.name
    assert "/feed" in paths and "/feed.txt" in paths and "/live" in paths


def test_openapi_is_deterministic() -> None:
    first = json.dumps(apidocs.build_openapi(), ensure_ascii=False, sort_keys=True)
    second = json.dumps(apidocs.build_openapi(), ensure_ascii=False, sort_keys=True)
    assert first == second


# --------------------------------------------------------------------------
# 接入指南
# --------------------------------------------------------------------------


def _json_blocks(text: str) -> list[str]:
    return re.findall(r"```json\n(.*?)```", text, re.S)


def test_connect_guide_json_blocks_are_valid() -> None:
    """生成的配置必须能直接粘贴使用——手拼字符串很容易塞进不存在的键。"""
    blocks = _json_blocks(apidocs.render_connect_markdown())
    assert len(blocks) >= 2, "至少要有 stdio 与 HTTP 两份配置"
    for block in blocks:
        data = json.loads(block)  # 解析失败即测试失败
        server = data["mcpServers"]["ppt"]
        assert set(server) <= {"command", "args", "cwd", "env", "url", "headers"}, server


def test_connect_guide_uses_this_machines_real_paths() -> None:
    from pptd.config import host_interpreter, settings as cfg

    text = apidocs.render_connect_markdown()
    exe, extra_env = host_interpreter()
    assert exe in text
    assert str(cfg.port) in text
    assert str(cfg.mcp_bind_port) in text

    # PYTHONPATH 在 JSON 里是转义过的（`\\`），所以要比对**解析后**的值
    stdio = json.loads(_json_blocks(text)[0])["mcpServers"]["ppt"]
    assert stdio["command"] == exe
    assert Path(stdio["cwd"]).is_dir(), "cwd 必须指向真实存在的仓库根"
    if extra_env.get("PYTHONPATH"):
        assert stdio["env"]["PYTHONPATH"] == extra_env["PYTHONPATH"]


def test_connect_guide_warns_chatgpt_cannot_reach_localhost() -> None:
    text = apidocs.render_connect_markdown()
    assert "127.0.0.1" in text and "访问不到" in text
    assert "令牌不是可选项" in text, "必须把鉴权是硬要求这件事说清楚"


def test_public_docs_do_not_depend_on_host_interpreter(monkeypatch, tmp_path: Path) -> None:
    def unexpected_host_lookup():
        pytest.fail("公开文档不应包含作者本机的解释器路径")

    monkeypatch.setattr(apidocs, "host_interpreter", unexpected_host_lookup)
    apidocs.write_api_docs(tmp_path)
    assert apidocs.check_api_docs(tmp_path) == []
    stdio = json.loads(_json_blocks((tmp_path / "CONNECT.md").read_text(encoding="utf-8"))[0])["mcpServers"]["ppt"]
    assert stdio["cwd"].replace("\\", "/") == "C:/path/to/ppt-agent"


def test_local_docs_can_still_export_real_paths(tmp_path: Path) -> None:
    from pptd.config import host_interpreter

    apidocs.write_api_docs(tmp_path, local_paths=True)
    assert apidocs.check_api_docs(tmp_path, local_paths=True) == []
    stdio = json.loads(_json_blocks((tmp_path / "CONNECT.md").read_text(encoding="utf-8"))[0])["mcpServers"]["ppt"]
    assert stdio["command"] == host_interpreter()[0]


# --------------------------------------------------------------------------
# 落盘与漂移
# --------------------------------------------------------------------------


def test_write_then_check_has_no_drift(tmp_path: Path) -> None:
    apidocs.write_api_docs(tmp_path)
    assert apidocs.check_api_docs(tmp_path) == []


def test_check_reports_missing_file(tmp_path: Path) -> None:
    problems = apidocs.check_api_docs(tmp_path)
    assert problems and "还没有生成" in problems[0]


def test_check_detects_edited_file(tmp_path: Path) -> None:
    apidocs.write_api_docs(tmp_path)
    (tmp_path / "tools.md").write_text("被手改了\n", encoding="utf-8")
    problems = apidocs.check_api_docs(tmp_path)
    assert any("tools.md" in p and "过期" in p for p in problems)


def test_shipped_docs_are_up_to_date() -> None:
    """仓库里那份文档必须和当前代码一致——否则模型会照着旧 schema 调用。"""
    problems = apidocs.check_api_docs()
    assert problems == [], "接口文档已漂移，重跑 `pptctl export-api`：" + "; ".join(problems)


# --------------------------------------------------------------------------
# 技能
# --------------------------------------------------------------------------


def test_skill_frontmatter_is_valid() -> None:
    text = (skill.skill_source() / "SKILL.md").read_text(encoding="utf-8")
    meta = skill._parse_frontmatter(text)
    assert meta["name"] == "ppt-agent"
    assert len(meta["description"]) > 200, "描述要足够长，宿主才能判断'什么时候该用'"
    # 描述里要带上用户可能说的中文说法，否则中文提问触发不到
    assert "PPT" in meta["description"]


def test_skill_body_states_the_undo_caveat_and_verification_rule() -> None:
    """技能的价值在于把'用错会怎样'提前讲清楚。"""
    body = (skill.skill_source() / "SKILL.md").read_text(encoding="utf-8")
    assert "撤销" in body and "回退" in body
    assert "绝对路径" in body
    assert "改完不验证就等于没做完" in body
    assert "不要自己再起一个 PowerPoint" in body


def test_skill_install_into_temp_dir(tmp_path: Path) -> None:
    result = skill.install(tmp_path / "ppt-agent")
    assert result["ok"] is True
    installed = tmp_path / "ppt-agent" / "SKILL.md"
    assert installed.is_file()
    assert installed.read_bytes() == (skill.skill_source() / "SKILL.md").read_bytes()


def test_skill_install_backs_up_previous(tmp_path: Path) -> None:
    target = tmp_path / "ppt-agent"
    target.mkdir(parents=True)
    (target / "SKILL.md").write_text("旧版本", encoding="utf-8")
    (target / "extra.txt").write_text("用户自己加的", encoding="utf-8")

    result = skill.install(target)
    assert result["ok"] is True
    backup = Path(result["backup"])
    assert backup.is_dir()
    assert (backup / "SKILL.md").read_text(encoding="utf-8") == "旧版本"
    assert (backup / "extra.txt").is_file(), "覆盖前要把旧目录整个留一份"


def test_skill_install_rejects_bad_frontmatter(tmp_path: Path, monkeypatch) -> None:
    broken = tmp_path / "src"
    broken.mkdir()
    (broken / "SKILL.md").write_text("---\nname: 别的名字\n---\n内容\n", encoding="utf-8")
    monkeypatch.setattr(skill, "skill_source", lambda: broken)

    result = skill.install(tmp_path / "out")
    assert result["ok"] is False
    assert any("name" in p for p in result["problems"])


def test_skill_status_reports_not_installed(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(skill, "installed_path", lambda: tmp_path / "nope")
    report = skill.status()
    assert report["frontmatter_ok"] is True
    assert "尚未安装" in report["verdict"]


def test_skill_status_detects_drift(tmp_path: Path, monkeypatch) -> None:
    target = tmp_path / "ppt-agent"
    skill.install(target)
    (target / "SKILL.md").write_text("被改过了", encoding="utf-8")
    monkeypatch.setattr(skill, "installed_path", lambda: target)
    report = skill.status()
    assert report["installed_matches_source"] is False
    assert "不一致" in report["verdict"]
