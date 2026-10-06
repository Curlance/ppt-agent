"""DSH 插件 bundle：清单契约、宿主半边、以及**浏览器半边的契约探针**。

浏览器半边（`client.js`）在真机上要等 DSH 重启才会被加载，所以我没法端到端验证它。
但**契约**可以验，而且最容易写错的就是契约：注册 id 是否等于包名、注入的 slot 对不对、
工厂有没有副作用、组件卸载后定时器有没有清掉。`tests/dsh_client_probe.mjs` 用桩把这些
逐条查一遍——同 `.frm` 的 CLSID 校验一个思路。
"""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest
import yaml

from pptd import dsh_bundle

ROOT = Path(__file__).resolve().parent.parent
PROBE = ROOT / "tests" / "dsh_client_probe.mjs"
#: DSH 自带的 Node；PATH 上没有时用它
DSH_NODE = Path.home() / ".dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe"


def _node() -> str | None:
    found = shutil.which("node")
    if found:
        return found
    return str(DSH_NODE) if DSH_NODE.is_file() else None


# --------------------------------------------------------------------------
# 生成
# --------------------------------------------------------------------------


def test_bundle_writes_four_files(tmp_path: Path) -> None:
    written = dsh_bundle.write_bundle(out=tmp_path)
    assert set(written) == {"package", "patch", "index", "client"}
    for path in written.values():
        assert path.is_file() and path.stat().st_size > 0


def test_package_json_declares_both_halves(tmp_path: Path) -> None:
    dsh_bundle.write_bundle(out=tmp_path)
    package = json.loads((tmp_path / "package.json").read_text(encoding="utf-8"))
    assert package["type"] == "module"
    assert package["exports"] == {".": "./index.js", "./client": "./client.js"}
    assert package["dsh"]["bundle"]["patch"] == "./cordis.patch.yml"

    client = package["dsh"]["client"]
    assert client["platform"] == "web"
    assert client["immediately"] is True
    assert client["inject"] == ["@deepseek-ai/dsh-client-ui-conversation"]


def test_patch_inserts_both_rows(tmp_path: Path) -> None:
    written = dsh_bundle.write_bundle(out=tmp_path)
    patch = yaml.safe_load(written["patch"].read_text(encoding="utf-8"))
    rows = patch[0]["insert"]
    names = [row["name"] for row in rows]
    assert "@deepseek-ai/dsh-mcp-client" in names, "少了 MCP 客户端这一行就没有工具"
    assert dsh_bundle.PACKAGE_NAME in names, "少了这一行就没有状态条"

    mcp = next(r for r in rows if r["name"] == "@deepseek-ai/dsh-mcp-client")
    assert mcp["config"]["serverName"] == "ppt"
    assert mcp["config"]["transport"] == "stdio"
    assert mcp["config"]["failOnStartupError"] is False, "起不来只该少工具，不该拖垮 harness"


def test_host_half_export_func_apply_is_empty(tmp_path: Path) -> None:
    """宿主半边刻意什么都不做——空的 apply 意味着零失败点。"""
    written = dsh_bundle.write_bundle(out=tmp_path)
    text = written["index"].read_text(encoding="utf-8")
    assert "export function apply() {}" in text
    body = text.split("export function apply() {}")[0]
    # 只有注释，没有可执行语句
    code_lines = [
        line for line in body.splitlines()
        if line.strip() and not line.strip().startswith(("*", "/**", "//"))
    ]
    assert code_lines == [], f"宿主半边不该有可执行代码：{code_lines}"


def test_client_js_is_deterministic_and_uses_the_text_endpoint(tmp_path: Path) -> None:
    first = dsh_bundle.write_bundle(out=tmp_path / "a")
    second = dsh_bundle.write_bundle(out=tmp_path / "b")
    assert first["client"].read_text(encoding="utf-8") == second["client"].read_text(encoding="utf-8")
    text = first["client"].read_text(encoding="utf-8")
    assert "/strip.txt" in text
    assert "window.__ModuleLoader__.load" in text
    # 状态条不该自己去读需要令牌的端点
    assert "/feed" not in text


def test_install_hint_mentions_both_deliverables() -> None:
    hint = dsh_bundle.install_hint()
    assert "mcp__ppt__" in hint
    assert "状态条" in hint
    assert "重启 DSH" in hint


# --------------------------------------------------------------------------
# 浏览器半边的契约探针
# --------------------------------------------------------------------------


@pytest.fixture(scope="module")
def probe_result(tmp_path_factory) -> dict:
    node = _node()
    if node is None:
        pytest.skip("本机没有可用的 node，跳过浏览器半边契约探针")

    bundle = tmp_path_factory.mktemp("dsh-bundle")
    dsh_bundle.write_bundle(out=bundle)
    proc = subprocess.run(
        [node, str(PROBE), str(bundle / "client.js")],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=90,
    )
    lines = [line for line in (proc.stdout or "").splitlines() if line.strip()]
    assert lines, f"探针没有输出。stderr：\n{proc.stderr[:1500]}"
    return json.loads(lines[-1])


def test_client_probe_all_checks_pass(probe_result: dict) -> None:
    failed = [c for c in probe_result["checks"] if not c["pass"]]
    assert not failed, "浏览器半边不符合契约：" + "; ".join(f"{c['name']}（{c['detail']}）" for c in failed)
    assert not probe_result["errors"], probe_result["errors"]


def test_client_probe_covers_the_contract(probe_result: dict) -> None:
    """探针本身也要有覆盖度——别让它悄悄退化成"什么都没查"。

    每条对应一个真会踩的坑：注册 id 写错宿主就找不到模块；slot 名写错就挂不上；
    工厂有副作用会在装载阶段炸；忘了清理定时器会让面板重复轮询。
    """
    # 名字和详情都算——有些检查的证据在 detail 里（比如提示文案的具体内容）
    haystack = " ".join(f"{c['name']} {c.get('detail') or ''}" for c in probe_result["checks"])
    for must in (
        "注册 id",                    # 宿主按 id 找模块
        "注入 slots",                 # 不注入就拿不到 slot 服务
        "conversation.composer.dock",  # slot 名写错就挂不上
        "factory 调用本身不渲染",      # 工厂必须无副作用
        "清掉了定时器",                # 卸载不清理会重复轮询
        "pptctl serve",               # 连不上时要有可执行的提示
    ):
        assert must in haystack, f"探针少了「{must}」这一项检查"


# --------------------------------------------------------------------------
# 安装状态
# --------------------------------------------------------------------------


def test_profile_status_reports_client_files(tmp_path: Path) -> None:
    profile = tmp_path / "desktop"
    (profile / "node_modules" / dsh_bundle.PACKAGE_NAME).mkdir(parents=True)
    (profile / "package.json").write_text(
        json.dumps({"dsh": {"profile": {"bundles": [dsh_bundle.PACKAGE_NAME]}}}, ensure_ascii=False),
        encoding="utf-8",
    )
    linked = profile / "node_modules" / dsh_bundle.PACKAGE_NAME
    (linked / "cordis.patch.yml").write_text("- insert: []\n", encoding="utf-8")

    status = dsh_bundle.profile_status(profile)
    assert status["listed"] is True and status["linked"] is True and status["patch"] is True
    assert status["client"] is False, "缺 index.js/client.js 时应认为浏览器半边没装好"

    (linked / "index.js").write_text("export function apply() {}\n", encoding="utf-8")
    (linked / "client.js").write_text("// x\n", encoding="utf-8")
    assert dsh_bundle.profile_status(profile)["client"] is True
