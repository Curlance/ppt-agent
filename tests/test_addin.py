"""PowerPoint 加载项：`.frm` 校验、编码导出、构建前置条件。

其中 ``.frm`` 的控件 CLSID 校验是**真验证**：它拿注册表里"指向 FM20.DLL 的那些
CLSID"来比对，能挡住"凭记忆写 GUID 导致导入失败"这一类错误。
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from pptd import addin


needs_windows = pytest.mark.skipif(os.name != "nt", reason="注册表校验只在 Windows 上有意义")

#: 自测专用的注册表位置——**绝不碰**用户真实的 AccessVBOM。
SELFTEST_KEY = r"Software\ppt-agent-selftest\PowerPoint\Security"


# --------------------------------------------------------------------------
# CLSID 目录
# --------------------------------------------------------------------------


@needs_windows
def test_msforms_catalog_is_populated() -> None:
    catalog = addin.msforms_clsids()
    assert catalog, "枚举不到 FM20.DLL 注册的控件，校验会退化成空转"
    names = " ".join(catalog.values()).lower()
    assert "form" in names and "commandbutton" in names and "listbox" in names


@needs_windows
def test_catalog_keys_are_normalised() -> None:
    """花括号/大小写不一致会让查找永远失败——这正是第一次跑出来的 bug。"""
    catalog = addin.msforms_clsids()
    for key in catalog:
        assert "{" not in key and "}" not in key
        assert key == key.lower()


@needs_windows
def test_userform_clsid_is_known() -> None:
    """窗体本身的 CLSID 必须也在目录里，否则 .frm 连文件头都过不了。"""
    catalog = addin.msforms_clsids()
    assert "c62a69f0-16dc-11ce-9e98-00aa00574a4f" in catalog


# --------------------------------------------------------------------------
# .frm 校验
# --------------------------------------------------------------------------


@needs_windows
def test_shipped_frm_is_valid() -> None:
    """我们实际分发的那份 .frm 必须零问题。"""
    text = (addin.SOURCE_DIR / "PPTAgentPanel.frm").read_text(encoding="utf-8")
    assert addin.validate_frm(text) == []


def test_frm_missing_header_is_flagged() -> None:
    problems = addin.validate_frm("Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} X\nEnd\n")
    assert any("VERSION 5.00" in p for p in problems)


def test_frm_with_frx_reference_is_rejected() -> None:
    """引用 .frx 但只分发 .frm —— 导入必然失败，必须拦下。"""
    text = (
        "VERSION 5.00\n"
        "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTAgentPanel \n"
        '   OleObjectBlob   =   "PPTAgentPanel.frx":0000\n'
        "End\n"
        "Attribute VB_Name = \"PPTAgentPanel\"\n"
    )
    problems = addin.validate_frm(text, known={})
    assert any(".frx" in p for p in problems)
    assert any("OleObjectBlob" in p for p in problems)


def test_frm_with_unknown_clsid_is_rejected() -> None:
    text = (
        "VERSION 5.00\n"
        "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTAgentPanel \n"
        "   Begin {DEADBEEF-0000-0000-0000-000000000000} bogus \n"
        "   End\n"
        "End\n"
    )
    problems = addin.validate_frm(text, known={"c62a69f0-16dc-11ce-9e98-00aa00574a4f": "Form"})
    assert any("DEADBEEF" in p for p in problems)


def test_frm_missing_expected_controls_is_flagged() -> None:
    text = "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PPTAgentPanel \nEnd\n"
    problems = addin.validate_frm(text, known={})
    assert any("lblStatus" in p for p in problems)
    assert any("btnStop" in p for p in problems)


# --------------------------------------------------------------------------
# 导出（UTF-8 → GBK）
# --------------------------------------------------------------------------


def test_export_writes_both_files(tmp_path: Path) -> None:
    written = addin.export_vba(tmp_path)
    assert {p.name for p in written} == set(addin.VBA_FILES)
    assert all(p.is_file() and p.stat().st_size > 0 for p in written)


def test_export_uses_gbk_so_vbe_reads_chinese(tmp_path: Path) -> None:
    """VBE 按系统 ANSI 代码页读文件；用 UTF-8 导出中文会变乱码。"""
    addin.export_vba(tmp_path)
    raw = (tmp_path / "PPTAgentPanel.frm").read_bytes()
    assert "监视台".encode("gbk") in raw
    text = raw.decode("gbk")
    assert "ppt-agent 监视台" in text


def test_export_round_trips_to_source(tmp_path: Path) -> None:
    addin.export_vba(tmp_path)
    for name in addin.VBA_FILES:
        source = (addin.SOURCE_DIR / name).read_text(encoding="utf-8")
        exported = (tmp_path / name).read_bytes().decode("gbk")
        assert exported == source, f"{name} 转码后与源码不一致"


def test_export_refuses_characters_the_encoding_cannot_represent(tmp_path: Path, monkeypatch) -> None:
    """静默替换成 `?` 会让面板上冒出莫名其妙的问号，必须直接报错。

    这条是被真实 bug 逼出来的：`✗`(U+2717) 不在 GBK 里。
    """
    broken = tmp_path / "src"
    broken.mkdir()
    (broken / "PPTAgent.bas").write_text("Attribute VB_Name = \"PPTAgent\"\n' ✗\n", encoding="utf-8")
    (broken / "PPTAgentPanel.frm").write_text("VERSION 5.00\n", encoding="utf-8")
    monkeypatch.setattr(addin, "SOURCE_DIR", broken)

    with pytest.raises(ValueError) as exc:
        addin.export_vba(tmp_path / "out")
    assert "U+2717" in str(exc.value)
    assert "×" in str(exc.value), "错误信息要给出可用的替代字符"


def test_export_missing_source_raises(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr(addin, "SOURCE_DIR", tmp_path / "nope")
    with pytest.raises(FileNotFoundError):
        addin.export_vba(tmp_path / "out")


# --------------------------------------------------------------------------
# 构建前置条件与状态
# --------------------------------------------------------------------------


def test_vbom_enabled_returns_bool() -> None:
    assert isinstance(addin.vbom_enabled(), bool)


def test_build_without_trust_explains_how_to_enable() -> None:
    """没开信任时要给出**可执行**的指引，而不是一句失败。"""
    if addin.vbom_enabled():  # pragma: no cover - 本机默认未开启
        pytest.skip("本机已开启 VBA 工程信任，这条走的是另一分支")
    result = addin.build_ppam()
    assert result["ok"] is False
    assert result["error"] == "vbom-not-trusted"
    assert "信任中心" in result["hint"]
    assert "信任对 VBA 工程对象模型的访问" in result["hint"]


def test_status_shape() -> None:
    report = addin.status()
    assert report["sources"]["PPTAgent.bas"] is True
    assert report["sources"]["PPTAgentPanel.frm"] is True
    assert isinstance(report["frm_ok"], bool)
    assert isinstance(report["vbom_enabled"], bool)
    assert report["verdict"]
    assert Path(report["install_doc"]).is_file(), "安装说明必须随源码一起存在"


# --------------------------------------------------------------------------
# VBA 工程信任开关
#
# 「PowerPoint 运行时改这个键无效、退出时会被覆盖回去」是微软文档写明的行为，
# 所以这里重点测两件事：**能不能正确写**、以及**运行时会不会如实说"这次不生效"**。
# 全部在自测注册表位置进行，绝不碰用户真实的 AccessVBOM。
# --------------------------------------------------------------------------


@pytest.fixture()
def self_test_key(monkeypatch):
    """把注册表路径指向自测位置，用完删掉。"""
    if os.name != "nt":
        pytest.skip("注册表只在 Windows 上有意义")
    import winreg

    monkeypatch.setattr(addin, "vbom_key", lambda: SELFTEST_KEY)

    def cleanup() -> None:
        for path in (SELFTEST_KEY, r"Software\ppt-agent-selftest\PowerPoint", r"Software\ppt-agent-selftest"):
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
            except OSError:
                pass

    cleanup()
    yield SELFTEST_KEY
    cleanup()


@needs_windows
def test_set_vbom_writes_and_reads_back(self_test_key, monkeypatch) -> None:
    import winreg

    # 桩掉进程检查：本机 PowerPoint 可能正开着，而"开着时"的情况有专门的测试
    monkeypatch.setattr(addin, "powerpoint_running", lambda: False)

    result = addin.set_vbom(True)
    assert result["ok"] is True, result
    assert result["enabled"] is True
    assert result["effective"] is True

    # 真的落到了注册表里，而不是只改了内存
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self_test_key) as key:
        value, kind = winreg.QueryValueEx(key, "AccessVBOM")
    assert value == 1
    assert kind == winreg.REG_DWORD

    assert addin.vbom_enabled() is True
    assert addin.set_vbom(False)["enabled"] is False
    assert addin.vbom_enabled() is False


@needs_windows
def test_set_vbom_creates_a_missing_key(self_test_key) -> None:
    """键不存在时也要能建出来——全新机器上本来就没有这个键。"""
    import winreg

    with pytest.raises(OSError):
        winreg.OpenKey(winreg.HKEY_CURRENT_USER, self_test_key)

    assert addin.set_vbom(True)["ok"] is True
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self_test_key) as key:
        assert winreg.QueryValueEx(key, "AccessVBOM")[0] == 1


@needs_windows
def test_set_vbom_reports_not_effective_while_powerpoint_runs(self_test_key, monkeypatch) -> None:
    """PowerPoint 在跑时必须如实说"这次不生效"。

    微软文档：这个设置只在应用程序启动时读取；运行期间改它，退出时会被丢弃。
    如果这里不如实上报，用户会以为开好了、其实白开——比报错更糟。
    """
    monkeypatch.setattr(addin, "powerpoint_running", lambda: True)
    result = addin.set_vbom(True)
    assert result["ok"] is True
    assert result["powerpoint_was_running"] is True
    assert result["effective"] is False, "运行时改动必须标记为不生效"


@needs_windows
def test_vbom_key_detects_installed_version() -> None:
    key = addin.vbom_key()
    assert key.startswith("Software\\Microsoft\\Office\\")
    assert key.endswith("\\PowerPoint\\Security")
    assert "Office" in key and "PowerPoint" in key
    # 本机装了 Office 16.0（换成别的版本也不该失败，只要求能探到某个版本号）
    assert any(part and part[0].isdigit() for part in key.split("\\")), key


@needs_windows
def test_powerpoint_running_returns_bool() -> None:
    assert isinstance(addin.powerpoint_running(), bool)


def test_vbom_hint_offers_the_command() -> None:
    hint = addin.vbom_hint()
    assert "信任对 VBA 工程对象模型的访问" in hint
    assert "addin-trust-vba --enable" in hint
    assert "--disable" in hint, "必须告诉用户怎么关掉"
    assert "信任中心" in hint, "手动路径也要留着"


def test_cli_refuses_to_write_while_powerpoint_runs(monkeypatch, capsys) -> None:
    """安全闸门：PowerPoint 在跑时，命令必须**拒绝**而不是静默写一个会被丢弃的值。"""
    from pptd import cli

    monkeypatch.setattr(addin, "powerpoint_running", lambda: True)
    monkeypatch.setattr(addin, "vbom_enabled", lambda: False)
    written: list[bool] = []
    monkeypatch.setattr(addin, "set_vbom", lambda flag: written.append(flag) or {"ok": True})

    code = cli.main(["addin-trust-vba", "--enable"])
    output = capsys.readouterr().out
    assert code == 1, "应当以失败退出"
    assert written == [], "绝不能真的去写注册表"
    assert "不会生效" in output
    assert "关掉 PowerPoint" in output


def test_cli_enable_reports_security_tradeoff(monkeypatch, capsys) -> None:
    """开启时要明说有安全代价——微软文档也强调这事必须由用户自己决定。"""
    from pptd import cli

    monkeypatch.setattr(addin, "powerpoint_running", lambda: False)
    monkeypatch.setattr(addin, "vbom_enabled", lambda: False)
    monkeypatch.setattr(
        addin, "set_vbom", lambda flag: {"ok": True, "key": SELFTEST_KEY, "value": 1, "enabled": True}
    )

    code = cli.main(["addin-trust-vba", "--enable"])
    output = capsys.readouterr().out
    assert code == 0
    assert "宏病毒" in output or "任何**程序**" in output.replace("*", "")
    assert "addin-trust-vba --disable" in output, "必须给出关闭方法"
