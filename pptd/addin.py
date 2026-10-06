"""PowerPoint 侧的加载项（`.ppam`）。

为什么是 VBA 而不是 Office Web Add-in
------------------------------------
Office 加载项**即使在开发期也强制 HTTPS**，而 localhost 并不豁免——要跑起来就得往
用户的受信任根存储里装一张开发证书。VBA 加载项没有这些约束：它跑在 PowerPoint
进程内，直接访问 ``http://127.0.0.1:8791`` 即可，不需要证书、不需要旁加载目录。

代价是：VBA 没有 CustomTaskPane API，所以拿不到"停靠式任务窗格"，
只能给一个**无模式浮窗**（UserForm）。这是取舍，不是疏漏。

为什么控件 CLSID 要现查而不是写死
--------------------------------
``.frm`` 里的控件用 CLSID 标识，写错一个字节导入就失败。所以
:func:`msforms_clsids` 直接从注册表枚举"指向 FM20.DLL 的那些 CLSID"，
:func:`validate_frm` 再拿它校验我们发出去的 ``.frm``。
（实测：凭印象写的 ``DFD0A2C2-…`` 系列**根本不在** FM20 的注册表里，
正确的是 ``D7053240-…``（按钮）与 ``8BD21D20-…``（列表框）。）

为什么导出要转 GBK
------------------
VBE 的导入按系统 ANSI 代码页解读 ``.bas`` / ``.frm``。仓库里存 UTF-8
（便于阅读与 diff），导出时转成 GBK，否则中文会变乱码。
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from .config import Settings, settings as default_settings

__all__ = [
    "SOURCE_DIR",
    "VBA_FILES",
    "msforms_clsids",
    "validate_frm",
    "export_vba",
    "vbom_enabled",
    "vbom_hint",
    "vbom_key",
    "set_vbom",
    "powerpoint_running",
    "build_ppam",
    "status",
]

#: 仓库里的 VBA 源码（UTF-8）
SOURCE_DIR = Path(__file__).resolve().parent.parent / "addin" / "vba"
if not SOURCE_DIR.is_dir():
    SOURCE_DIR = Path(__file__).resolve().parent / "resources" / "addin" / "vba"
#: 需要导入 VBE 的文件，顺序无所谓
VBA_FILES = ("PPTAgent.bas", "PPTAgentPanel.frm")
#: .ppam 里的加载项在 PowerPoint 中显示的名字
ADDIN_NAME = "ppt-agent 监视台"


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def addin_dir() -> Path:
    source = _repo_root() / "addin"
    return source if source.is_dir() else default_settings.home / "addin"


# --------------------------------------------------------------------------
# MSForms 控件 CLSID：现查注册表，不靠记忆
# --------------------------------------------------------------------------


def _norm_clsid(value: str) -> str:
    """CLSID 归一化：统一去掉花括号、转小写。否则 ``{ABC}`` ≠ ``abc``。"""
    return value.strip().strip("{}").lower()


def msforms_clsids() -> dict[str, str]:
    """枚举 FM20.DLL 注册的控件：``归一化后的 CLSID -> 可读名``。

    非 Windows 或枚举失败时返回空字典（校验会退化成"跳过"）。
    """
    if os.name != "nt":
        return {}
    import winreg

    found: dict[str, str] = {}
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(root, r"SOFTWARE\Classes\CLSID")
        except OSError:
            continue
        index = 0
        while True:
            try:
                clsid = winreg.EnumKey(key, index)
            except OSError:
                break
            index += 1
            try:
                server_key = winreg.OpenKey(key, clsid + r"\InprocServer32")
            except OSError:
                continue
            try:
                server = str(winreg.QueryValueEx(server_key, "")[0])
            except OSError:
                continue
            if "FM20.DLL" not in server.upper():
                continue
            name = ""
            try:
                own = winreg.OpenKey(key, clsid)
                name = str(winreg.QueryValueEx(own, "")[0])
            except OSError:
                pass
            found[_norm_clsid(clsid)] = name or "(未命名)"
    return found


#: 面板用到的五个控件（校验失败时会指出是哪个）
EXPECTED_CONTROLS = ("PPTAgentPanel", "lblStatus", "lstFeed", "chkAuto", "btnFast", "btnNormal", "btnSlow", "btnRefresh", "btnStop")


def validate_frm(text: str, known: dict[str, str] | None = None) -> list[str]:
    """校验 ``.frm`` 能不能被 VBE 顺利导入。返回问题列表（空 = 没问题）。"""
    problems: list[str] = []
    if not text.startswith("VERSION 5.00"):
        problems.append("缺少 `VERSION 5.00` 头部")

    # 引用了 .frx 但分发里没有 .frx —— 导入会直接失败
    if ".frx" in text:
        problems.append("引用了 .frx 二进制资源，但我们只分发 .frm，导入会失败")
    if "OleObjectBlob" in text:
        problems.append("OleObjectBlob 指向 .frx，应删除")

    guids = re.findall(r"Begin \{([0-9A-Fa-f-]{36})\}", text)
    if not guids:
        problems.append("没有解析到任何 Begin {...} 控件块")

    catalog = known if known is not None else msforms_clsids()
    if catalog:
        for guid in guids:
            if _norm_clsid(guid) not in catalog:
                problems.append(f"控件 CLSID 不是本机注册的 MSForms 控件：{guid}")
    else:
        problems.append("（警告）枚举不到 FM20.DLL 注册表项，CLSID 未校验")

    for name in ("lblStatus", "lstFeed", "chkAuto", "btnFast", "btnStop"):
        if name not in text:
            problems.append(f"缺少控件 {name}")

    for attribute in ("Attribute VB_Name", "Attribute VB_PredeclaredId = True", "Option Explicit"):
        if attribute not in text:
            problems.append(f"缺少 {attribute}")

    return problems


# --------------------------------------------------------------------------
# 导出（UTF-8 → GBK）
# --------------------------------------------------------------------------


def export_vba(out: Path | None = None, *, encoding: str = "gbk") -> list[Path]:
    """把 VBA 源码导出成 VBE 能正确读入的编码。

    默认写到 ``addin/export/``，用 ``pptctl addin-export`` 可以指定别处。

    **遇到编不进目标编码的字符会直接报错**，而不是静默替换成 ``?``——
    实测 ``✗``(U+2717) 就不在 GBK 里，静默替换会让面板上出现莫名其妙的问号。
    """
    target = out or (addin_dir() / "export")
    target.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for name in VBA_FILES:
        source = SOURCE_DIR / name
        if not source.is_file():
            raise FileNotFoundError(f"缺少 VBA 源码：{source}")
        text = source.read_text(encoding="utf-8")
        try:
            payload = text.encode(encoding)
        except UnicodeEncodeError as exc:
            bad = sorted({text[exc.start: exc.end], text[exc.start]})
            shown = "、".join(f"{ch!r}(U+{ord(ch):04X})" for ch in bad)
            raise ValueError(
                f"{name} 里有 {encoding} 编不了的字符：{shown}。"
                f"换成 {encoding} 里存在的字符（例如把 ✗ 换成 ×），否则 VBE 里会显示成问号。"
            ) from exc
        destination = target / name
        destination.write_bytes(payload)
        written.append(destination)
    return written


# --------------------------------------------------------------------------
# 自动构建 .ppam（需要一次性开启 VBA 工程信任）
# --------------------------------------------------------------------------

VBOM_KEY = r"Software\Microsoft\Office\16.0\PowerPoint\Security"


def vbom_key() -> str:
    """当前机器上 PowerPoint 的 ``Security`` 键路径。

    版本号从注册表里探，而不是写死 16.0——换了 Office 版本也不用改代码。
    取版本号最大的那个（字符串比较对 8.0/14.0/16.0 这种也是对的）。
    """
    if os.name != "nt":
        return VBOM_KEY
    import winreg

    best = ""
    try:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office")
        index = 0
        while True:
            try:
                version = winreg.EnumKey(root, index)
            except OSError:
                break
            index += 1
            if version and version[0].isdigit() and version > best:
                try:
                    winreg.OpenKey(root, version + r"\PowerPoint\Security")
                    best = version
                except OSError:
                    continue
    except OSError:
        return VBOM_KEY
    return rf"Software\Microsoft\Office\{best}\PowerPoint\Security" if best else VBOM_KEY


def vbom_enabled() -> bool:
    """是否已信任"对 VBA 工程对象模型的访问"。

    没有它，``Application.VBE`` 会直接抛
    ``Programmatic access to Visual Basic Project is not trusted``。
    """
    if os.name != "nt":
        return False
    import winreg

    for path in dict.fromkeys((vbom_key(), VBOM_KEY)):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
            value, _ = winreg.QueryValueEx(key, "AccessVBOM")
            return int(value) == 1
        except (OSError, ValueError, TypeError):
            continue
    return False


def powerpoint_running() -> bool:
    """PowerPoint 进程在不在。

    这个判断是必需的，不是锦上添花：**PowerPoint 在运行时改 ``AccessVBOM`` 是无效的。**
    微软文档写得很明确——"the setting will not be taken into account until the application
    restarts"，而且运行期间用任何方式（宏、脚本、手改注册表）改它，
    退出时"your change will be discarded"。所以不检查就直接写，用户会以为开了、其实白开。
    """
    if os.name != "nt":
        return False
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return "POWERPNT.EXE" in (result.stdout or "")


def set_vbom(enabled: bool) -> dict[str, Any]:
    """改 ``AccessVBOM``。**调用方必须先确认 PowerPoint 没在跑。**

    返回结果里带上 ``effective``：只有 PowerPoint 不在运行时，这次改动才会在下次启动生效。
    """
    if os.name != "nt":
        return {"ok": False, "reason": "只有 Windows 上才有这个注册表键"}
    import winreg

    path = vbom_key()
    running = powerpoint_running()
    try:
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE)
    except OSError as exc:
        return {"ok": False, "reason": f"打不开注册表键：{exc}", "key": path}
    try:
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1 if enabled else 0)
    except OSError as exc:
        return {"ok": False, "reason": f"写注册表失败：{exc}", "key": path}
    finally:
        winreg.CloseKey(key)

    return {
        "ok": True,
        "key": path,
        "value": 1 if enabled else 0,
        "enabled": vbom_enabled(),
        "powerpoint_was_running": running,
        # 运行时改 = 无效，退出时会被覆盖回去
        "effective": not running,
    }


def vbom_hint() -> str:
    return (
        "需要一次性开启「信任对 VBA 工程对象模型的访问」：\n"
        "  PowerPoint → 文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置\n"
        "  → 勾选「信任对 VBA 工程对象模型的访问」→ 确定 → 重启 PowerPoint\n"
        "或者用一条命令代劳（会先要求你关掉 PowerPoint）：\n"
        "  pptctl addin-trust-vba --enable\n"
        "这是微软为 VBA 开发提供的标准开关，随时可以取消（--disable）。\n"
        "不想开也可以手动导入（见 addin/README.md），两条路等价。"
    )


def build_ppam(out: Path | None = None, *, cfg: Settings | None = None) -> dict[str, Any]:
    """用 PowerPoint 自己把 VBA 源码编译成 ``.ppam``。

    **需要 VBA 工程信任**。流程：新建一个临时演示承载 VBA 工程 → 导入 .bas/.frm
    → 另存为 .ppam → 关掉临时演示（不保存 pptx）。
    """
    if not vbom_enabled():
        return {"ok": False, "error": "vbom-not-trusted", "hint": vbom_hint()}

    import win32com.client

    target = out or (addin_dir() / "ppt-agent.ppam")
    target.parent.mkdir(parents=True, exist_ok=True)

    staging = Path(tempfile.mkdtemp(prefix="ppt-agent-vba-"))
    try:
        files = export_vba(staging)
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        app.DisplayAlerts = 1  # ppAlertsNone
        pres = app.Presentations.Add(True)
        try:
            vbe = app.VBE
            project = _pick_project(vbe, pres)
            imported: list[str] = []
            for path in files:
                component = project.VBComponents.Import(str(path))
                imported.append(str(component.Name))
            # 另存为 .ppam：扩展名会决定格式，不必记枚举值
            pres.SaveAs(str(target))
        finally:
            try:
                pres.Close()
            except Exception:  # noqa: BLE001 - 关掉临时演示失败不该吞掉构建结果
                pass
        return {
            "ok": target.is_file(),
            "ppam": str(target),
            "bytes": target.stat().st_size if target.is_file() else 0,
            "imported": imported,
        }
    finally:
        shutil.rmtree(staging, ignore_errors=True)


def _pick_project(vbe: Any, pres: Any) -> Any:
    """找到承载这份演示的 VBA 工程。"""
    full = str(getattr(pres, "FullName", "") or "").casefold()
    projects = vbe.VBProjects
    for index in range(1, projects.Count + 1):
        project = projects(index)
        if str(getattr(project, "FileName", "") or "").casefold() == full:
            return project
    return projects(projects.Count)


# --------------------------------------------------------------------------
# 状态
# --------------------------------------------------------------------------


def status(cfg: Settings | None = None) -> dict[str, Any]:
    """加载项当前处于什么状态，给 ``pptctl addin-status`` 用。"""
    sources = {name: (SOURCE_DIR / name).is_file() for name in VBA_FILES}
    exported = sorted(p.name for p in (addin_dir() / "export").glob("*") if p.suffix in {".bas", ".frm"})
    ppam = addin_dir() / "ppt-agent.ppam"

    frm_path = SOURCE_DIR / "PPTAgentPanel.frm"
    problems: list[str] = []
    if frm_path.is_file():
        problems = validate_frm(frm_path.read_text(encoding="utf-8"))

    return {
        "source_dir": str(SOURCE_DIR),
        "sources": sources,
        "export_dir": str(addin_dir() / "export"),
        "exported": exported,
        "ppam": str(ppam) if ppam.is_file() else None,
        "ppam_bytes": ppam.stat().st_size if ppam.is_file() else 0,
        "vbom_enabled": vbom_enabled(),
        "frm_problems": problems,
        "frm_ok": not problems,
        "install_doc": str(SOURCE_DIR.parent / "README.md"),
        "verdict": _verdict(sources, exported, ppam, problems),
    }


def _verdict(sources: dict[str, bool], exported: list[str], ppam: Path, problems: list[str]) -> str:
    if not all(sources.values()):
        return "VBA 源码不完整"
    if problems:
        return f".frm 有问题：{problems[0]}"
    if ppam.is_file():
        return "已构建 .ppam：在 PowerPoint 里 文件→选项→加载项 勾选即可"
    if not vbom_enabled():
        return "源码就绪；开启 VBA 工程信任可自动构建，或按 addin/README.md 手动导入"
    return "源码就绪且已信任 VBA 工程：跑 `pptctl addin-build` 构建 .ppam"
