"""命令行：``pptctl``。

人最常用的几个动作要能一行敲完：自检、看状态、打开演示、看观察台、停掉守护进程。
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from typing import Any

from . import __version__
from .config import ensure_dirs, interpreter_report, read_runtime, settings as default_settings
from .daemon import ensure, health, observer_url, runtime_info, serve, spawn, stop, wait_ready
from .registry import all_tools

__all__ = ["main"]


def _force_utf8() -> None:
    """Windows 控制台默认 GBK，中文会乱码；显式切到 UTF-8。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
        except Exception:  # noqa: BLE001 - 非 tty 或旧版本
            pass


def _dump(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2, default=str))


def _kv_pairs(items: list[str] | None) -> dict[str, Any]:
    """把 ``-a path=D:\\a.pptx`` 解析成字典；值一律当字符串，由注册表按 schema 转型。"""
    out: dict[str, Any] = {}
    for item in items or []:
        if "=" not in item:
            raise SystemExit(f"参数格式应为 key=value，收到：{item}")
        key, _, value = item.partition("=")
        out[key.strip()] = value
    return out


# --------------------------------------------------------------------------
# 子命令
# --------------------------------------------------------------------------


def cmd_doctor(args: argparse.Namespace) -> int:
    """自检：解释器 / pywin32 / PowerPoint / 守护进程 / COM 实连。"""
    cfg = default_settings
    ensure_dirs(cfg)
    lines: list[str] = [f"ppt-agent {__version__} 自检", "=" * 52]

    interp = interpreter_report()
    lines.append(f"解释器      : {interp['executable']}")
    lines.append(f"Python      : {interp['version']}（venv={interp['in_venv']}）")
    lines.append(f"pywin32     : {'可用' if interp.get('pywin32') else '不可用 —— ' + str(interp.get('pywin32_error'))}")
    lines.append(f"python-pptx : {'可用' if interp.get('python_pptx') else '不可用'}")

    # PowerPoint 是否安装
    try:
        import winreg

        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application\\CLSID")
        clsid = winreg.QueryValueEx(key, "")[0]
        lines.append(f"PowerPoint  : 已注册（CLSID {clsid}）")
    except Exception:  # noqa: BLE001
        lines.append("PowerPoint  : 未在注册表中找到 PowerPoint.Application")

    # 守护进程
    info = runtime_info(cfg)
    if info:
        lines.append(f"守护进程    : 运行中 pid={info.get('pid')} 端口={info.get('port')}")
        url = observer_url(cfg)
        lines.append(f"观察台      : {url}")
        lines.append(f"MCP over HTTP: {cfg.mcp_url}（需要令牌）")
    else:
        lines.append("守护进程    : 未运行（pptctl serve 启动，或 pptctl status 自动拉起）")
        stale = read_runtime(cfg)
        if stale:
            lines.append(f"              （发现陈旧的 runtime.json，pid={stale.get('pid')}，health 不通）")

    # DSH 插件：装了没有 / 加载了没有（加载只能在 DSH 启动时发生）
    lines.append("-" * 52)
    try:
        from .dsh_bundle import status as dsh_status

        dsh = dsh_status()
        lines.append(f"DSH 插件    : {dsh['verdict']}")
        for item in dsh["profiles"]:
            name = Path(item["profile"]).name
            mark = "就绪" if item.get("ready") else "未就绪"
            lines.append(
                f"  profile {name}: {mark}"
                f"（列为 bundle={item.get('listed')} 已链接={item.get('linked')} patch={item.get('patch')}）"
            )
        for proc in dsh["processes"]:
            lines.append(f"  已加载的 MCP 进程：pid={proc.get('ProcessId')}")
    except Exception as exc:  # noqa: BLE001 - 诊断不该自己炸
        lines.append(f"DSH 插件    : 检查失败 {type(exc).__name__}: {exc}")

    # Office 加载项：两条路各自到哪一步了
    try:
        from .addin import powerpoint_running, status as vba_status, vbom_enabled
        from .addin_https import cert_status, trust_command

        report = vba_status()
        lines.append(f"加载项(.ppam) : {report['verdict']}")
        if not vbom_enabled():
            where = "（PowerPoint 正开着，要先关掉）" if powerpoint_running() else ""
            lines.append(f"               开启 VBA 工程信任即可自动构建 {where}：pptctl addin-trust-vba --enable")
        certs = cert_status(cfg)
        if certs["ready"]:
            trust = "已受信任" if certs["trusted"] else "**CA 未装进受信任根**"
            lines.append(
                f"加载项(窗格)  : HTTPS 就绪，{trust}，证书还剩 {certs['days_left']} 天"
            )
            if not certs["trusted"]:
                lines.append(f"               {trust_command(cfg)}")
        else:
            lines.append("加载项(窗格)  : 未启用（可选）—— 跑 `pptctl addin-https` 生成证书")
    except Exception as exc:  # noqa: BLE001 - 诊断不该自己炸
        lines.append(f"加载项        : 检查失败 {type(exc).__name__}: {exc}")

    # agent 技能：装在哪、和仓库一不一致
    try:
        from .skill import status as skill_status

        sk = skill_status()
        lines.append(f"agent 技能  : {sk['verdict']}")
        if sk.get("installed"):
            lines.append(f"  {sk['installed']}")
    except Exception as exc:  # noqa: BLE001 - 诊断不该自己炸
        lines.append(f"agent 技能  : 检查失败 {type(exc).__name__}: {exc}")

    # 接口文档是否已过期
    try:
        from .apidocs import check_api_docs

        stale = check_api_docs()
        lines.append("接口文档    : " + ("最新" if not stale else "已过期 —— 重跑 `pptctl export-api`：" + stale[0]))
    except Exception as exc:  # noqa: BLE001
        lines.append(f"接口文档    : 检查失败 {type(exc).__name__}: {exc}")

    # 自检也走唯一守护进程，不能再创建一条 COM 通道。
    lines.append("-" * 52)
    verdict = "失败"
    if interp.get("pywin32"):
        try:
            from .mcp_server import ForwardingCaller
            info = ensure(cfg, timeout=40)
            response = ForwardingCaller(cfg, info).call("ppt_status", {})
            payload = response.get("result") or {}
            com = payload.get("com") or {}
            if response.get("ok") and com.get("alive") and not payload.get("com_error"):
                verdict = "可用"
                lines.append(f"COM 实连    : 成功，PowerPoint {com.get('version')}（经守护进程）")
                lines.append(f"              窗口可见 = {com.get('visible')}（可见性是硬要求）")
                lines.append(f"当前打开的演示：{len(payload.get('presentations') or [])} 个")
            else:
                error = response.get("error") or {}
                lines.append(f"COM 实连    : 失败 —— {payload.get('com_error') or error.get('message') or com.get('last_error')}")
        except Exception as exc:
            lines.append(f"COM 实连    : 失败 —— {exc}")
    else:
        lines.append("COM 实连    : 跳过（pywin32 不可用）")

    lines.append("=" * 52)
    lines.append(f"结论：{verdict}")
    if verdict == "可用" and not info:
        lines.append("下一步：pptctl serve     # 启动守护进程，然后打开观察台")
    elif verdict == "可用":
        lines.append("下一步：pptctl url --launch    # 打开观察台看 agent 操作")
    else:
        lines.append("下一步：确认 PowerPoint 已安装，且 pptctl 跑在项目 venv 的解释器里")
    print("\n".join(lines))
    return 0 if verdict == "可用" else 1


def cmd_serve(args: argparse.Namespace) -> int:
    return serve()


def cmd_status(args: argparse.Namespace) -> int:
    if args.no_spawn:
        info = runtime_info(default_settings)
        if info is None:
            _dump({"ok": False, "error": "守护进程未运行"})
            return 1
    else:
        info = ensure(default_settings)
    _dump({"ok": True, "runtime": info, "observer": observer_url(default_settings), "health": health()})
    return 0


def cmd_tools(args: argparse.Namespace) -> int:
    import pptd.ops  # noqa: F401  导入即注册

    if args.json:
        _dump([t.to_manifest() for t in all_tools()])
        return 0
    print(f"共 {len(all_tools())} 个工具\n")
    for t in all_tools():
        print(f"  {t.name:<16} {t.behavior:<11} {t.summary}")
        for p in t.params:
            flag = "*" if p.required else " "
            print(f"      {flag} {p.name} ({p.type})  {p.description}")
    return 0


def cmd_call(args: argparse.Namespace) -> int:
    payload: dict[str, Any] = {}
    if args.json:
        try:
            payload = json.loads(args.json)
        except ValueError as exc:
            raise SystemExit(f"--json 不是合法 JSON：{exc}") from exc
    payload.update(_kv_pairs(args.arg))

    info = ensure(default_settings)
    import httpx

    with httpx.Client(timeout=max(30.0, default_settings.call_timeout + 15)) as client:
        resp = client.post(
            f"http://{default_settings.host}:{default_settings.port}/call",
            json={"name": args.name, "args": payload},
            headers={"X-PPT-Token": str(info.get("token", ""))},
        )
    _dump(resp.json())
    return 0 if resp.status_code == 200 else 1


def cmd_open(args: argparse.Namespace) -> int:
    ns = argparse.Namespace(name="ppt_open", json=None, arg=[f"path={args.path}"])
    return cmd_call(ns)


def cmd_url(args: argparse.Namespace) -> int:
    info = runtime_info(default_settings)
    if info is None:
        if args.no_spawn:
            raise SystemExit("守护进程未运行")
        info = ensure(default_settings)
    url = observer_url(default_settings)
    print(f"观察台   : {url}")
    print(f"MCP HTTP : {default_settings.mcp_url}   （需要令牌，见 --mcp-token）")
    if args.mcp_token and info:
        print(f"           token={info.get('token', '')}")
    if args.launch and url:
        import webbrowser

        webbrowser.open(url)
    return 0


def cmd_stop(args: argparse.Namespace) -> int:
    result = stop(default_settings, force=args.force)
    _dump(result)
    if not result.get("stopped"):
        print("提示：COM 卡死时用 pptctl stop --force", file=sys.stderr)
        return 1
    return 0


def cmd_addin_https(args: argparse.Namespace) -> int:
    """为 Office 加载项准备 HTTPS 证书。

    **生成证书是程序的事；装进受信任根是用户的决定**，所以 `--trust` 必须显式指定。
    """
    from .addin_https import cert_status, generate_certs, trust_command, trust_hint, write_manifest

    result = generate_certs(force=args.force)
    _dump(result)

    if args.write_manifest or not Path("addin/web/manifest.xml").is_file():
        print(f"\n已写出清单：{write_manifest()}")
    else:
        print(f"\n清单已存在：{write_manifest()}")

    if args.trust:
        if result["trusted"]:
            print("\nCA 已经在受信任根里，无需重复安装。")
        else:
            import subprocess

            command = trust_command()
            print(f"\n执行：{command}")
            proc = subprocess.run(command, shell=True, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace")
            print((proc.stdout or "").strip()[-500:])
            after = cert_status()
            print(f"\n结果：{'已信任 ✓' if after['trusted'] else '仍未信任 ✗'}")
            return 0 if after["trusted"] else 1
    else:
        print()
        if result["trusted"]:
            print("CA 已受信任 ✓ —— 可以直接旁加载清单了")
        else:
            print(trust_hint())
    return 0


def cmd_addin_trust_vba(args: argparse.Namespace) -> int:
    """一次性开关：「信任对 VBA 工程对象模型的访问」。

    必须显式 ``--enable`` / ``--disable``：这是一个**有安全代价**的设置
    （开启后任何程序都能读写你文档里的 VBA 代码，宏病毒正是靠它扩散），
    微软文档里的原话是 "don't try to force/sneak it through behind their back"。
    """
    from .addin import powerpoint_running, set_vbom, vbom_enabled, vbom_hint, vbom_key

    want = None if args.enable == args.disable else args.enable  # 两个都没给或都给了 -> 只看状态
    current = vbom_enabled()
    print(f"注册表键 : HKCU\\{vbom_key()}\\AccessVBOM")
    print(f"当前状态 : {'已开启' if current else '未开启'}")

    if want is None:
        if current:
            print("\n已经开着，`pptctl addin-build` 可以直接用。")
        else:
            print()
            print(vbom_hint())
        return 0

    if want == current:
        print(f"\n已经是{'开启' if want else '关闭'}状态，无需改动。")
        return 0

    # 关键前置条件：PowerPoint 运行时改这个键**无效**，退出时会被覆盖回去
    if powerpoint_running():
        print()
        print("✗ PowerPoint 正在运行，**这次改动不会生效**。")
        print("  微软文档明确写着：这个设置只在 PowerPoint 启动时读取，运行期间用任何方式")
        print("  （宏、脚本、手改注册表）改它，退出时都会被丢弃。")
        print()
        print("  请这样做：")
        print("    1) 关掉 PowerPoint（先存好你的文件）")
        print(f"    2) pptctl addin-trust-vba {'--enable' if want else '--disable'}")
        print("    3) 重新打开 PowerPoint，再跑 pptctl addin-build")
        return 1

    result = set_vbom(want)
    if not result.get("ok"):
        print(f"\n✗ {result.get('reason')}")
        return 1
    print(f"\n✓ 已{'开启' if want else '关闭'}（写入 {result['key']}）")
    if want:
        print()
        print("提醒：这个开关允许**任何**程序读写你文档里的 VBA 代码——宏病毒正是靠它扩散的。")
        print("构建完 .ppam 之后可以随时关掉：pptctl addin-trust-vba --disable")
    print("\n下一步：pptctl addin-build")
    return 0


def cmd_skill_status(args: argparse.Namespace) -> int:
    """报告 agent 技能装没装、和仓库是否一致。"""
    from .skill import status as skill_status

    report = skill_status()
    _dump(report)
    return 0 if report["frontmatter_ok"] else 1


def cmd_skill_install(args: argparse.Namespace) -> int:
    """把技能装进 DSH 的用户技能目录。"""
    from .skill import install

    out = Path(args.out).expanduser() if args.out else None
    result = install(out)
    _dump(result)
    return 0 if result.get("ok") else 1


def cmd_export_api(args: argparse.Namespace) -> int:
    """导出接口文档：OpenAPI 契约、工具参考、各宿主接入配置。"""
    from .apidocs import check_api_docs, write_api_docs

    out = Path(args.out).expanduser() if args.out else None
    if args.check:
        stale = check_api_docs(out, local_paths=args.local_paths)
        if stale:
            print("接口文档已过期：")
            for item in stale:
                print(f"  - {item}")
            return 1
        print("接口文档是最新的 ✓")
        return 0

    written = write_api_docs(out, local_paths=args.local_paths)
    for path in written:
        print(f"已写出：{path}（{path.stat().st_size} 字节）")
    return 0


def cmd_addin_status(args: argparse.Namespace) -> int:
    """报告 PowerPoint 加载项的状态。"""
    from .addin import status as addin_status

    report = addin_status()
    _dump(report)
    return 0 if report["frm_ok"] else 1


def cmd_addin_export(args: argparse.Namespace) -> int:
    """把 VBA 源码导出成 VBE 能读的编码（GBK）。"""
    from .addin import export_vba

    out = Path(args.out).expanduser() if args.out else None
    written = export_vba(out)
    print("已导出（GBK，供 VBE 导入）：")
    for path in written:
        print(f"  {path}")
    print()
    from .addin import SOURCE_DIR
    print(f"导入步骤见：{SOURCE_DIR.parent / 'README.md'}")
    return 0


def cmd_addin_build(args: argparse.Namespace) -> int:
    """用 PowerPoint 自己把 VBA 源码编译成 .ppam（需要 VBA 工程信任）。"""
    from .addin import build_ppam

    out = Path(args.out).expanduser() if args.out else None
    result = build_ppam(out)
    _dump(result)
    return 0 if result.get("ok") else 1


def cmd_dsh_status(args: argparse.Namespace) -> int:
    """只报 DSH 插件的状态：装好没有、加载没有。"""
    from .dsh_bundle import status as dsh_status

    report = dsh_status()
    _dump(report)
    return 0 if report.get("loaded") else (0 if report.get("any_ready") else 1)


def cmd_mcp(args: argparse.Namespace) -> int:
    from .mcp_stdio import main as mcp_main

    return mcp_main()


def cmd_dsh_bundle(args: argparse.Namespace) -> int:
    """生成 DSH 配置型 bundle（无宿主代码，只挂 MCP 客户端）。"""
    from .dsh_bundle import install_hint, write_bundle

    out = Path(args.out).expanduser() if args.out else None
    written = write_bundle(out=out)
    for path in written.values():
        print(f"已写出：{path}")
    print()
    print(install_hint(out=out))
    if args.show:
        print("\n--- cordis.patch.yml ---")
        print(written["patch"].read_text(encoding="utf-8"))
    return 0


def cmd_restart(args: argparse.Namespace) -> int:
    """卡死后的标准收场：强杀 → 重新拉起。"""
    stop(default_settings, force=True)
    time.sleep(0.5)
    spawn(default_settings)
    info = wait_ready(default_settings)
    # 注意：venv 的 python.exe 在 Windows 上可能是个再执行的壳，spawn 返回的 pid
    # 不一定是真正的守护进程，所以以 runtime.json 里的 pid 为准。
    _dump({"ok": True, "pid": info.get("pid"), "runtime": info, "observer": observer_url(default_settings)})
    return 0


# --------------------------------------------------------------------------
# 入口
# --------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptctl",
        description="让 agent 实时控制 PowerPoint（观察台 + MCP + HTTP 三面同源）",
    )
    parser.add_argument("--version", action="version", version=f"ppt-agent {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("doctor", help="自检解释器、PowerPoint、守护进程与 COM 实连")
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("serve", help="在前台运行守护进程（阻塞）")
    p.set_defaults(func=cmd_serve)

    p = sub.add_parser("status", help="查看守护进程与 PowerPoint 状态")
    p.add_argument("--no-spawn", action="store_true", help="不要自动拉起守护进程")
    p.set_defaults(func=cmd_status)

    p = sub.add_parser("tools", help="列出全部工具及其参数")
    p.add_argument("--json", action="store_true", help="输出完整 JSON Schema")
    p.set_defaults(func=cmd_tools)

    p = sub.add_parser("call", help="调用任意工具")
    p.add_argument("name", help="工具名，例如 ppt_open")
    p.add_argument("--json", help="以 JSON 传入参数")
    p.add_argument("-a", "--arg", action="append", help="key=value，可重复")
    p.set_defaults(func=cmd_call)

    p = sub.add_parser("open", help="打开一个演示并让它显示出来")
    p.add_argument("path", help="演示文件绝对路径")
    p.set_defaults(func=cmd_open)

    p = sub.add_parser("url", help="打印观察台与 MCP 地址")
    p.add_argument("--launch", action="store_true", help="顺便在浏览器里打开观察台")
    p.add_argument("--no-spawn", action="store_true", help="守护进程没跑就报错，不自动拉起")
    p.add_argument("--mcp-token", action="store_true", help="顺便打印 MCP over HTTP 需要的令牌")
    p.set_defaults(func=cmd_url)

    p = sub.add_parser("stop", help="停止守护进程")
    p.add_argument("--force", action="store_true", help="强杀（COM 卡死时用）")
    p.set_defaults(func=cmd_stop)

    p = sub.add_parser("restart", help="强杀并重新拉起守护进程")
    p.set_defaults(func=cmd_restart)

    p = sub.add_parser("mcp", help="以 stdio 方式运行 MCP 服务器（供 DSH 等宿主调用）")
    p.set_defaults(func=cmd_mcp)

    p = sub.add_parser("dsh-bundle", help="生成 DSH 插件 bundle（配置型，挂 MCP 客户端）")
    p.add_argument("--out", help="输出目录（默认仓库下的 dsh-bundle/）")
    p.add_argument("--show", action="store_true", help="顺便打印生成的 patch 内容")
    p.set_defaults(func=cmd_dsh_bundle)

    p = sub.add_parser("dsh-status", help="检查 DSH 插件是否已装、是否已被加载")
    p.set_defaults(func=cmd_dsh_status)

    p = sub.add_parser("addin-status", help="检查 PowerPoint 加载项（.ppam）的状态")
    p.set_defaults(func=cmd_addin_status)

    p = sub.add_parser("export-api", help="导出接口文档（OpenAPI / 工具参考 / 接入指南）")
    p.add_argument("--out", help="输出目录（默认 docs/api/）")
    p.add_argument("--check", action="store_true", help="只检查文档是否已过期，不写文件")
    p.add_argument("--local-paths", action="store_true", help="接入配置填写本机路径（默认使用公开占位路径）")
    p.set_defaults(func=cmd_export_api)

    p = sub.add_parser("addin-https", help="为 Office 加载项准备 HTTPS 证书与清单")
    p.add_argument("--force", action="store_true", help="重新生成证书（覆盖已有的）")
    p.add_argument("--trust", action="store_true",
                   help="把生成的 CA 装进当前用户的受信任根（需要你显式指定）")
    p.add_argument("--write-manifest", action="store_true", help="重新生成加载项清单 XML")
    p.set_defaults(func=cmd_addin_https)

    p = sub.add_parser("addin-trust-vba",
                       help="开关「信任对 VBA 工程对象模型的访问」（构建 .ppam 需要）")
    p.add_argument("--enable", action="store_true", help="开启（需要先关掉 PowerPoint）")
    p.add_argument("--disable", action="store_true", help="关闭（建议构建完就关掉）")
    p.set_defaults(func=cmd_addin_trust_vba)

    p = sub.add_parser("skill-status", help="检查 agent 技能是否已安装、是否与仓库一致")
    p.set_defaults(func=cmd_skill_status)

    p = sub.add_parser("skill-install", help="把 agent 技能装进 ~/.dsh/skills/")
    p.add_argument("--out", help="安装到指定目录（默认 DSH 的用户技能目录）")
    p.set_defaults(func=cmd_skill_install)

    p = sub.add_parser("addin-export", help="导出 VBA 源码（GBK），供 VBE 手动导入")
    p.add_argument("--out", help="输出目录（默认 addin/export/）")
    p.set_defaults(func=cmd_addin_export)

    p = sub.add_parser("addin-build", help="自动构建 .ppam（需要开启 VBA 工程信任）")
    p.add_argument("--out", help="输出文件（默认 addin/ppt-agent.ppam）")
    p.set_defaults(func=cmd_addin_build)

    return parser


def main(argv: list[str] | None = None) -> int:
    _force_utf8()
    parser = build_parser()
    args = parser.parse_args(argv)
    ensure_dirs(default_settings)
    return int(args.func(args))


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
