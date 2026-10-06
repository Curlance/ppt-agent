"""生成 DSH 配置型 bundle。

DSH 自带 ``@deepseek-ai/dsh-mcp-client``，所以"DSH 插件"不需要写 JavaScript——
只要一个 ``package.json``（声明 ``dsh.bundle.patch``）加一份 ``cordis.patch.yml``，
把我们的 stdio MCP 服务器挂上去，工具就会以 ``mcp__ppt__<工具名>`` 出现在 DSH 里。

这里由 CLI 生成而**不是**手写死路径，原因有二：
1. 解释器必须用 ``host_interpreter()`` 返回的基础解释器（venv 的 python.exe 是再执行
   的壳，宿主杀掉壳会留下孤儿并占住 stdio）；
2. 解释器与仓库路径随机器而变，生成比手改可靠。
"""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path
from typing import Any

from . import __version__
from .config import Settings, host_interpreter, settings as default_settings

__all__ = [
    "bundle_dir",
    "render_package_json",
    "render_index_js",
    "render_client_js",
    "render_patch",
    "write_bundle",
    "install_hint",
    "discover_profiles",
    "profile_status",
    "mcp_processes",
    "status",
]

SERVER_NAME = "ppt"
BUNDLE_ID = "ppt-agent-mcp"
PACKAGE_NAME = "@local/ppt-agent-mcp"

#: 守护进程侧的 COM 调用超时（秒）。留给大文件与导出用。
CALL_TIMEOUT_S = 120
#: MCP 客户端侧的调用超时（毫秒）。**必须大于守护进程的**，这样超时先由我们
#: 自己报出「可能弹了模态框」这种可诊断的原因，而不是被客户端一刀切断。
TOOL_CALL_TIMEOUT_MS = (CALL_TIMEOUT_S + 30) * 1000


def bundle_dir(cfg: Settings | None = None) -> Path:
    repo = Path(__file__).resolve().parent.parent
    return repo / "dsh-bundle"


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def render_package_json() -> str:
    payload: dict[str, Any] = {
        "name": PACKAGE_NAME,
        "version": __version__,
        "private": True,
        "type": "module",
        "description": "ppt-agent：把 PowerPoint 控制台挂进 DSH（MCP 工具 + 输入框下方的一行状态条）",
        # 宿主半边（MCP 客户端那行由 patch 插入）+ 浏览器半边（状态条）
        "exports": {".": "./index.js", "./client": "./client.js"},
        "dsh": {
            "bundle": {"patch": "./cordis.patch.yml"},
            "client": {
                "platform": "web",
                "immediately": True,
                "inject": ["@deepseek-ai/dsh-client-ui-conversation"],
            },
        },
    }
    return json.dumps(payload, ensure_ascii=False, indent=2) + "\n"


def render_index_js() -> str:
    """宿主半边。

    工具全部由 ``@deepseek-ai/dsh-mcp-client`` 提供（见 patch 的第一行），
    这个模块只负责**让浏览器半边被装载**——客户端的入口由 ``dsh.client`` 声明。
    所以它是空的，这也意味着它不可能在宿主启动时搞出副作用。
    """
    return (
        "/** ppt-agent 插件的宿主半边。\n"
        " *\n"
        " * 这里刻意什么都不做：\n"
        " * - 工具由 @deepseek-ai/dsh-mcp-client 提供（守护进程独占 COM，见 README）\n"
        " * - 界面全部在浏览器半边（client.js），由 package.json 的 dsh.client 装载\n"
        " *\n"
        " * 空的 apply 意味着这个插件在宿主启动阶段不会引入任何副作用或失败点。\n"
        " */\n"
        "export function apply() {}\n"
    )


def render_client_js(settings: Settings | None = None) -> str:
    """浏览器半边：输入框下方的一行状态条。

    为什么是"一行"而不是一个面板：DSH 的 ``conversation.composer.dock`` 是
    输入框下方**已分配空间的窄条**（渲染时传的是空 props，宽度受 dock 约束），
    DSH 的插件开发指引也明确要求"第一版保持在这个 slot 的流里"。
    硬塞一个高面板进去只会挤坏宿主布局。

    数据来自守护进程的 ``/strip.txt``——**不含任何密钥**，所以浏览器半边不需要令牌，
    也就不需要跟宿主做"取令牌"的握手（那会引入一堆我验证不了的契约）。
    """
    s = settings or default_settings
    endpoint = f"http://{s.host}:{s.port}/strip.txt"
    observer = f"http://{s.host}:{s.port}/"
    refresh = 3000
    return f"""window.__ModuleLoader__.load({{
  id: '{PACKAGE_NAME}',
  factory(require) {{
    const React = require('react');
    const h = React.createElement;

    const ENDPOINT = '{endpoint}';
    const OBSERVER = '{observer}';
    const REFRESH_MS = {refresh};

    function StatusStrip() {{
      const [text, setText] = React.useState('○ ppt-agent 正在连接…');
      const [down, setDown] = React.useState(false);

      React.useEffect(() => {{
        let alive = true;

        async function tick() {{
          try {{
            const resp = await fetch(ENDPOINT, {{ cache: 'no-store' }});
            if (!resp.ok) throw new Error('HTTP ' + resp.status);
            const line = (await resp.text()).trim();
            if (alive) {{ setText(line || '○ ppt-agent 无状态'); setDown(false); }}
          }} catch (err) {{
            if (alive) {{ setText('○ ppt-agent 未连接 —— 先运行 pptctl serve'); setDown(true); }}
          }}
        }}

        tick();
        const timer = setInterval(tick, REFRESH_MS);
        return () => {{ alive = false; clearInterval(timer); }};
      }}, []);

      return h(
        'a',
        {{
          href: OBSERVER,
          target: '_blank',
          rel: 'noreferrer',
          title: '在浏览器里打开 ppt-agent 观察台',
          style: {{
            display: 'block',
            width: '100%',
            boxSizing: 'border-box',
            padding: '2px 4px',
            fontSize: 12,
            lineHeight: '16px',
            textDecoration: 'none',
            // 继承宿主主题：连上了就用前景色，断了才变淡
            color: down ? 'var(--dsw-alias-text-3, #8b93a7)' : 'inherit',
            opacity: down ? 0.75 : 1,
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            cursor: 'pointer',
          }},
        }},
        text
      );
    }}

    return {{
      inject: ['slots'],
      apply(ctx) {{
        ctx.slots.inject('conversation.composer.dock', () => ctx.slots.register({{
          name: 'conversation.composer.dock',
          id: 'ppt-agent',
          order: 20,
        }}, StatusStrip));
      }},
    }};
  }},
}});
"""


def render_patch(cfg: Settings | None = None) -> str:
    """生成 cordis.patch.yml。Windows 路径用 YAML 单引号包住，反斜杠保持字面量。"""
    exe, extra_env = host_interpreter()
    repo = _repo_root()
    pythonpath = extra_env.get("PYTHONPATH", "")
    lines = [
        "# 由 `pptctl dsh-bundle` 生成，请勿手改；重跑该命令即可覆盖。",
        "# 两行各行其职：",
        "#   1) ppt-agent-mcp —— 挂 DSH 自带的 MCP 客户端，工具以 mcp__ppt__<工具名> 出现；",
        f"#   2) ppt-agent     —— 插件自身：宿主半边是空的（index.js），",
        "#                        浏览器半边在 client.js，给输入框下方加一行状态条。",
        "# 改动本文件后需要重启 harness 才会生效（契约在启动时组装）。",
        "- insert:",
        f"    - id: {BUNDLE_ID}",
        "      name: '@deepseek-ai/dsh-mcp-client'",
        "      config:",
        f"        serverName: {SERVER_NAME}",
        "        transport: stdio",
        f"        command: '{exe}'",
        "        args:",
        "          - '-X'",
        "          - 'utf8'",
        "          - '-m'",
        "          - 'pptd'",
        "          - 'mcp'",
        f"        cwd: '{repo}'",
    ]
    if pythonpath:
        lines += ["        env:", f"          PYTHONPATH: '{pythonpath}'"]
    lines += [
        f"        toolCallTimeoutMs: {TOOL_CALL_TIMEOUT_MS}",
        # 保持 schema 默认值 false（registrationFailure: "contain"）：万一 MCP 服务器
        # 起不来，只是工具缺失，不会让整个 harness 启动失败。诊断用 pptctl doctor。
        "        failOnStartupError: false",
        # 第二行：插件自身（宿主半边是空的，界面在浏览器半边）。
        # 少了它，输入框下方那条状态条就不会被装载。
        f"    - id: ppt-agent",
        f"      name: '{PACKAGE_NAME}'",
    ]
    return "\n".join(lines) + "\n"


def write_bundle(cfg: Settings | None = None, out: Path | None = None) -> dict[str, Path]:
    """把四个文件写到目标目录，返回写出的路径。

    - ``package.json`` / ``cordis.patch.yml``：插件清单与装载补丁
    - ``index.js`` / ``client.js``：宿主半边（空）与浏览器半边（状态条）
    """
    s = cfg or default_settings
    target = out or bundle_dir(s)
    target.mkdir(parents=True, exist_ok=True)
    files = {
        "package": (target / "package.json", render_package_json()),
        "patch": (target / "cordis.patch.yml", render_patch(s)),
        "index": (target / "index.js", render_index_js()),
        "client": (target / "client.js", render_client_js(s)),
    }
    written: dict[str, Path] = {}
    for key, (path, content) in files.items():
        path.write_text(content, encoding="utf-8")
        written[key] = path
    return written


def install_hint(cfg: Settings | None = None, out: Path | None = None) -> str:
    """给人和 agent 的安装提示。"""
    target = out or bundle_dir(cfg)
    return (
        f"bundle 已生成：{target}\n"
        "把这个目录作为本地包挂进当前 profile（两种方式任选）：\n"
        f"  1) 在 ~/.dsh/profiles/desktop/package.json 的 dsh.profile.bundles 里加上 \"{PACKAGE_NAME}\"，\n"
        f"     并在 dependencies 里加 \"{PACKAGE_NAME}\": \"link:{target}\"，然后在该目录执行 pnpm install。\n"
        "  2) 或让 DSH 的 plugin_manager 安装这个目录。\n\n"
        "装好后重启 DSH（契约在启动时组装），你会得到两样东西：\n"
        f"  · 工具：mcp__{SERVER_NAME}__ppt_status 等（先跑 pptctl serve 把守护进程起起来）\n"
        "  · 界面：输入框下方一行状态条，显示 PowerPoint 状态与最近一步操作\n\n"
        "改回配置：删掉 bundles 里那一行即可，不影响其它插件。"
    )


# --------------------------------------------------------------------------
# 自检：装了没有 / 加载了没有
# --------------------------------------------------------------------------


def dsh_home(cfg: Settings | None = None) -> Path:
    override = os.environ.get("DSH_HOME")
    if override:
        return Path(override).expanduser()
    return Path(os.path.expanduser("~")) / ".dsh"


def discover_profiles(home: Path | None = None) -> list[Path]:
    """找出所有 profile 目录（含 package.json 的那些）。"""
    root = (home or dsh_home()) / "profiles"
    if not root.is_dir():
        return []
    return sorted(p for p in root.iterdir() if (p / "package.json").is_file())


def profile_status(profile: Path) -> dict[str, Any]:
    """一个 profile 里我们的 bundle 装到什么程度了。"""
    status: dict[str, Any] = {
        "profile": str(profile),
        "listed": False,
        "linked": False,
        "patch": False,
        "client": False,
    }
    try:
        package = json.loads((profile / "package.json").read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        status["error"] = f"读不到 package.json：{exc}"
        return status

    bundles = (((package.get("dsh") or {}).get("profile") or {}).get("bundles")) or []
    deps = package.get("dependencies") or {}
    status["listed"] = PACKAGE_NAME in bundles
    status["declared"] = PACKAGE_NAME in deps
    status["dependency"] = deps.get(PACKAGE_NAME)

    linked = profile / "node_modules" / PACKAGE_NAME
    status["linked"] = linked.is_dir()
    status["linked_path"] = str(linked) if status["linked"] else None
    if status["linked"]:
        status["patch"] = (linked / "cordis.patch.yml").is_file()
        status["client"] = (linked / "client.js").is_file() and (linked / "index.js").is_file()
        status["generated_from"] = bundle_dir().as_posix()
    status["ready"] = bool(status["listed"] and status["linked"] and status["patch"])
    return status


def _alive_pids() -> set[int]:
    """当前活着的进程号集合。用来区分"宿主拉起的子进程"和"没人管的孤儿"。"""
    if os.name != "nt":
        return set()
    script = "Get-Process | Select-Object -ExpandProperty Id"
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=15.0,
        )
    except (OSError, subprocess.TimeoutExpired):
        return set()
    pids: set[int] = set()
    for line in (result.stdout or "").splitlines():
        line = line.strip()
        if line.isdigit():
            pids.add(int(line))
    return pids


def mcp_processes(timeout: float = 15.0) -> list[dict[str, Any]]:
    """找出正在运行的 ``pptd mcp`` 进程——这是"DSH 真的加载了"最硬的证据。

    DSH 启动时会把 mcp-client 那一行装配进去，随后拉起 stdio 子进程；
    没有这个进程，就说明 bundle 没被组合进来。

    每条记录额外带 ``orphan``：**父进程已经不在**的那种才是残留。
    这个区分是必需的——DSH 正常运行时本来就该有一个这样的进程，
    早期版本把它当成"孤儿"来断言，等插件真的被加载之后就误报了。
    """
    if os.name != "nt":
        return []
    script = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | "
        "Where-Object { $_.CommandLine -match 'pptd mcp' } | "
        "Select-Object ProcessId, ParentProcessId, CommandLine | ConvertTo-Json -Compress"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except (OSError, subprocess.TimeoutExpired):
        return []
    raw = (result.stdout or "").strip()
    if not raw:
        return []
    try:
        data = json.loads(raw)
    except ValueError:
        return []
    if isinstance(data, dict):
        data = [data]

    alive = _alive_pids()
    processes: list[dict[str, Any]] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        parent = int(item.get("ParentProcessId") or 0)
        item["orphan"] = parent not in alive if alive else False
        processes.append(item)
    return processes


def status(cfg: Settings | None = None) -> dict[str, Any]:
    """给 ``pptctl doctor`` 用的完整自检结果。"""
    profiles = discover_profiles()
    report: dict[str, Any] = {
        "dsh_home": str(dsh_home()),
        "profiles": [profile_status(p) for p in profiles],
        "processes": mcp_processes(),
    }
    report["any_ready"] = any(p.get("ready") for p in report["profiles"])
    report["loaded"] = bool(report["processes"])
    if report["loaded"]:
        report["verdict"] = "DSH 已经加载插件：工具以 mcp__ppt__* 出现"
    elif report["any_ready"]:
        report["verdict"] = "已装好但尚未加载 —— 重启 DSH 后生效（工具列表在启动时组装）"
    else:
        report["verdict"] = "尚未安装：先跑 `pptctl dsh-bundle` 再按提示挂进 profile"
    return report
