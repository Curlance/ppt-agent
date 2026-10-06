"""把接口导出成文档：OpenAPI 契约、工具参考、各宿主的接入配置。

为什么要生成而不是手写
----------------------
工具是模型直接看的契约。手写的文档一定会和代码漂移——加了工具忘了改文档，
模型就会照着旧 schema 调用。所以这里全部**从注册表和真实环境生成**，
并且带一个漂移检查（``pptctl export-api --check``），文档过期就让 CI/自检失败。

生成过程**不启动 COM、不起服务**：只构造 FastAPI 应用取它的 ``openapi()``，
以及读注册表。所以它可以在任何机器上安全地跑。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from . import __version__
from .config import Settings, host_interpreter, settings as default_settings
from .registry import Tool, all_tools

__all__ = [
    "docs_dir",
    "build_openapi",
    "render_tools_markdown",
    "render_connect_markdown",
    "write_api_docs",
    "check_api_docs",
]

BEHAVIOR_LABELS = {
    "read": "看（只读）",
    "write": "做（写入）",
    "destroy": "删（破坏性，动手前先备份）",
    "idempotent": "切（幂等，重复执行结果相同）",
}
BEHAVIOR_ORDER = ("read", "idempotent", "write", "destroy")


def docs_dir(cfg: Settings | None = None) -> Path:
    return Path(__file__).resolve().parent.parent / "docs" / "api"


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


# --------------------------------------------------------------------------
# OpenAPI
# --------------------------------------------------------------------------


def build_openapi(cfg: Settings | None = None) -> dict[str, Any]:
    """构造应用并取出 OpenAPI 契约。不起服务、不碰 COM。"""
    import pptd.ops  # noqa: F401  导入即注册全部工具
    from .engine import Engine
    from .http_api import create_app

    engine = Engine.create(cfg or default_settings, token="docs")
    app = create_app(engine)
    return app.openapi()


# --------------------------------------------------------------------------
# 工具参考
# --------------------------------------------------------------------------


def _escape_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ").strip()


def render_tools_markdown() -> str:
    """按行为分组渲染全部工具。顺序即注册顺序（稳定，便于 diff）。"""
    import pptd.ops  # noqa: F401

    tools = all_tools()
    lines = [
        "# ppt-agent 工具参考",
        "",
        "> 本文件由 `pptctl export-api` 生成，**请勿手改**。",
        "> 改了代码就重跑生成，`pptctl export-api --check` 会在文档过期时报错。",
        "",
        f"共 **{len(tools)} 个工具**。同一份注册表同时暴露为 MCP 与 HTTP/OpenAPI，",
        "所以这里的 schema 与 `openapi.json`、以及 MCP `tools/list` 给出的**完全一致**。",
        "",
        "每个工具在 HTTP 上都有一个同名端点：`POST /tools/<工具名>`（需要令牌），",
        "也可以用 `POST /call` 带 `{\"name\": ..., \"args\": {...}}` 统一调用。",
        "",
    ]

    for behavior in BEHAVIOR_ORDER:
        group = [t for t in tools if t.behavior == behavior]
        if not group:
            continue
        lines += [f"## {BEHAVIOR_LABELS[behavior]}", ""]
        for t in group:
            lines += _render_tool(t)
    return "\n".join(lines).rstrip() + "\n"


def _render_tool(t: Tool) -> list[str]:
    out = [f"### `{t.name}` — {t.title}", "", t.summary, "", f"- 返回：{t.returns}"]
    if t.needs_com:
        out.append("- 依赖：需要 PowerPoint（走 COM）")
    else:
        out.append("- 依赖：不需要 PowerPoint（COM 挂掉也能用）")
    if t.tags:
        out.append(f"- 标签：{', '.join(t.tags)}")
    out.append("")

    if not t.params:
        out += ["*无参数。*", ""]
        return out

    out += ["| 参数 | 类型 | 必填 | 默认 | 说明 |", "|---|---|---|---|---|"]
    for p in t.params:
        required = "是" if p.required else ""
        default = "" if p.default is None else f"`{json.dumps(p.default, ensure_ascii=False)}`"
        desc = _escape_cell(p.description)
        if p.enum:
            desc += f"（可选：{' / '.join(p.enum)}）"
        if p.minimum is not None or p.maximum is not None:
            lo = "" if p.minimum is None else f"{p.minimum:g}"
            hi = "" if p.maximum is None else f"{p.maximum:g}"
            desc += f"（范围 {lo}..{hi}）"
        out.append(f"| `{p.name}` | {p.type} | {required} | {default} | {desc} |")
    out.append("")
    return out


# --------------------------------------------------------------------------
# 接入指南（含本机真实路径）
# --------------------------------------------------------------------------


def render_connect_markdown(cfg: Settings | None = None, *, portable: bool = False) -> str:
    """生成接入配置；公开文档使用占位路径，本机导出可填实际路径。"""
    s = cfg or default_settings
    if portable:
        exe, extra_env = "C:/Python313/python.exe", {
            "PYTHONPATH": "C:/path/to/ppt-agent/.venv/Lib/site-packages"
        }
        repo = Path("C:/path/to/ppt-agent")
    else:
        exe, extra_env = host_interpreter()
        repo = _repo_root()
    path_note = (
        "路径为示例占位值，请替换为自己的解释器和仓库目录。\n"
        "> 可运行 `pptctl export-api --local-paths --out local-docs` 生成本机配置。"
        if portable else "路径是本机的真实值，分享前请检查其中的个人信息。"
    )

    # 用字典构造再 dump，保证生成的是**合法 JSON**（手拼字符串容易漏键或多键）
    stdio_server: dict[str, Any] = {
        "command": exe,
        "args": ["-X", "utf8", "-m", "pptd", "mcp"],
        "cwd": str(repo),
    }
    if extra_env.get("PYTHONPATH"):
        stdio_server["env"] = {"PYTHONPATH": extra_env["PYTHONPATH"]}
    claude_config = (
        "```json\n"
        + json.dumps({"mcpServers": {"ppt": stdio_server}}, ensure_ascii=False, indent=2)
        + "\n```\n"
    )
    http_config = (
        "```json\n"
        + json.dumps(
            {"mcpServers": {"ppt": {"url": s.mcp_url, "headers": {"Authorization": "Bearer <token>"}}}},
            ensure_ascii=False,
            indent=2,
        )
        + "\n```\n"
    )

    return f"""# ppt-agent 接入指南

> 本文件由 `pptctl export-api` 生成；{path_note}

三个面，同一份工具注册表：

| 面 | 地址 | 给谁 |
|---|---|---|
| HTTP + OpenAPI | `http://{s.host}:{s.port}` | 人、ChatGPT Actions、加载项、脚本 |
| MCP over HTTP | `{s.mcp_url}` | DSH / ChatGPT 的 MCP 连接器 / 任何 MCP 宿主 |
| MCP over stdio | `{exe} -X utf8 -m pptd mcp` | 本机宿主（DSH 用配置型 bundle 挂它） |

契约文件：[`openapi.json`](openapi.json) · 工具参考：[`tools.md`](tools.md)

---

## 鉴权

**工具调用与状态、截图接口需要令牌。** `GET /health` 免令牌；
少数展示入口（如 `/strip.txt`、`/addin/`）仅允许本机 Host 免令牌访问。
令牌在守护进程的
`%LOCALAPPDATA%\\ppt-agent\\runtime.json` 里，用 `pptctl url --mcp-token` 可以打印。

三种传法任选：

```
X-PPT-Token: <token>
Authorization: Bearer <token>
?token=<token>
```

> **这个令牌不是可选项。** 一旦把端口隧道到公网而没有鉴权，任何人都能驱动你的
> PowerPoint——能读你打开的演示、也能改它。

---

## ① DSH（本机，stdio）

```powershell
pptctl dsh-bundle      # 生成 dsh-bundle/（配置型 bundle，无宿主代码）
pptctl dsh-status      # 检查装没装、加载没加载
```

DSH 自带 `@deepseek-ai/dsh-mcp-client`，所以只需在 profile 的 `dsh.profile.bundles`
里加一行、在 `dependencies` 里加一个 `link:`，再 `pnpm install`。
**改完必须重启 DSH**——工具列表在启动时组装。工具以 `mcp__ppt__<工具名>` 出现。

---

## ② ChatGPT

ChatGPT **访问不到 `127.0.0.1`**，所以无论走哪条路，都需要一个公网 HTTPS 端点。
使用远程客户端时，需要自己配置隧道或反向代理，并确认客户端支持所选协议与鉴权。

### 路线 A：支持 MCP 的远程客户端

1. 把 `{s.mcp_url}` 暴露到公网 HTTPS（隧道或反向代理）
2. 在 ChatGPT 里添加 MCP 连接器，URL 填那个公网地址
3. 让请求带上令牌（`Authorization: Bearer <token>`，或在 URL 上加 `?token=<token>`）

### 路线 B：Custom GPT Actions（只需要 OpenAPI）

1. 下载 [`openapi.json`](openapi.json)
2. 把里面的 `servers` 改成你的公网 HTTPS 地址
3. 在 GPT 的 Actions 里导入这份 schema
4. 认证方式选 API Key，Header 名填 `X-PPT-Token`

> 两条路的工具是同一批：路线 A 走 MCP 协议，路线 B 走 `/tools/<名字>` 这些 REST 端点。

### 隧道怎么起（示例）

```powershell
# Cloudflare Tunnel（需自行安装 cloudflared）
cloudflared tunnel --url http://127.0.0.1:{s.port}      # 给 Actions 用
cloudflared tunnel --url http://127.0.0.1:{s.mcp_bind_port}   # 给 MCP 连接器用
```

**务必只在需要时开隧道，用完就关。**

---

## ③ 其它 MCP 宿主（Claude Desktop / Cline / 自研）

stdio 一处配置，处处可用：

{claude_config}
HTTP 传输的等价配置：

{http_config}
---

## ④ 命令行 / 脚本

```powershell
# 通用调用
curl.exe -H "X-PPT-Token: <token>" -H "Content-Type: application/json" `
  --data-binary "@body.json" http://{s.host}:{s.port}/tools/ppt_status

# 面板数据（纯文本，一行一条）
curl.exe -H "X-PPT-Token: <token>" "http://{s.host}:{s.port}/feed.txt?limit=40"

# 实时取景（读 PowerPoint 当前状态，覆盖同一个文件）
curl.exe -H "X-PPT-Token: <token>" "http://{s.host}:{s.port}/live?width=960" -o live.jpg

# 用自带 CLI
pptctl tools                     # 看有哪些工具
pptctl call ppt_shot -a slide=3  # 截第 3 页
```

---

## 端点一览

| 方法 | 路径 | 作用 |
|---|---|---|
| GET | `/health` | 存活探针（免令牌） |
| GET | `/state` | 守护进程与 PowerPoint 状态 |
| GET | `/tools` | 全部工具及其 JSON Schema |
| POST | `/call` | 按名字调用任意工具 |
| POST | `/tools/<名字>` | 每个工具一个端点，供 OpenAPI 客户端使用 |
| GET | `/events` | SSE 事件流（网页用） |
| GET | `/feed` | 面板快照：状态 + 最近事件（JSON） |
| GET | `/feed.txt` | 同上，纯文本（VBA / PowerShell 用） |
| GET | `/strip.txt` | **一行**状态（宿主界面里的窄条用，本机 Host 免令牌） |
| GET | `/live` | 实时取景：现拍当前页 JPEG |
| GET | `/shots/<文件名>` | 取截图文件 |
| GET | `/addin/` | Office 加载项任务窗格页面（本机 Host 免令牌） |
| GET | `/addin/manifest.xml` | Office 加载项清单（按当前端口生成） |
| POST | `/pace` | 设置跟速 |
| POST | `/shutdown` | 急停（停止守护进程，不关 PowerPoint） |
| GET | `/docs` | 交互式 API 文档 |
| GET | `/openapi.json` | OpenAPI 契约 |

版本：ppt-agent {__version__}
"""


# --------------------------------------------------------------------------
# 落盘与漂移检查
# --------------------------------------------------------------------------


def render_all(cfg: Settings | None = None, *, local_paths: bool = False) -> dict[str, str]:
    """返回 ``文件名 -> 内容``。写入与检查共用这一份，保证两者不会不一致。"""
    spec = build_openapi(cfg)
    return {
        "openapi.json": json.dumps(spec, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        "tools.md": render_tools_markdown(),
        "CONNECT.md": render_connect_markdown(cfg, portable=not local_paths),
    }


def write_api_docs(out: Path | None = None, cfg: Settings | None = None, *, local_paths: bool = False) -> list[Path]:
    target = out or docs_dir(cfg)
    target.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for name, content in render_all(cfg, local_paths=local_paths).items():
        path = target / name
        path.write_text(content, encoding="utf-8")
        written.append(path)
    return written


def check_api_docs(out: Path | None = None, cfg: Settings | None = None, *, local_paths: bool = False) -> list[str]:
    """漂移检查：返回"过期了"的文件说明；空列表表示都是最新的。"""
    target = out or docs_dir(cfg)
    stale: list[str] = []
    for name, content in render_all(cfg, local_paths=local_paths).items():
        path = target / name
        if not path.is_file():
            stale.append(f"{name}：还没有生成")
            continue
        if path.read_text(encoding="utf-8") != content:
            stale.append(f"{name}：内容已过期（重跑 pptctl export-api）")
    return stale
