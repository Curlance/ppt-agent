# ppt-agent 接入指南

> 本文件由 `pptctl export-api` 生成；路径为示例占位值，请替换为自己的解释器和仓库目录。
> 可运行 `pptctl export-api --local-paths --out local-docs` 生成本机配置。

三个面，同一份工具注册表：

| 面 | 地址 | 给谁 |
|---|---|---|
| HTTP + OpenAPI | `http://127.0.0.1:8791` | 人、ChatGPT Actions、加载项、脚本 |
| MCP over HTTP | `http://127.0.0.1:8792/mcp` | DSH / ChatGPT 的 MCP 连接器 / 任何 MCP 宿主 |
| MCP over stdio | `C:/Python313/python.exe -X utf8 -m pptd mcp` | 本机宿主（DSH 用配置型 bundle 挂它） |

契约文件：[`openapi.json`](openapi.json) · 工具参考：[`tools.md`](tools.md)

---

## 鉴权

**工具调用与状态、截图接口需要令牌。** `GET /health` 免令牌；
少数展示入口（如 `/strip.txt`、`/addin/`）仅允许本机 Host 免令牌访问。
令牌在守护进程的
`%LOCALAPPDATA%\ppt-agent\runtime.json` 里，用 `pptctl url --mcp-token` 可以打印。

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

1. 把 `http://127.0.0.1:8792/mcp` 暴露到公网 HTTPS（隧道或反向代理）
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
cloudflared tunnel --url http://127.0.0.1:8791      # 给 Actions 用
cloudflared tunnel --url http://127.0.0.1:8792   # 给 MCP 连接器用
```

**务必只在需要时开隧道，用完就关。**

---

## ③ 其它 MCP 宿主（Claude Desktop / Cline / 自研）

stdio 一处配置，处处可用：

```json
{
  "mcpServers": {
    "ppt": {
      "command": "C:/Python313/python.exe",
      "args": [
        "-X",
        "utf8",
        "-m",
        "pptd",
        "mcp"
      ],
      "cwd": "C:\\path\\to\\ppt-agent",
      "env": {
        "PYTHONPATH": "C:/path/to/ppt-agent/.venv/Lib/site-packages"
      }
    }
  }
}
```

HTTP 传输的等价配置：

```json
{
  "mcpServers": {
    "ppt": {
      "url": "http://127.0.0.1:8792/mcp",
      "headers": {
        "Authorization": "Bearer <token>"
      }
    }
  }
}
```

---

## ④ 命令行 / 脚本

```powershell
# 通用调用
curl.exe -H "X-PPT-Token: <token>" -H "Content-Type: application/json" `
  --data-binary "@body.json" http://127.0.0.1:8791/tools/ppt_status

# 面板数据（纯文本，一行一条）
curl.exe -H "X-PPT-Token: <token>" "http://127.0.0.1:8791/feed.txt?limit=40"

# 实时取景（读 PowerPoint 当前状态，覆盖同一个文件）
curl.exe -H "X-PPT-Token: <token>" "http://127.0.0.1:8791/live?width=960" -o live.jpg

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

版本：ppt-agent 0.1.0
