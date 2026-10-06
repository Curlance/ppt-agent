"""MCP 面：同一份工具注册表，装配成 stdio 与 streamable-http 两种传输。

两种传输的区别只在"调用怎么落到守护进程"：

- **stdio**（``pptd mcp``）：进程间转发 —— MCP 进程是宿主拉起的子进程，
  它不碰 COM，只把调用转给守护进程。好处是宿主重启 MCP 不会丢 COM 会话。
- **streamable-http**（守护进程自己的第二个端口）：进程内直调 —— 已经在守护进程里了，
  再绕一圈 HTTP 没有意义，直接调 engine。

两者共用同一个 :func:`build_server` 与同一份注册表，所以工具的 schema 必然一致。
"""

from __future__ import annotations

import base64
import inspect
import json
import logging
from pathlib import Path
from typing import Any, Protocol

import httpx
import mcp.types as types
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError as McpToolError

from . import __version__
from .config import Settings, read_runtime, settings as default_settings
from .registry import Tool, all_tools

__all__ = [
    "Caller",
    "DirectCaller",
    "ForwardingCaller",
    "build_server",
    "build_http_app",
    "guard_app",
    "INSTRUCTIONS",
]

#: httpx 默认按 INFO 记录每条请求，会把 stderr 刷得很吵。
logging.getLogger("httpx").setLevel(logging.WARNING)

INSTRUCTIONS = """\
这是对**正在运行的 PowerPoint** 的实时控制接口。

几条纪律：
- 每次写操作前，先用 `ppt_status` / `ppt_decks` 看清当前有哪些演示打开着、停在第几页。
- 打开文件或插图用**绝对路径**；相对路径会按 PowerPoint 自己的进程目录解析，报错还很有误导性。
- 所有操作都会广播到用户的观察台，用户看得见你在做什么，所以每一步都要有明确目的。
- `ppt_undo` 用的是 PowerPoint 自己的撤销栈：粒度由它决定，一次撤销可能回退得比上一步更多。
  动一批破坏性操作前先 `ppt_save(mode="copy")` 留备份。
- 失败时返回结构化错误码（如 `com/timeout`）。`com/timeout` 通常意味着 PowerPoint 弹出了
  模态对话框，需要先让用户处理，而不是反复重试。
"""


class Caller(Protocol):
    """把一次工具调用落到守护进程。"""

    def call(self, name: str, args: dict[str, Any]) -> dict[str, Any]: ...

    def fetch_image(self, url: str) -> bytes | None: ...


class ForwardingCaller:
    """stdio 进程用：把调用转成对守护进程的一次 HTTP 请求。

    **令牌必须每次调用现读，不能启动时写死。** 这个进程由宿主拉起、会活很久，
    而用户随时可能 ``pptctl restart``——重启会换一个新令牌。第一版把令牌固定塞进
    ``httpx.Client`` 的默认头里，结果**重启一次就只剩"未知错误"**，
    所有 MCP 工具失效，非得再重启一次宿主才能恢复。

    顺带把两类响应错误讲清楚：401（令牌过期）与 FastAPI 的 ``{"detail": ...}``
    （以前会被上层变成毫无信息的"未知错误"）。
    """

    def __init__(
        self,
        settings: Settings,
        info: dict[str, Any] | None = None,
        *,
        ensure_hook: Any = None,
    ) -> None:
        self.settings = settings
        self.base_url = f"http://{settings.host}:{settings.port}"
        self.timeout = max(30.0, float(settings.call_timeout) + 15.0)
        self._info = dict(info or {})
        self._client: httpx.Client | None = None
        self._client_token: str | None = None
        #: 连不上守护进程时调它试着拉起来（stdio 面传 ``daemon.ensure``）。
        self._ensure_hook = ensure_hook

    # -- 令牌 ---------------------------------------------------------------

    def _record(self) -> dict[str, Any]:
        """现读 ``runtime.json``；读不到就退回上一次已知的那份。"""
        try:
            current = read_runtime(self.settings)
        except Exception:  # noqa: BLE001 - 读文件失败不该影响工具调用
            current = None
        if isinstance(current, dict) and current.get("token"):
            self._info = current
        return self._info

    def _client_for(self, token: str) -> httpx.Client:
        if self._client is None or self._client_token != token:
            if self._client is not None:
                self._client.close()
            self._client = httpx.Client(timeout=self.timeout, headers={"X-PPT-Token": token})
            self._client_token = token
        return self._client

    def _forget(self) -> None:
        """丢掉缓存的令牌与连接，下次调用重新读。"""
        if self._client is not None:
            self._client.close()
        self._client = None
        self._client_token = None
        self._info = {k: v for k, v in self._info.items() if k != "token"}

    def _request(self, method: str, path: str, **kwargs: Any) -> tuple[Any, dict[str, Any] | None]:
        """发一次请求。401 或连不上时刷新一次再试，最多两轮。"""
        last: dict[str, Any] | None = None
        for attempt in (1, 2):
            token = str(self._record().get("token") or "")
            try:
                resp = self._client_for(token).request(method, f"{self.base_url}{path}", **kwargs)
            except httpx.HTTPError as exc:
                # 超时/读断连时，写入可能已经落到 PowerPoint；重发 POST 会重复创建或删除。
                if method.upper() != "GET" and not isinstance(exc, (httpx.ConnectError, httpx.ConnectTimeout, httpx.PoolTimeout)):
                    return None, {
                        "code": "daemon/result-unknown",
                        "message": f"响应中断，操作可能已执行。请读取演示确认结果，勿直接重试写入：{exc}",
                        "details": {"result_unknown": True},
                    }
                last = {"code": "daemon/unreachable", "message": f"连不上守护进程：{exc}"}
                self._forget()
                if attempt == 1 and self._ensure_hook is not None:
                    try:
                        self._ensure_hook()
                    except Exception as hook_exc:  # noqa: BLE001
                        last["message"] += f"（自动拉起也失败：{hook_exc}）"
                    continue
                return None, last

            if resp.status_code == 401 and attempt == 1:
                # 守护进程重启过 → 令牌换了。重读 runtime.json 再试一次。
                self._forget()
                continue
            return resp, None
        return None, last

    # -- Caller 协议 --------------------------------------------------------

    def call(self, name: str, args: dict[str, Any]) -> dict[str, Any]:
        resp, error = self._request("POST", "/call", json={"name": name, "args": args})
        if error is not None:
            return {"ok": False, "tool": name, "error": error}
        if resp.status_code == 401:
            return {
                "ok": False,
                "tool": name,
                "error": {
                    "code": "auth/unauthorized",
                    "message": (
                        "守护进程拒绝了令牌（通常是它刚重启过、换了新令牌）。"
                        "重试一次一般就好了；若持续失败，跑 pptctl status 看一眼。"
                    ),
                },
            }
        try:
            data = resp.json()
        except ValueError:
            return {
                "ok": False,
                "tool": name,
                "error": {
                    "code": "daemon/bad-response",
                    "message": f"HTTP {resp.status_code}: {resp.text[:400]}",
                },
            }
        if not isinstance(data, dict) or "ok" not in data:
            # FastAPI 的错误体长这样：{"detail": "..."}。
            # 不做这层翻译的话，上层只会显示"未知错误"，等于没有信息。
            detail = data.get("detail") if isinstance(data, dict) else resp.text[:300]
            return {
                "ok": False,
                "tool": name,
                "error": {
                    "code": f"api/http-{resp.status_code}",
                    "message": f"守护进程返回 HTTP {resp.status_code}：{detail}",
                },
            }
        return data

    def fetch_image(self, url: str) -> bytes | None:
        resp, error = self._request("GET", url)
        if error is not None or resp is None or resp.status_code != 200:
            return None
        return resp.content


class DirectCaller:
    """守护进程内用：直接调 engine，不再绕一圈 HTTP。"""

    def __init__(self, engine: Any) -> None:
        self.engine = engine

    def call(self, name: str, args: dict[str, Any]) -> dict[str, Any]:
        return self.engine.invoke(name, args, source="mcp")

    def fetch_image(self, url: str) -> bytes | None:
        # url 形如 /shots/xxx.jpg —— 直接读硬盘，别为了内联一张图再打一次自己的 HTTP。
        name = url.rsplit("/", 1)[-1]
        if not name or name != Path(name).name:
            return None
        path = self.engine.settings.shot_dir / name
        try:
            return path.read_bytes() if path.is_file() else None
        except OSError:
            return None


def _make_handler(t: Tool, caller: Caller):
    """合成一个签名精确的处理函数，让 SDK 生成与注册表一致的 inputSchema。"""

    def handler(**kwargs: Any) -> Any:
        result = caller.call(t.name, kwargs)
        if not result.get("ok"):
            error = result.get("error") or {}
            # 用 SDK 的 ToolError：模型能看到真实原因，服务端只记一行 INFO，
            # 不会把长 traceback 写进 stderr 把管道堵死。
            raise McpToolError(f"{error.get('code', 'error')}: {error.get('message', '未知错误')}")
        payload = result.get("result")
        text = payload if isinstance(payload, str) else json.dumps(payload, ensure_ascii=False, indent=2, default=str)
        blocks: list[Any] = [types.TextContent(type="text", text=text)]

        # 截图：把缩略图作为图片内容块内联，模型才能真正"看见"这一页。
        # 高清图留在观察台，不塞进上下文（一张 1280 宽的 PNG 有 1 MB 以上）。
        if isinstance(payload, dict):
            for shot in payload.get("shots") or []:
                thumb = shot.get("thumb") or {}
                url = thumb.get("url")
                if not url:
                    continue
                data = caller.fetch_image(str(url))
                if data:
                    blocks.append(
                        types.ImageContent(
                            type="image",
                            data=base64.b64encode(data).decode("ascii"),
                            mimeType="image/jpeg",
                        )
                    )
        return blocks if len(blocks) > 1 else blocks[0]

    handler.__name__ = t.name
    handler.__doc__ = t.summary
    handler.__signature__ = inspect.Signature([p.to_parameter() for p in t.params])
    return handler


def build_server(caller: Caller, *, name: str = "powerpoint") -> MCPServer:
    """按注册表装配 MCP 服务器。stdio 与 streamable-http 共用这一个。"""
    import pptd.ops  # noqa: F401  导入即注册全部工具

    server = MCPServer(
        name,
        title="PowerPoint 实时控制",
        version=__version__,
        instructions=INSTRUCTIONS,
    )
    for t in all_tools():
        server.add_tool(
            _make_handler(t, caller),
            name=t.name,
            title=t.title,
            description=f"{t.summary}\n\n行为：{t.behavior}。返回：{t.returns}",
        )
    return server


class TokenGuard:
    """给 MCP over HTTP 加令牌门。

    **为什么必须加**：这个端口一旦被隧道到公网，没有令牌就等于任何人都能驱动用户的
    PowerPoint。令牌支持三种传法（``Authorization: Bearer``、``X-PPT-Token``、``?token=``），
    方便各种 MCP 客户端接入。
    """

    def __init__(self, app: Any, token: str) -> None:
        self.app = app
        self.token = token

    async def __call__(self, scope: dict, receive: Any, send: Any) -> None:
        if scope.get("type") != "http" or not self.token:
            await self.app(scope, receive, send)
            return
        headers = {k.decode("latin-1").lower(): v.decode("latin-1") for k, v in scope.get("headers") or []}
        supplied = headers.get("x-ppt-token")
        if not supplied:
            auth = headers.get("authorization", "")
            if auth.lower().startswith("bearer "):
                supplied = auth[7:].strip()
        if not supplied:
            from urllib.parse import parse_qs

            query = parse_qs((scope.get("query_string") or b"").decode("latin-1"))
            values = query.get("token") or []
            supplied = values[0] if values else None

        if supplied != self.token:
            body = json.dumps(
                {"jsonrpc": "2.0", "error": {"code": -32001, "message": "缺少或错误的访问令牌"}, "id": None},
                ensure_ascii=False,
            ).encode("utf-8")
            await send(
                {
                    "type": "http.response.start",
                    "status": 401,
                    "headers": [
                        (b"content-type", b"application/json; charset=utf-8"),
                        (b"content-length", str(len(body)).encode()),
                        (b"www-authenticate", b'Bearer realm="ppt-agent"'),
                    ],
                }
            )
            await send({"type": "http.response.body", "body": body})
            return
        await self.app(scope, receive, send)


def build_http_app(caller: Caller, token: str, *, host: str = "127.0.0.1") -> Any:
    """装配 streamable-http 的 ASGI 应用（已包好令牌门）。"""
    server = build_server(caller)
    return guard_app(server.streamable_http_app(host=host), token)


def guard_app(app: Any, token: str) -> Any:
    """给任意 ASGI 应用套上令牌门。"""
    return TokenGuard(app, token)
