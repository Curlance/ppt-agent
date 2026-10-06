"""MCP stdio 面：DSH 通过它拿到 ``mcp__ppt__*`` 工具。

为什么是"转发器"而不是"再来一个 COM 客户端"
------------------------------------------
PowerPoint 是单 UI 线程的进程外服务器。如果 MCP 进程自己也 ``Dispatch``，
就会出现两个进程同时驱动同一个 PowerPoint，代理调用会互相把对方的
调用打成 ``RPC_E_CALL_REJECTED``。所以这里只做转发，COM 归守护进程独占。

好处：stdio 进程被宿主重启也不会丢 COM 会话，观察台地址也保持稳定。

装配逻辑在 :mod:`pptd.mcp_server`，与 streamable-http 面共用同一份注册表——
所以两种传输给出的工具 schema 必然一致。

两条踩过的坑（改代码前先读）
----------------------------
1. **stdout 是协议通道**：任何调试输出必须走 stderr。
2. **异常要用 SDK 的 ``ToolError``**：抛普通异常会被判为"意外崩溃"，
   模型只看到 ``Error executing tool <name>``，我们的错误码与中文原因会被丢掉；
   而且 SDK 会把整条富文本 traceback 写进 stderr——Windows 匿名管道缓冲区很小，
   宿主若没及时抽干 stderr，服务端就会阻塞在写上，表现成"客户端永久等待"。
"""

from __future__ import annotations

import sys

from .config import settings as default_settings
from .daemon import ensure
from .mcp_server import ForwardingCaller, build_server
from .registry import all_tools

__all__ = ["main"]


def main(argv: list[str] | None = None) -> int:
    """stdio 入口。注意：stdout 是协议通道，任何调试输出都必须走 stderr。"""
    s = default_settings
    info = ensure(s)
    # ensure_hook：守护进程被停掉之后，下一次工具调用会顺手把它拉起来，
    # 而不是把"连不上"直接甩给模型。
    server = build_server(ForwardingCaller(s, info, ensure_hook=lambda: ensure(s)))
    print(
        f"[pptd] MCP stdio 已就绪，共 {len(all_tools())} 个工具（守护进程 pid={info.get('pid')}）",
        file=sys.stderr,
        flush=True,
    )
    print(
        "[pptd] 提示：令牌每次调用现读，所以 pptctl restart 之后本进程无需重启。",
        file=sys.stderr,
        flush=True,
    )
    server.run(transport="stdio")
    return 0
