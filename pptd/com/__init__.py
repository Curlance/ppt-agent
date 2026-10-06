"""COM 桥：把 PowerPoint 的调用收敛到单条 STA 线程。

为什么必须这样
--------------
1. PowerPoint 是单 UI 线程的进程外 COM 服务器。所有调用必须**串行**，
   并发调用会拿到 ``RPC_E_CALL_REJECTED``（"调用被拒绝"）。
2. 任何一次调用都可能被模态对话框永久卡住。因此：
   - 调用在**独立进程**里执行（守护进程），卡死可被强杀而不影响 agent；
   - 每个请求都有超时，超时后把会话标记为 degraded，交由 ``pptctl stop --force`` 收场。
3. 可见性是硬要求：拿到实例后立刻 ``Visible = True``，绝不隐身操作。
"""

from __future__ import annotations

from .session import ComInfo, ComSession

__all__ = ["ComSession", "ComInfo"]
