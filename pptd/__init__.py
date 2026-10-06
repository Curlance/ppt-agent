"""pptd —— 让 agent 实时控制 PowerPoint 的守护进程。

三层结构：
- ``pptd.com``    COM 桥：单条 STA 线程独占 PowerPoint Application，独立进程隔离挂起。
- ``pptd.registry`` 唯一事实源：工具的 JSON Schema 与处理器。
- 三个面从注册表生成：MCP（stdio）、HTTP+OpenAPI、CLI。
"""

from __future__ import annotations

__all__ = ["__version__"]

__version__ = "0.1.0"
