"""用 DevTools 协议驱动无头浏览器：确定性等待、精确视口、按需截图。

为什么不继续用 `--dump-dom` / `--screenshot`
------------------------------------------
那两个开关只等 **load 事件**，而页面的内容来自 load 之后才发的异步请求（fetch / SSE）。
于是同一个命令有时拿到有内容的 DOM、有时拿到空壳——**时快时慢，靠运气**。
我就因此误判过：以为按钮被布局挤掉了，其实是截图把宽布局裁掉了；
又以为 `dump-dom` 稳定可靠，结果换一轮就失败。

另外 headless 有约 **518px 的最小窗口宽度**，`--window-size=380` 得到的仍是 518 的布局、
只是把截图裁成 380。窄版式的验证必须靠 CDP 的 `Emulation.setDeviceMetricsOverride`。

这个模块给的是三件确定的事：

- :meth:`Browser.wait_for` —— 轮询到条件成立再继续，而不是赌它已经好了；
- :meth:`Browser.evaluate` —— 直接读运行后的 DOM/几何；
- :meth:`Browser.screenshot` —— 指定**真实**视口尺寸截图（含整页高度）。
"""

from __future__ import annotations

import base64
import json
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path
from typing import Any

__all__ = ["Browser", "find_browser", "CDPError"]

EDGE_CANDIDATES = [
    Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
]


class CDPError(RuntimeError):
    pass


def find_browser() -> Path | None:
    import os

    override = os.environ.get("PPT_AGENT_BROWSER")
    if override and Path(override).is_file():
        return Path(override)
    for candidate in EDGE_CANDIDATES:
        if candidate.is_file():
            return candidate
    return None


class Browser:
    """一个无头浏览器会话。用 ``with`` 保证进程和临时目录都被收掉。"""

    def __init__(self, port: int = 9333, timeout: float = 30.0) -> None:
        browser = find_browser()
        if browser is None:
            raise CDPError("找不到 Edge/Chrome，可用 PPT_AGENT_BROWSER 指定")
        self._profile = tempfile.TemporaryDirectory(prefix="ppt-agent-cdp-")
        self._port = port
        self._timeout = timeout
        self._next_id = 0
        self._ws = None

        self._process = subprocess.Popen(
            [
                str(browser),
                "--headless=new",
                "--disable-gpu",
                "--no-first-run",
                "--no-default-browser-check",
                f"--remote-debugging-port={port}",
                f"--user-data-dir={self._profile.name}",
                "about:blank",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        self._connect()

    # -- 连接 ---------------------------------------------------------------

    def _endpoint(self) -> str:
        deadline = time.monotonic() + self._timeout
        while time.monotonic() < deadline:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{self._port}/json/list", timeout=2) as resp:
                    targets = json.loads(resp.read().decode("utf-8"))
                for target in targets:
                    if target.get("type") == "page" and target.get("webSocketDebuggerUrl"):
                        return str(target["webSocketDebuggerUrl"])
            except Exception:  # noqa: BLE001 - 浏览器还没起来
                pass
            time.sleep(0.25)
        raise CDPError(f"CDP 端点 {self._port} 在 {self._timeout:.0f}s 内没起来")

    def _connect(self) -> None:
        from websockets.sync.client import connect

        # legacy=True：明确用"直接连接"而不是 with 上下文管理器，避免库的弃用警告
        self._ws = connect(
            self._endpoint(), max_size=64 * 1024 * 1024, open_timeout=self._timeout, legacy=True
        )
        self.call("Page.enable")
        self.call("Runtime.enable")

    # -- 调用 ---------------------------------------------------------------

    def call(self, method: str, params: dict[str, Any] | None = None, timeout: float | None = None) -> dict[str, Any]:
        """发一条 CDP 命令并等它的响应（跳过中间的事件通知）。"""
        if self._ws is None:
            raise CDPError("连接已关闭")
        self._next_id += 1
        message_id = self._next_id
        self._ws.send(json.dumps({"id": message_id, "method": method, "params": params or {}}))
        deadline = time.monotonic() + (timeout or self._timeout)
        while time.monotonic() < deadline:
            remaining = max(0.1, deadline - time.monotonic())
            raw = self._ws.recv(timeout=remaining)
            data = json.loads(raw)
            if data.get("id") != message_id:
                continue  # 事件通知，忽略
            if "error" in data:
                raise CDPError(f"{method} 失败：{data['error']}")
            return data.get("result") or {}
        raise CDPError(f"{method} 超时")

    # -- 页面操作 -----------------------------------------------------------

    def goto(self, url: str, *, settle: float = 0.4) -> None:
        self.call("Page.navigate", {"url": url})
        # 等 document.readyState 到 complete
        self.wait_for("document.readyState === 'complete'", timeout=20)
        if settle:
            time.sleep(settle)

    def evaluate(self, expression: str, *, timeout: float | None = None) -> Any:
        result = self.call(
            "Runtime.evaluate",
            {"expression": expression, "returnByValue": True, "awaitPromise": True},
            timeout,
        )
        if result.get("exceptionDetails"):
            raise CDPError(f"JS 抛异常：{result['exceptionDetails'].get('text')}")
        return (result.get("result") or {}).get("value")

    def wait_for(self, expression: str, *, timeout: float = 20.0, interval: float = 0.25) -> Any:
        """轮询到表达式为真再返回它的值。

        这就是 `--dump-dom` 给不了的东西：**确定性地等页面准备好**，而不是赌。
        """
        deadline = time.monotonic() + timeout
        last: Any = None
        while time.monotonic() < deadline:
            last = self.evaluate(expression)
            if last:
                return last
            time.sleep(interval)
        raise CDPError(f"等待超时（{timeout:.0f}s）：{expression}（最后的值 {last!r}）")

    def html(self) -> str:
        return str(self.evaluate("document.documentElement.outerHTML"))

    def set_viewport(self, width: int, height: int) -> None:
        """设定**真实**的布局视口。

        headless 的 `--window-size` 有约 518px 的下限，`--force-device-scale-factor` 也不
        改变 CSS 视口；只有这个 API 能真正拿到 360/380 这类窄视口。
        """
        self.call(
            "Emulation.setDeviceMetricsOverride",
            {"width": width, "height": height, "deviceScaleFactor": 1, "mobile": False},
        )

    def screenshot(self, out: Path, *, width: int = 1440, height: int = 900, full_page: bool = True) -> Path:
        self.set_viewport(width, height)
        time.sleep(0.3)  # 让重排后的布局稳定下来
        params: dict[str, Any] = {"format": "png"}
        if full_page:
            params["captureBeyondViewport"] = True
        data = self.call("Page.captureScreenshot", params)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_bytes(base64.b64decode(data["data"]))
        return out

    def close(self) -> None:
        try:
            if self._ws is not None:
                self._ws.close()
        except Exception:  # noqa: BLE001
            pass
        self._ws = None
        try:
            self._process.terminate()
            self._process.wait(timeout=10)
        except Exception:  # noqa: BLE001
            self._process.kill()
        self._profile.cleanup()

    def __enter__(self) -> "Browser":
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()
