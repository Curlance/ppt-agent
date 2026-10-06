"""开发用：用 DevTools 协议给页面截图 / 取渲染后的 DOM。

**为什么不用 `msedge --screenshot` / `--dump-dom`**：那两个开关只等 load 事件，
而页面内容来自 load 之后的异步请求，于是同一个命令时灵时不灵。我为此误判过两次
（一次以为按钮被布局挤掉、其实是截图把宽布局裁掉了；一次以为 dump-dom 稳定，
换一轮就拿到空壳）。第一版这个脚本就是这么写的，现在改成走 CDP —— 只有一套机制。

用法：
    python tests/probe_page.py shot <URL> <输出PNG> [宽] [高]
    python tests/probe_page.py dom  <URL> [输出HTML]
    python tests/probe_page.py wait <URL> "<JS 表达式>"
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from cdp import Browser, CDPError, find_browser

__all__ = ["find_browser", "Browser", "CDPError"]


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    mode, url = sys.argv[1], sys.argv[2]

    with Browser() as browser:
        browser.goto(url)
        if mode == "shot":
            target = Path(sys.argv[3]) if len(sys.argv) > 3 else Path("page.png")
            width = int(sys.argv[4]) if len(sys.argv) > 4 else 1440
            height = int(sys.argv[5]) if len(sys.argv) > 5 else 900
            path = browser.screenshot(target, width=width, height=height)
            print(f"已截图（视口 {width}x{height}，整页）：{path}（{path.stat().st_size} 字节）")
            return 0
        if mode == "dom":
            html = browser.html()
            if len(sys.argv) > 3:
                out = Path(sys.argv[3])
                out.parent.mkdir(parents=True, exist_ok=True)
                out.write_text(html, encoding="utf-8")
                print(f"DOM 已写入 {out}（{len(html)} 字符）")
            else:
                print(html)
            return 0
        if mode == "wait":
            expression = sys.argv[3] if len(sys.argv) > 3 else "document.readyState === 'complete'"
            value = browser.wait_for(expression)
            print(json.dumps(value, ensure_ascii=False))
            return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    import os as _os

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    _os.chdir(Path(__file__).resolve().parent.parent)
    raise SystemExit(main())
