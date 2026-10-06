"""开发时用的探针输出格式化器：把 dsh_client_probe.mjs 的 JSON 打成人看的样子。

pytest 里的断言在 tests/test_dsh_bundle.py；这个小工具只是让手动排查方便。
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
NODE = Path.home() / ".dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe"


def main() -> int:
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "dsh-bundle" / "client.js"
    node = NODE if NODE.is_file() else Path("node")
    proc = subprocess.run(
        [str(node), str(ROOT / "tests" / "dsh_client_probe.mjs"), str(target)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    raw = (proc.stdout or "").strip().splitlines()
    if not raw:
        print("探针没有输出")
        print(proc.stderr[:2000])
        return 1
    data = json.loads(raw[-1])
    print(f"目标：{target}")
    print(f"结论：{'全部通过 ✓' if data['ok'] else '有问题 ✗'}\n")
    for item in data["checks"]:
        mark = "OK  " if item["pass"] else "FAIL"
        detail = f"  ← {item['detail']}" if item["detail"] and not item["pass"] else ""
        print(f"  {mark} {item['name']}{detail}")
    if data["errors"]:
        print("\n错误：")
        for err in data["errors"]:
            print(f"  - {err}")
    return 0 if data["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
