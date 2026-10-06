"""真机集成测试：在**全新空白演示**上跑完整编辑流程。

刻意用 ``ppt_new`` 现场造一份演示，而不是打开用户已有的文件——
真机测试绝不该改动用户的东西。

    .venv\\Scripts\\python.exe -m pytest -m com -v
"""

from __future__ import annotations

import json
import os
import sys
import urllib.error
import urllib.request
from pathlib import Path

import pytest

pytestmark = pytest.mark.com

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pptd.config import settings as cfg  # noqa: E402


def _call(daemon: dict, name: str, args: dict | None = None, timeout: float = 120.0) -> dict:
    body = json.dumps({"name": name, "args": args or {}}, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        f"http://{cfg.host}:{cfg.port}/call",
        data=body,
        method="POST",
        headers={"Content-Type": "application/json", "X-PPT-Token": daemon["token"]},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:  # 工具报错会走 400，body 里才是真正的错误
        return json.loads(exc.read().decode("utf-8"))


@pytest.fixture(scope="module")
def daemon():
    from pptd.daemon import ensure, is_alive, stop

    already_running = is_alive(cfg)
    info = ensure(cfg, timeout=40)
    yield info
    if not already_running:
        stop(cfg, force=True)


@pytest.fixture(scope="module")
def original_pace(daemon: dict):
    """测试期间把跟速归零，跑完还原——顺便验证 /pace 在真机守护进程上可用。"""
    import httpx

    with httpx.Client(timeout=10) as client:
        before = client.get(f"http://{cfg.host}:{cfg.port}/state", headers={"X-PPT-Token": daemon["token"]}).json()
        client.post(
            f"http://{cfg.host}:{cfg.port}/pace",
            json={"pace_ms": 0},
            headers={"X-PPT-Token": daemon["token"]},
        )
    yield before.get("pace_ms")
    import httpx as _httpx

    with _httpx.Client(timeout=10) as client:
        client.post(
            f"http://{cfg.host}:{cfg.port}/pace",
            json={"pace_ms": before.get("pace_ms") or 0},
            headers={"X-PPT-Token": daemon["token"]},
        )


@pytest.fixture(scope="module")
def fresh_deck(daemon: dict, original_pace):
    """新建一份演示，测试结束无论成败都关掉且不保存。"""
    created = _call(daemon, "ppt_new", {})
    assert created["ok"] is True, created
    name = created["result"]["presentation"]["name"]
    yield name
    _call(daemon, "ppt_close", {"deck": name, "save": False})


def _ensure_slide_count(daemon: dict, deck: str, want: int) -> int:
    """保证演示至少有 want 页。

    为什么每个用例都要自保：PowerPoint 的撤销可能把新增的幻灯片一并回退掉，
    模块级的 fresh_deck 是共享的，用例之间不该互相踩。
    """
    info = _call(daemon, "ppt_decks", {})
    current = next((d["slides"] for d in info["result"]["decks"] if d["name"] == deck), 0)
    while current < want:
        added = _call(daemon, "ppt_add_slide", {"deck": deck, "layout": 1})
        assert added["ok"] is True, added
        current += 1
    return current


def test_new_deck_is_visible_and_listed(daemon: dict, fresh_deck: str) -> None:
    state = _call(daemon, "ppt_decks", {})
    assert state["ok"] is True
    names = [d["name"] for d in state["result"]["decks"]]
    assert fresh_deck in names, f"新建的演示没出现在列表里：{names}"


def test_add_slide_and_read_back(daemon: dict, fresh_deck: str) -> None:
    added = _call(daemon, "ppt_add_slide", {"deck": fresh_deck, "layout": 1})
    assert added["ok"] is True, added
    index = added["result"]["slide"]
    assert index >= 1

    read = _call(daemon, "ppt_read_slide", {"deck": fresh_deck, "slide": index})
    assert read["ok"] is True
    assert read["result"]["slide"] == index
    assert read["result"]["shape_count"] >= 0


def test_textbox_set_text_geometry_and_notes(daemon: dict, fresh_deck: str) -> None:
    added = _call(
        daemon,
        "ppt_add_textbox",
        {
            "deck": fresh_deck,
            "slide": 1,
            "text": "ppt-agent 真机测试",
            "left": 80,
            "top": 90,
            "width": 500,
            "height": 90,
            "font_size": 30,
        },
    )
    assert added["ok"] is True, added
    shape = added["result"]["shape"]

    written = _call(
        daemon,
        "ppt_set_text",
        {"deck": fresh_deck, "slide": 1, "shape": shape, "text": "改过的文字", "bold": True},
    )
    assert written["ok"] is True, written
    assert written["result"]["after"] == "改过的文字"
    assert written["result"]["applied"]["bold"] is True

    moved = _call(
        daemon,
        "ppt_set_geometry",
        {"deck": fresh_deck, "slide": 1, "shape": shape, "left": 140.0, "width": 620.0},
    )
    assert moved["ok"] is True, moved
    assert moved["result"]["after"]["left"] == 140.0

    notes = _call(daemon, "ppt_set_notes", {"deck": fresh_deck, "slide": 1, "text": "这一步是讲重点"})
    assert notes["ok"] is True, notes
    assert notes["result"]["after"] == "这一步是讲重点"

    # 读回来核对（COM 写入之后必须能读回同样内容）
    read = _call(daemon, "ppt_read_slide", {"deck": fresh_deck, "slide": 1})
    texts = [s.get("text") for s in read["result"]["shapes"] if s.get("text")]
    assert "改过的文字" in texts, texts
    assert read["result"]["notes"] == "这一步是讲重点"


def test_add_picture_on_real_deck(daemon: dict, fresh_deck: str, tmp_path_factory) -> None:
    from PIL import Image

    png = tmp_path_factory.mktemp("pic") / "logo.png"
    Image.new("RGB", (240, 120), (77, 107, 254)).save(png)

    result = _call(
        daemon,
        "ppt_add_picture",
        {"deck": fresh_deck, "slide": 1, "path": str(png), "left": 200, "top": 320, "width": 240, "height": 120},
    )
    assert result["ok"] is True, result
    assert result["result"]["added"] is True
    assert result["result"]["geometry"]["width"] == 240.0


def test_shot_view_of_edited_deck(daemon: dict, fresh_deck: str, tmp_path: Path) -> None:
    """截的图必须是真图：尺寸对得上、不是纯色空图。"""
    _ensure_slide_count(daemon, fresh_deck, 1)
    result = _call(daemon, "ppt_shot", {"deck": fresh_deck, "slide": 1, "width": 800, "thumb_width": 640})
    assert result["ok"] is True, result
    shot = result["result"]["shots"][0]
    image_path = Path(shot["image"]["path"])
    thumb_path = Path(shot["thumb"]["path"])
    assert image_path.is_file() and thumb_path.is_file()

    from PIL import Image

    with Image.open(image_path) as img:
        assert img.size == (800, round(800 * 540 / 960))
        colors = img.convert("RGB").getcolors(maxcolors=1_000_000)
        assert len(colors) > 5, "截图几乎只有一种颜色，可能是空白页"
    with Image.open(thumb_path) as thumb:
        assert thumb.size[0] == 640
        assert thumb.size[0] < 800

    # 刻意**不**比较字节大小：空白页是纯色，PNG 能压到 2 KB 以下，反而比
    # 带 JPEG 噪声的缩略图还小。字节大小关系只在照片页上成立
    # （真机实测 991 KB vs 120 KB，见 README）。


def test_undo_on_real_deck(daemon: dict, fresh_deck: str) -> None:
    """撤销：只断言"刚加进去的东西确实没了"。

    PowerPoint 的撤销粒度**不按我们的工具调用切分**——实测一次 Undo 可能把之前
    好几步（甚至新增的幻灯片）一起回退。所以这里不假设形状数精确回到操作前，
    只断言我们写的那段文字消失了；撤销粒度本身在 ``caveat`` 里如实返回。
    """
    _ensure_slide_count(daemon, fresh_deck, 1)

    added = _call(daemon, "ppt_add_textbox", {"deck": fresh_deck, "slide": 1, "text": "待撤销的记号"})
    assert added["ok"] is True, added
    before_undo = _call(daemon, "ppt_read_slide", {"deck": fresh_deck, "slide": 1})
    texts = [s.get("text") for s in before_undo["result"]["shapes"] if s.get("text")]
    assert "待撤销的记号" in texts, texts

    undone = _call(daemon, "ppt_undo", {"deck": fresh_deck})
    assert undone["ok"] is True, undone
    assert undone["result"]["undone"] >= 1
    assert "撤销粒度" in undone["result"]["caveat"]
    assert "ppt_save" in undone["result"]["safety"], "破坏性操作前要先备份的提醒不能丢"

    _ensure_slide_count(daemon, fresh_deck, 1)  # 可能被一并回退，补回来
    after = _call(daemon, "ppt_read_slide", {"deck": fresh_deck, "slide": 1})
    remaining = [s.get("text") for s in after["result"]["shapes"] if s.get("text")]
    assert "待撤销的记号" not in remaining, remaining


def test_save_copy_makes_backup(daemon: dict, fresh_deck: str) -> None:
    result = _call(daemon, "ppt_save", {"deck": fresh_deck, "mode": "copy"})
    assert result["ok"] is True, result
    copy_path = Path(result["result"]["copy_path"])
    assert copy_path.is_file()
    assert copy_path.stat().st_size > 1000, "备份文件太小，不像一份演示"


def test_live_view_tracks_current_slide(daemon: dict, fresh_deck: str) -> None:
    """实时取景必须跟着 PowerPoint 的当前页走。"""
    import httpx

    _ensure_slide_count(daemon, fresh_deck, 2)
    _call(daemon, "ppt_goto", {"deck": fresh_deck, "slide": 1})
    with httpx.Client(timeout=40) as client:
        first = client.get(
            f"http://{cfg.host}:{cfg.port}/live",
            params={"width": 640},
            headers={"X-PPT-Token": daemon["token"]},
        )
        assert first.status_code == 200
        assert first.headers["content-type"].startswith("image/jpeg")
        assert first.headers["x-ppt-slide"] == "1"
        assert len(first.content) > 2000

        # 用户自己翻页 / agent 跳页，取景都要跟上
        _call(daemon, "ppt_goto", {"deck": fresh_deck, "slide": 2})
        second = client.get(
            f"http://{cfg.host}:{cfg.port}/live",
            params={"width": 640},
            headers={"X-PPT-Token": daemon["token"]},
        )
        assert second.headers["x-ppt-slide"] == "2"
