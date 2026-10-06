"""agent 技能：让另一个 agent 知道怎么正确用这套工具。

技能是**给人机共读的说明**，不是代码。它的价值在于把"用错会怎样"提前讲清楚——
比如 `ppt_undo` 的粒度限制、相对路径的坑、以及"改完必须截图验证"。
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Any

__all__ = ["SKILL_NAME", "skill_source", "dsh_skills_dir", "installed_path", "install", "status"]

SKILL_NAME = "ppt-agent"


def skill_source() -> Path:
    """仓库里的技能目录。"""
    source = Path(__file__).resolve().parent.parent / "skills" / SKILL_NAME
    return source if source.is_dir() else Path(__file__).resolve().parent / "resources" / "skills" / SKILL_NAME


def dsh_skills_dir() -> Path:
    """DSH 读取用户技能的目录（与 github-publish / ppt-master 那些同级）。"""
    home = os.environ.get("DSH_HOME") or (Path(os.path.expanduser("~")) / ".dsh")
    return Path(home) / "skills"


def installed_path() -> Path:
    return dsh_skills_dir() / SKILL_NAME


def _parse_frontmatter(text: str) -> dict[str, str]:
    """极简 frontmatter 解析：够用来取出 name/description。"""
    if not text.startswith("---"):
        return {}
    end = text.find("\n---", 3)
    if end < 0:
        return {}
    out: dict[str, str] = {}
    key = ""
    for raw in text[3:end].splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        if raw[:1].isspace() and key:          # 续行（description: > 的折行）
            out[key] = (out[key] + " " + raw.strip()).strip()
            continue
        if ":" in raw:
            key, _, value = raw.partition(":")
            key = key.strip()
            out[key] = value.strip().lstrip(">").strip()
    return out


def install(target: Path | None = None, *, force: bool = True) -> dict[str, Any]:
    """把技能装到 DSH 的用户技能目录。

    装之前会校验 frontmatter —— 没有 ``name`` / ``description`` 的技能，
    宿主不会在正确的时机把它交给模型，等于白装。
    """
    source = skill_source()
    skill_file = source / "SKILL.md"
    if not skill_file.is_file():
        return {"ok": False, "error": f"找不到技能文件：{skill_file}"}

    text = skill_file.read_text(encoding="utf-8")
    meta = _parse_frontmatter(text)
    problems: list[str] = []
    if meta.get("name") != SKILL_NAME:
        problems.append(f"frontmatter 的 name 应为 {SKILL_NAME!r}，实际是 {meta.get('name')!r}")
    if not meta.get("description"):
        problems.append("frontmatter 缺少 description（宿主靠它决定何时把技能交给模型）")
    if problems:
        return {"ok": False, "error": "技能 frontmatter 不合法", "problems": problems}

    destination = target or installed_path()
    previous: Path | None = None
    if destination.exists():
        if not force:
            return {"ok": True, "skipped": True, "path": str(destination)}
        # 覆盖前先留一份旧的，装坏了能退回去
        previous = destination.with_name(destination.name + ".bak")
        if previous.exists():
            shutil.rmtree(previous, ignore_errors=True)
        shutil.copytree(destination, previous)

    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, destination, dirs_exist_ok=True)
    return {
        "ok": True,
        "path": str(destination),
        "backup": str(previous) if previous else None,
        "description_chars": len(meta["description"]),
        "hint": "宿主实时监视这个目录，装完当场生效，不用重启。（实测确认）",
    }


def status() -> dict[str, Any]:
    source = skill_source()
    installed = installed_path()
    meta: dict[str, str] = {}
    if (source / "SKILL.md").is_file():
        meta = _parse_frontmatter((source / "SKILL.md").read_text(encoding="utf-8"))

    same = False
    if installed.is_dir() and (installed / "SKILL.md").is_file() and (source / "SKILL.md").is_file():
        same = (installed / "SKILL.md").read_bytes() == (source / "SKILL.md").read_bytes()

    return {
        "source": str(source),
        "installed": str(installed) if installed.is_dir() else None,
        "installed_matches_source": same,
        "frontmatter_ok": meta.get("name") == SKILL_NAME and bool(meta.get("description")),
        "skills_dir": str(dsh_skills_dir()),
        "verdict": _verdict(installed, same, meta),
    }


def _verdict(installed: Path, same: bool, meta: dict[str, str]) -> str:
    if meta.get("name") != SKILL_NAME or not meta.get("description"):
        return "技能 frontmatter 不合法，装了也不会在正确的时机触发"
    if not installed.is_dir():
        return "尚未安装：跑 `pptctl skill-install`"
    if not same:
        return "已安装，但内容与仓库里的不一致：重跑 `pptctl skill-install` 覆盖"
    return "已安装且与仓库一致（宿主实时监视该目录，无需重启）"
