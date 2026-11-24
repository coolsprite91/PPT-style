#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""install.py — PPT-style 单技能全局安装器。

铁律:
  1. 目标全局目录已存在同名旧技能时，必须把旧版移入 ~/.ppt-style/backups/<timestamp>/，
     绝不把备份目录留在全局 skills/ 同级目录下（防止客户端误扫描并冲突触发）。
  2. 支持 --target-dir、--home-dir、--uninstall 与 --status。
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


SKILL_NAME = "PPT-style"
SKILL_SOURCE = Path(__file__).resolve().parent
DIST_FILES = ("SKILL.md", "VERSION", "README.md", "install.py", "install_windows.bat", "install_mac.command")
DIST_DIRS = ("references", "scripts")


def _default_target_dir() -> Path:
    candidates = [
        Path.home() / ".agents" / "skills",
        Path.home() / ".claude" / "skills",
        Path.home() / ".cursor" / "skills",
    ]
    for c in candidates:
        if c.is_dir():
            return c
    return candidates[0]


def _default_home_dir(override: str | None = None) -> Path:
    if override:
        return Path(override)
    env_home = os.environ.get("PPT_STYLE_HOME")
    if env_home:
        return Path(env_home)
    return Path.home() / ".ppt-style"


def install_skill(target_dir: Path, home_dir: Path) -> dict[str, Any]:
    target_dir.mkdir(parents=True, exist_ok=True)
    dest = target_dir / SKILL_NAME
    backup_path: str | None = None

    if dest.exists():
        backups_root = home_dir / "backups"
        backups_root.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        b_dest = backups_root / f"{SKILL_NAME}-{ts}"
        shutil.move(str(dest), str(b_dest))
        backup_path = str(b_dest)

    dest.mkdir(parents=True, exist_ok=True)
    for fname in DIST_FILES:
        src_f = SKILL_SOURCE / fname
        if src_f.is_file():
            shutil.copy2(src_f, dest / fname)
    for dname in DIST_DIRS:
        src_d = SKILL_SOURCE / dname
        if src_d.is_dir():
            shutil.copytree(src_d, dest / dname, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))

    # Copy bundled-profile if this package was exported with one
    bundled = SKILL_SOURCE / "bundled-profile"
    if bundled.is_dir():
        shutil.copytree(bundled, dest / "bundled-profile")

    version = (SKILL_SOURCE / "VERSION").read_text(encoding="utf-8").strip() if (SKILL_SOURCE / "VERSION").is_file() else "V0.1"

    return {
        "status": "INSTALLED",
        "skill_name": SKILL_NAME,
        "version": version,
        "installed_path": str(dest),
        "backup_path": backup_path,
        "human_hint": (
            f"已成功安装 {SKILL_NAME}（版本 {version}）！"
            "请新开一个对话窗口（或重启你的 AI 软件），直接说“帮我建 PPT 风格工坊”即可开始使用。"
        ),
    }


def uninstall_skill(target_dir: Path, home_dir: Path) -> dict[str, Any]:
    dest = target_dir / SKILL_NAME
    if not dest.exists():
        return {"status": "NOT_INSTALLED", "human_hint": "未在目标目录检测到已安装的 PPT-style 技能。"}
    backups_root = home_dir / "backups"
    backups_root.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    b_dest = backups_root / f"{SKILL_NAME}-uninstalled-{ts}"
    shutil.move(str(dest), str(b_dest))
    return {
        "status": "UNINSTALLED",
        "removed_from": str(dest),
        "backup_path": str(b_dest),
        "human_hint": "已卸载全局 PPT-style 技能（你的个人风格档案仍安全保留，日后重装即可恢复）。",
    }


def check_status(target_dir: Path) -> dict[str, Any]:
    dest = target_dir / SKILL_NAME
    if not dest.is_dir():
        return {"installed": False, "target_dir": str(target_dir)}
    ver_file = dest / "VERSION"
    ver = ver_file.read_text(encoding="utf-8").strip() if ver_file.is_file() else "unknown"
    return {"installed": True, "installed_path": str(dest), "version": ver}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PPT-style 全局安装器")
    parser.add_argument("--target-dir", default=None, help="全局 skills 目标目录（默认自动探测 ~/.agents/skills）")
    parser.add_argument("--home-dir", default=None, help="全局档案与备份目录（默认 ~/.ppt-style）")
    parser.add_argument("--uninstall", action="store_true", help="卸载全局技能")
    parser.add_argument("--status", action="store_true", help="查看安装状态")

    args = parser.parse_args(argv)
    target = Path(args.target_dir) if args.target_dir else _default_target_dir()
    home = _default_home_dir(args.home_dir)

    if args.status:
        print(json.dumps(check_status(target), ensure_ascii=False, indent=2))
        return 0
    if args.uninstall:
        print(json.dumps(uninstall_skill(target, home), ensure_ascii=False, indent=2))
        return 0
    print(json.dumps(install_skill(target, home), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
