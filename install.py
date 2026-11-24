#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""install.py — PPT-style 多端全嗅探全局安装器（机读目录名统一小写 ppt-style）。

铁律:
  1. 机读技能目录名统一为全小写 `ppt-style`，兼容所有严格校验小写命名的桌面 AI 客户端。
  2. 未显式指定 `--target-dir` 时，自动扫描本机已存在的主流客户端全局目录并同步写入；
     若一台都未探测到，回落至 ~/.agents/skills 并发出明确告警提示由 AI 自查真实目录。
  3. 发现同名旧目录或旧版大写 `PPT-style` 时，一律移入 ~/.ppt-style/backups/<timestamp>/，
     绝不把备份目录留在全局 skills/ 同级目录下（防止客户端误扫描并冲突触发）。
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


SKILL_NAME = "ppt-style"
LEGACY_SKILL_NAMES = ("PPT-style",)
SKILL_SOURCE = Path(__file__).resolve().parent
DIST_FILES = ("SKILL.md", "VERSION", "README.md", "install.py", "install_windows.bat", "install_mac.command")
DIST_DIRS = ("references", "scripts")

# (客户端根目录相对路径, 对应的全局 skills 子目录相对路径, 客户端展示名称)
CLIENT_DIR_SPECS: tuple[tuple[tuple[str, ...], tuple[str, ...], str], ...] = (
    ((".agents",), (".agents", "skills"), "通用 Agents (~/.agents/skills)"),
    ((".gemini", "antigravity"), (".gemini", "antigravity", "skills"), "Antigravity"),
    ((".claude",), (".claude", "skills"), "Claude Code"),
    ((".cursor",), (".cursor", "skills"), "Cursor"),
    ((".trae-cn",), (".trae-cn", "skills"), "Trae CN"),
    ((".trae",), (".trae", "skills"), "Trae"),
    ((".config", "opencode"), (".config", "opencode", "skills"), "OpenCode"),
    ((".codeium", "windsurf"), (".codeium", "windsurf", "skills"), "Windsurf"),
    ((".qoder",), (".qoder", "skills"), "Qoder"),
)


def _default_user_home(override: str | None = None) -> Path:
    return Path(override) if override else Path.home()


def _default_home_dir(override: str | None = None, user_home: Path | None = None) -> Path:
    if override:
        return Path(override)
    env_home = os.environ.get("PPT_STYLE_HOME")
    if env_home:
        return Path(env_home)
    base = user_home or Path.home()
    return base / ".ppt-style"


def detect_client_skill_dirs(user_home: Path) -> list[tuple[Path, str]]:
    """扫描当前电脑下真实存在的主流桌面 AI 客户端根目录，返回去重后的待安装 skills 目录列表。"""
    detected: list[tuple[Path, str]] = []
    seen: set[str] = set()
    for root_parts, skills_parts, label in CLIENT_DIR_SPECS:
        client_root = user_home.joinpath(*root_parts)
        skills_dir = user_home.joinpath(*skills_parts)
        if client_root.is_dir() or skills_dir.is_dir():
            key = str(skills_dir.resolve()) if skills_dir.exists() else str(skills_dir)
            if key not in seen:
                seen.add(key)
                detected.append((skills_dir, label))
    return detected


def _matching_skill_entries(target_dir: Path) -> list[Path]:
    """按实际磁盘目录名匹配 ppt-style 或历史旧名，避免 Windows 大小写不敏感导致重复匹配。"""
    if not target_dir.is_dir():
        return []
    targets_set = {SKILL_NAME, *LEGACY_SKILL_NAMES}
    return [entry for entry in target_dir.iterdir() if entry.name in targets_set]


def _install_into_single_dir(target_dir: Path, home_dir: Path) -> tuple[str, list[str]]:
    target_dir.mkdir(parents=True, exist_ok=True)
    backups: list[str] = []
    backups_root = home_dir / "backups"

    for existing_path in _matching_skill_entries(target_dir):
        backups_root.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        b_dest = backups_root / f"{existing_path.name}-{ts}"
        shutil.move(str(existing_path), str(b_dest))
        backups.append(str(b_dest))

    dest = target_dir / SKILL_NAME
    dest.mkdir(parents=True, exist_ok=True)
    for fname in DIST_FILES:
        src_f = SKILL_SOURCE / fname
        if src_f.is_file():
            shutil.copy2(src_f, dest / fname)
    for dname in DIST_DIRS:
        src_d = SKILL_SOURCE / dname
        if src_d.is_dir():
            shutil.copytree(src_d, dest / dname, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))

    bundled = SKILL_SOURCE / "bundled-profile"
    if bundled.is_dir():
        shutil.copytree(bundled, dest / "bundled-profile")

    return str(dest), backups


def install_skill(
    target_dir: Path | None,
    home_dir: Path,
    user_home: Path | None = None,
) -> dict[str, Any]:
    u_home = user_home or Path.home()
    fallback_used = False

    if target_dir is not None:
        targets = [(target_dir, "指定目录")]
    else:
        targets = detect_client_skill_dirs(u_home)
        if not targets:
            fallback_used = True
            targets = [(u_home / ".agents" / "skills", "默认通用目录 (~/.agents/skills)")]

    installed_paths: list[str] = []
    matched_clients: list[str] = []
    all_backups: list[str] = []

    for t_dir, label in targets:
        dest_str, backups = _install_into_single_dir(t_dir, home_dir)
        installed_paths.append(dest_str)
        matched_clients.append(label)
        all_backups.extend(backups)

    version = (
        (SKILL_SOURCE / "VERSION").read_text(encoding="utf-8").strip()
        if (SKILL_SOURCE / "VERSION").is_file()
        else "V0.1.1"
    )

    hint_lines = [
        f"已成功安装 PPT-style（机读目录 `{SKILL_NAME}`，版本 {version}）到 {len(installed_paths)} 个位置：",
        *[f"  - {lbl}: {p}" for lbl, p in zip(matched_clients, installed_paths)],
    ]
    if fallback_used:
        hint_lines.append("【注意】未自动探测到已知客户端目录，已先装入通用目录；如果新开对话喊不出来，请把仓库链接发给你的 AI 让它装到自己实际读取的目录。")
    hint_lines.append("请【新开一个对话窗口】（或重启软件），直接发送 `检查 PPT-style 是否就绪` 即可 3 秒验证是否真的加载成功。")

    return {
        "status": "INSTALLED",
        "skill_name": SKILL_NAME,
        "version": version,
        "installed_paths": installed_paths,
        "installed_path": installed_paths[0],
        "matched_clients": matched_clients,
        "fallback_used": fallback_used,
        "backup_paths": all_backups,
        "backup_path": all_backups[0] if all_backups else None,
        "human_hint": "\n".join(hint_lines),
    }


def uninstall_skill(
    target_dir: Path | None,
    home_dir: Path,
    user_home: Path | None = None,
) -> dict[str, Any]:
    u_home = user_home or Path.home()
    if target_dir is not None:
        targets = [(target_dir, "指定目录")]
    else:
        targets = detect_client_skill_dirs(u_home)
        default_agents = u_home / ".agents" / "skills"
        if all(p != default_agents for p, _ in targets):
            targets.append((default_agents, "默认通用目录"))

    removed_from: list[str] = []
    backups: list[str] = []
    backups_root = home_dir / "backups"

    for t_dir, _label in targets:
        for dest in _matching_skill_entries(t_dir):
            backups_root.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            b_dest = backups_root / f"{dest.name}-uninstalled-{ts}"
            shutil.move(str(dest), str(b_dest))
            removed_from.append(str(dest))
            backups.append(str(b_dest))

    if not removed_from:
        return {"status": "NOT_INSTALLED", "human_hint": "未在已探测目录检测到已安装的 ppt-style 技能。"}

    return {
        "status": "UNINSTALLED",
        "removed_from": removed_from,
        "backup_paths": backups,
        "human_hint": f"已从 {len(removed_from)} 个目录卸载 ppt-style（你的个人风格档案仍安全保留，日后重装即可恢复）。",
    }


def check_status(target_dir: Path | None, user_home: Path | None = None) -> dict[str, Any]:
    u_home = user_home or Path.home()
    targets = [(target_dir, "指定目录")] if target_dir is not None else detect_client_skill_dirs(u_home)
    found: list[dict[str, Any]] = []
    for t_dir, label in targets:
        for dest in _matching_skill_entries(t_dir):
            if dest.is_dir():
                ver_file = dest / "VERSION"
                ver = ver_file.read_text(encoding="utf-8").strip() if ver_file.is_file() else "unknown"
                found.append({"client": label, "dir_name": dest.name, "installed_path": str(dest), "version": ver})
    return {
        "installed": bool(found),
        "installations": found,
        "detected_client_dirs": [str(p) for p, _ in targets],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PPT-style 多端全嗅探全局安装器")
    parser.add_argument("--target-dir", default=None, help="显式指定单个全局 skills 目录（不传则自动嗅探所有已装客户端）")
    parser.add_argument("--user-home", default=None, help="覆盖用户家目录（便于隔离测试）")
    parser.add_argument("--home-dir", default=None, help="覆盖个人档案与备份目录（默认 ~/.ppt-style）")
    parser.add_argument("--uninstall", action="store_true", help="卸载全局技能")
    parser.add_argument("--status", action="store_true", help="查看各客户端安装状态")

    args = parser.parse_args(argv)
    u_home = _default_user_home(args.user_home)
    target = Path(args.target_dir) if args.target_dir else None
    home = _default_home_dir(args.home_dir, user_home=u_home)

    if args.status:
        print(json.dumps(check_status(target, user_home=u_home), ensure_ascii=False, indent=2))
        return 0
    if args.uninstall:
        print(json.dumps(uninstall_skill(target, home, user_home=u_home), ensure_ascii=False, indent=2))
        return 0
    print(json.dumps(install_skill(target, home, user_home=u_home), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
