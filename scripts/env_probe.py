#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""env_probe.py — 零依赖运行环境探针。

检测当前 Python 是否为有效运行环境（显式识别 WindowsApps 空壳）、
是否安装可选写稿组件（python-pptx / Pillow），以及本机可用的常用中英文字体。
"""

from __future__ import annotations

import importlib.util
import json
import os
import platform
import shutil
import sys
from pathlib import Path
from typing import Any

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


FONT_CATALOG: list[tuple[str, tuple[str, ...]]] = [
    ("微软雅黑", ("msyh.ttc", "msyh.ttf", "msyhbd.ttc")),
    ("黑体", ("simhei.ttf",)),
    ("楷体", ("simkai.ttf", "kaiti.ttf")),
    ("宋体", ("simsun.ttc", "simsunb.ttf")),
    ("仿宋", ("simfang.ttf",)),
    ("等线", ("deng.ttf", "dengb.ttf")),
    ("苹方", ("pingfang.ttc", "pingfang sc.ttc")),
    ("思源黑体", ("sourcehansans", "notosanscjk", "notosanssc")),
    ("Arial", ("arial.ttf", "arial.ttc")),
    ("Calibri", ("calibri.ttf",)),
    ("Segoe UI", ("segoeui.ttf",)),
]


def _is_windows_apps_stub(exe_path: str) -> bool:
    norm = exe_path.replace("/", "\\").lower()
    return "windowsapps" in norm


def _find_real_pythons() -> list[str]:
    candidates: list[str] = []
    if not _is_windows_apps_stub(sys.executable):
        candidates.append(sys.executable)

    system = platform.system().lower()
    if system == "windows":
        local_app = Path(os.environ.get("LOCALAPPDATA", ""))
        prog_files = Path(os.environ.get("ProgramFiles", "C:\\Program Files"))
        search_patterns = [
            local_app / "Python" / "bin" / "python.exe",
            local_app / "Programs" / "Python",
            prog_files,
        ]
        for base in search_patterns:
            if base.is_file():
                candidates.append(str(base))
            elif base.is_dir():
                for child in sorted(base.glob("Python*/python.exe"), reverse=True):
                    candidates.append(str(child))
    else:
        for p in (
            "/usr/local/bin/python3",
            "/opt/homebrew/bin/python3",
            "/usr/bin/python3",
        ):
            if Path(p).is_file():
                candidates.append(p)
        which_py = shutil.which("python3")
        if which_py:
            candidates.append(which_py)

    seen: set[str] = set()
    unique: list[str] = []
    for item in candidates:
        resolved = str(Path(item))
        if resolved not in seen and Path(item).exists() and not _is_windows_apps_stub(resolved):
            seen.add(resolved)
            unique.append(resolved)
    return unique


def _scan_installed_fonts() -> list[str]:
    system = platform.system().lower()
    font_dirs: list[Path] = []
    if system == "windows":
        windir = Path(os.environ.get("WINDIR", "C:\\Windows"))
        local_app = Path(os.environ.get("LOCALAPPDATA", ""))
        font_dirs.extend([windir / "Fonts", local_app / "Microsoft" / "Windows" / "Fonts"])
    elif system == "darwin":
        font_dirs.extend([
            Path("/System/Library/Fonts"),
            Path("/Library/Fonts"),
            Path.home() / "Library" / "Fonts",
        ])
    else:
        font_dirs.extend([
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
            Path.home() / ".local" / "share" / "fonts",
        ])

    existing_names: set[str] = set()
    for folder in font_dirs:
        if not folder.is_dir():
            continue
        try:
            for entry in folder.rglob("*"):
                if entry.is_file():
                    existing_names.add(entry.name.lower())
        except Exception:
            continue

    found: list[str] = []
    for display_name, markers in FONT_CATALOG:
        if any(any(m in fname for fname in existing_names) for m in markers):
            found.append(display_name)
    return found


def probe_environment() -> dict[str, Any]:
    """探测本机 Python、可选库与常用字体情况，返回分级就绪状态。"""
    has_pptx = importlib.util.find_spec("pptx") is not None
    has_pillow = importlib.util.find_spec("PIL") is not None
    installed_fonts = _scan_installed_fonts()

    system = platform.system().lower()
    cjk_priority = ["微软雅黑", "苹方", "思源黑体", "黑体", "等线", "楷体", "宋体"]
    cjk_fallback = next((f for f in cjk_priority if f in installed_fonts), "黑体" if system == "windows" else "苹方")
    latin_fallback = "Calibri" if "Calibri" in installed_fonts else ("Arial" if "Arial" in installed_fonts else "sans-serif")

    real_pythons = _find_real_pythons()
    stub_warning = _is_windows_apps_stub(sys.executable)
    active_python = real_pythons[0] if (stub_warning and real_pythons) else sys.executable

    return {
        "platform": platform.system(),
        "platform_release": platform.release(),
        "python_executable": sys.executable,
        "python_version": platform.python_version(),
        "is_windows_apps_stub": stub_warning,
        "real_python_candidates": real_pythons,
        "has_python_pptx": has_pptx,
        "has_pillow": has_pillow,
        "tier_read_plan_write_text": "READY",
        "tier_build_pptx": "READY" if has_pptx else "NEEDS_OPTIONAL_INSTALL",
        "optional_install_command": f'"{active_python}" -m pip install python-pptx Pillow',
        "optional_install_human_prompt": (
            "本机已可直接读取旧稿、规划页数并撰写逐页文案；"
            "如需自动生成可编辑的 .pptx 成品，请确认是否一键安装轻量写稿组件（拒装也可先拿到完整逐页文案稿）。"
            if not has_pptx
            else "所有读写与生成组件均已就绪。"
        ),
        "installed_fonts": installed_fonts,
        "recommended_cjk_fallback": cjk_fallback,
        "recommended_latin_fallback": latin_fallback,
    }


def main() -> int:
    result = probe_environment()
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
