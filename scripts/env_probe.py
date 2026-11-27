#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""env_probe.py — 零依赖运行环境探针（V0.2.0）。

检测当前 Python 是否为真实可用环境（实测规避 macOS Xcode 假桩与 WindowsApps 空壳）、
OneDrive 同步文档路径、可选写稿组件（python-pptx / Pillow）、跨平台中文字体回退表，
以及三级显式降级状态与三句话确认提示。
"""

from __future__ import annotations

import importlib.util
import json
import os
import platform
import shutil
import subprocess
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
    ("楷体", ("simkai.ttf", "kaiti.ttf", "stkaiti.ttf")),
    ("宋体", ("simsun.ttc", "simsunb.ttf", "songti.ttc")),
    ("仿宋", ("simfang.ttf", "stfangso.ttf")),
    ("等线", ("deng.ttf", "dengb.ttf")),
    ("苹方", ("pingfang.ttc", "pingfang sc.ttc")),
    ("思源黑体", ("sourcehansans", "notosanscjk", "notosanssc")),
    ("Arial", ("arial.ttf", "arial.ttc")),
    ("Calibri", ("calibri.ttf",)),
    ("Segoe UI", ("segoeui.ttf",)),
]

FONT_EQUIVALENCE_MAP: dict[str, list[str]] = {
    "微软雅黑": ["微软雅黑", "苹方", "思源黑体", "黑体", "等线"],
    "Microsoft YaHei": ["微软雅黑", "苹方", "思源黑体", "黑体", "等线"],
    "苹方": ["苹方", "微软雅黑", "思源黑体", "黑体"],
    "PingFang SC": ["苹方", "微软雅黑", "思源黑体", "黑体"],
    "思源黑体": ["思源黑体", "苹方", "微软雅黑", "黑体"],
    "黑体": ["黑体", "微软雅黑", "苹方", "思源黑体"],
    "SimHei": ["黑体", "微软雅黑", "苹方", "思源黑体"],
    "等线": ["等线", "微软雅黑", "苹方", "思源黑体"],
    "楷体": ["楷体", "仿宋", "微软雅黑", "苹方"],
    "KaiTi": ["楷体", "仿宋", "微软雅黑", "苹方"],
    "华文楷体": ["楷体", "仿宋", "苹方"],
    "宋体": ["宋体", "仿宋", "苹方", "微软雅黑"],
    "SimSun": ["宋体", "仿宋", "苹方", "微软雅黑"],
    "华文宋体": ["宋体", "仿宋", "苹方"],
    "Calibri": ["Calibri", "Arial", "Segoe UI"],
    "Arial": ["Arial", "Calibri", "Segoe UI"],
}

FONT_CANONICAL_ALIAS: dict[str, str] = {
    "microsoft yahei": "微软雅黑",
    "microsoft yahei ui": "微软雅黑",
    "pingfang sc": "苹方",
    "pingfang tc": "苹方",
    "source han sans sc": "思源黑体",
    "noto sans cjk sc": "思源黑体",
    "noto sans sc": "思源黑体",
    "simhei": "黑体",
    "dengxian": "等线",
    "kaiti": "楷体",
    "stkaiti": "楷体",
    "simsun": "宋体",
    "songti sc": "宋体",
    "fangsong": "仿宋",
    "stfangsong": "仿宋",
}


def _is_windows_apps_stub(exe_path: str) -> bool:
    norm = exe_path.replace("/", "\\").lower()
    return "windowsapps" in norm


def _is_working_python(exe_path: str) -> bool:
    """实际启动检查 Python 3 是否真实可用，拦截 macOS /usr/bin/python3 Xcode 假桩与空壳。"""
    if not exe_path or _is_windows_apps_stub(exe_path):
        return False
    p = Path(exe_path)
    if not p.exists() or p.is_dir():
        return False
    try:
        proc = subprocess.run(
            [str(p), "-c", "import sys; print(sys.version_info[0])"],
            capture_output=True,
            text=True,
            timeout=3,
        )
        return proc.returncode == 0 and proc.stdout.strip().startswith("3")
    except Exception:
        return False


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
            "/opt/homebrew/bin/python3",
            "/usr/local/bin/python3",
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
        if resolved not in seen and _is_working_python(resolved):
            seen.add(resolved)
            unique.append(resolved)
    return unique


def detect_documents_dir(user_home: Path | None = None) -> dict[str, Any]:
    """探测默认文档目录，Windows 下优先识别 OneDrive 同步后的文档目录。"""
    u_home = user_home or Path.home()
    system = platform.system().lower()
    if system == "windows":
        onedrive_env = os.environ.get("OneDrive") or os.environ.get("OneDriveCommercial") or os.environ.get("OneDriveConsumer")
        onedrive_roots: list[Path] = []
        if onedrive_env:
            onedrive_roots.append(Path(onedrive_env))
        onedrive_roots.append(u_home / "OneDrive")
        for od_root in onedrive_roots:
            for sub_name in ("文档", "Documents"):
                cand = od_root / sub_name
                if cand.is_dir():
                    return {"path": str(cand), "is_onedrive": True, "source": f"onedrive_{sub_name}"}
    return {"path": str(u_home / "Documents"), "is_onedrive": False, "source": "local_documents"}


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


def resolve_safe_font(preferred_font: str, installed_fonts: list[str] | None = None) -> dict[str, Any]:
    """根据本机已装字体映射跨平台安全字体，避免换电脑后文字错行。"""
    fonts = installed_fonts if installed_fonts is not None else _scan_installed_fonts()
    pref = (preferred_font or "").strip()
    canon = FONT_CANONICAL_ALIAS.get(pref.lower(), pref)

    if canon and canon in fonts:
        return {
            "requested_font": preferred_font,
            "resolved_font": canon,
            "mapped": canon != preferred_font,
            "reason": "" if canon == preferred_font else f"本机使用对应名称「{canon}」",
        }

    eq_list = FONT_EQUIVALENCE_MAP.get(preferred_font) or FONT_EQUIVALENCE_MAP.get(canon) or []
    for cand in eq_list:
        if cand in fonts:
            return {
                "requested_font": preferred_font,
                "resolved_font": cand,
                "mapped": True,
                "reason": f"本机未装「{preferred_font or '指定字体'}」，已自动映射为相近字体「{cand}」",
            }

    fallback = next((f for f in ("微软雅黑", "苹方", "思源黑体", "黑体", "Arial", "Calibri") if f in fonts), canon or "微软雅黑")
    return {
        "requested_font": preferred_font,
        "resolved_font": fallback,
        "mapped": fallback != preferred_font,
        "reason": f"本机未找到「{preferred_font or '指定字体'}」，使用安全字体「{fallback}」",
    }


def build_degradation_notice(level: str, reason: str, impact: str, fix_hint: str) -> dict[str, Any]:
    """构造标准三句话显式降级确认：原因 + 影响 + 询问是否接受降级效果。"""
    confirm_q = "请问你能接受先按这个降级效果继续吗？（A. 接受，直接继续 / B. 暂不继续，先按步骤补齐环境）"
    msg = (
        f"1. 原因：{reason}\n"
        f"2. 影响：{impact}\n"
        f"3. 确认：{confirm_q}"
    )
    return {
        "level": level,
        "reason": reason,
        "impact": impact,
        "confirm_question": confirm_q,
        "fix_hint": fix_hint,
        "formatted_message": msg,
    }


def _check_dir_writable_no_create(target_dir: Path) -> bool:
    cur = target_dir
    while not cur.exists():
        if cur.parent == cur:
            return False
        cur = cur.parent
    return cur.is_dir() and os.access(cur, os.W_OK | os.X_OK)


def probe_environment() -> dict[str, Any]:
    """探测本机 Python、可选库、文档路径、可用字体与三级降级状态。"""
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
    docs_info = detect_documents_dir()
    home_writable = _check_dir_writable_no_create(Path.home() / ".ppt-style")

    install_cmd = f'"{active_python}" -m pip install python-pptx Pillow'
    install_user_cmd = f'"{active_python}" -m pip install --user python-pptx Pillow'
    install_mirror_cmd = f'"{active_python}" -m pip install --user -i https://pypi.tuna.tsinghua.edu.cn/simple python-pptx Pillow'

    degradation_tier = "FULL"
    degradation_notice: dict[str, Any] | None = None
    if not home_writable:
        degradation_tier = "TIER2_SANDBOX"
        degradation_notice = build_degradation_notice(
            "TIER2_SANDBOX",
            "当前软件权限受限，无法写入用户个人全局档案目录。",
            "建档、写稿和出图都还能在当前打开的文件夹内正常用，只是换到别的文件夹时需要你手动指一下档案位置。",
            "如需跨文件夹通用，请放开当前软件对个人主目录的写入权限。",
        )
    elif not has_pptx:
        degradation_tier = "TIER1_LIGHT"
        degradation_notice = build_degradation_notice(
            "TIER1_LIGHT",
            "这台电脑还没有安装用来自动生成可编辑 PPT 文件的小组件。",
            "读旧稿、提炼风格、定页面大纲、写逐页内容和比对修改都能正常用；只是没法直接导出 .pptx 文件，会改成交给你带清晰排版说明的逐页文案稿。",
            f"标准安装：{install_cmd} ；免管理员：{install_user_cmd} ；国内网络加速：{install_mirror_cmd}",
        )

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
        "optional_install_command": install_cmd,
        "optional_install_user_command": install_user_cmd,
        "optional_install_mirror_command": install_mirror_cmd,
        "optional_install_human_prompt": (
            degradation_notice["formatted_message"]
            if degradation_notice is not None
            else "所有读写与生成组件均已就绪。"
        ),
        "documents_dir": docs_info,
        "home_dir_writable": home_writable,
        "degradation_tier": degradation_tier,
        "degradation_notice": degradation_notice,
        "installed_fonts": installed_fonts,
        "font_equivalence_map": FONT_EQUIVALENCE_MAP,
        "recommended_cjk_fallback": cjk_fallback,
        "recommended_latin_fallback": latin_fallback,
    }


def main() -> int:
    result = probe_environment()
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
