#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""workspace_manager.py — PPT-style 工坊操作中心、状态机、多档案与反悔管理器（零第三方依赖）。

子命令:
  init            初始化工坊目录（支持空白对话默认落点 文档/我的PPT风格工坊/）并登记全局指针
  status          读取当前或最近活跃工坊进度（支持跨窗口“继续”）
  scan-dir        对本地整个 PPT 文件夹先做快速只读分诊（识别规模、多作者署名、重复文件并推荐引用/复制模式）
  intake          双通道只读归档或轻量引用建档（精确去重、格式拦截、>50MB剥离视频流、合写连续页L1识别）
  scan            运行本地深层体检并保存浓缩摘要
  plan-questions  按4A/双保险规则拆出“核心≤2轮必答题”与“剩余长尾题（精确报数）”
  next-batch      从状态机取下一批（脚本硬限每次最多5题）
  record-answer   保存回答并写入反悔栈；核心轮结束返回 CHECKPOINT_CORE_DONE 与精确剩余题数
  choose-path     记录4A用户拍板：ship_now（先开用，细节留改稿学）或 continue_optional
  undo-last       撤回上一批回答（修改上一题）
  profile         多档案管理（list/switch/split-mask/swap-visual-shell/delete-rule/promote-rule/rollback/reset/export-share）
  render-hub      刷新 00-操作中心.md 与 03-我的PPT风格说明书.md
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import output_guard
import pptx_inspector
import env_probe


CATEGORY_DIR_MAP = {
    "full_personal": "完整个人代表作",
    "partial_collab": "局部合写代表作",
    "company_template": "公司固定模板",
}
MEDIA_AV_EXTENSIONS = {".mp4", ".mov", ".avi", ".wmv", ".mkv", ".mp3", ".wav", ".m4a"}


def _home_dir(override: str | None = None) -> Path:
    if override:
        return Path(override)
    env_home = os.environ.get("PPT_STYLE_HOME")
    if env_home:
        return Path(env_home)
    return Path.home() / ".ppt-style"


def _registry_path(home_dir: Path) -> Path:
    return home_dir / "registry.json"


def _load_json(path: Path, default: Any) -> Any:
    if path.is_file():
        return json.loads(path.read_text(encoding="utf-8"))
    return copy.deepcopy(default)


def _save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _parse_slide_ranges(spec: str | None) -> list[int]:
    if not spec:
        return []
    pages: set[int] = set()
    for part in re.split(r"[,，\s]+", spec.strip()):
        if not part:
            continue
        m = re.match(r"^(\d+)[-~～到](\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a <= b:
                pages.update(range(a, b + 1))
        elif part.isdigit():
            pages.add(int(part))
    return sorted(pages)


def _has_contiguous_run(pages: list[int], min_len: int = 5) -> bool:
    if len(pages) < min_len:
        return False
    run = 1
    for i in range(1, len(pages)):
        if pages[i] == pages[i - 1] + 1:
            run += 1
            if run >= min_len:
                return True
        else:
            run = 1
    return False


def _ensure_default_profile(home_dir: Path, profile_id: str = "default") -> Path:
    pdir = home_dir / "profiles" / profile_id
    (pdir / "masks" / "templates").mkdir(parents=True, exist_ok=True)
    (pdir / "history").mkdir(parents=True, exist_ok=True)
    l1_p = pdir / "core_l1_planning.json"
    l2_p = pdir / "core_l2_content.json"
    mask_p = pdir / "masks" / "default.json"
    if not l1_p.is_file():
        _save_json(l1_p, {"version": "2.0", "rules": [], "scenarios": [], "settings": {}})
    if not l2_p.is_file():
        _save_json(l2_p, {"version": "2.0", "rules": [], "frameworks": [], "settings": {"notes_dual_channel": False}})
    if not mask_p.is_file():
        _save_json(mask_p, {
            "mask_id": "default",
            "mask_name": "默认汇报风格",
            "version": "2.0",
            "rules": [],
            "template_pptx": "",
            "theme_tokens": {},
            "binding_mode": "bound",
            "bound_scenarios": [],
            "settings": {"clean_grid_enabled": True},
        })
    return pdir


def _snapshot_profile(home_dir: Path, profile_id: str, reason: str) -> str:
    pdir = _ensure_default_profile(home_dir, profile_id)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    snap_data = {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "reason": reason,
        "core_l1": _load_json(pdir / "core_l1_planning.json", {}),
        "core_l2": _load_json(pdir / "core_l2_content.json", {}),
        "masks": {
            mp.stem: _load_json(mp, {})
            for mp in sorted((pdir / "masks").glob("*.json"))
        },
    }
    snap_path = pdir / "history" / f"snap-{ts}.json"
    _save_json(snap_path, snap_data)
    return str(snap_path)


def _find_git_root(path: Path) -> Path | None:
    cur = path.resolve() if path.exists() else path
    for cand in (cur, *cur.parents):
        if (cand / ".git").exists():
            return cand
    return None


def _is_temp_upload_path(src_path: Path) -> bool:
    norm = str(src_path).replace("\\", "/").lower()
    markers = ("/uploads/", "/browser_sandbox/uploads/", "/chat_uploads/", "/temp_uploads/", "/.gemini/antigravity/")
    return any(m in norm for m in markers)


def _workshop_dir(workspace: str | None, default_home: bool = False) -> Path:
    if default_home or not workspace:
        docs = env_probe.detect_documents_dir()
        return Path(docs["path"]) / "我的PPT风格工坊"
    ws = Path(workspace)
    if ws.name in ("PPT风格工坊", "我的PPT风格工坊"):
        return ws
    return ws / "PPT风格工坊"


def init_workshop(workspace: str | None, default_home: bool, home_dir: Path) -> dict[str, Any]:
    wdir = _workshop_dir(workspace, default_home=default_home)
    if not env_probe._check_dir_writable_no_create(wdir):
        return {
            "status": "ERROR_NOT_WRITABLE",
            "workshop_dir": str(wdir),
            "human_hint": f"无法在「{wdir}」创建工坊（没有写入权限），请换一个你有写入权限的文件夹，或直接用默认个人文档目录。",
        }

    archive = wdir / "01-原始代表作归档"
    for sub in CATEGORY_DIR_MAP.values():
        (archive / sub).mkdir(parents=True, exist_ok=True)
    (wdir / "02-试炼样稿与调优").mkdir(parents=True, exist_ok=True)

    git_root = _find_git_root(wdir)
    inside_git = git_root is not None
    git_warning = ""
    if inside_git:
        gi_path = wdir / ".gitignore"
        if not gi_path.is_file():
            gi_path.write_text("*\n!.gitignore\n", encoding="utf-8")
        git_warning = (
            "检测到你指定的目录位于项目版本库内，我已在工坊目录自动配好忽略清单防误上传；"
            "如果只是想放个人工作区，也随时可以改放个人文档目录。"
        )

    manifest_p = archive / ".archive-manifest.json"
    if not manifest_p.is_file():
        _save_json(manifest_p, {"items": []})

    state_p = wdir / ".wizard-state.json"
    if not state_p.is_file():
        _save_json(state_p, {
            "step": 1,
            "step_label": "收集代表作",
            "active_profile": "default",
            "active_mask": "default",
            "preview_done": False,
            "core_rounds_done": 0,
            "optional_decision": None,
            "auto_confirmed_rules": [],
            "core_queue": [],
            "optional_queue": [],
            "answered_questions": [],
            "undo_stack": [],
            "scan_summary": None,
        })

    _ensure_default_profile(home_dir, "default")
    reg_p = _registry_path(home_dir)
    reg = _load_json(reg_p, {"active_workspace": "", "active_profile": "default", "active_mask": "default", "recent_workshops": []})
    reg["active_workspace"] = str(wdir.resolve())
    if str(wdir.resolve()) not in reg["recent_workshops"]:
        reg["recent_workshops"].append(str(wdir.resolve()))
    _save_json(reg_p, reg)

    render_hub(wdir, home_dir)
    return {
        "status": "INITIALIZED",
        "workshop_dir": str(wdir.resolve()),
        "registry_path": str(reg_p),
        "inside_git_repo": inside_git,
        "git_warning_hint": git_warning,
    }


def _copy_with_media_strip(src: Path, dst: Path, strip_av: bool) -> bool:
    if not strip_av:
        shutil.copy2(src, dst)
        return False
    stripped = False
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename.startswith("ppt/media/") and Path(item.filename).suffix.lower() in MEDIA_AV_EXTENSIONS:
                stripped = True
                continue
            zout.writestr(item, zin.read(item.filename))
    return stripped


AUTHOR_STOPWORDS = {
    "方案", "推广", "营销", "策划", "视频", "文化", "合作", "年度", "品牌",
    "内容", "美妆", "平台", "系列", "更新", "部分", "双十一", "第二版", "最终版",
    "结案", "招商", "活动", "测试", "站内", "站外", "汇总", "报告", "简介", "大纲",
    "新品", "超级", "话题", "栏目", "直播", "传播", "战略", "规划", "创意", "日报", "周报", "月报",
}


def triage_paths(paths_list: list[str]) -> dict[str, Any]:
    """对目录或文件列表做多信号预分诊：分出绿组(纯个人)/黄组(疑似合写)/灰组(纯模板)/待处理格式异常。"""
    files: list[Path] = []
    for raw in paths_list:
        p = Path(raw)
        if p.is_dir() and p.suffix.lower() == ".key":
            files.append(p)
        elif p.is_dir():
            files.extend(sorted(x for x in p.iterdir() if x.is_file() or (x.is_dir() and x.suffix.lower() == ".key")))
        elif p.exists():
            files.append(p)

    green_group: list[dict[str, Any]] = []
    yellow_group: list[dict[str, Any]] = []
    gray_group: list[dict[str, Any]] = []
    needs_attention: list[dict[str, Any]] = []

    for fp in files:
        chk = pptx_inspector.check_file(str(fp))
        st = chk.get("status")
        if st != "VALID_PPTX":
            needs_attention.append({"file": fp.name, "path": str(fp), "status": st, "human_hint": chk.get("human_hint", "")})
            continue
        insp = pptx_inspector.inspect_deck(str(fp))
        tg = insp.get("triage_group", "green_likely_personal")
        item = {
            "file": fp.name,
            "path": str(fp),
            "total_slides": insp.get("total_slides", 0),
            "triage_group": tg,
            "triage_reason": insp.get("triage_reason", ""),
            "collab_signals": insp.get("collab_signals", []),
            "chapter_blocks": insp.get("chapter_blocks", [])[:6],
        }
        if tg == "yellow_suspected_collab":
            yellow_group.append(item)
        elif tg == "gray_template":
            gray_group.append(item)
        else:
            green_group.append(item)

    yellow_prompt = ""
    if yellow_group:
        names = "、".join(f"《{x['file']}》（{x['triage_reason']}）" for x in yellow_group[:3])
        yellow_prompt = (
            f"检测到 {len(yellow_group)} 份疑似多人合写稿：{names}。\n"
            "提醒：提炼个人风格尽量不要把同事写的页面混进来，否则容易学偏。\n"
            "请确认怎么处理这些合写稿（共 1 个问题）：\n"
            "A. 一键去掉，这次暂不纳入提炼（推荐最省心）\n"
            "B. 保留但标出哪些页是我写的（直接报页码或关键词，不拆原文件）\n"
            "C. 我确认整份都是我自己写的"
        )

    return {
        "status": "TRIAGED",
        "total_scanned": len(files),
        "green_likely_personal": green_group,
        "yellow_suspected_collab": yellow_group,
        "gray_template": gray_group,
        "needs_attention": needs_attention,
        "human_triage_prompt": yellow_prompt,
    }


def scan_directory(dir_path_str: str) -> dict[str, Any]:
    """对本地整目录做只读分诊：统计格式分布、总大小、精确重复、多作者署名与合写预分诊。"""
    target = Path(dir_path_str)
    if not target.is_dir():
        return {"status": "ERROR", "message": f"找不到目录：{dir_path_str}"}

    files = sorted(p for p in target.iterdir() if p.is_file() or (p.is_dir() and p.suffix.lower() == ".key"))
    valid_pptx = 0
    visual_ref = 0
    needs_convert = 0
    suspected_drm = 0
    other_files = 0
    total_bytes = 0
    sha_seen: dict[str, str] = {}
    exact_duplicates: list[dict[str, str]] = []
    heavy_av_files: list[str] = []
    heavy_image_only_files: list[str] = []
    author_counts: dict[str, int] = {}

    for p in files:
        if p.is_file():
            total_bytes += p.stat().st_size
        chk = pptx_inspector.check_file(str(p))
        st = chk.get("status")
        if st == "VALID_PPTX":
            valid_pptx += 1
            sha = chk.get("sha256", "")
            if sha:
                if sha in sha_seen:
                    exact_duplicates.append({"duplicate": p.name, "existing": sha_seen[sha]})
                else:
                    sha_seen[sha] = p.name
            if chk.get("has_av_to_strip"):
                heavy_av_files.append(p.name)
            elif chk.get("has_heavy_media"):
                heavy_image_only_files.append(p.name)
        elif st == "VISUAL_REF_ONLY":
            visual_ref += 1
        elif st == "NEEDS_USER_CONVERT":
            needs_convert += 1
        elif st == "SUSPECTED_DRM_ENCRYPTED":
            suspected_drm += 1
        else:
            other_files += 1

        for m in re.finditer(r"(?<=[-_—])([\u4e00-\u9fa5]{2,3})(?=[-_—\.\d\s]|$)", p.stem):
            cand = m.group(1)
            if cand not in AUTHOR_STOPWORDS and not any(sw in cand for sw in AUTHOR_STOPWORDS):
                author_counts[cand] = author_counts.get(cand, 0) + 1

    total_mb = round(total_bytes / (1024 * 1024), 1)
    rec_mode = "reference" if (total_mb > 500.0 or valid_pptx > 20) else "copy"
    detected_authors = sorted(author_counts.items(), key=lambda kv: (-kv[1], kv[0]))
    triage_res = triage_paths([str(target)]) if valid_pptx <= 30 else None

    hint_parts = [
        f"共扫描 {len(files)} 个文件（合计约 {total_mb} MB）：其中 {valid_pptx} 份可提炼的演示稿、{visual_ref} 份 PDF/图片参考、{needs_convert} 份需先另存为格式。"
    ]
    if suspected_drm:
        hint_parts.append(f"另有 {suspected_drm} 份文件疑似被公司加密软件锁定。")
    if exact_duplicates:
        hint_parts.append(f"发现 {len(exact_duplicates)} 对完全重复文件。")
    if len(detected_authors) >= 2:
        names_preview = "、".join(f"{n}({c}份)" for n, c in detected_authors[:5])
        hint_parts.append(f"文件名里出现多位不同撰写人署名（{names_preview}），建议按本人署名筛选或先挑代表作，避免多人风格混杂。")
    if rec_mode == "reference":
        hint_parts.append("由于样本数量或体积较大，建议使用只读引用模式建档（不往工坊重复复制几个 G 原文件）。")

    return {
        "status": "DIR_TRIAGED",
        "directory": str(target.resolve()),
        "total_files": len(files),
        "total_mb": total_mb,
        "valid_pptx_count": valid_pptx,
        "visual_ref_count": visual_ref,
        "needs_convert_count": needs_convert,
        "suspected_drm_count": suspected_drm,
        "other_count": other_files,
        "exact_duplicates": exact_duplicates,
        "heavy_av_count": len(heavy_av_files),
        "heavy_image_only_count": len(heavy_image_only_files),
        "detected_authors": [{"name": n, "count": c} for n, c in detected_authors],
        "recommended_copy_mode": rec_mode,
        "collab_triage": triage_res,
        "human_hint": " ".join(hint_parts),
    }


def intake_file(
    wdir: Path,
    src_file: str,
    category: str,
    slides_spec: str | None,
    external_toc: bool,
    home_dir: Path,
    copy_mode: str = "copy",
    from_upload: bool = False,
) -> dict[str, Any]:
    if category not in CATEGORY_DIR_MAP:
        return {"status": "ERROR", "message": f"未知分类：{category}"}
    if copy_mode not in ("copy", "reference"):
        return {"status": "ERROR", "message": f"不支持的收样存储模式：{copy_mode}"}

    chk = pptx_inspector.check_file(src_file)
    if chk.get("status") in ("MISSING", "NEEDS_USER_CONVERT", "SUSPECTED_DRM_ENCRYPTED", "UNSUPPORTED", "CORRUPTED"):
        return {**chk, "status": "REJECTED", "reject_reason": chk.get("status")}

    manifest_p = wdir / "01-原始代表作归档" / ".archive-manifest.json"
    manifest = _load_json(manifest_p, {"items": []})
    items: list[dict[str, Any]] = manifest["items"]

    src_path = Path(src_file).resolve()
    forced_copy = bool(from_upload or _is_temp_upload_path(src_path))
    eff_copy_mode = "copy" if forced_copy else copy_mode
    file_sha = chk.get("sha256")
    if not file_sha:
        h = hashlib.sha256()
        h.update(src_path.read_bytes())
        file_sha = h.hexdigest()

    for existing in items:
        if existing.get("sha256") == file_sha:
            return {
                "status": "DUPLICATE_SKIPPED",
                "existing_name": existing["archived_name"],
                "human_hint": f"《{src_path.name}》与已归档的《{existing['archived_name']}》内容完全相同，已自动跳过，不会重复算权重。",
            }

    if eff_copy_mode == "reference":
        dst_path = src_path
        if chk.get("status") == "VISUAL_REF_ONLY":
            entry = {
                "file_name": src_path.name,
                "archived_name": src_path.name,
                "path": str(src_path),
                "sha256": file_sha,
                "category": category,
                "storage_mode": "reference",
                "forced_copy_from_upload": forced_copy,
                "format_status": "VISUAL_REF_ONLY",
                "selected_slides": [],
                "module_l1_eligible": False,
                "external_toc": False,
            }
            items.append(entry)
            _save_json(manifest_p, manifest)
            render_hub(wdir, home_dir)
            return {"status": "ARCHIVED_VISUAL_REF", **entry, "human_hint": chk["human_hint"]}
        stripped = False
    else:
        sub_dir = wdir / "01-原始代表作归档" / CATEGORY_DIR_MAP[category]
        sub_dir.mkdir(parents=True, exist_ok=True)
        dst_path = sub_dir / src_path.name
        if dst_path.exists():
            dst_path = sub_dir / f"{src_path.stem}-{file_sha[:6]}{src_path.suffix}"

        if chk.get("status") == "VISUAL_REF_ONLY":
            shutil.copy2(src_path, dst_path)
            entry = {
                "file_name": src_path.name,
                "archived_name": dst_path.name,
                "path": str(dst_path),
                "sha256": file_sha,
                "category": category,
                "storage_mode": "copy",
                "forced_copy_from_upload": forced_copy,
                "format_status": "VISUAL_REF_ONLY",
                "selected_slides": [],
                "module_l1_eligible": False,
                "external_toc": False,
            }
            items.append(entry)
            _save_json(manifest_p, manifest)
            render_hub(wdir, home_dir)
            return {"status": "ARCHIVED_VISUAL_REF", **entry, "human_hint": chk["human_hint"]}

        strip_av = bool(chk.get("has_av_to_strip", False))
        stripped = _copy_with_media_strip(src_path, dst_path, strip_av=strip_av)

    selected = _parse_slide_ranges(slides_spec)
    module_l1 = (category == "partial_collab") and _has_contiguous_run(selected, 5)
    deck_info = pptx_inspector.inspect_deck(str(dst_path), selected_slides=set(selected) if selected else None)

    entry = {
        "file_name": src_path.name,
        "archived_name": dst_path.name,
        "path": str(dst_path),
        "sha256": file_sha,
        "category": category,
        "storage_mode": eff_copy_mode,
        "forced_copy_from_upload": forced_copy,
        "format_status": "VALID_PPTX",
        "media_stripped": stripped,
        "total_slides": deck_info.get("total_slides", 0),
        "selected_slides": selected,
        "module_l1_eligible": module_l1,
        "external_toc": external_toc,
        "chapter_blocks": deck_info.get("chapter_blocks", []),
        "anomaly_candidates": deck_info.get("anomaly_candidates", []),
    }
    items.append(entry)
    _save_json(manifest_p, manifest)
    render_hub(wdir, home_dir)
    return {"status": "ARCHIVED", **entry}


def scan_workshop(wdir: Path, home_dir: Path) -> dict[str, Any]:
    manifest = _load_json(wdir / "01-原始代表作归档" / ".archive-manifest.json", {"items": []})
    pptx_specs = [
        {
            "path": it["path"],
            "category": it["category"],
            "selected_slides": it.get("selected_slides") or None,
            "module_l1_eligible": it.get("module_l1_eligible", False),
        }
        for it in manifest["items"]
        if it.get("format_status") == "VALID_PPTX"
    ]
    if not pptx_specs:
        return {
            "status": "NO_VALID_PPTX",
            "human_hint": "建档至少需要 1 份 .pptx 格式的代表作，请先放入 1 份你写过或认可的 .pptx 文件。",
        }

    scan_res = pptx_inspector.scan_corpus(pptx_specs)
    state_p = wdir / ".wizard-state.json"
    state = _load_json(state_p, {})
    state["scan_summary"] = scan_res["summary"]
    state["step"] = max(state.get("step", 1), 2)
    state["step_label"] = "体检完成与风格预览"
    _save_json(state_p, state)
    render_hub(wdir, home_dir)
    return {"status": "SCANNED", **scan_res}


def plan_questions(wdir: Path, home_dir: Path) -> dict[str, Any]:
    state_p = wdir / ".wizard-state.json"
    state = _load_json(state_p, {})
    summary = state.get("scan_summary")
    if not summary:
        scan_out = scan_workshop(wdir, home_dir)
        if scan_out.get("status") != "SCANNED":
            return scan_out
        state = _load_json(state_p, {})
        summary = state["scan_summary"]

    indep_decks = summary.get("independent_active_decks", 0)
    auto_confirmed: list[dict[str, Any]] = []
    core_queue: list[dict[str, Any]] = []
    optional_queue: list[dict[str, Any]] = []

    # 1. Double-insurance auto-confirm check: ratio >= 0.85 AND independent_decks >= 5
    title_ratio = summary.get("conclusion_title_ratio", 0.0)
    samples = summary.get("conclusion_title_samples", [])
    ref_str = (
        f"如《{samples[0]['file']}》第 {samples[0]['slide_no']} 页"
        if samples
        else "代表作正文页"
    )

    if title_ratio >= 0.85 and indep_decks >= 5:
        auto_confirmed.append({
            "rule_id": "R-TITLE-CONCLUSION",
            "layer": "L2",
            "human_text": f"标题习惯采用“主题＋核心结论”句式（在 {indep_decks} 份样本中出现率达 {int(title_ratio * 100)}%，{ref_str}）",
        })
    else:
        core_queue.append({
            "id": "Q_CORE_TITLE",
            "layer": "L2",
            "title": f"我看到{ref_str}等页面用了带结论的长标题（占比 {int(title_ratio * 100)}%），平时你更喜欢哪种标题写法？",
            "options": [
                "A. 标题直接写明结论（主题＋结论，推荐）",
                "B. 短词标签标题，把结论放正文第一句",
                "C. 按汇报对象切换：日常汇报用短标题，领导复盘写结论句",
                "D. 都不太像，我口语补充一句",
            ],
            "recommended": "A" if title_ratio >= 0.5 else "C",
        })

    # L1 Planning skeleton questions
    l1_decks = summary.get("l1_eligible_decks", 0)
    if l1_decks == 0:
        core_queue.append({
            "id": "Q_CORE_L1_COLD",
            "layer": "L1",
            "title": "由于本次样本里没有整份个人独立完整稿，日常做一套新汇报时，你通常怎么定整体骨架？",
            "options": [
                "A. 先抛核心结论与关键数字，再展开原因与行动计划（推荐）",
                "B. 先讲背景痛点与现状，最后引出方案与预期价值",
                "C. 按汇报类型自动切换：复盘先摆数据、方案先立痛点",
                "D. 都不太像，我口语补充一句",
            ],
            "recommended": "A",
        })
    else:
        core_queue.append({
            "id": "Q_CORE_L1_SPLIT",
            "layer": "L1",
            "title": f"在整份规划节奏上（已参考 {l1_decks} 份完整章节/稿件，中位数约 {summary.get('l1_median_pages', 5)} 页），遇到内容较多时你通常怎么拆页？",
            "options": [
                "A. 宁可一页把一个完整事项讲透，不轻易拆碎",
                "B. 把数据明细页和管理结论页拆开，一页专讲一件事",
                "C. 按场景切换：例会一页闭环，专题立项拆多页（推荐）",
                "D. 都不太像，我口语补充一句",
            ],
            "recommended": "C",
        })

    # L2 Evidence & Tone
    core_queue.append({
        "id": "Q_CORE_L2_TONE",
        "layer": "L2",
        "title": "写单页文字和摆论据时，你对数字严谨度和语言风格最看重哪条？",
        "options": [
            "A. 数据页必须带齐时间和口径，对外客户提案放行感染力文案（推荐）",
            "B. 全部页面一律极简克制大白话，严禁任何夸大修饰词",
            "C. 偏视觉与金字塔观点为主，少贴大段明细数字",
            "D. 都不太像，我口语补充一句",
        ],
        "recommended": "A",
    })

    # L3 Visual grid / hygiene (3A)
    if summary.get("recommend_clean_grid"):
        h_samples = summary.get("hygiene_samples", [])
        h_ref = f"如《{h_samples[0]['file']}》第 {h_samples[0]['slide_no']} 页" if h_samples else "部分旧稿"
        core_queue.append({
            "id": "Q_CORE_L3_GRID",
            "layer": "L3",
            "title": f"体检发现{h_ref}存在部分字号偏小或靠边太紧的情况，以后按你的风格排版时优先采用哪种？",
            "options": [
                "A. 保留原有配色与分区骨架，自动用干净安全留白微调（推荐）",
                "B. 100% 完全照搬旧稿原本的字号和紧凑位置",
                "C. 都不太像，我口语补充一句",
            ],
            "recommended": "A",
        })

    # L3 Template shell vs personal style
    bg_cn = summary.get("dominant_bg_color_cn", "纯白")
    hdr_cn = summary.get("top_table_header_color_cn") or "默认"
    core_queue.append({
        "id": "Q_CORE_L3_SHELL",
        "layer": "L3",
        "title": f"样本中常见{bg_cn}底色和{hdr_cn}表格样式，这属于你个人偏爱还是公司固定模板要求？",
        "options": [
            "A. 配色和页眉属于公司硬规定，新稿继续套用但内部排版按我习惯（推荐）",
            "B. 这就是我平时最顺手的个人日常配色，换模板也可沿用",
            "C. 仅作历史包袱剥离，新稿改用更清爽通用配色",
            "D. 都不太像，我口语补充一句",
        ],
        "recommended": "A",
    })

    # Optional split-mask check ONLY when bimodal detected
    if summary.get("bimodal_style_detected"):
        core_queue.append({
            "id": "Q_CORE_SPLIT_MASK",
            "layer": "L3",
            "title": f"系统检测到{summary.get('bimodal_reason', '样本明显分成两派')}，需要拆成两套日常风格面具吗？",
            "options": [
                "A. 拆成两套（共用同一套写作逻辑，仅按场景切换视觉皮肤，推荐）",
                "B. 先按一套主风格走，以后需要再单独新建",
                "C. 都不太像，我口语补充一句",
            ],
            "recommended": "A",
        })

    # Optional detail questions (4A: explicitly counted & user chooses to continue or ship!)
    optional_queue = [
        {
            "id": "Q_OPT_CAPTION",
            "layer": "L2",
            "title": "页面里放系统截图或实景照片时，旁边的说明标签习惯怎么写？",
            "options": [
                "A. 直接写清这张图证明什么结论（任务型短标签，推荐）",
                "B. 只写界面名称，把解释放在正文大块里",
                "C. 不加额外标签，靠图本身说话",
            ],
            "recommended": "A",
        },
        {
            "id": "Q_OPT_SPARSE",
            "layer": "L2",
            "title": "遇到某一页只有一两句短说明或核心观点时，怎么处理版面不显得空？",
            "options": [
                "A. 直接把短说明放大成醒目章节大字号居中/左对齐（推荐）",
                "B. 配一张弱装饰配图或极简边框卡片撑起版面",
                "C. 用较小浅灰字放在角落当备注",
            ],
            "recommended": "A",
        },
        {
            "id": "Q_OPT_CHART_DATA",
            "layer": "L3",
            "title": "展示柱状图或折线图时，图上习惯标出具体数据数字吗？",
            "options": [
                "A. 只在关键峰值或最新一期标数字，保持画面干净（推荐）",
                "B. 每个柱子/节点都标出具体数值",
                "C. 完全不标数字，只看趋势高低",
            ],
            "recommended": "A",
        },
        {
            "id": "Q_OPT_BIG_NUM",
            "layer": "L3",
            "title": "呈现核心变化数字（如效率翻倍或金额突破）时，喜欢用“大数字＋方向箭头”看板吗？",
            "options": [
                "A. 仅在管理总结页用 1～2 组醒目大数字，常规页嵌在句子或表里（推荐）",
                "B. 每页顶部都喜欢摆一排三四个大数字指标卡",
                "C. 很少用大数字块，统一用表格对比",
            ],
            "recommended": "A",
        },
    ]

    state["auto_confirmed_rules"] = auto_confirmed
    state["core_queue"] = core_queue
    state["optional_queue"] = optional_queue
    state["core_rounds_done"] = 0
    state["optional_decision"] = None
    state["step"] = 3
    state["step_label"] = "核心风格问答"
    _save_json(state_p, state)
    render_hub(wdir, home_dir)

    return {
        "status": "PLANNED",
        "auto_confirmed_count": len(auto_confirmed),
        "auto_confirmed_rules": auto_confirmed,
        "core_question_count": len(core_queue),
        "core_estimated_rounds": math.ceil(len(core_queue) / 5),
        "optional_question_count": len(optional_queue),
        "optional_estimated_rounds": math.ceil(len(optional_queue) / 5),
    }


def next_batch(wdir: Path) -> dict[str, Any]:
    state = _load_json(wdir / ".wizard-state.json", {})
    core_q = state.get("core_queue", [])
    opt_q = state.get("optional_queue", [])
    opt_dec = state.get("optional_decision")

    if core_q:
        batch = core_q[:5]
        return {
            "status": "BATCH_READY",
            "phase": "core",
            "question_count": len(batch),
            "questions": batch,
            "remaining_after_batch": max(0, len(core_q) - len(batch)),
        }

    if opt_q and opt_dec is None:
        return {
            "status": "CHECKPOINT_CORE_DONE",
            "remaining_optional_count": len(opt_q),
            "estimated_extra_rounds": math.ceil(len(opt_q) / 5),
            "human_prompt": (
                f"核心大调已定，现在就可以直接生成试炼样稿并开始用。"
                f"除此之外，还剩 {len(opt_q)} 个偏细枝末节的小习惯待确认（预计还需 {math.ceil(len(opt_q) / 5)} 轮）。"
                "请你确认：A. 先出样稿直接开用，这些细节以后改稿时再顺手学（推荐） / B. 继续把剩余问题答完"
            ),
        }

    if opt_q and opt_dec == "continue_optional":
        batch = opt_q[:5]
        return {
            "status": "BATCH_READY",
            "phase": "optional",
            "question_count": len(batch),
            "questions": batch,
            "remaining_after_batch": max(0, len(opt_q) - len(batch)),
        }

    return {"status": "ALL_QUESTIONS_DONE"}


ANSWER_RULE_MAP: dict[str, dict[str, str]] = {
    "Q_CORE_TITLE": {
        "A": "单页标题采用“主题＋核心结论”句式，直接写出该页核心观点",
        "B": "单页标题使用精炼短词标签，把结论放在正文开头首句",
        "C": "标题风格按汇报对象切换：日常同步用短词标签，正式复盘用结论句",
    },
    "Q_CORE_L1_SPLIT": {
        "A": "页数规划优先让每页讲透一个完整事项，不轻易拆碎",
        "B": "页数规划将数据明细页与管理结论页分开，一页专讲一件事",
        "C": "页数规划按场景切换：例会偏向一页闭环，专题立项拆成多页展开",
    },
    "Q_CORE_L1_COLD": {
        "A": "整稿按结论先行组织：先给核心结论与关键数字，再展开原因与行动计划",
        "B": "整稿按递进说服组织：先讲背景痛点与现状，最后引出方案与预期价值",
        "C": "整稿骨架按汇报类型切换：复盘先摆数据结论，方案先立痛点机会",
    },
    "Q_CORE_L2_TONE": {
        "A": "数据页必须带齐时间区间与统计口径，对外提案页允许更有感染力的主张表达",
        "B": "全部页面保持极简克制大白话，不放主观修饰或夸大表达",
        "C": "内容表达以金字塔核心观点和图示结构为主，减少大段明细数字堆砌",
    },
    "Q_CORE_L3_GRID": {
        "A": "保留原有配色与分区骨架，自动套用干净安全留白网格微调防拥挤重叠",
        "B": "严格按历史代表作原有的字号层级与紧凑位置排版",
    },
    "Q_CORE_L3_SHELL": {
        "A": "沿用当前模板固定底色与表头风格作为统一外壳，内部排版按个人习惯组织",
        "B": "保留当前样本的常用底色与表格样式作为个人日常视觉偏好",
        "C": "剥离历史模板视觉包袱，新稿改用更清爽通用的浅底配色",
    },
    "Q_CORE_SPLIT_MASK": {
        "A": "共用同一套页面规划与表达内核，按汇报场景拆出独立视觉皮肤面具",
        "B": "先沿用一套主视觉风格，日后需要时再拆分新面具",
    },
    "Q_OPT_CAPTION": {
        "A": "配图旁边的短标签直接写清这张图证明什么结论",
        "B": "配图标签只标界面或对象名称，解释内容留在正文主区",
        "C": "页面插图尽量靠图本身表达，不加额外文字标签",
    },
    "Q_OPT_SPARSE": {
        "A": "极简观点页直接放大核心短句居中或左对齐突出呈现",
        "B": "极简观点页搭配弱装饰配图或极简边框卡片撑起版面重心",
        "C": "极简说明文字采用克制小字号放置在版面角落作补充",
    },
    "Q_OPT_CHART_DATA": {
        "A": "柱状图与折线图只在关键峰值或最新一期标出数字，保持画面干净",
        "B": "柱状图与折线图的每个柱子或数据节点都标出具体数值",
        "C": "图表仅保留趋势高低与对比关系，不在图内叠写详细数字",
    },
    "Q_OPT_BIG_NUM": {
        "A": "仅在管理总结页使用 1～2 组醒目大数字看板，常规页嵌在句子或表格里",
        "B": "每页顶部高频使用横向一排大数字指标卡提炼核心数据",
        "C": "变化对比优先采用规整表格，较少单独放大数字指标块",
    },
}


def _answer_to_rule_text(qid: str, raw_answer: str) -> str:
    txt = (raw_answer or "").strip()
    m = re.match(r"^\s*([A-D])(?:\s*[\.、\)：:]\s*(.*))?$", txt)
    if m:
        letter = m.group(1).upper()
        rest = (m.group(2) or "").strip()
        if letter == "D":
            if rest and not rest.startswith("都不太像"):
                return rest
            return ""
        mapped = ANSWER_RULE_MAP.get(qid, {}).get(letter)
        if mapped:
            return mapped
        return rest or txt
    return txt


def _upsert_layer_rule(target_p: Path, layer: str, source_id: str, text: str) -> bool:
    san = output_guard.sanitize_rule_text(text, auto_abstract=True)
    clean_text = san["sanitized_text"].strip()
    if not clean_text:
        return False
    data = _load_json(target_p, {"version": "1.0", "rules": [], "settings": {}})
    rules: list[dict[str, Any]] = data.setdefault("rules", [])
    updated = False
    for r in rules:
        if r.get("source_id") == source_id:
            if r.get("text") != clean_text:
                r["text"] = clean_text
                r["updated_at"] = datetime.now().isoformat(timespec="seconds")
            updated = True
            break
    if not updated:
        rules.append({
            "rule_no": len(rules) + 1,
            "layer": layer,
            "source_id": source_id,
            "text": clean_text,
            "added_at": datetime.now().isoformat(timespec="seconds"),
        })
    for i, r in enumerate(rules, 1):
        r["rule_no"] = i
    _save_json(target_p, data)
    return True


def _sync_wizard_answers_to_profile(wdir: Path, state: dict[str, Any], home_dir: Path) -> None:
    """将向导中用户已确认的规律、L1场景、L2多框架及L3多母版皮肤自动沉淀入激活档案。"""
    pid = state.get("active_profile", "default")
    mid = state.get("active_mask", "default")
    pdir = _ensure_default_profile(home_dir, pid)
    layer_paths = {
        "L1": pdir / "core_l1_planning.json",
        "L2": pdir / "core_l2_content.json",
        "L3": pdir / "masks" / f"{mid}.json",
    }

    for auto_r in state.get("auto_confirmed_rules", []):
        layer = auto_r.get("layer", "L2")
        sid = auto_r.get("rule_id") or f"AUTO_{layer}"
        _upsert_layer_rule(layer_paths.get(layer, layer_paths["L2"]), layer, sid, auto_r.get("human_text", ""))

    for ans in state.get("answered_questions", []):
        layer = ans.get("layer", "L2")
        qid = ans.get("id", "")
        rule_txt = _answer_to_rule_text(qid, ans.get("answer", ""))
        if rule_txt and qid:
            _upsert_layer_rule(layer_paths.get(layer, layer_paths["L2"]), layer, qid, rule_txt)

    summary = state.get("scan_summary") or {}
    if summary:
        # 1. L1 场景库同步
        l1_data = _load_json(layer_paths["L1"], {"version": "2.0", "rules": [], "scenarios": [], "settings": {}})
        existing_sc_map = {s["scenario_id"]: s for s in l1_data.get("scenarios", []) if s.get("scenario_id")}
        for sc in summary.get("l1_scenarios", []):
            sc_id = sc["scenario_id"]
            if sc_id in existing_sc_map:
                existing_sc_map[sc_id].update({
                    "deck_count": sc.get("deck_count", 0),
                    "sample_files": sc.get("sample_files", []),
                    "median_pages": sc.get("median_pages", 0),
                })
            else:
                existing_sc_map[sc_id] = dict(sc)
        l1_data["scenarios"] = list(existing_sc_map.values())
        _save_json(layer_paths["L1"], l1_data)

        # 2. L2 按场景多框架同步
        l2_data = _load_json(layer_paths["L2"], {"version": "2.0", "rules": [], "frameworks": [], "settings": {}})
        l2_data.setdefault("settings", {})["notes_dual_channel"] = bool(summary.get("notes_dual_channel_recommended", False))
        existing_fw_map = {f["framework_id"]: f for f in l2_data.get("frameworks", []) if f.get("framework_id")}
        for fw in summary.get("l2_frameworks", []):
            fw_id = fw["framework_id"]
            if fw_id in existing_fw_map:
                # 保留用户已经 confirm-framework 拍板过的状态
                if existing_fw_map[fw_id].get("status") in ("confirmed", "adjusted_confirmed", "dropped"):
                    continue
                existing_fw_map[fw_id].update(fw)
            else:
                existing_fw_map[fw_id] = dict(fw)
        l2_data["frameworks"] = list(existing_fw_map.values())
        _save_json(layer_paths["L2"], l2_data)

        # 3. L3 多套真实母版皮肤与绑定关系同步
        clean_grid_pref = True
        for ans in state.get("answered_questions", []):
            if ans.get("id") == "Q_CORE_L3_GRID":
                clean_grid_pref = not ans.get("answer", "").strip().upper().startswith("B")

        skins = summary.get("l3_visual_skins") or []
        all_sc_ids = [sc["scenario_id"] for sc in l1_data.get("scenarios", [])]
        if skins:
            templates_dir = pdir / "masks" / "templates"
            templates_dir.mkdir(parents=True, exist_ok=True)
            for s_idx, skin in enumerate(skins):
                skin_id = mid if s_idx == 0 else skin.get("skin_id", f"skin_{s_idx + 1}")
                m_path = pdir / "masks" / f"{skin_id}.json"
                m_data = _load_json(m_path, {
                    "mask_id": skin_id,
                    "mask_name": skin.get("label_cn", "汇报皮肤"),
                    "version": "2.0",
                    "rules": [],
                    "template_pptx": "",
                    "theme_tokens": {},
                    "binding_mode": "bound" if s_idx == 0 else "independent",
                    "bound_scenarios": all_sc_ids if s_idx == 0 else [],
                    "settings": {"clean_grid_enabled": clean_grid_pref},
                })
                m_data["mask_name"] = m_data.get("mask_name") or skin.get("label_cn", "汇报皮肤")
                m_data["theme_tokens"] = skin.get("theme_tokens") or m_data.get("theme_tokens", {})
                if not m_data.get("bound_scenarios") and m_data.get("binding_mode") == "bound":
                    m_data["bound_scenarios"] = all_sc_ids
                sample_src = skin.get("sample_deck_path", "")
                if sample_src and Path(sample_src).is_file():
                    tpl_out = templates_dir / f"base_{skin_id}.pptx"
                    ext_res = pptx_inspector.extract_clean_master_pptx(sample_src, str(tpl_out))
                    if ext_res.get("template_path"):
                        m_data["template_pptx"] = ext_res["template_path"]
                m_settings = m_data.setdefault("settings", {})
                m_settings["clean_grid_enabled"] = clean_grid_pref
                for k in (
                    "top_title_font",
                    "top_body_font",
                    "median_title_size_pt",
                    "median_body_size_pt",
                    "dominant_bg_color_cn",
                    "top_accent_colors_cn",
                ):
                    if k in summary:
                        m_settings[k] = summary[k]
                _save_json(m_path, m_data)
        else:
            mask_data = _load_json(layer_paths["L3"], {"mask_id": mid, "mask_name": "默认汇报风格", "version": "2.0", "rules": [], "settings": {}})
            m_settings = mask_data.setdefault("settings", {})
            m_settings["clean_grid_enabled"] = clean_grid_pref
            for k in (
                "top_title_font",
                "top_body_font",
                "median_title_size_pt",
                "median_body_size_pt",
                "dominant_bg_color_cn",
                "top_accent_colors_cn",
            ):
                if k in summary:
                    m_settings[k] = summary[k]
            _save_json(layer_paths["L3"], mask_data)


def record_answer(wdir: Path, answers: dict[str, str], home_dir: Path) -> dict[str, Any]:
    state_p = wdir / ".wizard-state.json"
    state = _load_json(state_p, {})
    core_q: list[dict[str, Any]] = state.get("core_queue", [])
    opt_q: list[dict[str, Any]] = state.get("optional_queue", [])

    # Save full pre-answer snapshot to undo_stack
    undo_entry = {
        "core_queue": copy.deepcopy(core_q),
        "optional_queue": copy.deepcopy(opt_q),
        "answered_questions": copy.deepcopy(state.get("answered_questions", [])),
        "core_rounds_done": state.get("core_rounds_done", 0),
        "answers_batch": answers,
    }
    state.setdefault("undo_stack", []).append(undo_entry)

    answered_ids = set(answers.keys())
    consumed_core = [q for q in core_q if q["id"] in answered_ids]
    consumed_opt = [q for q in opt_q if q["id"] in answered_ids]

    for q in consumed_core + consumed_opt:
        state.setdefault("answered_questions", []).append({
            "id": q["id"],
            "layer": q["layer"],
            "title": q["title"],
            "answer": answers[q["id"]],
        })

    if consumed_core:
        state["core_rounds_done"] = state.get("core_rounds_done", 0) + 1
    state["core_queue"] = [q for q in core_q if q["id"] not in answered_ids]
    state["optional_queue"] = [q for q in opt_q if q["id"] not in answered_ids]

    ready_for_trial = not state["core_queue"] and (not state["optional_queue"] or state.get("optional_decision") == "ship_now")
    if ready_for_trial:
        state["step"] = max(state.get("step", 3), 4)
        state["step_label"] = "组合试炼与手改验收"
        _sync_wizard_answers_to_profile(wdir, state, home_dir)

    _save_json(state_p, state)
    render_hub(wdir, home_dir)

    if not state["core_queue"] and state["optional_queue"] and state.get("optional_decision") is None:
        rem = len(state["optional_queue"])
        rounds = math.ceil(rem / 5)
        return {
            "status": "CHECKPOINT_CORE_DONE",
            "remaining_optional_count": rem,
            "estimated_extra_rounds": rounds,
            "human_prompt": (
                f"核心大调已定，现在就可以直接进入试炼并投入使用。"
                f"另外还有 {rem} 个偏细枝末节的问题可以继续确认（预计还需 {rounds} 轮）。"
                "请你拍板：A. 先生成试炼样稿直接开用（推荐） / B. 继续把这几道细节答完"
            ),
        }

    if ready_for_trial:
        return {"status": "READY_FOR_TRIAL"}

    return {
        "status": "RECORDED",
        "remaining_core": len(state["core_queue"]),
        "remaining_optional": len(state["optional_queue"]),
    }


def choose_path(wdir: Path, choice: str, home_dir: Path) -> dict[str, Any]:
    state_p = wdir / ".wizard-state.json"
    state = _load_json(state_p, {})
    state["optional_decision"] = choice
    if choice == "ship_now":
        state["step"] = 4
        state["step_label"] = "组合试炼与手改验收"
        _sync_wizard_answers_to_profile(wdir, state, home_dir)
    _save_json(state_p, state)
    render_hub(wdir, home_dir)
    return {
        "status": "CHOICE_SAVED",
        "choice": choice,
        "deferred_optional_count": len(state.get("optional_queue", [])) if choice == "ship_now" else 0,
    }


def undo_last(wdir: Path, home_dir: Path) -> dict[str, Any]:
    state_p = wdir / ".wizard-state.json"
    state = _load_json(state_p, {})
    stack = state.get("undo_stack", [])
    if not stack:
        return {"status": "NOTHING_TO_UNDO", "human_hint": "当前没有可撤回的上一题记录。"}
    last = stack.pop()
    state["core_queue"] = last["core_queue"]
    state["optional_queue"] = last["optional_queue"]
    state["answered_questions"] = last["answered_questions"]
    state["core_rounds_done"] = last["core_rounds_done"]
    state["optional_decision"] = None
    _save_json(state_p, state)
    render_hub(wdir, home_dir)
    return {"status": "UNDONE", "reopened_ids": list(last["answers_batch"].keys())}


def handle_profile(args: argparse.Namespace, home_dir: Path) -> dict[str, Any]:
    reg_p = _registry_path(home_dir)
    reg = _load_json(reg_p, {"active_workspace": "", "active_profile": "default", "active_mask": "default", "recent_workshops": []})
    pid = args.profile or reg.get("active_profile", "default")
    pdir = _ensure_default_profile(home_dir, pid)
    action = args.subaction

    if action == "list":
        profiles_root = home_dir / "profiles"
        out = []
        if profiles_root.is_dir():
            for d in sorted(profiles_root.iterdir()):
                if d.is_dir():
                    masks = [m.stem for m in sorted((d / "masks").glob("*.json"))]
                    out.append({"profile_id": d.name, "masks": masks})
        return {"status": "OK", "active_profile": reg.get("active_profile"), "active_mask": reg.get("active_mask"), "profiles": out}

    if action == "switch":
        target_mask = args.mask or "default"
        mask_p = pdir / "masks" / f"{target_mask}.json"
        if not mask_p.is_file():
            return {"status": "NOT_FOUND", "human_hint": f"未找到名为「{target_mask}」的视觉风格面具。"}
        reg["active_profile"] = pid
        reg["active_mask"] = target_mask
        _save_json(reg_p, reg)
        return {"status": "SWITCHED", "active_profile": pid, "active_mask": target_mask}

    if action == "split-mask":
        new_mask = args.mask or "external_pitch"
        _snapshot_profile(home_dir, pid, f"split_mask:{new_mask}")
        new_p = pdir / "masks" / f"{new_mask}.json"
        base_m = _load_json(pdir / "masks" / "default.json", {})
        base_m["mask_id"] = new_mask
        base_m["mask_name"] = args.label or new_mask
        _save_json(new_p, base_m)
        reg["active_mask"] = new_mask
        _save_json(reg_p, reg)
        return {"status": "MASK_CREATED", "profile_id": pid, "mask_id": new_mask}

    if action == "swap-visual-shell":
        target_mask = args.mask or reg.get("active_mask", "default")
        snap_p = _snapshot_profile(home_dir, pid, f"swap_visual_shell:{target_mask}")
        mask_p = pdir / "masks" / f"{target_mask}.json"
        _save_json(mask_p, {
            "mask_id": target_mask,
            "mask_name": args.label or f"{target_mask}（新换公司壳）",
            "version": "1.0",
            "rules": [],
            "settings": {"clean_grid_enabled": True, "shell_swapped_at": datetime.now().isoformat(timespec="seconds")},
        })
        return {
            "status": "SHELL_SWAPPED",
            "snapshot_before": snap_p,
            "preserved_layers": ["L1_planning", "L2_content"],
            "reset_mask": target_mask,
            "human_hint": "已完整锁死保留你的第 1、2 层规划与写作习惯，并留好旧版备份；现在可以放入新公司模板重采第 3 层视觉样式。",
        }

    if action == "promote-rule":
        layer = args.layer or "L2"
        raw_rule = args.rule_text or ""
        san = output_guard.sanitize_rule_text(raw_rule, auto_abstract=True)
        abstracted = san["sanitized_text"]
        _snapshot_profile(home_dir, pid, f"promote_rule:{layer}")
        if layer == "L1":
            target_p = pdir / "core_l1_planning.json"
        elif layer == "L2":
            target_p = pdir / "core_l2_content.json"
        else:
            target_p = pdir / "masks" / f"{args.mask or reg.get('active_mask', 'default')}.json"
        data = _load_json(target_p, {"version": "1.0", "rules": []})
        next_idx = len(data.get("rules", [])) + 1
        data.setdefault("rules", []).append({
            "rule_no": next_idx,
            "layer": layer,
            "text": abstracted,
            "added_at": datetime.now().isoformat(timespec="seconds"),
        })
        _save_json(target_p, data)
        return {"status": "PROMOTED", "layer": layer, "sanitized_rule": abstracted, "had_raw_specifics": san["had_raw_specifics"]}

    if action == "delete-rule":
        idx_to_del = int(args.rule_no or 0)
        _snapshot_profile(home_dir, pid, f"delete_rule:{args.layer}:{idx_to_del}")
        if args.layer == "L1":
            target_p = pdir / "core_l1_planning.json"
        elif args.layer == "L2":
            target_p = pdir / "core_l2_content.json"
        else:
            target_p = pdir / "masks" / f"{args.mask or reg.get('active_mask', 'default')}.json"
        data = _load_json(target_p, {"rules": []})
        rules = data.get("rules", [])
        if idx_to_del < 1 or idx_to_del > len(rules):
            return {"status": "OUT_OF_RANGE", "human_hint": f"该层共有 {len(rules)} 条规则，请指定 1～{len(rules)} 之间的编号。"}
        removed = rules.pop(idx_to_del - 1)
        for i, r in enumerate(rules, 1):
            r["rule_no"] = i
        _save_json(target_p, data)
        return {"status": "DELETED", "removed": removed}

    if action == "rollback":
        snaps = sorted((pdir / "history").glob("snap-*.json"))
        if not snaps:
            return {"status": "NO_HISTORY", "human_hint": "暂时没有更早的历史版本可以回滚。"}
        latest_snap = _load_json(snaps[-1], {})
        _save_json(pdir / "core_l1_planning.json", latest_snap.get("core_l1", {"rules": []}))
        _save_json(pdir / "core_l2_content.json", latest_snap.get("core_l2", {"rules": []}))
        for m_name, m_val in latest_snap.get("masks", {}).items():
            _save_json(pdir / "masks" / f"{m_name}.json", m_val)
        snaps[-1].unlink()
        return {"status": "ROLLED_BACK", "restored_from": latest_snap.get("saved_at")}

    if action == "reset":
        snap_p = _snapshot_profile(home_dir, pid, "factory_reset")
        for fn in ("core_l1_planning.json", "core_l2_content.json"):
            _save_json(pdir / fn, {"version": "1.0", "rules": [], "settings": {}})
        for mp in (pdir / "masks").glob("*.json"):
            if mp.stem != "default":
                mp.unlink()
        _save_json(pdir / "masks" / "default.json", {"mask_id": "default", "mask_name": "默认汇报风格", "version": "1.0", "rules": [], "settings": {"clean_grid_enabled": True}})
        reg["active_mask"] = "default"
        _save_json(reg_p, reg)
        return {"status": "RESET", "backup_snapshot": snap_p}

    if action == "bind-mask":
        target_mask = args.mask or reg.get("active_mask", "default")
        mask_p = pdir / "masks" / f"{target_mask}.json"
        if not mask_p.is_file():
            return {"status": "NOT_FOUND", "human_hint": f"未找到名为「{target_mask}」的视觉母版皮肤。"}
        _snapshot_profile(home_dir, pid, f"bind_mask:{target_mask}")
        m_data = _load_json(mask_p, {})
        if getattr(args, "independent", False):
            m_data["binding_mode"] = "independent"
            m_data["bound_scenarios"] = []
        else:
            sc_list = [s.strip() for s in (getattr(args, "scenarios", None) or "").split(",") if s.strip()]
            m_data["binding_mode"] = "bound"
            m_data["bound_scenarios"] = sc_list
        _save_json(mask_p, m_data)
        return {
            "status": "MASK_BINDING_UPDATED",
            "mask_id": target_mask,
            "binding_mode": m_data["binding_mode"],
            "bound_scenarios": m_data["bound_scenarios"],
        }

    if action == "confirm-scenario":
        sc_name = (getattr(args, "scenario_name", None) or "").strip()
        sc_id = (getattr(args, "scenario_id", None) or "").strip()
        if not sc_name and not sc_id:
            return {"status": "ERROR", "message": "请提供 --scenario-name 或 --scenario-id"}
        _snapshot_profile(home_dir, pid, "confirm_scenario")
        l1_p = pdir / "core_l1_planning.json"
        l1_data = _load_json(l1_p, {"version": "2.0", "rules": [], "scenarios": [], "settings": {}})
        scenarios = l1_data.setdefault("scenarios", [])
        if not sc_id:
            sc_id = f"sc_custom_{len(scenarios) + 1}"
        matched = next((s for s in scenarios if s.get("scenario_id") == sc_id or (sc_name and s.get("scenario_name") == sc_name)), None)
        if matched:
            if sc_name:
                matched["scenario_name"] = sc_name
            if getattr(args, "skeleton", None):
                matched["typical_skeleton"] = args.skeleton
            matched["status"] = "confirmed"
            matched["confidence"] = "high"
        else:
            scenarios.append({
                "scenario_id": sc_id,
                "scenario_name": sc_name or sc_id,
                "deck_count": 0,
                "sample_files": [],
                "median_pages": 5,
                "typical_skeleton": getattr(args, "skeleton", None) or "核心结论先行 → 分块论据展开 → 下一步行动",
                "confidence": "high",
                "status": "confirmed",
            })
        _save_json(l1_p, l1_data)
        return {"status": "SCENARIO_CONFIRMED", "scenarios": scenarios}

    if action == "confirm-framework":
        fw_id = (getattr(args, "framework_id", None) or "").strip()
        decision = getattr(args, "decision", None) or "accept"
        if not fw_id:
            return {"status": "ERROR", "message": "请提供 --framework-id"}
        _snapshot_profile(home_dir, pid, f"confirm_framework:{fw_id}")
        l2_p = pdir / "core_l2_content.json"
        l2_data = _load_json(l2_p, {"version": "2.0", "rules": [], "frameworks": [], "settings": {}})
        frameworks = l2_data.setdefault("frameworks", [])
        target_fw = next((f for f in frameworks if f.get("framework_id") == fw_id), None)
        if target_fw is None:
            return {"status": "NOT_FOUND", "human_hint": f"未找到编号为「{fw_id}」的页面框架。"}
        if decision == "drop":
            target_fw["status"] = "dropped"
        elif decision == "adjust":
            target_fw["status"] = "adjusted_confirmed"
            if getattr(args, "note", None):
                target_fw["applied_adjustment"] = args.note
            elif target_fw.get("suggested_adjustment"):
                target_fw["applied_adjustment"] = target_fw["suggested_adjustment"]
        else:
            target_fw["status"] = "confirmed"
        _save_json(l2_p, l2_data)
        return {"status": "FRAMEWORK_UPDATED", "framework": target_fw}

    if action == "resolve-combo":
        sc_query = (getattr(args, "scenario", None) or "").strip()
        req_mask = getattr(args, "mask", None)
        l1_data = _load_json(pdir / "core_l1_planning.json", {"rules": [], "scenarios": []})
        l2_data = _load_json(pdir / "core_l2_content.json", {"rules": [], "frameworks": [], "settings": {}})
        scenarios = l1_data.get("scenarios", [])
        matched_sc = next(
            (s for s in scenarios if s.get("scenario_id") == sc_query or sc_query in s.get("scenario_name", "")),
            scenarios[0] if scenarios else None,
        )
        sc_id = matched_sc["scenario_id"] if matched_sc else ""
        active_fws = [
            f for f in l2_data.get("frameworks", [])
            if f.get("status") != "dropped" and (not sc_id or f.get("scenario_id") == sc_id)
        ]
        if not active_fws:
            active_fws = [f for f in l2_data.get("frameworks", []) if f.get("status") != "dropped"]

        masks_dir = pdir / "masks"
        mask_objs = [_load_json(mp, {}) for mp in sorted(masks_dir.glob("*.json"))]
        chosen_mask = None
        if req_mask:
            chosen_mask = next((m for m in mask_objs if m.get("mask_id") == req_mask or req_mask in m.get("mask_name", "")), None)
        if chosen_mask is None and sc_id:
            chosen_mask = next(
                (m for m in mask_objs if m.get("binding_mode") == "bound" and sc_id in m.get("bound_scenarios", [])),
                None,
            )
        if chosen_mask is None:
            active_mid = reg.get("active_mask", "default")
            chosen_mask = next((m for m in mask_objs if m.get("mask_id") == active_mid), mask_objs[0] if mask_objs else {})

        return {
            "status": "COMBO_RESOLVED",
            "profile_id": pid,
            "l1_scenario": matched_sc,
            "l1_rules": l1_data.get("rules", []),
            "l2_frameworks": active_fws,
            "l2_rules": l2_data.get("rules", []),
            "l2_settings": l2_data.get("settings", {}),
            "l3_mask": chosen_mask,
            "combo_mode": "manual_mix" if req_mask else "scenario_default",
        }

    if action == "export-share":
        out_dir = Path(args.out or "./PPT-style-share")
        if out_dir.exists():
            shutil.rmtree(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        for item_name in ("SKILL.md", "VERSION", "README.md", "install.py", "install_windows.bat", "install_mac.command"):
            src_item = SKILL_ROOT / item_name
            if src_item.is_file():
                shutil.copy2(src_item, out_dir / item_name)
        for folder in ("references", "scripts"):
            src_f = SKILL_ROOT / folder
            if src_f.is_dir():
                shutil.copytree(src_f, out_dir / folder, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))

        # Bundle sanitized profile preset only (zero raw business sentences)
        preset_dir = out_dir / "bundled-profile"
        preset_dir.mkdir(parents=True, exist_ok=True)
        for name, src_json in (
            ("core_l1_planning.json", pdir / "core_l1_planning.json"),
            ("core_l2_content.json", pdir / "core_l2_content.json"),
            ("mask.json", pdir / "masks" / f"{args.mask or reg.get('active_mask', 'default')}.json"),
        ):
            obj = _load_json(src_json, {"rules": []})
            for r in obj.get("rules", []):
                r["text"] = output_guard.sanitize_rule_text(r.get("text", ""), auto_abstract=True)["sanitized_text"]
            _save_json(preset_dir / name, obj)
        return {"status": "EXPORTED", "share_dir": str(out_dir.resolve())}

    return {"status": "ERROR", "message": f"不支持的档案操作：{action}"}


def generate_usage_guide(
    home_dir: Path,
    wdir: Path | None = None,
    profile_id: str | None = None,
) -> dict[str, Any]:
    """生成固定 5 段范式的大白话调用指南（供提炼完成回执教学与用户随时查询使用）。"""
    reg = _load_json(_registry_path(home_dir), {})
    pid = profile_id or reg.get("active_profile", "default")
    active_mid = reg.get("active_mask", "default")
    pdir = _ensure_default_profile(home_dir, pid)
    l1 = _load_json(pdir / "core_l1_planning.json", {"rules": [], "scenarios": []})
    l2 = _load_json(pdir / "core_l2_content.json", {"rules": [], "frameworks": []})

    masks_dir = pdir / "masks"
    masks = [_load_json(mp, {}) for mp in sorted(masks_dir.glob("*.json"))]
    scenarios = l1.get("scenarios", [])
    frameworks = [f for f in l2.get("frameworks", []) if f.get("status") != "dropped"]

    sc_lines = []
    prompt_lines = []
    for sc in scenarios:
        sc_id = sc.get("scenario_id", "")
        sc_name = sc.get("scenario_name", sc_id)
        sc_fws = [f for f in frameworks if f.get("scenario_id") == sc_id]
        bound_m = next(
            (m for m in masks if m.get("binding_mode") == "bound" and sc_id in m.get("bound_scenarios", [])),
            None,
        )
        m_label = bound_m.get("mask_name", bound_m.get("mask_id")) if bound_m else "默认模板皮肤"
        sc_lines.append(f"- {sc_name}：已配 {len(sc_fws)} 套页面框架，默认搭配「{m_label}」")
        prompt_lines.append(f"- 做「{sc_name}」时，直接说：“帮我按我的{sc_name}风格做一套 PPT”，然后把资料丢给我")

    if not sc_lines:
        sc_lines.append("- 综合汇报：暂未区分细分场景，按通用汇报习惯启动")
        prompt_lines.append("- 直接说：“帮我按我的个人风格做一套汇报 PPT”，然后把资料丢给我")

    mask_lines = []
    for m in masks:
        m_name = m.get("mask_name") or m.get("mask_id", "默认风格")
        b_mode = "独立通用（随时可切）" if m.get("binding_mode") == "independent" else "已绑定对应场景"
        cur_tag = "（当前默认）" if m.get("mask_id") == active_mid else ""
        mask_lines.append(f"- {m_name}{cur_tag}：{b_mode}")

    alt_mask = next((m for m in masks if m.get("mask_id") != active_mid), masks[0] if masks else {})
    alt_mask_name = alt_mask.get("mask_name") or "另一套视觉模板"
    first_sc_name = scenarios[0].get("scenario_name", "汇报") if scenarios else "汇报"

    guide_text = "\n".join([
        "【你的 PPT 风格资产速览】",
        f"当前使用档案：{pid}（已沉淀 {len(scenarios)} 个主要场景、{len(frameworks)} 套单页框架、{len(masks)} 套视觉母版皮肤）",
        *sc_lines,
        *mask_lines,
        "",
        "【按你的场景直接发起】",
        *prompt_lines,
        "",
        "【日常 5 档怎么开口（按需点菜）】",
        "1. 全套代写：“帮我基于这些资料做一套汇报 PPT”（先定大纲，再写内容，最后排版出稿）",
        "2. 只做页面规划：“先帮我规划这套汇报分几页、每页讲什么”（先出页数和每页大纲，不用先给全部细数）",
        "3. 按大纲填内容：“这几页大纲定了，帮我把每页具体内容写出来”（锁死页数和结构，写标题、正文和备注）",
        "4. 定稿只美化排版：“文字已经定稿了，一字不改帮我按我的风格排版”（数字和文字一字不改，按安全留白排成成品）",
        "5. 旧稿局部加改页：“在这份 PPT 第 X 页后加两页（或只改第 X 页），别的页别动”（其余页原封不动，仅改指定位置）",
        "",
        "【自由混搭与临时借参考】",
        f"- 场景和皮肤自由混搭：说“用我的{first_sc_name}框架，但视觉换成「{alt_mask_name}」”",
        "- 临时借用外部参考：直接丢进参考文件说“样式按这份参考”或“结构按这份参考”，我会先跟你确认只借哪一层且只对本轮生效",
        "",
        "【手改进化、反悔与随时查口令】",
        "- 改完保存即学习：在成品里改完按保存后说一句“我改好了”，我会自动抓差异请你勾选要不要记为长期习惯",
        "- 随时反悔调整：直接说“修改上一题”“删除第 2 条习惯”或“回滚到上一个版本”",
        "- 忘记怎么开口时：随时问一句“怎么发起写 PPT”或“怎么调用我的风格”，我会再发一遍这份速查给你",
    ])

    guard_res = output_guard.check_reply_text(guide_text)
    return {
        "status": "OK",
        "profile_id": pid,
        "active_mask": active_mid,
        "scenario_count": len(scenarios),
        "framework_count": len(frameworks),
        "mask_count": len(masks),
        "guard_ok": guard_res["ok"],
        "guide_text": guide_text,
    }


def render_hub(wdir: Path, home_dir: Path) -> None:
    manifest = _load_json(wdir / "01-原始代表作归档" / ".archive-manifest.json", {"items": []})
    state = _load_json(wdir / ".wizard-state.json", {})
    pid = state.get("active_profile", "default")
    mid = state.get("active_mask", "default")
    pdir = _ensure_default_profile(home_dir, pid)
    l1 = _load_json(pdir / "core_l1_planning.json", {"rules": [], "scenarios": []})
    l2 = _load_json(pdir / "core_l2_content.json", {"rules": [], "frameworks": []})
    m_data = _load_json(pdir / "masks" / f"{mid}.json", {"rules": [], "mask_name": "默认汇报风格"})
    all_masks = [_load_json(mp, {}) for mp in sorted((pdir / "masks").glob("*.json"))]

    sample_lines = []
    for it in manifest["items"]:
        cat_cn = CATEGORY_DIR_MAP.get(it["category"], it["category"])
        pages = f"锁定第 {it['selected_slides']} 页" if it.get("selected_slides") else "全稿"
        mode_tag = " · 只读引用" if it.get("storage_mode") == "reference" else ""
        sample_lines.append(f"- 《{it['archived_name']}》（{cat_cn} · {pages}{mode_tag}）")
    if not sample_lines:
        sample_lines.append("- （暂未归档代表作，请放入至少 1 份 .pptx 开始提炼）")

    scenarios = l1.get("scenarios", [])
    frameworks = [f for f in l2.get("frameworks", []) if f.get("status") != "dropped"]
    sc_summary_lines = [
        f"- **{sc.get('scenario_name')}**（样本 {sc.get('deck_count', 0)} 份 · 常见节奏：{sc.get('typical_skeleton', '')}）"
        for sc in scenarios
    ] or ["- （待提炼或由你补充主要汇报场景）"]

    fw_summary_lines = [
        f"- [{f.get('scenario_name', '综合')}] **{f.get('framework_name')}**（{f.get('confidence', 'medium')}置信 · 状态：{f.get('status')}）"
        + (f"｜建议：{f['suggested_adjustment']}" if f.get("suggested_adjustment") else "")
        for f in frameworks
    ] or ["- （待提炼或从手改中积累）"]

    mask_summary_lines = [
        f"- **{m.get('mask_name', m.get('mask_id'))}**：{'独立通用' if m.get('binding_mode') == 'independent' else '默认绑定场景 ' + str(m.get('bound_scenarios', []))}"
        for m in all_masks
    ]

    ug = generate_usage_guide(home_dir, wdir=wdir, profile_id=pid)

    hub_md = f"""# 00 · PPT 风格工坊操作中心

- **当前步骤**：第 {state.get('step', 1)} / 6 步（{state.get('step_label', '准备就绪')}）
- **当前激活皮肤**：{m_data.get('mask_name', mid)}
- **已归档样本清单**：
{chr(10).join(sample_lines)}

## 已沉淀的三层活资产摘要
- **第一层｜写稿场景（{len(scenarios)} 个）** & 规划规则（{len(l1.get('rules', []))} 条）：
{chr(10).join(sc_summary_lines)}
- **第二层｜场景框架（{len(frameworks)} 套）** & 内容规则（{len(l2.get('rules', []))} 条）：
{chr(10).join(fw_summary_lines)}
- **第三层｜视觉母版皮肤（{len(all_masks)} 套）** & 排版规则（{len(m_data.get('rules', []))} 条）：
{chr(10).join(mask_summary_lines)}

## 怎么喊我写 PPT（速查指南）
```text
{ug['guide_text']}
```
"""
    (wdir / "00-操作中心.md").write_text(hub_md, encoding="utf-8")

    guide_md = f"""# 我的 PPT 风格说明书

> 本说明书为您保留清晰的个人习惯摘要，可随时查阅或要求删改指定编号条目。

## 一、写稿场景与页面规划习惯（第一层）
### 1.1 主要写稿场景
{chr(10).join(sc_summary_lines)}

### 1.2 页面规划原则
{chr(10).join(f"{r['rule_no']}. {r['text']}" for r in l1.get('rules', [])) or "（尚在积累中，可通过日常手改回流补充）"}

## 二、单页内容框架与表达习惯（第二层）
### 2.1 场景配套单页框架
{chr(10).join(fw_summary_lines)}

### 2.2 内容表达原则
{chr(10).join(f"{r['rule_no']}. {r['text']}" for r in l2.get('rules', [])) or "（尚在积累中，可通过日常手改回流补充）"}

## 三、视觉母版与排版习惯（第三层 · {m_data.get('mask_name', mid)}）
### 3.1 母版皮肤与绑定关系
{chr(10).join(mask_summary_lines)}

### 3.2 视觉排版原则
{chr(10).join(f"{r['rule_no']}. {r['text']}" for r in m_data.get('rules', [])) or "（尚在积累中，可通过日常手改回流补充）"}
"""
    (wdir / "03-我的PPT风格说明书.md").write_text(guide_md, encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PPT-style 工坊操作中心与状态机管理器")
    parser.add_argument("--home-dir", default=None, help="覆盖全局档案目录（默认 ~/.ppt-style）")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_init = sub.add_parser("init")
    p_init.add_argument("--workspace", default=None)
    p_init.add_argument("--default-home", action="store_true")

    p_status = sub.add_parser("status")
    p_status.add_argument("--workspace", default=None)

    p_sd = sub.add_parser("scan-dir")
    p_sd.add_argument("--dir", required=True)

    p_tr = sub.add_parser("triage")
    p_tr.add_argument("--path", action="append", required=True, help="目录或文件路径（可多次传入）")

    p_intake = sub.add_parser("intake")
    p_intake.add_argument("--workspace", required=True)
    p_intake.add_argument("--file", required=True)
    p_intake.add_argument("--category", choices=tuple(CATEGORY_DIR_MAP.keys()), required=True)
    p_intake.add_argument("--slides", default=None)
    p_intake.add_argument("--external-toc", action="store_true")
    p_intake.add_argument("--copy-mode", choices=("copy", "reference"), default="copy")
    p_intake.add_argument("--from-upload", action="store_true", help="来自对话上传的文件强制拷贝入库")

    p_scan = sub.add_parser("scan")
    p_scan.add_argument("--workspace", required=True)

    p_pq = sub.add_parser("plan-questions")
    p_pq.add_argument("--workspace", required=True)

    p_nb = sub.add_parser("next-batch")
    p_nb.add_argument("--workspace", required=True)

    p_rec = sub.add_parser("record-answer")
    p_rec.add_argument("--workspace", required=True)
    p_rec.add_argument("--answers", required=True, help="JSON 字符串或文件路径")

    p_cp = sub.add_parser("choose-path")
    p_cp.add_argument("--workspace", required=True)
    p_cp.add_argument("--choice", choices=("ship_now", "continue_optional"), required=True)

    p_undo = sub.add_parser("undo-last")
    p_undo.add_argument("--workspace", required=True)

    p_prof = sub.add_parser("profile")
    p_prof.add_argument(
        "subaction",
        choices=(
            "list", "switch", "split-mask", "swap-visual-shell",
            "promote-rule", "delete-rule", "rollback", "reset", "export-share",
            "bind-mask", "confirm-scenario", "confirm-framework", "resolve-combo",
        ),
    )
    p_prof.add_argument("--profile", default=None)
    p_prof.add_argument("--mask", default=None)
    p_prof.add_argument("--label", default=None)
    p_prof.add_argument("--layer", choices=("L1", "L2", "L3"), default="L2")
    p_prof.add_argument("--rule-text", default="")
    p_prof.add_argument("--rule-no", type=int, default=0)
    p_prof.add_argument("--out", default=None)
    p_prof.add_argument("--scenarios", default="")
    p_prof.add_argument("--independent", action="store_true")
    p_prof.add_argument("--scenario-name", default="")
    p_prof.add_argument("--scenario-id", default="")
    p_prof.add_argument("--skeleton", default="")
    p_prof.add_argument("--framework-id", default="")
    p_prof.add_argument("--decision", choices=("accept", "adjust", "drop"), default="accept")
    p_prof.add_argument("--note", default="")
    p_prof.add_argument("--scenario", default="")

    p_ug = sub.add_parser("usage-guide")
    p_ug.add_argument("--workspace", default=None)
    p_ug.add_argument("--profile", default=None)

    p_rh = sub.add_parser("render-hub")
    p_rh.add_argument("--workspace", required=True)

    args = parser.parse_args(argv)
    home = _home_dir(args.home_dir)

    if args.cmd == "init":
        print(json.dumps(init_workshop(args.workspace, args.default_home, home), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "status":
        reg = _load_json(_registry_path(home), {})
        ws = args.workspace or reg.get("active_workspace")
        if not ws or not Path(ws).is_dir():
            print(json.dumps({"status": "NO_ACTIVE_WORKSHOP", "human_hint": "还没有找到建档记录，说一句“帮我建 PPT 风格工坊”即可开始。"}, ensure_ascii=False, indent=2))
            return 0
        wdir = _workshop_dir(ws)
        st = _load_json(wdir / ".wizard-state.json", {})
        print(json.dumps({"status": "FOUND", "workshop_dir": str(wdir), "wizard_state": st, "registry": reg}, ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "scan-dir":
        res = scan_directory(args.dir)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0 if res.get("status") != "ERROR" else 1

    if args.cmd == "triage":
        res = triage_paths(args.path)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "intake":
        wdir = _workshop_dir(args.workspace)
        res = intake_file(
            wdir, args.file, args.category, args.slides, args.external_toc, home,
            copy_mode=args.copy_mode, from_upload=args.from_upload,
        )
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0 if res["status"] not in ("ERROR", "REJECTED") else 2

    if args.cmd == "scan":
        print(json.dumps(scan_workshop(_workshop_dir(args.workspace), home), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "plan-questions":
        print(json.dumps(plan_questions(_workshop_dir(args.workspace), home), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "next-batch":
        print(json.dumps(next_batch(_workshop_dir(args.workspace)), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "record-answer":
        ans_raw = Path(args.answers).read_text(encoding="utf-8") if Path(args.answers).is_file() else args.answers
        res = record_answer(_workshop_dir(args.workspace), json.loads(ans_raw), home)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "choose-path":
        print(json.dumps(choose_path(_workshop_dir(args.workspace), args.choice, home), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "undo-last":
        print(json.dumps(undo_last(_workshop_dir(args.workspace), home), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "profile":
        print(json.dumps(handle_profile(args, home), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "usage-guide":
        wdir = _workshop_dir(args.workspace) if args.workspace else None
        print(json.dumps(generate_usage_guide(home, wdir=wdir, profile_id=args.profile), ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "render-hub":
        render_hub(_workshop_dir(args.workspace), home)
        print(json.dumps({"status": "RENDERED"}, ensure_ascii=False))
        return 0

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
