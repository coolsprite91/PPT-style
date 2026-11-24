#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""snapshot_diff.py — 全局集中快照库、跨目录隐形编号认领、语义级手改比对与双通道一字不改核对锁（零第三方依赖）。

子命令:
  save-snapshot     出稿时将底稿留存在 ~/.ppt-style/snapshots/（>50MB 自动剥离音视频流）
  diff              自动识别原地覆盖或改名移走的手改文件，过滤时间戳/WPS拆run/≤1pt抖动输出改动清单
  verify-text-lock  第3层“一字不改核对锁”：允许美观分段，但对数字/单位/金额/百分比/符号零容忍全等校验
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "dc": "http://purl.org/dc/elements/1.1/",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
}

EMU_PER_PT = 12700.0
MEDIA_AV_EXTENSIONS = {".mp4", ".mov", ".avi", ".wmv", ".mkv", ".mp3", ".wav", ".m4a"}
DOC_ID_RE = re.compile(r"\[ppt-style-id:([A-Za-z0-9_-]+)\]")
NUMERIC_TOKEN_RE = re.compile(
    r"(?:[+\-±≥≤><≈]?\s*\d+(?:,\d{3})*(?:\.\d+)?\s*(?:%|‰|个百分点|万元|亿元|万美元|元|万|亿|倍|次|家|户|人|页|项|天|周|月|年|小时|分钟|秒|pt|磅)?)"
    r"|(?:[+\-±≥≤><])"
)


def default_ppt_style_home(override: str | None = None) -> Path:
    if override:
        return Path(override)
    env_home = os.environ.get("PPT_STYLE_HOME")
    if env_home:
        return Path(env_home)
    return Path.home() / ".ppt-style"


def _clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _norm_text(text: str) -> str:
    return re.sub(r"[^0-9A-Za-z\u3400-\u9fff]+", "", text or "").lower()


def _bigrams(text: str) -> set[str]:
    n = _norm_text(text)
    if len(n) < 2:
        return {n} if n else set()
    return {n[i : i + 2] for i in range(len(n) - 1)}


def _jaccard(a: str, b: str) -> float:
    ba, bb = _bigrams(a), _bigrams(b)
    if not ba or not bb:
        return 0.0
    return len(ba & bb) / len(ba | bb)


def _posix_norm(base_dir: str, target: str) -> str:
    parts = (base_dir.rstrip("/") + "/" + target.lstrip("/")).split("/")
    out: list[str] = []
    for part in parts:
        if not part or part == ".":
            continue
        if part == "..":
            if out:
                out.pop()
        else:
            out.append(part)
    return "/".join(out)


def _read_xml(zf: zipfile.ZipFile, member: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(member))
    except Exception:
        return None


def _slide_paths_in_order(zf: zipfile.ZipFile) -> list[str]:
    pres = _read_xml(zf, "ppt/presentation.xml")
    rels = _read_xml(zf, "ppt/_rels/presentation.xml.rels")
    if pres is None or rels is None:
        return []
    rel_map: dict[str, str] = {}
    for rel in rels.findall(f"{{{NS['rel']}}}Relationship"):
        rid = rel.attrib.get("Id", "")
        target = rel.attrib.get("Target", "")
        if rid and target:
            rel_map[rid] = _posix_norm("ppt", target) if not target.startswith("/") else target.lstrip("/")
    out: list[str] = []
    for sld in pres.findall(f".//{{{NS['p']}}}sldIdLst/{{{NS['p']}}}sldId"):
        rid = sld.attrib.get(f"{{{NS['r']}}}id", "")
        if rid in rel_map:
            out.append(rel_map[rid])
    return out


def extract_doc_id(pptx_path: Path) -> str | None:
    """从 .pptx 内部属性读取隐形 doc_id。"""
    try:
        with zipfile.ZipFile(pptx_path, "r") as zf:
            core = _read_xml(zf, "docProps/core.xml")
            if core is None:
                return None
            desc = core.find(f"{{{NS['dc']}}}description")
            if desc is not None and desc.text:
                m = DOC_ID_RE.search(desc.text)
                if m:
                    return m.group(1)
    except Exception:
        return None
    return None


@dataclass
class SemanticSlide:
    slide_no: int
    title: str
    paragraphs: list[str]
    full_text: str
    font_sizes: list[float]
    shape_boxes_pt: list[tuple[float, float, float, float]]
    table_count: int
    table_cells: list[str]
    diagram_shapes: int
    pic_hashes: list[str]
    notes_text: str


def extract_semantic_slides(pptx_path: Path) -> list[SemanticSlide]:
    """抽取忽略时间戳、忽略 <a:r> 碎片切分、量化坐标的语义幻灯片快照。"""
    slides: list[SemanticSlide] = []
    with zipfile.ZipFile(pptx_path, "r") as zf:
        slide_paths = _slide_paths_in_order(zf)
        for idx, sld_path in enumerate(slide_paths, 1):
            root = _read_xml(zf, sld_path)
            if root is None:
                continue

            # Read slide rels for notes & images
            parent, name = sld_path.rsplit("/", 1)
            rel_xml = _read_xml(zf, f"{parent}/_rels/{name}.rels")
            notes_path = ""
            media_targets: list[str] = []
            if rel_xml is not None:
                for rel in rel_xml.findall(f"{{{NS['rel']}}}Relationship"):
                    rt = rel.attrib.get("Type", "")
                    tg = rel.attrib.get("Target", "")
                    resolved = _posix_norm(parent, tg) if not tg.startswith("/") else tg.lstrip("/")
                    if rt.endswith("/notesSlide"):
                        notes_path = resolved
                    elif rt.endswith("/image"):
                        media_targets.append(resolved)

            title = ""
            paragraphs: list[str] = []
            font_sizes: set[float] = set()
            boxes: list[tuple[float, float, float, float]] = []
            table_count = 0
            table_cells: list[str] = []

            for sp in root.findall(f".//{{{NS['p']}}}sp"):
                ph = sp.find(f".//{{{NS['p']}}}ph")
                ph_type = ph.attrib.get("type", "") if ph is not None else ""
                xfrm = sp.find(f"./{{{NS['p']}}}spPr/{{{NS['a']}}}xfrm")
                if xfrm is not None:
                    off = xfrm.find(f"{{{NS['a']}}}off")
                    ext = xfrm.find(f"{{{NS['a']}}}ext")
                    if off is not None and ext is not None:
                        try:
                            boxes.append((
                                round(int(off.attrib.get("x", "0")) / EMU_PER_PT, 1),
                                round(int(off.attrib.get("y", "0")) / EMU_PER_PT, 1),
                                round(int(ext.attrib.get("cx", "0")) / EMU_PER_PT, 1),
                                round(int(ext.attrib.get("cy", "0")) / EMU_PER_PT, 1),
                            ))
                        except ValueError:
                            pass

                tx = sp.find(f"{{{NS['p']}}}txBody")
                if tx is None:
                    continue
                sp_paras: list[str] = []
                for p_el in tx.findall(f"{{{NS['a']}}}p"):
                    # Concatenate all <a:t> inside paragraph to ignore WPS/Office run fragmentation!
                    p_txt = _clean_text("".join(t.text or "" for t in p_el.findall(f".//{{{NS['a']}}}t")))
                    if p_txt:
                        sp_paras.append(p_txt)
                    for rpr in p_el.findall(f".//{{{NS['a']}}}rPr"):
                        if "sz" in rpr.attrib:
                            try:
                                # Quantize to 0.5pt to ignore minor rendering jitter
                                font_sizes.add(round(int(rpr.attrib["sz"]) / 100.0 * 2) / 2.0)
                            except ValueError:
                                pass
                if sp_paras:
                    paragraphs.extend(sp_paras)
                    if ph_type in ("title", "ctrTitle") and not title:
                        title = sp_paras[0]

            if not title and paragraphs:
                title = paragraphs[0][:80]

            for tbl in root.findall(f".//{{{NS['a']}}}tbl"):
                table_count += 1
                for tc in tbl.findall(f".//{{{NS['a']}}}tc"):
                    ct = _clean_text("".join(t.text or "" for t in tc.findall(f".//{{{NS['a']}}}t")))
                    if ct:
                        table_cells.append(ct)
                        paragraphs.append(ct)

            pic_hashes: list[str] = []
            for m_path in sorted(set(media_targets)):
                try:
                    pic_hashes.append(hashlib.sha256(zf.read(m_path)).hexdigest()[:12])
                except Exception:
                    pass

            notes_text = ""
            if notes_path:
                n_xml = _read_xml(zf, notes_path)
                if n_xml is not None:
                    notes_parts = []
                    for sp in n_xml.findall(f".//{{{NS['p']}}}sp"):
                        ph = sp.find(f".//{{{NS['p']}}}ph")
                        if ph is not None and ph.attrib.get("type") == "body":
                            notes_parts.append(_clean_text("".join(t.text or "" for t in sp.findall(f".//{{{NS['a']}}}t"))))
                    notes_text = _clean_text(" ".join(p for p in notes_parts if p))

            diagram_shapes = len(root.findall(f".//{{{NS['p']}}}cxnSp")) + len(root.findall(f".//{{{NS['p']}}}grpSp"))
            slides.append(
                SemanticSlide(
                    slide_no=idx,
                    title=title,
                    paragraphs=paragraphs,
                    full_text=" | ".join(paragraphs),
                    font_sizes=sorted(font_sizes),
                    shape_boxes_pt=boxes,
                    table_count=table_count,
                    table_cells=table_cells,
                    diagram_shapes=diagram_shapes,
                    pic_hashes=pic_hashes,
                    notes_text=notes_text,
                )
            )
    return slides


def semantic_fingerprint(slides: list[SemanticSlide]) -> str:
    """计算完全脱离保存时间戳和 run 拆分的语义指纹。"""
    payload = [
        {
            "title": _norm_text(s.title),
            "paragraphs": [_norm_text(p) for p in s.paragraphs],
            "font_sizes": s.font_sizes,
            # Quantize coordinates to 2pt bins so <=1pt jitter never flips fingerprint
            "boxes": [[round(v / 2.0) * 2 for v in b] for b in s.shape_boxes_pt],
            "tables": s.table_count,
            "pics": s.pic_hashes,
            "notes": _norm_text(s.notes_text),
        }
        for s in slides
    ]
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _copy_snapshot_file(src: Path, dst: Path) -> bool:
    """复制底稿到快照库；若文件 >50MB，自动剥离音视频流防止占满系统盘。返回是否执行了媒体剥离。"""
    size_mb = src.stat().st_size / (1024 * 1024)
    if size_mb <= 50.0:
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


def save_snapshot(
    pptx_path_str: str,
    snapshot_root: str | None = None,
    doc_id: str | None = None,
    source_label: str = "ai_generated",
) -> dict[str, Any]:
    """保存只读快照到全局集中目录，并登记文档隐形编号与原始路径索引。"""
    src = Path(pptx_path_str).resolve()
    if not src.is_file():
        return {"status": "ERROR", "message": f"找不到文件：{src}"}

    base_root = Path(snapshot_root) if snapshot_root else (default_ppt_style_home() / "snapshots")
    base_root.mkdir(parents=True, exist_ok=True)
    index_path = base_root / "index.json"
    index: dict[str, Any] = json.loads(index_path.read_text(encoding="utf-8")) if index_path.is_file() else {"docs": {}, "paths": {}}

    detected_id = doc_id or extract_doc_id(src) or index["paths"].get(str(src)) or uuid.uuid4().hex[:12]
    doc_dir = base_root / detected_id
    doc_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    snap_file = doc_dir / f"snap-{ts}.pptx"
    stripped = _copy_snapshot_file(src, snap_file)

    slides = extract_semantic_slides(src)
    sem_fp = semantic_fingerprint(slides)

    meta = {
        "doc_id": detected_id,
        "latest_snapshot": str(snap_file),
        "source_path": str(src),
        "source_name": src.name,
        "source_label": source_label,
        "semantic_fingerprint": sem_fp,
        "slide_count": len(slides),
        "media_stripped": stripped,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    index["docs"][detected_id] = meta
    index["paths"][str(src)] = detected_id
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"status": "SAVED", **meta}


def _slide_similarity(a: SemanticSlide, h: SemanticSlide) -> float:
    if _norm_text(a.full_text) == _norm_text(h.full_text) and len(a.pic_hashes) == len(h.pic_hashes):
        return 1.0
    title_sim = 1.0 if (_norm_text(a.title) and _norm_text(a.title) == _norm_text(h.title)) else _jaccard(a.title, h.title)
    text_sim = _jaccard(a.full_text, h.full_text)
    pa, ph = set(a.pic_hashes), set(h.pic_hashes)
    img_sim = (len(pa & ph) / len(pa | ph)) if (pa or ph) else 0.0
    return 0.4 * title_sim + 0.45 * text_sim + 0.15 * img_sim


def _align_slides(base: list[SemanticSlide], edited: list[SemanticSlide]) -> list[tuple[int | None, int | None]]:
    m, n = len(base), len(edited)
    sim = [[_slide_similarity(base[i], edited[j]) for j in range(n)] for i in range(m)]
    dp = [[0.0] * (n + 1) for _ in range(m + 1)]
    back: list[list[tuple[int, int] | None]] = [[None] * (n + 1) for _ in range(m + 1)]
    for i in range(m + 1):
        for j in range(n + 1):
            if i < m and dp[i + 1][j] <= dp[i][j] - 0.15:
                dp[i + 1][j] = dp[i][j] - 0.15
                back[i + 1][j] = (i, j)
            if j < n and dp[i][j + 1] <= dp[i][j] - 0.15:
                dp[i][j + 1] = dp[i][j] - 0.15
                back[i][j + 1] = (i, j)
            if i < m and j < n:
                cand = dp[i][j] + sim[i][j] + 0.05
                if cand >= dp[i + 1][j + 1]:
                    dp[i + 1][j + 1] = cand
                    back[i + 1][j + 1] = (i, j)

    pairs: list[tuple[int | None, int | None]] = []
    ci, cj = m, n
    while ci > 0 or cj > 0:
        prev = back[ci][cj]
        if prev is None:
            break
        pi, pj = prev
        if pi == ci - 1 and pj == cj - 1:
            pairs.append((pi, pj))
        elif pi == ci - 1:
            pairs.append((pi, None))
        else:
            pairs.append((None, pj))
        ci, cj = pi, pj
    pairs.reverse()
    return pairs


def _significant_box_shift(boxes_a: list[tuple[float, float, float, float]], boxes_b: list[tuple[float, float, float, float]]) -> bool:
    if len(boxes_a) != len(boxes_b):
        return False
    for ba, bb in zip(boxes_a, boxes_b):
        if any(abs(x - y) > 1.0 for x, y in zip(ba, bb)):
            return True
    return False


def diff_against_snapshot(
    human_path_str: str,
    base_snapshot_str: str | None = None,
    snapshot_root: str | None = None,
) -> dict[str, Any]:
    """对用户原地保存或改名移动的手改文件进行语义级差异提取。"""
    hum_path = Path(human_path_str).resolve()
    if not hum_path.is_file():
        return {"status": "ERROR", "message": f"找不到文件：{hum_path}"}

    base_root = Path(snapshot_root) if snapshot_root else (default_ppt_style_home() / "snapshots")
    index_path = base_root / "index.json"
    snap_path: Path | None = Path(base_snapshot_str) if base_snapshot_str else None
    found_doc_id = extract_doc_id(hum_path)

    if snap_path is None and index_path.is_file():
        index = json.loads(index_path.read_text(encoding="utf-8"))
        if found_doc_id and found_doc_id in index.get("docs", {}):
            snap_path = Path(index["docs"][found_doc_id]["latest_snapshot"])
        elif str(hum_path) in index.get("paths", {}):
            did = index["paths"][str(hum_path)]
            if did in index.get("docs", {}):
                found_doc_id = did
                snap_path = Path(index["docs"][did]["latest_snapshot"])

    if snap_path is None or not snap_path.is_file():
        return {
            "status": "NO_SNAPSHOT_FOUND",
            "human_hint": "没有找到这份 PPT 对应的生成底稿记录，请告诉我原始初稿是哪一个文件，我帮你直接比对。",
        }

    base_slides = extract_semantic_slides(snap_path)
    hum_slides = extract_semantic_slides(hum_path)
    base_fp = semantic_fingerprint(base_slides)
    hum_fp = semantic_fingerprint(hum_slides)

    if base_fp == hum_fp:
        return {
            "status": "NO_CHANGES_DETECTED",
            "doc_id": found_doc_id,
            "human_hint": "还没有读到新的修改记录，请先确认在 PPT 软件里按了保存（Ctrl+S）后再告诉我。",
        }

    aligned = _align_slides(base_slides, hum_slides)
    changes: list[dict[str, Any]] = []
    structural_deletions: list[dict[str, Any]] = []

    for bi, hi in aligned:
        if bi is None and hi is not None:
            hs = hum_slides[hi]
            changes.append({
                "slide_ref": f"第 {hs.slide_no} 页（新增）",
                "kind": "新增页面",
                "summary": f"新增了主题为《{hs.title or '无标题'}》的页面",
                "likely_category": "内容增补",
            })
        elif hi is None and bi is not None:
            bs = base_slides[bi]
            changes.append({
                "slide_ref": f"原第 {bs.slide_no} 页（已删）",
                "kind": "删除页面",
                "summary": f"删除了原第 {bs.slide_no} 页《{bs.title or '无标题'}》",
                "likely_category": "页面规划",
            })
        elif bi is not None and hi is not None:
            bs, hs = base_slides[bi], hum_slides[hi]
            if _norm_text(bs.title) != _norm_text(hs.title):
                changes.append({
                    "slide_ref": f"第 {hs.slide_no} 页",
                    "kind": "修改标题",
                    "summary": f"将第 {hs.slide_no} 页标题调整为《{hs.title}》",
                    "likely_category": "内容组织",
                })
            if _norm_text(bs.full_text) != _norm_text(hs.full_text):
                b_set = {_norm_text(p) for p in bs.paragraphs}
                h_set = {_norm_text(p) for p in hs.paragraphs}
                if b_set != h_set:
                    changes.append({
                        "slide_ref": f"第 {hs.slide_no} 页",
                        "kind": "正文增删改",
                        "summary": f"调整了第 {hs.slide_no} 页的正文内容或数据表述",
                        "likely_category": "内容增补",
                    })
            if bs.font_sizes != hs.font_sizes and hs.font_sizes:
                changes.append({
                    "slide_ref": f"第 {hs.slide_no} 页",
                    "kind": "字号调整",
                    "summary": f"将第 {hs.slide_no} 页文字字号从 {bs.font_sizes} 调整为 {hs.font_sizes}",
                    "likely_category": "视觉样式",
                })
            if _significant_box_shift(bs.shape_boxes_pt, hs.shape_boxes_pt):
                changes.append({
                    "slide_ref": f"第 {hs.slide_no} 页",
                    "kind": "版面位置移动",
                    "summary": f"重新调整了第 {hs.slide_no} 页内部区块的摆放位置或大小",
                    "likely_category": "视觉样式",
                })
            if bs.table_count > hs.table_count or bs.diagram_shapes > hs.diagram_shapes + 1:
                structural_deletions.append({
                    "slide_ref": f"第 {hs.slide_no} 页",
                    "deleted_element": "表格" if bs.table_count > hs.table_count else "流程图/图形结构",
                    "follow_up_question": (
                        f"注意到你把第 {hs.slide_no} 页的"
                        f"{'表格' if bs.table_count > hs.table_count else '流程图结构'}整块删掉了，"
                        "这是因为以后尽量少用这类表达，还是觉得我这次画得不够好？"
                    ),
                })
            if bs.pic_hashes != hs.pic_hashes:
                changes.append({
                    "slide_ref": f"第 {hs.slide_no} 页",
                    "kind": "配图变化",
                    "summary": f"替换或增删了第 {hs.slide_no} 页的配图",
                    "likely_category": "素材调整",
                })

    if not changes:
        return {
            "status": "NO_CHANGES_DETECTED",
            "doc_id": found_doc_id,
            "human_hint": "对比后未发现实质文字、字号或明显排版变化（仅保存格式微调），无需额外确认。",
        }

    # Build user checklist always starting with 0 (anti-harassment emergency override)
    checklist = [
        {
            "item_no": 0,
            "label": "0. 这次是临时应急调整，千万别记入我的长期风格",
            "default_selected": False,
        }
    ]
    for idx, ch in enumerate(changes[:5], 1):
        checklist.append({
            "item_no": idx,
            "label": f"{idx}. [{ch['likely_category']}] {ch['summary']}（是否记为以后默认习惯？）",
            "slide_ref": ch["slide_ref"],
            "summary": ch["summary"],
        })

    return {
        "status": "CHANGES_DETECTED",
        "doc_id": found_doc_id,
        "snapshot_used": str(snap_path),
        "total_changes": len(changes),
        "changes": changes,
        "structural_deletions": structural_deletions,
        "confirmation_checklist": checklist,
    }


def _extract_numeric_tokens(text: str) -> list[str]:
    tokens: list[str] = []
    for m in NUMERIC_TOKEN_RE.finditer(text):
        tok = re.sub(r"\s+", "", m.group(0))
        if tok:
            tokens.append(tok)
    return tokens


def verify_text_lock(
    source_texts: list[str],
    output_pptx_str: str,
    exempt_blocks: list[str] | None = None,
) -> dict[str, Any]:
    """第3层双通道一字不改核对锁：
    通道1：正文规范化一致性检查（允许为排版拆分成多行/多个框，严禁丢字改字）；
    通道2：全部数字、单位、金额、百分比与大小/正负号零容忍全等校验（代码块与公式白名单豁免）。
    """
    exempt_norm = {_norm_text(b) for b in (exempt_blocks or []) if _norm_text(b)}

    def _filter_exempt(items: list[str]) -> list[str]:
        out: list[str] = []
        for it in items:
            nt = _norm_text(it)
            if nt and any(ex in nt or nt in ex for ex in exempt_norm):
                continue
            out.append(it)
        return out

    src_filtered = _filter_exempt(source_texts)
    slides = extract_semantic_slides(Path(output_pptx_str))
    out_paras = [p for s in slides for p in s.paragraphs]
    out_filtered = _filter_exempt(out_paras)

    src_joined = _norm_text("".join(src_filtered))
    out_joined = _norm_text("".join(out_filtered))

    src_nums = _extract_numeric_tokens(" ".join(src_filtered))
    out_nums = _extract_numeric_tokens(" ".join(out_filtered))

    text_match = src_joined == out_joined
    numeric_match = src_nums == out_nums

    issues: list[str] = []
    if not text_match:
        issues.append("检测到定稿文字内容与排前原文不完全一致（存在增删或改写）。")
    if not numeric_match:
        issues.append(f"关键数字或单位/符号发生变动：排前 {src_nums} vs 排后 {out_nums}。")

    return {
        "ok": text_match and numeric_match,
        "text_channel_ok": text_match,
        "numeric_channel_ok": numeric_match,
        "exempt_count": len(exempt_norm),
        "issues": issues,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PPT-style 全局快照与语义级手改比对工具")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_save = sub.add_parser("save-snapshot", help="保存底稿快照")
    p_save.add_argument("--pptx", required=True)
    p_save.add_argument("--snapshot-root", default=None)
    p_save.add_argument("--doc-id", default=None)
    p_save.add_argument("--label", default="ai_generated")

    p_diff = sub.add_parser("diff", help="比对原地或另存修改稿与底稿差异")
    p_diff.add_argument("--pptx", required=True, help="用户修改后的 PPTX 路径")
    p_diff.add_argument("--base", default=None, help="显式指定底稿路径（可选，默认按隐形编号找回）")
    p_diff.add_argument("--snapshot-root", default=None)

    p_lock = sub.add_parser("verify-text-lock", help="第3层定稿纯排版一字不改核对锁")
    p_lock.add_argument("--source-json", required=True, help="包含 source_texts 与可选 exempt_blocks 的 JSON")
    p_lock.add_argument("--pptx", required=True, help="排版输出后的 PPTX")

    args = parser.parse_args(argv)
    if args.cmd == "save-snapshot":
        res = save_snapshot(args.pptx, snapshot_root=args.snapshot_root, doc_id=args.doc_id, source_label=args.label)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0 if res.get("status") == "SAVED" else 1

    if args.cmd == "diff":
        res = diff_against_snapshot(args.pptx, base_snapshot_str=args.base, snapshot_root=args.snapshot_root)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "verify-text-lock":
        data = json.loads(Path(args.source_json).read_text(encoding="utf-8"))
        res = verify_text_lock(
            source_texts=data.get("source_texts", []),
            output_pptx_str=args.pptx,
            exempt_blocks=data.get("exempt_blocks", []),
        )
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0 if res["ok"] else 2

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
