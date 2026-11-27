#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pptx_inspector.py — 零第三方依赖 PPTX 深层结构与样式解析器。

基于 Python 标准库直接解析 OOXML 压缩包：
  - check-file: 真实文件头校验（拦截 .ppt / .dps / 加密文件，识别 PDF/图片，检测 >50MB 视频流）
  - scan:       穿透 slide → slideLayout → slideMaster → theme1.xml 颜色与中文字体继承，
                统计图表/备注/表格/粗糙排版硬伤、章节块聚类、同源改稿折叠与双峰风格探测，
                仅输出本地算完的浓缩摘要 JSON，严禁撑爆对话上下文。
"""

from __future__ import annotations

import argparse
import colorsys
import hashlib
import importlib.util
import json
import re
import statistics
import sys
import zipfile
from collections import Counter, defaultdict
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
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dc": "http://purl.org/dc/elements/1.1/",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
}

EMU_PER_INCH = 914400.0
EMU_PER_PT = 12700.0
OLE_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
ZIP_MAGIC = b"PK\x03\x04"
PDF_MAGIC = b"%PDF"
IMG_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"}
MEDIA_AV_EXTENSIONS = {".mp4", ".mov", ".avi", ".wmv", ".mkv", ".mp3", ".wav", ".m4a"}
TEMPLATE_PLACEHOLDER_RE = re.compile(r"(?:点击添加|单击此处添加|请输入|占位文本|示例文字|示例内容|\bTBD\b|\bTODO\b|Click to add)", re.IGNORECASE)
COLLAB_FILENAME_RE = re.compile(r"(汇总|月会|月例会|周例会|联合汇报|部门汇报|合并版|汇编|合集|各组)")
SPEAKER_SIGN_RE = re.compile(r"(?:汇报人|负责人|主讲人|撰写人|Owner)\s*[:：]\s*([\u4e00-\u9fa5A-Za-z0-9_]{2,12})")


def _clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _norm_text(text: str) -> str:
    return re.sub(r"[^0-9A-Za-z\u3400-\u9fff]+", "", text or "").lower()


def _bigrams(text: str) -> set[str]:
    norm = _norm_text(text)
    if len(norm) < 2:
        return {norm} if norm else set()
    return {norm[i : i + 2] for i in range(len(norm) - 1)}


def _jaccard(a: str, b: str) -> float:
    ba, bb = _bigrams(a), _bigrams(b)
    if not ba or not bb:
        return 0.0
    return len(ba & bb) / len(ba | bb)


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


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


def _rels_path_for(part_path: str) -> str:
    if "/" not in part_path:
        return f"_rels/{part_path}.rels"
    parent, name = part_path.rsplit("/", 1)
    return f"{parent}/_rels/{name}.rels"


def _read_xml(zf: zipfile.ZipFile, member: str) -> ET.Element | None:
    try:
        data = zf.read(member)
        return ET.fromstring(data)
    except Exception:
        return None


def _read_rels(zf: zipfile.ZipFile, part_path: str) -> dict[str, tuple[str, str]]:
    """Return {rId: (type_uri, resolved_target_path)}."""
    rel_xml = _read_xml(zf, _rels_path_for(part_path))
    if rel_xml is None:
        return {}
    base_dir = part_path.rsplit("/", 1)[0] if "/" in part_path else ""
    out: dict[str, tuple[str, str]] = {}
    for rel in rel_xml.findall(f"{{{NS['rel']}}}Relationship"):
        rid = rel.attrib.get("Id", "")
        rtype = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if target and not target.startswith(("http://", "https://", "mailto:")):
            resolved = _posix_norm(base_dir, target) if not target.startswith("/") else target.lstrip("/")
        else:
            resolved = target
        out[rid] = (rtype, resolved)
    return out


def _apply_lum_mod_off(hex_rgb: str, lum_mod: int | None, lum_off: int | None, tint: int | None, shade: int | None) -> str:
    hex_rgb = hex_rgb.strip().lstrip("#").upper()
    if len(hex_rgb) != 6:
        return "000000"
    r = int(hex_rgb[0:2], 16) / 255.0
    g = int(hex_rgb[2:4], 16) / 255.0
    b = int(hex_rgb[4:6], 16) / 255.0

    if shade is not None:
        factor = shade / 100000.0
        r, g, b = r * factor, g * factor, b * factor
    if tint is not None:
        factor = tint / 100000.0
        r = r * factor + (1.0 - factor)
        g = g * factor + (1.0 - factor)
        b = b * factor + (1.0 - factor)

    if lum_mod is not None or lum_off is not None:
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        if lum_mod is not None:
            l = l * (lum_mod / 100000.0)
        if lum_off is not None:
            l = l + (lum_off / 100000.0)
        l = max(0.0, min(1.0, l))
        r, g, b = colorsys.hls_to_rgb(h, l, s)

    ri = max(0, min(255, round(r * 255)))
    gi = max(0, min(255, round(g * 255)))
    bi = max(0, min(255, round(b * 255)))
    return f"{ri:02X}{gi:02X}{bi:02X}"


def rgb_to_cn_name(hex_rgb: str) -> str:
    """将 RGB 十六进制值翻译为对用户友好的中文自然色名。"""
    hex_rgb = hex_rgb.strip().lstrip("#").upper()
    if len(hex_rgb) != 6:
        return "默认色"
    r, g, b = int(hex_rgb[0:2], 16) / 255.0, int(hex_rgb[2:4], 16) / 255.0, int(hex_rgb[4:6], 16) / 255.0
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    hue_deg = h * 360.0

    if s < 0.10:
        if l >= 0.94:
            return "纯白"
        if l >= 0.80:
            return "极浅灰"
        if l >= 0.45:
            return "中灰"
        if l >= 0.18:
            return "深灰"
        return "纯黑"

    if l < 0.25:
        prefix = "深"
    elif l > 0.72:
        prefix = "浅"
    else:
        prefix = ""

    if hue_deg < 15 or hue_deg >= 340:
        base = "红"
    elif hue_deg < 42:
        base = "橙"
    elif hue_deg < 68:
        base = "黄"
    elif hue_deg < 160:
        base = "绿"
    elif hue_deg < 195:
        base = "青"
    elif hue_deg < 255:
        base = "蓝"
    elif hue_deg < 300:
        base = "紫"
    else:
        base = "玫红"
    return f"{prefix}{base}"


def _parse_theme(zf: zipfile.ZipFile, theme_path: str) -> dict[str, Any]:
    root = _read_xml(zf, theme_path)
    colors: dict[str, str] = {}
    fonts: dict[str, str] = {"mj_ea": "", "mn_ea": "", "mj_lt": "", "mn_lt": ""}
    if root is None:
        return {"colors": colors, "fonts": fonts}

    clr_scheme = root.find(f".//{{{NS['a']}}}clrScheme")
    if clr_scheme is not None:
        for child in clr_scheme:
            slot = child.tag.split("}")[-1]
            srgb = child.find(f"{{{NS['a']}}}srgbClr")
            sysc = child.find(f"{{{NS['a']}}}sysClr")
            if srgb is not None and srgb.attrib.get("val"):
                colors[slot] = srgb.attrib["val"].upper()
            elif sysc is not None and sysc.attrib.get("lastClr"):
                colors[slot] = sysc.attrib["lastClr"].upper()

    font_scheme = root.find(f".//{{{NS['a']}}}fontScheme")
    if font_scheme is not None:
        maj = font_scheme.find(f"{{{NS['a']}}}majorFont")
        min_f = font_scheme.find(f"{{{NS['a']}}}minorFont")
        if maj is not None:
            lt = maj.find(f"{{{NS['a']}}}latin")
            ea = maj.find(f"{{{NS['a']}}}ea")
            fonts["mj_lt"] = lt.attrib.get("typeface", "") if lt is not None else ""
            fonts["mj_ea"] = ea.attrib.get("typeface", "") if ea is not None else ""
            if not fonts["mj_ea"]:
                for sf in maj.findall(f"{{{NS['a']}}}font"):
                    if sf.attrib.get("script") in ("Hans", "Hant") and sf.attrib.get("typeface"):
                        fonts["mj_ea"] = sf.attrib["typeface"]
                        break
        if min_f is not None:
            lt = min_f.find(f"{{{NS['a']}}}latin")
            ea = min_f.find(f"{{{NS['a']}}}ea")
            fonts["mn_lt"] = lt.attrib.get("typeface", "") if lt is not None else ""
            fonts["mn_ea"] = ea.attrib.get("typeface", "") if ea is not None else ""
            if not fonts["mn_ea"]:
                for sf in min_f.findall(f"{{{NS['a']}}}font"):
                    if sf.attrib.get("script") in ("Hans", "Hant") and sf.attrib.get("typeface"):
                        fonts["mn_ea"] = sf.attrib["typeface"]
                        break

    return {"colors": colors, "fonts": fonts}


def _resolve_color_elem(
    parent_elem: ET.Element | None,
    clr_map: dict[str, str],
    theme_colors: dict[str, str],
) -> str | None:
    """Resolve solidFill or color-holding element to RRGGBB hex string."""
    if parent_elem is None:
        return None
    srgb = parent_elem.find(f"{{{NS['a']}}}srgbClr")
    sysc = parent_elem.find(f"{{{NS['a']}}}sysClr")
    scheme = parent_elem.find(f"{{{NS['a']}}}schemeClr")

    base_hex: str | None = None
    mod_holder: ET.Element | None = None
    if srgb is not None and srgb.attrib.get("val"):
        base_hex = srgb.attrib["val"].upper()
        mod_holder = srgb
    elif sysc is not None and sysc.attrib.get("lastClr"):
        base_hex = sysc.attrib["lastClr"].upper()
        mod_holder = sysc
    elif scheme is not None and scheme.attrib.get("val"):
        key = scheme.attrib["val"]
        mapped = clr_map.get(key, key)
        base_hex = theme_colors.get(mapped) or theme_colors.get(key)
        mod_holder = scheme

    if not base_hex or mod_holder is None:
        return None

    def _int_child(tag: str) -> int | None:
        el = mod_holder.find(f"{{{NS['a']}}}{tag}")
        if el is not None and el.attrib.get("val"):
            try:
                return int(el.attrib["val"])
            except ValueError:
                return None
        return None

    return _apply_lum_mod_off(
        base_hex,
        _int_child("lumMod"),
        _int_child("lumOff"),
        _int_child("tint"),
        _int_child("shade"),
    )


def _resolve_typeface(name: str, theme_fonts: dict[str, str]) -> str:
    if not name:
        return ""
    mapping = {
        "+mj-ea": theme_fonts.get("mj_ea", ""),
        "+mn-ea": theme_fonts.get("mn_ea", ""),
        "+mj-lt": theme_fonts.get("mj_lt", ""),
        "+mn-lt": theme_fonts.get("mn_lt", ""),
    }
    return mapping.get(name, name) or name


def _extract_rpr_style(
    rpr: ET.Element | None,
    clr_map: dict[str, str],
    theme: dict[str, Any],
) -> dict[str, Any]:
    out: dict[str, Any] = {}
    if rpr is None:
        return out
    if "sz" in rpr.attrib:
        try:
            out["size_pt"] = round(int(rpr.attrib["sz"]) / 100.0, 1)
        except ValueError:
            pass
    if "b" in rpr.attrib:
        out["bold"] = rpr.attrib["b"] in ("1", "true", "True")

    ea = rpr.find(f"{{{NS['a']}}}ea")
    lt = rpr.find(f"{{{NS['a']}}}latin")
    fonts = theme["fonts"]
    if ea is not None and ea.attrib.get("typeface"):
        out["font"] = _resolve_typeface(ea.attrib["typeface"], fonts)
    elif lt is not None and lt.attrib.get("typeface"):
        out["font"] = _resolve_typeface(lt.attrib["typeface"], fonts)

    solid = rpr.find(f"{{{NS['a']}}}solidFill")
    if solid is not None:
        c = _resolve_color_elem(solid, clr_map, theme["colors"])
        if c:
            out["color"] = c
    return out


def _collect_placeholder_styles(
    root_elem: ET.Element | None,
    clr_map: dict[str, str],
    theme: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    """Collect default style per placeholder type ('title', 'ctrTitle', 'body', 'subTitle', '')."""
    styles: dict[str, dict[str, Any]] = {}
    if root_elem is None:
        return styles

    # Master txStyles fallback
    tx_styles = root_elem.find(f".//{{{NS['p']}}}txStyles")
    if tx_styles is not None:
        for xml_tag, ph_keys in (
            ("titleStyle", ("title", "ctrTitle")),
            ("bodyStyle", ("body", "subTitle", "")),
            ("otherStyle", ("other",)),
        ):
            style_el = tx_styles.find(f"{{{NS['p']}}}{xml_tag}")
            if style_el is not None:
                lvl1 = style_el.find(f"{{{NS['a']}}}lvl1pPr")
                rpr = lvl1.find(f"{{{NS['a']}}}defRPr") if lvl1 is not None else None
                parsed = _extract_rpr_style(rpr, clr_map, theme)
                for k in ph_keys:
                    styles.setdefault(k, {}).update(parsed)

    # Shape placeholders on layout/master
    for sp in root_elem.findall(f".//{{{NS['p']}}}sp"):
        ph = sp.find(f".//{{{NS['p']}}}ph")
        if ph is None:
            continue
        ph_type = ph.attrib.get("type", "body")
        lst = sp.find(f".//{{{NS['a']}}}lstStyle")
        lvl1 = lst.find(f"{{{NS['a']}}}lvl1pPr") if lst is not None else None
        rpr = lvl1.find(f"{{{NS['a']}}}defRPr") if lvl1 is not None else None
        parsed = _extract_rpr_style(rpr, clr_map, theme)
        if not parsed:
            # Check first run inside layout placeholder
            first_rpr = sp.find(f".//{{{NS['a']}}}rPr")
            parsed = _extract_rpr_style(first_rpr, clr_map, theme)
        if parsed:
            merged = dict(styles.get(ph_type, {}))
            merged.update(parsed)
            styles[ph_type] = merged
    return styles


def check_file(path_str: str) -> dict[str, Any]:
    """检查文件真实格式头、跨软件格式（Keynote/ODP）、企业透明加密与多媒体体积。"""
    path = Path(path_str)
    ext = path.suffix.lower()

    if path.is_dir() and ext == ".key":
        return {
            "path": str(path),
            "name": path.name,
            "status": "NEEDS_USER_CONVERT",
            "format_kind": "keynote",
            "size_mb": 0.0,
            "human_hint": (
                f"《{path.name}》是苹果 Keynote 演示文稿。"
                "请用 Keynote 打开它，点击菜单栏「文件 → 导出到 → PowerPoint…（选 .pptx）」保存后再发我即可。"
            ),
        }

    if not path.is_file():
        return {"path": str(path), "status": "MISSING", "human_hint": f"未找到文件：{path.name}"}

    size_bytes = path.stat().st_size
    size_mb = round(size_bytes / (1024 * 1024), 2)

    if ext == ".key":
        return {
            "path": str(path),
            "name": path.name,
            "status": "NEEDS_USER_CONVERT",
            "format_kind": "keynote",
            "size_mb": size_mb,
            "human_hint": (
                f"《{path.name}》是苹果 Keynote 演示文稿。"
                "请用 Keynote 打开它，点击「文件 → 导出到 → PowerPoint…（.pptx）」保存后再发我即可。"
            ),
        }

    if ext == ".odp":
        return {
            "path": str(path),
            "name": path.name,
            "status": "NEEDS_USER_CONVERT",
            "format_kind": "odp",
            "size_mb": size_mb,
            "human_hint": (
                f"《{path.name}》是 LibreOffice/OpenDocument 演示文稿。"
                "请在软件里点击「文件 → 另存为 → PowerPoint 2007-365 (.pptx)」后再发我。"
            ),
        }

    with path.open("rb") as fh:
        header = fh.read(16)

    if ext in IMG_EXTENSIONS:
        return {
            "path": str(path),
            "name": path.name,
            "status": "VISUAL_REF_ONLY",
            "size_mb": size_mb,
            "human_hint": "图片仅可作为本轮配色与版面参考，不能作为建档的基础代表作（建档至少需要 1 份 .pptx）。",
        }

    if header.startswith(PDF_MAGIC) or ext == ".pdf":
        return {
            "path": str(path),
            "name": path.name,
            "status": "VISUAL_REF_ONLY",
            "size_mb": size_mb,
            "human_hint": "PDF 仅可作为辅助视觉风格参考（能看版面与配色，无法读取精确可编辑母版和字体）。",
        }

    if header.startswith(OLE_MAGIC) or ext in (".ppt", ".dps"):
        label = "老式 .ppt 格式" if ext == ".ppt" else ("WPS 专属 .dps 格式" if ext == ".dps" else "加密或旧版演示文稿")
        return {
            "path": str(path),
            "name": path.name,
            "status": "NEEDS_USER_CONVERT",
            "format_kind": ext.lstrip(".") or "ole",
            "size_mb": size_mb,
            "human_hint": (
                f"《{path.name}》属于{label}格式，直接读取会丢掉样式或提示加密。"
                "请用电脑上的 PowerPoint 或 WPS 打开它，点击「文件 → 另存为 → 选择 .pptx 格式」保存后再发我即可。"
            ),
        }

    if not header.startswith(ZIP_MAGIC):
        if ext == ".pptx" and size_bytes >= 1024:
            return {
                "path": str(path),
                "name": path.name,
                "status": "SUSPECTED_DRM_ENCRYPTED",
                "suspected_drm": True,
                "size_mb": size_mb,
                "human_hint": (
                    f"《{path.name}》在你电脑软件里也许能正常打开，但底层读取时是加密乱码，大概率是公司装了文件透明加密软件。"
                    "请在非加密区重新保存一份，或先导出 PDF 发我走纯对话降级模式。"
                ),
            }
        return {
            "path": str(path),
            "name": path.name,
            "status": "UNSUPPORTED",
            "size_mb": size_mb,
            "human_hint": f"《{path.name}》不是有效的演示文稿文件。",
        }

    try:
        with zipfile.ZipFile(path, "r") as zf:
            names = zf.namelist()
            if "ppt/presentation.xml" not in names:
                return {
                    "path": str(path),
                    "name": path.name,
                    "status": "NEEDS_USER_CONVERT",
                    "size_mb": size_mb,
                    "human_hint": f"《{path.name}》内部不是标准 .pptx 结构，请先在软件里「另存为 .pptx」后再试。",
                }
            av_bytes = sum(
                zf.getinfo(n).file_size
                for n in names
                if n.startswith("ppt/media/") and Path(n).suffix.lower() in MEDIA_AV_EXTENSIONS
            )
            av_mb = round(av_bytes / (1024 * 1024), 2)
    except zipfile.BadZipFile:
        return {
            "path": str(path),
            "name": path.name,
            "status": "CORRUPTED",
            "size_mb": size_mb,
            "human_hint": f"《{path.name}》文件已损坏无法打开，请检查原文件是否完整。",
        }

    return {
        "path": str(path),
        "name": path.name,
        "status": "VALID_PPTX",
        "size_mb": size_mb,
        "media_av_mb": av_mb,
        "has_heavy_media": size_mb > 50.0 or av_mb > 10.0,
        "has_av_to_strip": (av_mb > 10.0) or (size_mb > 50.0 and av_mb > 0.0),
        "sha256": _sha256_file(path),
    }


def _extract_chart_kind(chart_xml: ET.Element | None) -> str:
    if chart_xml is None:
        return "其他图表"
    if chart_xml.find(f".//{{{NS['c']}}}barChart") is not None or chart_xml.find(f".//{{{NS['c']}}}bar3DChart") is not None:
        return "柱状图"
    if chart_xml.find(f".//{{{NS['c']}}}lineChart") is not None:
        return "折线图"
    if chart_xml.find(f".//{{{NS['c']}}}pieChart") is not None or chart_xml.find(f".//{{{NS['c']}}}doughnutChart") is not None:
        return "饼图"
    return "复杂统计图"


def _box_overlap_ratio(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    ax1, ay1, aw, ah = a
    bx1, by1, bw, bh = b
    ax2, ay2 = ax1 + aw, ay1 + ah
    bx2, by2 = bx1 + bw, by1 + bh
    ix1, iy1 = max(ax1, bx1), max(ay1, by1)
    ix2, iy2 = min(ax2, bx2), min(ay2, by2)
    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0
    inter = (ix2 - ix1) * (iy2 - iy1)
    min_area = min(max(aw * ah, 1e-6), max(bw * bh, 1e-6))
    return inter / min_area


def _get_xfrm_raw_emu(elem: ET.Element) -> tuple[int, int, int, int] | None:
    xfrm = elem.find(f"./{{{NS['p']}}}spPr/{{{NS['a']}}}xfrm")
    if xfrm is None:
        xfrm = elem.find(f"./{{{NS['p']}}}grpSpPr/{{{NS['a']}}}xfrm")
    if xfrm is None:
        xfrm = elem.find(f"./{{{NS['p']}}}xfrm")
    if xfrm is None:
        return None
    off = xfrm.find(f"{{{NS['a']}}}off")
    ext = xfrm.find(f"{{{NS['a']}}}ext")
    if off is None or ext is None:
        return None
    try:
        return (
            int(off.attrib.get("x", "0")),
            int(off.attrib.get("y", "0")),
            int(ext.attrib.get("cx", "0")),
            int(ext.attrib.get("cy", "0")),
        )
    except ValueError:
        return None


def _get_xfrm_inches(sp_elem: ET.Element) -> tuple[float, float, float, float] | None:
    raw = _get_xfrm_raw_emu(sp_elem)
    if raw is None:
        return None
    return (
        raw[0] / EMU_PER_INCH,
        raw[1] / EMU_PER_INCH,
        raw[2] / EMU_PER_INCH,
        raw[3] / EMU_PER_INCH,
    )


def _iter_shapes_with_slide_boxes(
    sp_tree: ET.Element | None,
) -> list[tuple[ET.Element, tuple[float, float, float, float] | None, bool]]:
    """递归展开 <p:grpSp>，把组内相对坐标准确换算成整页英寸绝对坐标。"""
    if sp_tree is None:
        return []
    out: list[tuple[ET.Element, tuple[float, float, float, float] | None, bool]] = []
    sp_tag = f"{{{NS['p']}}}sp"
    grp_tag = f"{{{NS['p']}}}grpSp"

    def _walk(
        container: ET.Element,
        tx: float,
        ty: float,
        sx: float,
        sy: float,
        in_group: bool,
    ) -> None:
        for child in container:
            if child.tag == sp_tag:
                raw = _get_xfrm_raw_emu(child)
                box = (tx + raw[0] * sx, ty + raw[1] * sy, raw[2] * sx, raw[3] * sy) if raw is not None else None
                out.append((child, box, in_group))
            elif child.tag == grp_tag:
                xfrm = child.find(f"./{{{NS['p']}}}grpSpPr/{{{NS['a']}}}xfrm")
                ntx, nty, nsx, nsy = tx, ty, sx, sy
                if xfrm is not None:
                    off = xfrm.find(f"{{{NS['a']}}}off")
                    ext = xfrm.find(f"{{{NS['a']}}}ext")
                    choff = xfrm.find(f"{{{NS['a']}}}chOff")
                    chext = xfrm.find(f"{{{NS['a']}}}chExt")
                    if off is not None and ext is not None and choff is not None and chext is not None:
                        try:
                            gx = int(off.attrib.get("x", "0"))
                            gy = int(off.attrib.get("y", "0"))
                            gcw = int(ext.attrib.get("cx", "0"))
                            gch = int(ext.attrib.get("cy", "0"))
                            chx = int(choff.attrib.get("x", "0"))
                            chy = int(choff.attrib.get("y", "0"))
                            chcw = int(chext.attrib.get("cx", "0"))
                            chch = int(chext.attrib.get("cy", "0"))
                            if chcw > 0 and chch > 0:
                                nsx = sx * (gcw / chcw)
                                nsy = sy * (gch / chch)
                                ntx = tx + gx * sx - chx * nsx
                                nty = ty + gy * sy - chy * nsy
                        except ValueError:
                            pass
                _walk(child, ntx, nty, nsx, nsy, True)

    _walk(sp_tree, 0.0, 0.0, 1.0 / EMU_PER_INCH, 1.0 / EMU_PER_INCH, False)
    return out


def _resolve_bg_info(
    root_elem: ET.Element | None,
    clr_map: dict[str, str],
    theme_colors: dict[str, str],
) -> tuple[str | None, str]:
    """解析单页、版式或母版自身的背景（纯色 / 渐变 / 底图 / 主题引用）。"""
    if root_elem is None:
        return (None, "none")
    bg_el = root_elem.find(f"./{{{NS['p']}}}cSld/{{{NS['p']}}}bg")
    if bg_el is None:
        return (None, "none")

    bg_pr = bg_el.find(f"{{{NS['p']}}}bgPr")
    if bg_pr is not None:
        solid = bg_pr.find(f"{{{NS['a']}}}solidFill")
        if solid is not None:
            return (_resolve_color_elem(solid, clr_map, theme_colors), "solid")
        grad = bg_pr.find(f"{{{NS['a']}}}gradFill")
        if grad is not None:
            gs = grad.find(f".//{{{NS['a']}}}gs")
            col = _resolve_color_elem(gs, clr_map, theme_colors) if gs is not None else None
            return (col, "gradient")
        blip = bg_pr.find(f"{{{NS['a']}}}blipFill")
        if blip is not None:
            return (None, "image")

    bg_ref = bg_el.find(f"{{{NS['p']}}}bgRef")
    if bg_ref is not None:
        col = _resolve_color_elem(bg_ref, clr_map, theme_colors)
        return (col, "bgRef")

    return (None, "none")


def _is_conclusion_title(title: str) -> bool:
    t = _clean_text(title)
    if not t or len(t) < 8:
        return False
    return bool(re.search(r"[:：，,]|(?:提升|下降|减少|缩短|增长|达成|实现|完成|推动|覆盖|建议|问题|原因|结论|方案)", t))


def inspect_deck(path_str: str, selected_slides: set[int] | None = None) -> dict[str, Any]:
    """对单份 .pptx 进行深层结构、主题继承、精确色板、多信号合写预分诊与硬伤扫描。"""
    path = Path(path_str)
    chk = check_file(path_str)
    if chk.get("status") != "VALID_PPTX":
        return {"file": path.name, "path": str(path), "error": chk}

    slides_data: list[dict[str, Any]] = []
    metadata: dict[str, str] = {"creator": "", "last_modified_by": ""}
    with zipfile.ZipFile(path, "r") as zf:
        core_xml = _read_xml(zf, "docProps/core.xml")
        if core_xml is not None:
            cr_el = core_xml.find(f"{{{NS['dc']}}}creator")
            lm_el = core_xml.find(f"{{{NS['cp']}}}lastModifiedBy")
            metadata["creator"] = _clean_text(cr_el.text) if (cr_el is not None and cr_el.text) else ""
            metadata["last_modified_by"] = _clean_text(lm_el.text) if (lm_el is not None and lm_el.text) else ""

        pres = _read_xml(zf, "ppt/presentation.xml")
        pres_rels = _read_rels(zf, "ppt/presentation.xml")
        if pres is None:
            return {"file": path.name, "error": {"status": "CORRUPTED", "human_hint": "缺少演示主索引"}}

        sld_sz = pres.find(f"{{{NS['p']}}}sldSz")
        slide_w_in = round(int(sld_sz.attrib.get("cx", "12192000")) / EMU_PER_INCH, 2) if sld_sz is not None else 13.33
        slide_h_in = round(int(sld_sz.attrib.get("cy", "6858000")) / EMU_PER_INCH, 2) if sld_sz is not None else 7.5

        sld_ids = pres.findall(f".//{{{NS['p']}}}sldIdLst/{{{NS['p']}}}sldId")
        slide_paths: list[str] = []
        for sld_id in sld_ids:
            rid = sld_id.attrib.get(f"{{{NS['r']}}}id", "")
            if rid in pres_rels:
                slide_paths.append(pres_rels[rid][1])

        theme_cache: dict[str, dict[str, Any]] = {}
        master_cache: dict[
            str,
            tuple[dict[str, str], dict[str, Any], dict[str, dict[str, Any]], tuple[str | None, str]],
        ] = {}
        layout_cache: dict[
            str,
            tuple[str, dict[str, dict[str, Any]], tuple[str | None, str]],
        ] = {}

        for idx, sld_path in enumerate(slide_paths, 1):
            sld_root = _read_xml(zf, sld_path)
            if sld_root is None:
                continue
            sld_rels = _read_rels(zf, sld_path)

            layout_path = ""
            notes_path = ""
            chart_paths: list[str] = []
            has_smartart = False
            for _rid, (rtype, target) in sld_rels.items():
                if rtype.endswith("/slideLayout"):
                    layout_path = target
                elif rtype.endswith("/notesSlide"):
                    notes_path = target
                elif rtype.endswith("/chart"):
                    chart_paths.append(target)
                elif "diagram" in rtype:
                    has_smartart = True

            master_path = ""
            clr_map: dict[str, str] = {}
            theme: dict[str, Any] = {"colors": {}, "fonts": {}}
            master_ph_styles: dict[str, dict[str, Any]] = {}
            layout_ph_styles: dict[str, dict[str, Any]] = {}
            m_bg_info: tuple[str | None, str] = (None, "none")
            l_bg_info: tuple[str | None, str] = (None, "none")

            if layout_path:
                if layout_path not in layout_cache:
                    lay_rels = _read_rels(zf, layout_path)
                    m_path = next((t for (_rt, t) in lay_rels.values() if _rt.endswith("/slideMaster")), "")
                    lay_xml = _read_xml(zf, layout_path)
                    if m_path and m_path not in master_cache:
                        m_xml = _read_xml(zf, m_path)
                        m_rels = _read_rels(zf, m_path)
                        t_path = next((t for (_rt, t) in m_rels.values() if _rt.endswith("/theme")), "")
                        if t_path and t_path not in theme_cache:
                            theme_cache[t_path] = _parse_theme(zf, t_path)
                        t_data = theme_cache.get(t_path, {"colors": {}, "fonts": {}})
                        cmap_el = m_xml.find(f"{{{NS['p']}}}clrMap") if m_xml is not None else None
                        cmap = dict(cmap_el.attrib) if cmap_el is not None else {}
                        m_styles = _collect_placeholder_styles(m_xml, cmap, t_data)
                        m_bg = _resolve_bg_info(m_xml, cmap, t_data["colors"])
                        master_cache[m_path] = (cmap, t_data, m_styles, m_bg)
                    if m_path in master_cache:
                        cmap, t_data, _, _ = master_cache[m_path]
                        l_styles = _collect_placeholder_styles(lay_xml, cmap, t_data)
                        l_bg = _resolve_bg_info(lay_xml, cmap, t_data["colors"])
                    else:
                        l_styles = {}
                        l_bg = (None, "none")
                    layout_cache[layout_path] = (m_path, l_styles, l_bg)

                master_path, layout_ph_styles, l_bg_info = layout_cache[layout_path]
                if master_path in master_cache:
                    clr_map, theme, master_ph_styles, m_bg_info = master_cache[master_path]

            canvas_scale = max(slide_w_in / 13.333, 0.25)

            title_text = ""
            body_texts: list[str] = []
            title_fonts: list[str] = []
            body_fonts: list[str] = []
            title_sizes: list[float] = []
            body_sizes: list[float] = []
            accent_colors: Counter[str] = Counter()
            accent_hexes: Counter[str] = Counter()
            title_color_hexes: Counter[str] = Counter()
            body_color_hexes: Counter[str] = Counter()
            table_count = 0
            table_header_colors: list[str] = []
            table_header_hexes: list[str] = []
            pic_count = 0
            pic_area = 0.0
            connector_count = len(sld_root.findall(f".//{{{NS['p']}}}cxnSp"))
            group_count = len(sld_root.findall(f".//{{{NS['p']}}}grpSp"))
            non_footer_boxes: list[tuple[float, float, float, float]] = []
            hygiene_issues: list[str] = []
            ph_types_present: set[str] = set()

            sp_tree = sld_root.find(f"./{{{NS['p']}}}cSld/{{{NS['p']}}}spTree")
            shape_entries = _iter_shapes_with_slide_boxes(sp_tree)
            if not shape_entries:
                shape_entries = [
                    (sp, _get_xfrm_inches(sp), False)
                    for sp in sld_root.findall(f".//{{{NS['p']}}}sp")
                ]

            for sp, box, in_group in shape_entries:
                ph = sp.find(f".//{{{NS['p']}}}ph")
                ph_type = ph.attrib.get("type", "") if ph is not None else ""
                if ph_type:
                    ph_types_present.add(ph_type)
                is_footer = ph_type in ("ftr", "sldNum", "dt")
                if box and box[1] > slide_h_in * 0.88 and box[3] < slide_h_in * 0.14:
                    is_footer = True

                tx_body = sp.find(f"{{{NS['p']}}}txBody")
                if tx_body is None:
                    continue

                para_texts: list[str] = []
                base_ph_style = dict(master_ph_styles.get(ph_type, {}))
                base_ph_style.update(layout_ph_styles.get(ph_type, {}))
                sp_lst_lvl1 = tx_body.find(f"./{{{NS['a']}}}lstStyle/{{{NS['a']}}}lvl1pPr/{{{NS['a']}}}defRPr")
                sp_style = dict(base_ph_style)
                sp_style.update(_extract_rpr_style(sp_lst_lvl1, clr_map, theme))

                is_title_shape = ph_type in ("title", "ctrTitle")
                for p_el in tx_body.findall(f"{{{NS['a']}}}p"):
                    p_def = p_el.find(f"./{{{NS['a']}}}pPr/{{{NS['a']}}}defRPr")
                    p_style = dict(sp_style)
                    p_style.update(_extract_rpr_style(p_def, clr_map, theme))
                    runs_txt: list[str] = []
                    for r_el in p_el.findall(f"{{{NS['a']}}}r"):
                        t_el = r_el.find(f"{{{NS['a']}}}t")
                        txt = t_el.text if (t_el is not None and t_el.text) else ""
                        if not txt:
                            continue
                        runs_txt.append(txt)
                        rpr = r_el.find(f"{{{NS['a']}}}rPr")
                        r_style = dict(p_style)
                        r_style.update(_extract_rpr_style(rpr, clr_map, theme))

                        sz = r_style.get("size_pt")
                        fn = r_style.get("font")
                        col = r_style.get("color")
                        if not sz and not is_footer:
                            sz = (18.0 if is_title_shape else 14.0) * canvas_scale
                        norm_sz = round(sz / canvas_scale, 1) if sz else None
                        if is_title_shape or (box and box[1] < slide_h_in * 0.22 and norm_sz and norm_sz >= 18.0 and not title_text):
                            if norm_sz:
                                title_sizes.append(norm_sz)
                            if fn:
                                title_fonts.append(fn)
                            if col:
                                title_color_hexes[col] += 1
                        elif not is_footer:
                            if norm_sz:
                                body_sizes.append(norm_sz)
                                if norm_sz < 10.0:
                                    hygiene_issues.append("small_body_font")
                            if fn:
                                body_fonts.append(fn)
                            if col:
                                body_color_hexes[col] += 1
                        if col and col not in ("FFFFFF", "000000", "333333", "595959"):
                            accent_colors[rgb_to_cn_name(col)] += 1
                            accent_hexes[col] += 1

                    joined_p = _clean_text("".join(runs_txt))
                    if joined_p:
                        para_texts.append(joined_p)

                full_sp_text = " ".join(para_texts).strip()
                norm_sp = _norm_text(full_sp_text)
                if full_sp_text and not is_footer:
                    if is_title_shape and not title_text:
                        if len(norm_sp) >= 2 or not norm_sp:
                            title_text = full_sp_text
                        else:
                            body_texts.append(full_sp_text)
                    elif (
                        not title_text
                        and box
                        and box[1] < slide_h_in * 0.24
                        and len(full_sp_text) <= 80
                        and len(norm_sp) >= 2
                        and not re.fullmatch(r"\d+", norm_sp)
                    ):
                        title_text = full_sp_text
                    else:
                        body_texts.append(full_sp_text)

                    if (not in_group) and box and box[2] < slide_w_in * 0.95 and box[3] < slide_h_in * 0.90:
                        if len(norm_sp) >= 3 and (
                            box[0] < 0.20 * canvas_scale
                            or (box[0] + box[2]) > (slide_w_in - 0.20 * canvas_scale)
                            or box[1] < 0.15 * canvas_scale
                            or (box[1] + box[3]) > (slide_h_in - 0.15 * canvas_scale)
                        ):
                            hygiene_issues.append("edge_overflow")
                        if len(norm_sp) >= 4:
                            for prev_box in non_footer_boxes:
                                if _box_overlap_ratio(box, prev_box) > 0.20:
                                    hygiene_issues.append("text_box_overlap")
                                    break
                            non_footer_boxes.append(box)

            # Tables
            for tbl in sld_root.findall(f".//{{{NS['a']}}}tbl"):
                table_count += 1
                rows = tbl.findall(f"{{{NS['a']}}}tr")
                if rows:
                    first_tc = rows[0].find(f"{{{NS['a']}}}tc")
                    if first_tc is not None:
                        fill = first_tc.find(f".//{{{NS['a']}}}tcPr/{{{NS['a']}}}solidFill")
                        tc_hex = _resolve_color_elem(fill, clr_map, theme["colors"])
                        if tc_hex:
                            table_header_colors.append(rgb_to_cn_name(tc_hex))
                            table_header_hexes.append(tc_hex)
                for tc in tbl.findall(f".//{{{NS['a']}}}tc"):
                    cell_txt = _clean_text("".join(t.text or "" for t in tc.findall(f".//{{{NS['a']}}}t")))
                    if cell_txt:
                        body_texts.append(cell_txt)

            # Pictures
            has_full_bleed_pic = False
            for pic in sld_root.findall(f".//{{{NS['p']}}}pic"):
                box = _get_xfrm_inches(pic)
                if box and box[2] >= slide_w_in * 0.95 and box[3] >= slide_h_in * 0.95:
                    has_full_bleed_pic = True
                    continue  # skip full-bleed background picture
                pic_count += 1
                if box:
                    pic_area += box[2] * box[3]

            # Inherited slide background resolution: slide -> layout -> master -> full-bleed pic -> theme
            s_bg_hex, s_bg_kind = _resolve_bg_info(sld_root, clr_map, theme["colors"])
            bg_hex: str | None = None
            bg_kind = "none"
            for cand_hex, cand_kind in ((s_bg_hex, s_bg_kind), l_bg_info, m_bg_info):
                if cand_kind != "none":
                    bg_hex, bg_kind = cand_hex, cand_kind
                    break

            if (
                bg_kind == "none"
                or (s_bg_kind == "none" and l_bg_info[1] == "none" and bg_kind == "bgRef" and (bg_hex or "FFFFFF") == "FFFFFF")
            ) and has_full_bleed_pic:
                bg_kind = "image"
                bg_hex = None
            elif bg_kind == "none":
                bg_hex = (
                    (clr_map.get("bg1") and theme["colors"].get(clr_map["bg1"]))
                    or theme["colors"].get("lt1")
                    or "FFFFFF"
                )
                bg_kind = "theme_default"

            is_dark_bg = False
            if bg_kind == "image" and bg_hex is None:
                bg_color_cn = "底图背景"
            else:
                eff_hex = bg_hex or "FFFFFF"
                bg_color_cn = rgb_to_cn_name(eff_hex)
                r, g, b = int(eff_hex[0:2], 16), int(eff_hex[2:4], 16), int(eff_hex[4:6], 16)
                is_dark_bg = (0.299 * r + 0.587 * g + 0.114 * b) < 128

            # Charts
            chart_kinds: list[str] = []
            for cp in chart_paths:
                chart_kinds.append(_extract_chart_kind(_read_xml(zf, cp)))

            # Notes
            notes_chars = 0
            if notes_path:
                n_xml = _read_xml(zf, notes_path)
                if n_xml is not None:
                    notes_texts = []
                    for sp in n_xml.findall(f".//{{{NS['p']}}}sp"):
                        ph = sp.find(f".//{{{NS['p']}}}ph")
                        if ph is not None and ph.attrib.get("type") == "body":
                            notes_texts.append("".join(t.text or "" for t in sp.findall(f".//{{{NS['a']}}}t")))
                    notes_chars = len(_clean_text(" ".join(notes_texts)))

            slide_char_count = len(_clean_text(title_text + " " + " ".join(body_texts)))
            full_text = _clean_text(title_text + " | " + " | ".join(body_texts))
            speakers = sorted(set(SPEAKER_SIGN_RE.findall(full_text)))

            # Classify special page roles (cover / toc / section / closing / content)
            page_role = "content"
            t_low = title_text.lower()
            if idx == 1 and table_count == 0 and not chart_kinds and (
                "ctrTitle" in ph_types_present
                or (len(slide_paths) >= 3 and slide_char_count <= 60 and not _is_conclusion_title(title_text))
            ):
                page_role = "cover"
            elif any(k in t_low for k in ("目录", "议程", "内容概览", "agenda", "contents")) and table_count == 0 and not chart_kinds:
                page_role = "toc"
            elif idx == len(slide_paths) and len(slide_paths) >= 3 and any(k in t_low for k in ("谢谢", "感谢", "q&a", "thank", "结束")):
                page_role = "closing"
            elif ("secTitle" in ph_types_present) or (
                len(slide_paths) >= 4
                and slide_char_count <= 35
                and not body_texts
                and table_count == 0
                and not chart_kinds
                and pic_count == 0
                and not _is_conclusion_title(title_text)
            ):
                page_role = "section_divider"

            slides_data.append({
                "slide_no": idx,
                "included": (selected_slides is None) or (idx in selected_slides),
                "page_role": page_role,
                "master_path": master_path,
                "layout_path": layout_path,
                "bg_color_cn": bg_color_cn,
                "bg_hex": (bg_hex or "FFFFFF") if bg_kind != "image" else "",
                "bg_kind": bg_kind,
                "is_dark_bg": is_dark_bg,
                "title": title_text,
                "snippet": _clean_text(" ".join(body_texts))[:60],
                "full_text": full_text,
                "speakers": speakers,
                "slide_char_count": slide_char_count,
                "notes_char_count": notes_chars,
                "title_fonts": sorted(set(title_fonts)),
                "body_fonts": sorted(set(body_fonts)),
                "title_sizes": sorted(set(title_sizes)),
                "body_sizes": sorted(set(body_sizes)),
                "accent_colors_cn": [c for c, _ in accent_colors.most_common(3)],
                "accent_hexes": [c for c, _ in accent_hexes.most_common(3)],
                "title_color_hexes": [c for c, _ in title_color_hexes.most_common(2)],
                "body_color_hexes": [c for c, _ in body_color_hexes.most_common(2)],
                "table_count": table_count,
                "table_header_colors_cn": table_header_colors,
                "table_header_hexes": table_header_hexes,
                "chart_kinds": chart_kinds,
                "has_smartart": has_smartart,
                "vector_diagram_signals": connector_count + group_count,
                "pic_count": pic_count,
                "pic_area_ratio": round(pic_area / max(slide_w_in * slide_h_in, 1.0), 3),
                "hygiene_issues": sorted(set(hygiene_issues)),
            })

    # Detect style anomaly content pages inside a single deck
    content_slides = [s for s in slides_data if s["page_role"] == "content"]
    master_counts = Counter(s["master_path"] for s in content_slides if s["master_path"])
    dominant_master = master_counts.most_common(1)[0][0] if master_counts else ""
    anomalies: list[dict[str, Any]] = []
    for s in content_slides:
        if dominant_master and s["master_path"] and s["master_path"] != dominant_master:
            anomalies.append({
                "slide_no": s["slide_no"],
                "title": s["title"] or s["snippet"],
                "reason": "母版风格与整份主调明显不同（疑似借用他人页面）",
            })

    # Build 4-6 readable chapter blocks skipping cover/toc/closing
    blocks: list[dict[str, Any]] = []
    current_slides: list[dict[str, Any]] = []
    current_title = "正文内容"
    for s in slides_data:
        if s["page_role"] in ("cover", "toc", "closing"):
            continue
        if s["page_role"] == "section_divider":
            if current_slides:
                blocks.append({
                    "block_title": current_title,
                    "slide_range": f"P{current_slides[0]['slide_no']}-P{current_slides[-1]['slide_no']}",
                    "slide_nos": [x["slide_no"] for x in current_slides],
                    "samples": [{"slide_no": x["slide_no"], "title": x["title"] or x["snippet"]} for x in current_slides[:5]],
                })
                current_slides = []
            current_title = s["title"] or f"第{s['slide_no']}页章节"
            continue
        current_slides.append(s)
    if current_slides:
        blocks.append({
            "block_title": current_title,
            "slide_range": f"P{current_slides[0]['slide_no']}-P{current_slides[-1]['slide_no']}",
            "slide_nos": [x["slide_no"] for x in current_slides],
            "samples": [{"slide_no": x["slide_no"], "title": x["title"] or x["snippet"]} for x in current_slides[:5]],
        })

    # If no section dividers produced >=2 blocks and deck is long, chunk into <=5 blocks
    if len(blocks) == 1 and len(blocks[0]["slide_nos"]) > 8:
        all_cs = [s for s in content_slides]
        chunk_size = max(3, (len(all_cs) + 4) // 5)
        blocks = []
        for i in range(0, len(all_cs), chunk_size):
            chunk = all_cs[i : i + chunk_size]
            blocks.append({
                "block_title": chunk[0]["title"] or f"第{chunk[0]['slide_no']}页起",
                "slide_range": f"P{chunk[0]['slide_no']}-P{chunk[-1]['slide_no']}",
                "slide_nos": [x["slide_no"] for x in chunk],
                "samples": [{"slide_no": x["slide_no"], "title": x["title"] or x["snippet"]} for x in chunk[:4]],
            })

    # Theme tokens for this deck
    deck_title_fonts = Counter(f for s in content_slides for f in s["title_fonts"])
    deck_body_fonts = Counter(f for s in content_slides for f in s["body_fonts"])
    deck_title_sizes = [sz for s in content_slides for sz in s["title_sizes"]]
    deck_body_sizes = [sz for s in content_slides for sz in s["body_sizes"]]
    deck_tbl_hex = Counter(h for s in content_slides for h in s["table_header_hexes"])
    deck_title_hex = Counter(h for s in content_slides for h in s["title_color_hexes"])
    deck_accent_hex = Counter(h for s in content_slides for h in s["accent_hexes"])
    deck_bg_hex = Counter(s["bg_hex"] for s in content_slides if s["bg_hex"])
    deck_bg_cn = Counter(s["bg_color_cn"] for s in content_slides)
    deck_dark_cnt = sum(1 for s in content_slides if s["is_dark_bg"])
    deck_is_dark = bool(content_slides) and (deck_dark_cnt / len(content_slides)) >= 0.5

    primary_hex = (
        (deck_tbl_hex.most_common(1)[0][0] if deck_tbl_hex else "")
        or (deck_title_hex.most_common(1)[0][0] if deck_title_hex else "")
        or (deck_accent_hex.most_common(1)[0][0] if deck_accent_hex else "")
        or ("3B82F6" if deck_is_dark else "1F4E79")
    )
    accent_hex = next((h for h, _ in deck_accent_hex.most_common(3) if h != primary_hex), "") or "D96B27"
    bg_hex_val = (deck_bg_hex.most_common(1)[0][0] if deck_bg_hex else "") or ("162238" if deck_is_dark else "FFFFFF")
    theme_tokens = {
        "primary_hex": primary_hex,
        "accent_hex": accent_hex,
        "bg_hex": bg_hex_val,
        "text_hex": "EBF0F5" if deck_is_dark else "222222",
        "table_header_hex": deck_tbl_hex.most_common(1)[0][0] if deck_tbl_hex else primary_hex,
        "title_font": deck_title_fonts.most_common(1)[0][0] if deck_title_fonts else "Microsoft YaHei",
        "body_font": deck_body_fonts.most_common(1)[0][0] if deck_body_fonts else "Microsoft YaHei",
        "title_size_pt": round(statistics.median(deck_title_sizes), 1) if deck_title_sizes else 22.0,
        "body_size_pt": round(statistics.median(deck_body_sizes), 1) if deck_body_sizes else 14.0,
        "is_dark_bg": deck_is_dark,
        "bg_color_cn": deck_bg_cn.most_common(1)[0][0] if deck_bg_cn else "纯白",
        "primary_color_cn": rgb_to_cn_name(primary_hex),
    }

    # Multi-signal collaboration / template pre-triage
    placeholder_hits = sum(1 for s in content_slides if TEMPLATE_PLACEHOLDER_RE.search(s["full_text"]))
    total_content_chars = sum(s["slide_char_count"] for s in content_slides)
    all_speakers = sorted({sp for s in slides_data for sp in s.get("speakers", [])})
    collab_signals: list[str] = []

    if len(master_counts) >= 2:
        collab_signals.append(f"文件内混用了 {len(master_counts)} 套不同母版")
    if len(all_speakers) >= 2:
        collab_signals.append(f"页内出现多位不同汇报人署名（{'、'.join(all_speakers[:4])}）")
    if COLLAB_FILENAME_RE.search(path.stem) and len(slides_data) >= 6:
        collab_signals.append("文件名含多人汇总或例会合并特征")

    # Same-master segment style shift check (fonts, table header colors, bg flip, density jump)
    if len(content_slides) >= 6:
        mid = len(content_slides) // 2
        seg_a, seg_b = content_slides[:mid], content_slides[mid:]
        shift_reasons: list[str] = []
        fa = Counter(f for s in seg_a for f in s["body_fonts"]).most_common(1)
        fb = Counter(f for s in seg_b for f in s["body_fonts"]).most_common(1)
        if fa and fb and fa[0][0] != fb[0][0]:
            shift_reasons.append("前后半段正文字体不同")
        ha = Counter(c for s in seg_a for c in s["table_header_colors_cn"]).most_common(1)
        hb = Counter(c for s in seg_b for c in s["table_header_colors_cn"]).most_common(1)
        if ha and hb and ha[0][0] != hb[0][0]:
            shift_reasons.append("前后半段表格表头配色不同")
        da = sum(1 for s in seg_a if s["is_dark_bg"]) / max(len(seg_a), 1)
        db = sum(1 for s in seg_b if s["is_dark_bg"]) / max(len(seg_b), 1)
        if abs(da - db) >= 0.7:
            shift_reasons.append("前后半段深浅背景反转")
        ca = sum(s["slide_char_count"] for s in seg_a) / max(len(seg_a), 1)
        cb = sum(s["slide_char_count"] for s in seg_b) / max(len(seg_b), 1)
        if max(ca, cb) >= max(min(ca, cb), 1.0) * 2.5 and abs(ca - cb) >= 120:
            shift_reasons.append("前后半段文字密度大幅跳变")
        if len(shift_reasons) >= 2:
            collab_signals.append("同母版下存在风格突变（" + "、".join(shift_reasons) + "）")

    if (len(content_slides) < 2 and total_content_chars < 60) or (
        content_slides and (placeholder_hits / len(content_slides)) >= 0.6
    ):
        triage_group = "gray_template"
        triage_reason = "页面多为空白占位框架，判定为模板壳文件"
    elif collab_signals:
        triage_group = "yellow_suspected_collab"
        triage_reason = "；".join(collab_signals)
    else:
        triage_group = "green_likely_personal"
        triage_reason = "画风与署名一致，较像完整个人作品"

    return {
        "file": path.name,
        "path": str(path),
        "slide_size_in": [slide_w_in, slide_h_in],
        "total_slides": len(slides_data),
        "metadata": metadata,
        "theme_tokens": theme_tokens,
        "triage_group": triage_group,
        "collab_signals": collab_signals,
        "triage_reason": triage_reason,
        "slides": slides_data,
        "chapter_blocks": blocks,
        "anomaly_candidates": anomalies,
    }


def extract_clean_master_pptx(src_pptx_path: str, out_template_path: str) -> dict[str, Any]:
    """从代表作中剥离全部内容页，保留真实母版/版式/主题色板生成干净模板底座。"""
    src = Path(src_pptx_path).resolve()
    out_p = Path(out_template_path).resolve()
    if not src.is_file():
        return {"status": "ERROR", "message": f"找不到源文件：{src}"}
    if importlib.util.find_spec("pptx") is None:
        return {
            "status": "FALLBACK_REFERENCE",
            "mode": "source_reference",
            "template_path": str(src),
        }
    from pptx import Presentation

    prs = Presentation(str(src))
    sld_id_lst = prs.slides._sldIdLst
    rel_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    for sld_el in list(sld_id_lst):
        rid = sld_el.attrib.get(rel_attr, "")
        sld_id_lst.remove(sld_el)
        if rid:
            try:
                prs.part.drop_rel(rid)
            except Exception:
                pass
    out_p.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_p))
    return {
        "status": "EXTRACTED",
        "mode": "clean_master_pptx",
        "template_path": str(out_p),
    }


SCENARIO_RULES: list[tuple[re.Pattern[str], str, str, str]] = [
    (re.compile(r"(月会|月例会|周会|周报|例会|进展|同步)"), "sc_progress", "例会与进展汇报", "核心进展与指标 → 问题卡点 → 下一步动作"),
    (re.compile(r"(复盘|总结|回顾|迭代|结项)"), "sc_review", "项目复盘与迭代", "现状基线与数据证据 → 根因诊断 → 取舍与改进计划"),
    (re.compile(r"(方案|立项|评审|规划|启动|架构)"), "sc_proposal", "方案评审与立项", "背景痛点 → 核心方案机制 → 角色落地与收益预期"),
    (re.compile(r"(客户|提案|宣讲|推介|合作|销售)"), "sc_pitch", "对外客户提案与宣讲", "客户痛点共鸣 → 解决方案与案例证据 → 合作价值"),
    (re.compile(r"(培训|分享|答辩|演讲|教学)"), "sc_sharing", "培训分享与答辩", "核心观点开篇 → 原理或案例拆解 → 实操落地清单"),
]


def _infer_deck_scenario(deck: dict[str, Any]) -> tuple[str, str, str]:
    text_probe = deck.get("file", "") + " " + " ".join(s.get("title", "") for s in deck.get("slides", [])[:5])
    for pat, sc_id, sc_name, skel in SCENARIO_RULES:
        if pat.search(text_probe):
            return sc_id, sc_name, skel
    return "sc_general", "综合汇报", "核心结论先行 → 分块论据展开 → 下一步行动"


def _classify_slide_framework(s: dict[str, Any]) -> tuple[str, str, str]:
    if s.get("chart_kinds"):
        return "fw_chart_trend", "图表趋势与对比页", "图表＋核心变化结论"
    if s.get("table_count", 0) > 0:
        if s.get("slide_char_count", 0) >= 220:
            return "fw_table_detail", "明细数据与口径对照页", "原生数据表＋补充说明"
        return "fw_table_summary", "结构化对比表页", "精简对照表＋核心结论"
    if s.get("vector_diagram_signals", 0) >= 2 or s.get("has_smartart"):
        return "fw_process_diagram", "流程与机制示意图页", "原生步骤流/结构图＋要点"
    if s.get("pic_count", 0) > 0:
        return "fw_visual_evidence", "截图/图片实证页", "主视觉截图＋任务型短标签"
    if s.get("slide_char_count", 0) <= 90:
        return "fw_core_statement", "极简核心观点页", "醒目大号核心结论＋极少修饰"
    return "fw_multi_column", "分栏要点阐述页", "多栏分块结构＋归纳结论"


def _suggest_framework_adjustment(fw_key: str, has_hygiene: bool, avg_chars: float) -> str:
    if has_hygiene or avg_chars >= 260:
        if fw_key == "fw_table_detail":
            return "原稿明细表和结论文字挤在一页里，建议拆成「1页完整明细表＋1页极简管理结论大数字页」，或适当删减背景字数"
        if fw_key == "fw_multi_column":
            return "原稿单页文字偏密或靠边太紧，建议按安全留白改用2～3栏清晰对齐，每栏控制3～4条短句"
        return "原稿局部存在字号偏小或靠边过紧，建议保留结构关系的同时放宽留白与字号层级"
    if fw_key == "fw_core_statement":
        return "仅在少数页出现的极简表达，建议明确归为过渡结论页或管理总结页，避免正文中段突然留白过多"
    return "样本量较少，建议确认是否按「主题＋结论标题配2～3个清晰分区」作为固定结构沉淀"


def scan_corpus(deck_specs: list[dict[str, Any]]) -> dict[str, Any]:
    """扫描多份样本并在本地汇总所有统计结论、同源改稿折叠与双峰探测。"""
    decks: list[dict[str, Any]] = []
    for spec in deck_specs:
        sel = set(spec["selected_slides"]) if spec.get("selected_slides") else None
        res = inspect_deck(spec["path"], selected_slides=sel)
        res["category"] = spec.get("category", "full_personal")  # full_personal | partial_collab | company_template
        res["module_l1_eligible"] = bool(spec.get("module_l1_eligible", False))
        decks.append(res)

    # Near-duplicate / iterative draft detection across non-template decks
    valid_decks = [d for d in decks if "slides" in d and d["category"] != "company_template"]
    deck_texts = {
        d["file"]: " | ".join(s["full_text"] for s in d["slides"] if s["included"] and s["page_role"] == "content")
        for d in valid_decks
    }
    duplicate_clusters: list[dict[str, Any]] = []
    visited: set[str] = set()
    for i, d1 in enumerate(valid_decks):
        if d1["file"] in visited:
            continue
        group = [d1["file"]]
        for d2 in valid_decks[i + 1 :]:
            if d2["file"] in visited:
                continue
            sim = _jaccard(deck_texts[d1["file"]], deck_texts[d2["file"]])
            if sim >= 0.72:
                group.append(d2["file"])
                visited.add(d2["file"])
        if len(group) > 1:
            visited.add(d1["file"])
            # Recommend the one with highest slide count or latest name indicator
            primary = max(group, key=lambda fn: (
                1 if any(k in fn for k in ("终", "定", "final", "最新")) else 0,
                next((len(d["slides"]) for d in valid_decks if d["file"] == fn), 0),
                fn,
            ))
            duplicate_clusters.append({
                "recommended_primary": primary,
                "members": group,
                "ignored_older_versions": [g for g in group if g != primary],
            })

    ignored_files = {f for c in duplicate_clusters for f in c["ignored_older_versions"]}

    # Aggregate personal slides across active (non-folded) decks
    active_slides: list[tuple[str, str, bool, dict[str, Any]]] = []
    for d in valid_decks:
        if d["file"] in ignored_files:
            continue
        for s in d["slides"]:
            if s["included"] and s["page_role"] == "content":
                active_slides.append((d["file"], d["category"], d["module_l1_eligible"], s))

    total_active = len(active_slides)
    independent_deck_count = len({fn for fn, _cat, _m1, _s in active_slides})

    # Compute structured statistics
    conclusion_title_hits = [(fn, s["slide_no"], s["title"]) for fn, _c, _m, s in active_slides if _is_conclusion_title(s["title"])]
    body_fonts = Counter(f for _fn, _c, _m, s in active_slides for f in s["body_fonts"])
    title_fonts = Counter(f for _fn, _c, _m, s in active_slides for f in s["title_fonts"])
    body_sizes_all = [sz for _fn, _c, _m, s in active_slides for sz in s["body_sizes"]]
    title_sizes_all = [sz for _fn, _c, _m, s in active_slides for sz in s["title_sizes"]]
    table_hdr_cn = Counter(c for _fn, _c, _m, s in active_slides for c in s["table_header_colors_cn"])
    accent_cn = Counter(c for _fn, _c, _m, s in active_slides for c in s["accent_colors_cn"])
    bg_cn = Counter(s["bg_color_cn"] for _fn, _c, _m, s in active_slides)
    bg_kinds = Counter(s.get("bg_kind", "solid") for _fn, _c, _m, s in active_slides)
    canvas_sizes = Counter(
        f"{d['slide_size_in'][0]}x{d['slide_size_in'][1]}"
        for d in valid_decks
        if d["file"] not in ignored_files and "slide_size_in" in d
    )
    chart_kinds = Counter(ck for _fn, _c, _m, s in active_slides for ck in s["chart_kinds"])
    table_slides = [s for _fn, _c, _m, s in active_slides if s["table_count"] > 0]
    diagram_slides = [s for _fn, _c, _m, s in active_slides if s["vector_diagram_signals"] >= 2 or s["has_smartart"]]
    pic_slides = [s for _fn, _c, _m, s in active_slides if s["pic_count"] > 0]

    # Hygiene / rough layout issues (3A)
    hygiene_counter = Counter(issue for _fn, _c, _m, s in active_slides for issue in s["hygiene_issues"])
    hygiene_samples = [
        {"file": fn, "slide_no": s["slide_no"], "issues": s["hygiene_issues"]}
        for fn, _c, _m, s in active_slides
        if s["hygiene_issues"]
    ][:4]

    # Dual-channel speaker notes detection (P1-1)
    notes_heavy_slides = [
        s for _fn, _c, _m, s in active_slides
        if s["notes_char_count"] >= 40 and s["slide_char_count"] <= 140
    ]
    notes_dual_channel = (len(notes_heavy_slides) / max(total_active, 1)) >= 0.35

    # Bimodal style detection (light vs dark background, or high density vs ultra-sparse)
    dark_cnt = sum(1 for _fn, _c, _m, s in active_slides if s["is_dark_bg"])
    light_cnt = total_active - dark_cnt
    dense_cnt = sum(1 for _fn, _c, _m, s in active_slides if s["slide_char_count"] >= 220)
    sparse_cnt = sum(1 for _fn, _c, _m, s in active_slides if s["slide_char_count"] <= 70)
    bimodal_bg = total_active >= 4 and (dark_cnt / total_active) >= 0.25 and (light_cnt / total_active) >= 0.25
    bimodal_density = total_active >= 4 and (dense_cnt / total_active) >= 0.25 and (sparse_cnt / total_active) >= 0.25

    # L1 page-planning stats (strictly from full_personal or module_l1_eligible)
    l1_decks = [d for d in valid_decks if d["file"] not in ignored_files and (d["category"] == "full_personal" or d["module_l1_eligible"])]
    l1_page_counts = [sum(1 for s in d["slides"] if s["included"]) for d in l1_decks]

    # L1 Scenarios + L2 Frameworks per Scenario + L3 Visual Skins
    active_valid_decks = [d for d in valid_decks if d["file"] not in ignored_files]
    sc_buckets: dict[str, dict[str, Any]] = {}
    fw_buckets: dict[tuple[str, str], dict[str, Any]] = {}
    skin_buckets: dict[tuple[bool, str, str], dict[str, Any]] = {}

    for d in active_valid_decks:
        sc_id, sc_name, sc_skel = _infer_deck_scenario(d)
        sc_b = sc_buckets.setdefault(
            sc_id,
            {
                "scenario_id": sc_id,
                "scenario_name": sc_name,
                "deck_files": [],
                "page_counts": [],
                "typical_skeleton": sc_skel,
            },
        )
        sc_b["deck_files"].append(d["file"])
        sc_b["page_counts"].append(sum(1 for s in d["slides"] if s["included"]))

        for s in d["slides"]:
            if not (s["included"] and s["page_role"] == "content"):
                continue
            fw_key, fw_name, fw_carrier = _classify_slide_framework(s)
            fb = fw_buckets.setdefault(
                (sc_id, fw_key),
                {
                    "framework_id": f"{sc_id}_{fw_key}",
                    "scenario_id": sc_id,
                    "scenario_name": sc_name,
                    "framework_name": fw_name,
                    "carrier_type": fw_carrier,
                    "deck_set": set(),
                    "occurrence_slides": 0,
                    "hygiene_hits": 0,
                    "char_counts": [],
                    "sample_refs": [],
                },
            )
            fb["deck_set"].add(d["file"])
            fb["occurrence_slides"] += 1
            if s["hygiene_issues"]:
                fb["hygiene_hits"] += 1
            fb["char_counts"].append(s["slide_char_count"])
            if len(fb["sample_refs"]) < 3:
                fb["sample_refs"].append({"file": d["file"], "slide_no": s["slide_no"], "title": s["title"] or s["snippet"]})

        tk = d.get("theme_tokens") or {}
        skin_key = (bool(tk.get("is_dark_bg")), str(tk.get("bg_color_cn", "纯白")), str(tk.get("primary_color_cn", "蓝")))
        sb = skin_buckets.setdefault(
            skin_key,
            {
                "skin_key": skin_key,
                "deck_files": [],
                "sample_deck_path": d["path"],
                "theme_tokens": tk,
                "bg_color_cn": skin_key[1],
                "primary_color_cn": skin_key[2],
            },
        )
        sb["deck_files"].append(d["file"])

    l1_scenarios = []
    for sc in sc_buckets.values():
        dc = len(sc["deck_files"])
        l1_scenarios.append({
            "scenario_id": sc["scenario_id"],
            "scenario_name": sc["scenario_name"],
            "deck_count": dc,
            "sample_files": sc["deck_files"][:5],
            "median_pages": int(statistics.median(sc["page_counts"])) if sc["page_counts"] else 0,
            "typical_skeleton": sc["typical_skeleton"],
            "confidence": "high" if dc >= 2 else "medium",
            "status": "confirmed" if dc >= 2 else "candidate",
        })

    l2_frameworks = []
    for fb in fw_buckets.values():
        occ = fb["occurrence_slides"]
        indep = len(fb["deck_set"])
        hygiene_rate = fb["hygiene_hits"] / max(occ, 1)
        avg_chars = sum(fb["char_counts"]) / max(len(fb["char_counts"]), 1)
        is_high = (indep >= 2 and occ >= 3 and hygiene_rate < 0.4) or (occ >= 4 and hygiene_rate < 0.25)
        conf = "high" if is_high else "medium"
        fw_key_raw = fb["framework_id"].split("_", 2)[-1]
        adj = _suggest_framework_adjustment(fw_key_raw, hygiene_rate >= 0.3, avg_chars) if (conf == "medium" or hygiene_rate >= 0.3) else ""
        l2_frameworks.append({
            "framework_id": fb["framework_id"],
            "scenario_id": fb["scenario_id"],
            "scenario_name": fb["scenario_name"],
            "framework_name": fb["framework_name"],
            "carrier_type": fb["carrier_type"],
            "occurrence_slides": occ,
            "independent_deck_count": indep,
            "sample_refs": fb["sample_refs"],
            "confidence": conf,
            "suggested_adjustment": adj,
            "status": "confirmed" if conf == "high" else "pending_confirm",
        })

    l3_visual_skins = []
    for idx_s, sb in enumerate(skin_buckets.values(), 1):
        is_dark = sb["skin_key"][0]
        skin_id = "default" if idx_s == 1 else f"skin_{idx_s}"
        label_cn = f"{'深色' if is_dark else '浅色'}{sb['primary_color_cn']}主题皮肤"
        l3_visual_skins.append({
            "skin_id": skin_id,
            "label_cn": label_cn,
            "sample_deck_path": sb["sample_deck_path"],
            "deck_files": sb["deck_files"],
            "theme_tokens": sb["theme_tokens"],
            "bg_color_cn": sb["bg_color_cn"],
            "primary_color_cn": sb["primary_color_cn"],
        })

    triage_overview = Counter(d.get("triage_group", "green_likely_personal") for d in decks if "slides" in d)

    # Summarize per deck compactly (never dump raw full texts)
    compact_decks = []
    for d in decks:
        if "slides" not in d:
            compact_decks.append(d)
            continue
        compact_decks.append({
            "file": d["file"],
            "category": d["category"],
            "total_slides": d["total_slides"],
            "included_content_slides": sum(1 for s in d["slides"] if s["included"] and s["page_role"] == "content"),
            "triage_group": d.get("triage_group", "green_likely_personal"),
            "triage_reason": d.get("triage_reason", ""),
            "chapter_blocks": d["chapter_blocks"],
            "anomaly_candidates": d["anomaly_candidates"],
        })

    return {
        "summary": {
            "total_decks_scanned": len(decks),
            "independent_active_decks": independent_deck_count,
            "total_active_content_slides": total_active,
            "l1_eligible_decks": len(l1_decks),
            "l1_median_pages": int(statistics.median(l1_page_counts)) if l1_page_counts else 0,
            "duplicate_clusters": duplicate_clusters,
            "conclusion_title_ratio": round(len(conclusion_title_hits) / max(total_active, 1), 3),
            "conclusion_title_hits": len(conclusion_title_hits),
            "conclusion_title_samples": [{"file": f, "slide_no": n, "title": t} for f, n, t in conclusion_title_hits[:3]],
            "top_title_font": title_fonts.most_common(1)[0][0] if title_fonts else "",
            "top_body_font": body_fonts.most_common(1)[0][0] if body_fonts else "",
            "median_title_size_pt": round(statistics.median(title_sizes_all), 1) if title_sizes_all else 22.0,
            "median_body_size_pt": round(statistics.median(body_sizes_all), 1) if body_sizes_all else 14.0,
            "dominant_bg_color_cn": bg_cn.most_common(1)[0][0] if bg_cn else "纯白",
            "bg_kind_counts": dict(bg_kinds),
            "canvas_sizes": dict(canvas_sizes),
            "top_accent_colors_cn": [c for c, _ in accent_cn.most_common(3)],
            "top_table_header_color_cn": table_hdr_cn.most_common(1)[0][0] if table_hdr_cn else "",
            "carrier_counts": {
                "tables": len(table_slides),
                "charts": dict(chart_kinds),
                "diagrams": len(diagram_slides),
                "pictures": len(pic_slides),
            },
            "hygiene_issues": dict(hygiene_counter),
            "hygiene_samples": hygiene_samples,
            "recommend_clean_grid": bool(hygiene_counter),
            "notes_dual_channel_recommended": notes_dual_channel,
            "bimodal_style_detected": bimodal_bg or bimodal_density,
            "bimodal_reason": (
                "样本中同时包含深色大字展示页与浅色常规页"
                if bimodal_bg
                else ("样本中同时包含高密详述页与极简少字页" if bimodal_density else "")
            ),
            "l1_scenarios": l1_scenarios,
            "l2_frameworks": l2_frameworks,
            "l3_visual_skins": l3_visual_skins,
            "triage_overview": {
                "green_likely_personal": triage_overview.get("green_likely_personal", 0),
                "yellow_suspected_collab": triage_overview.get("yellow_suspected_collab", 0),
                "gray_template": triage_overview.get("gray_template", 0),
            },
        },
        "decks": compact_decks,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PPT-style 零依赖 PPTX 体检与深层解析器")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_chk = sub.add_parser("check-file", help="检查单个文件格式头与体积")
    p_chk.add_argument("path", help="文件路径")

    p_scan = sub.add_parser("scan", help="扫描一份或多份 PPTX 并输出浓缩统计摘要")
    p_scan.add_argument("--pptx", action="append", default=[], help="PPTX 文件路径（可多次传入）")
    p_scan.add_argument("--manifest", help="包含待扫描清单的 JSON 文件")

    args = parser.parse_args(argv)
    if args.cmd == "check-file":
        res = check_file(args.path)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0

    if args.cmd == "scan":
        specs: list[dict[str, Any]] = []
        if args.manifest:
            data = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
            specs.extend(data.get("decks", data) if isinstance(data, dict) else data)
        for p in args.pptx:
            specs.append({"path": p, "category": "full_personal"})
        res = scan_corpus(specs)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
