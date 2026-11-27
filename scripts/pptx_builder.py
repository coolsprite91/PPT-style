#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pptx_builder.py — 分级可编辑 PPTX 构建器（V0.2.0：三句话降级 ＋ 真实母版皮肤预览 ＋ 多轮 doc_id 继承 ＋ 原地改既有页）。

运行分级底线:
  - 缺少 python-pptx 时绝不静默联网下载，也不强行拼写原生 XML，
    而是输出标准三句话降级确认 JSON（NEEDS_OPTIONAL_INSTALL）引导确认或降级文案稿。
  - 已安装 python-pptx 时支持：
    1) preview: 纯 L3 视觉皮肤预览页（优先克隆真实提取母版＋精确色板，支持默认固定示范内容）
    2) trial:   2~3 页三层组合试炼小套组（支持原生表格、矢量步骤流程、柱/折/饼图、等比插图、讲者备注）
    3) full:    第 3 层完整成稿生成（跨平台安全字体回退 ＋ 字宽安全余量 15% 超载停机检测）
    4) patch:   第 5 档局部增改（继承原 doc_id、AI 续改前自动补抓上轮人改、支持插入与 replace_slides 原地改页）
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
import os
import re
import struct
import sys
import uuid
from pathlib import Path
from typing import Any

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import env_probe
import snapshot_diff


SAFE_MARGIN_LEFT_IN = 0.60
SAFE_MARGIN_RIGHT_IN = 0.60
SAFE_MARGIN_TOP_IN = 0.45
SAFE_MARGIN_BOTTOM_IN = 0.45
REL_ID_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _probe_image_size_px(img_path: Path) -> tuple[int, int] | None:
    """获取图片像素宽高（优先 Pillow，无 Pillow 时纯标准库读取 PNG/JPEG 头）。"""
    if importlib.util.find_spec("PIL") is not None:
        try:
            from PIL import Image

            with Image.open(img_path) as im:
                return im.size
        except Exception:
            pass
    try:
        data = img_path.read_bytes()
        if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
            return struct.unpack(">II", data[16:24])
        if data[:2] == b"\xff\xd8":
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2):
                    h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                    return (w, h)
                seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
                i += 2 + seg_len
    except Exception:
        return None
    return None


def estimate_text_overflow(
    text: str,
    box_w_in: float,
    box_h_in: float,
    font_pt: float,
    safety_margin: float = 1.15,
) -> dict[str, Any]:
    """按本机字宽经验＋15%安全余量估算文本框是否会溢出。"""
    if not text.strip():
        return {"overflow": False, "estimated_lines": 0, "max_lines": 99}

    usable_w_pt = max((box_w_in - 0.16) * 72.0, font_pt * 2.0)
    usable_h_pt = max((box_h_in - 0.12) * 72.0, font_pt * 1.2)
    line_h_pt = font_pt * 1.35

    total_lines = 0
    for para in text.splitlines() or [text]:
        if not para.strip():
            total_lines += 1
            continue
        width_pt = 0.0
        for ch in para:
            width_pt += font_pt * (1.02 if ord(ch) > 127 else 0.56)
        width_pt *= safety_margin
        total_lines += max(1, math.ceil(width_pt / usable_w_pt))

    max_lines = max(1, int(usable_h_pt // line_h_pt))
    return {
        "overflow": total_lines > max_lines,
        "estimated_lines": total_lines,
        "max_lines": max_lines,
    }


def _hex_to_rgb_tuple(hex_str: str, fallback: tuple[int, int, int] = (31, 78, 121)) -> tuple[int, int, int]:
    cleaned = (hex_str or "").strip().lstrip("#")
    if len(cleaned) != 6:
        return fallback
    try:
        return (int(cleaned[0:2], 16), int(cleaned[2:4], 16), int(cleaned[4:6], 16))
    except ValueError:
        return fallback


def _save_with_lock_fallback(prs: Any, out_path: Path) -> tuple[Path, bool]:
    """写盘防占锁：若 Windows 下原文件被软件占用，自动递存 -v2.pptx。"""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(str(out_path))
        return out_path, False
    except PermissionError:
        stem = re.sub(r"-v\d+$", "", out_path.stem)
        for v in range(2, 50):
            cand = out_path.with_name(f"{stem}-v{v}{out_path.suffix}")
            try:
                prs.save(str(cand))
                return cand, True
            except PermissionError:
                continue
        raise


def _check_blueprint_overflow(blueprint: dict[str, Any]) -> list[dict[str, Any]]:
    default_body_pt = float(blueprint.get("theme", {}).get("body_size_pt", 14.0))
    all_specs = list(blueprint.get("slides", [])) + list(blueprint.get("replace_slides", []))
    overflows: list[dict[str, Any]] = []
    for idx, slide_spec in enumerate(all_specs, 1):
        blocks = slide_spec.get("blocks", [])
        n_blocks = max(1, len(blocks))
        box_w = (13.333 - SAFE_MARGIN_LEFT_IN - SAFE_MARGIN_RIGHT_IN - 0.25 * (n_blocks - 1)) / n_blocks
        box_h = 4.9
        for b_idx, blk in enumerate(blocks, 1):
            btype = blk.get("type", "bullets")
            if btype not in ("bullets", "text", "card"):
                continue
            txt = blk.get("text") or "\n".join(blk.get("items", []))
            pt = float(blk.get("font_size_pt", default_body_pt))
            est = estimate_text_overflow(txt, box_w, box_h, pt)
            if est["overflow"]:
                overflows.append({
                    "slide_no": slide_spec.get("slide_no") or slide_spec.get("target_slide_no") or idx,
                    "slide_title": slide_spec.get("title", f"第{idx}页"),
                    "block_index": b_idx,
                    "estimated_lines": est["estimated_lines"],
                    "max_lines": est["max_lines"],
                })
    return overflows


def _select_blank_layout(prs: Any) -> Any:
    """智能挑选真正的空白版式（规避第 7 个槽位被重排为图片/标题版式的企业模板陷阱）。"""
    layouts = list(prs.slide_layouts)
    if not layouts:
        raise ValueError("当前演示文稿缺少可用版式")

    def _is_blank_name(name: str) -> bool:
        low = (name or "").strip().lower()
        has_blank = ("blank" in low) or ("空白" in low)
        has_content_word = any(k in low for k in ("标题", "图片", "内容", "title", "pic", "content"))
        return has_blank and not has_content_word

    def _exact_blank(name: str) -> bool:
        low = (name or "").strip().lower()
        return low in ("blank", "空白", "空白页")

    zero_ph = [lay for lay in layouts if len(lay.placeholders) == 0]
    if zero_ph:
        return min(
            zero_ph,
            key=lambda lay: (0 if _is_blank_name(lay.name) else 1, len(lay.shapes)),
        )

    named_blank = [lay for lay in layouts if _is_blank_name(lay.name)]
    if named_blank:
        return min(
            named_blank,
            key=lambda lay: (0 if _exact_blank(lay.name) else 1, len(lay.placeholders), len(lay.shapes)),
        )

    return min(layouts, key=lambda lay: (len(lay.placeholders), len(lay.shapes)))


def _clear_empty_placeholders(slide: Any) -> None:
    """清除从母版版式继承下来的空占位框，防止叠出现影轮廓。"""
    for ph in list(slide.placeholders):
        has_txt = getattr(ph, "has_text_frame", False) and bool(ph.text_frame.text.strip())
        if not has_txt:
            sp_el = ph._element
            if sp_el.getparent() is not None:
                sp_el.getparent().remove(sp_el)


def _default_preview_slides(theme: dict[str, Any]) -> list[dict[str, Any]]:
    skin_label = theme.get("skin_label") or "提取的个人视觉皮肤"
    return [{
        "title": f"纯视觉皮肤预览：{skin_label}",
        "subtitle": "本页仅用于确认母版、配色、字体层级与干净留白，不涉及具体业务内容",
        "keep_existing_bg": bool(theme.get("preserve_template_bg", False)),
        "blocks": [
            {
                "type": "bullets",
                "heading": "字体与留白示范",
                "items": [
                    "标题与副标题保持清晰层次",
                    "正文留白充足，不贴边不重叠",
                    "重点内容用主配色与加粗突出",
                ],
            },
            {
                "type": "table",
                "headers": ["示范列", "样式基准"],
                "rows": [
                    ["表头配色", "继承提取主色"],
                    ["表格行高", "自动留出呼吸感"],
                ],
            },
            {
                "type": "diagram",
                "steps": ["风格确认", "场景规划", "生成成稿"],
            },
        ],
    }]


def build_pptx(
    blueprint: dict[str, Any],
    out_path_str: str,
    mode: str = "full",
    base_pptx_str: str | None = None,
    snapshot_root: str | None = None,
    force_shrink: bool = False,
) -> dict[str, Any]:
    """根据结构化蓝图构建可编辑 PPTX；缺依赖时返回标准三句话降级确认信号。"""
    if importlib.util.find_spec("pptx") is None:
        install_cmd = f'"{sys.executable}" -m pip install python-pptx Pillow'
        user_cmd = f'"{sys.executable}" -m pip install --user python-pptx Pillow'
        notice = env_probe.build_degradation_notice(
            "TIER1_LIGHT",
            "这台电脑还没有安装用来自动生成可编辑 PPT 文件的小组件。",
            "读旧稿、提炼风格、定页面大纲和写逐页内容完全正常；只是这次没法直接导出 .pptx 文件，会改成交给你一份带清晰排版说明的逐页文案稿。",
            f"标准安装：{install_cmd} ；免管理员：{user_cmd}",
        )
        return {
            "status": "NEEDS_OPTIONAL_INSTALL",
            "fallback_mode": "TEXT_DRAFT_ONLY",
            "install_command": install_cmd,
            "degradation_notice": notice,
            "human_hint": notice["formatted_message"],
        }

    eff_blueprint = dict(blueprint)
    theme = dict(eff_blueprint.get("theme", {}))
    if mode == "preview" and not eff_blueprint.get("slides"):
        if eff_blueprint.get("template_pptx") and "preserve_template_bg" not in theme:
            theme["preserve_template_bg"] = True
        eff_blueprint["slides"] = _default_preview_slides(theme)

    # Check capacity overflow before building unless user explicitly approved shrinking
    overflows = _check_blueprint_overflow(eff_blueprint)
    if overflows and not force_shrink:
        return {
            "status": "OVERFLOW_DETECTED",
            "overflow_count": len(overflows),
            "overflow_pages": overflows,
            "human_hint": (
                f"按你的常用字号和干净留白排版时，发现有 {len(overflows)} 处内容超出单页安全容量，"
                f"共 {len(overflows)} 个确认项：每处请选择「拆分成两页」或「特批缩小字号」。"
            ),
        }

    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt

    primary_rgb = RGBColor(*_hex_to_rgb_tuple(theme.get("primary_hex", "1F4E79")))
    accent_rgb = RGBColor(*_hex_to_rgb_tuple(theme.get("accent_hex", "D96B27"), (217, 107, 39)))
    bg_rgb = RGBColor(*_hex_to_rgb_tuple(theme.get("bg_hex", "FFFFFF"), (255, 255, 255)))
    text_rgb = RGBColor(*_hex_to_rgb_tuple(theme.get("text_hex", "222222"), (34, 34, 34)))
    raw_font = theme.get("font_name", "Microsoft YaHei")
    font_info = env_probe.resolve_safe_font(raw_font)
    font_name = font_info["resolved_font"]
    title_pt = float(theme.get("title_size_pt", 22.0))
    subtitle_pt = float(theme.get("subtitle_size_pt", 14.0))
    body_pt = float(theme.get("body_size_pt", 13.0))

    inherited_id: str | None = None
    captured_prior: dict[str, Any] | None = None
    if base_pptx_str and Path(base_pptx_str).is_file():
        inherited_id = snapshot_diff.extract_doc_id(Path(base_pptx_str))
        if inherited_id:
            captured_prior = snapshot_diff.capture_human_revision_if_changed(
                base_pptx_str,
                snapshot_root=snapshot_root,
                doc_id=inherited_id,
            )

    if mode == "patch" and base_pptx_str and Path(base_pptx_str).is_file():
        prs = Presentation(base_pptx_str)
    elif eff_blueprint.get("template_pptx") and Path(eff_blueprint["template_pptx"]).is_file():
        prs = Presentation(eff_blueprint["template_pptx"])
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    slide_w_in = float(prs.slide_width) / 914400.0 if prs.slide_width else 13.333
    slide_h_in = float(prs.slide_height) / 914400.0 if prs.slide_height else 7.5
    scale_w = slide_w_in / 13.333
    scale_h = slide_h_in / 7.5
    if abs(scale_w - 1.0) < 0.02 and abs(scale_h - 1.0) < 0.02:
        scale_w, scale_h = 1.0, 1.0
    font_scale = min(scale_w, scale_h)

    blank_layout = _select_blank_layout(prs)
    has_template_base = bool(
        (mode == "patch" and base_pptx_str and Path(base_pptx_str).is_file())
        or (eff_blueprint.get("template_pptx") and Path(eff_blueprint["template_pptx"]).is_file())
    )

    def _populate_slide(slide: Any, spec: dict[str, Any]) -> None:
        _clear_empty_placeholders(slide)

        keep_bg = spec.get("keep_existing_bg")
        if keep_bg is None:
            keep_bg = bool(has_template_base and theme.get("preserve_template_bg", False))
        if not keep_bg:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = bg_rgb

        margin_l = SAFE_MARGIN_LEFT_IN * scale_w
        margin_r = SAFE_MARGIN_RIGHT_IN * scale_w
        margin_t = SAFE_MARGIN_TOP_IN * scale_h
        content_w = slide_w_in - margin_l - margin_r

        title_box = slide.shapes.add_textbox(
            Inches(margin_l),
            Inches(margin_t),
            Inches(content_w),
            Inches(0.95 * scale_h),
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        p0.text = spec.get("title", "")
        p0.font.name = font_name
        p0.font.size = Pt(round(title_pt * font_scale, 1))
        p0.font.bold = True
        p0.font.color.rgb = primary_rgb

        if spec.get("subtitle"):
            p1 = tf.add_paragraph()
            p1.text = spec["subtitle"]
            p1.font.name = font_name
            p1.font.size = Pt(round(subtitle_pt * font_scale, 1))
            p1.font.color.rgb = text_rgb

        line_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(margin_l),
            Inches(1.48 * scale_h),
            Inches(content_w),
            Inches(0.025 * scale_h),
        )
        line_shape.fill.solid()
        line_shape.fill.fore_color.rgb = primary_rgb
        line_shape.line.fill.background()

        blocks = spec.get("blocks", [])
        n_blocks = max(1, len(blocks))
        gap_in = 0.25 * scale_w
        col_w = (content_w - gap_in * (n_blocks - 1)) / n_blocks
        top_y = 1.72 * scale_h
        area_h = 5.15 * scale_h

        for idx, blk in enumerate(blocks):
            left_x = margin_l + idx * (col_w + gap_in)
            btype = blk.get("type", "bullets")
            raw_eff_pt = max(10.5, body_pt - 1.5) if force_shrink else float(blk.get("font_size_pt", body_pt))
            eff_pt = round(raw_eff_pt * font_scale, 1)

            if btype in ("bullets", "text", "card"):
                if btype == "card":
                    bg_card = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        Inches(left_x),
                        Inches(top_y),
                        Inches(col_w),
                        Inches(area_h),
                    )
                    bg_card.fill.solid()
                    bg_card.fill.fore_color.rgb = RGBColor(246, 248, 251)
                    bg_card.line.color.rgb = RGBColor(210, 218, 228)
                    bg_card.line.width = Pt(1)
                tb = slide.shapes.add_textbox(
                    Inches(left_x + 0.1 * scale_w),
                    Inches(top_y + 0.1 * scale_h),
                    Inches(col_w - 0.2 * scale_w),
                    Inches(area_h - 0.2 * scale_h),
                )
                btf = tb.text_frame
                btf.word_wrap = True
                if blk.get("heading"):
                    hp = btf.paragraphs[0]
                    hp.text = blk["heading"]
                    hp.font.name = font_name
                    hp.font.size = Pt(round(eff_pt + 2 * font_scale, 1))
                    hp.font.bold = True
                    hp.font.color.rgb = primary_rgb
                    first = False
                else:
                    first = True

                lines = blk.get("items") or (blk.get("text", "").splitlines())
                for line in lines:
                    p = btf.paragraphs[0] if first else btf.add_paragraph()
                    first = False
                    p.text = line
                    p.font.name = font_name
                    p.font.size = Pt(eff_pt)
                    p.font.color.rgb = text_rgb
                    p.space_after = Pt(round(6 * font_scale, 1))

            elif btype == "table":
                headers = blk.get("headers", [])
                rows_data = blk.get("rows", [])
                all_rows = ([headers] if headers else []) + rows_data
                if all_rows:
                    n_r = len(all_rows)
                    n_c = max(len(r) for r in all_rows)
                    tbl_shape = slide.shapes.add_table(
                        n_r,
                        n_c,
                        Inches(left_x),
                        Inches(top_y),
                        Inches(col_w),
                        Inches(min(area_h, (0.48 * n_r + 0.3) * scale_h)),
                    )
                    tbl = tbl_shape.table
                    for r_i, row in enumerate(all_rows):
                        for c_i in range(n_c):
                            cell = tbl.cell(r_i, c_i)
                            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                            val = str(row[c_i]) if c_i < len(row) else ""
                            cell.text = val
                            p = cell.text_frame.paragraphs[0]
                            p.font.name = font_name
                            p.font.size = Pt(round(max(10.5 * font_scale, eff_pt - 1 * font_scale), 1))
                            if r_i == 0 and headers:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = primary_rgb
                                p.font.bold = True
                                p.font.color.rgb = RGBColor(255, 255, 255)
                            else:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = RGBColor(255, 255, 255) if r_i % 2 else RGBColor(246, 248, 252)
                                p.font.color.rgb = text_rgb

            elif btype == "diagram":
                steps = blk.get("steps", [])
                if steps:
                    step_n = len(steps)
                    arrow_w = 0.28 * scale_w
                    box_w = (col_w - arrow_w * (step_n - 1)) / step_n
                    box_h = 1.45 * scale_h
                    box_y = top_y + 0.6 * scale_h
                    for s_i, step_label in enumerate(steps):
                        sx = left_x + s_i * (box_w + arrow_w)
                        shp = slide.shapes.add_shape(
                            MSO_SHAPE.ROUNDED_RECTANGLE,
                            Inches(sx),
                            Inches(box_y),
                            Inches(box_w),
                            Inches(box_h),
                        )
                        shp.fill.solid()
                        shp.fill.fore_color.rgb = RGBColor(242, 246, 252)
                        shp.line.color.rgb = primary_rgb
                        stf = shp.text_frame
                        stf.word_wrap = True
                        sp = stf.paragraphs[0]
                        sp.text = step_label
                        sp.alignment = PP_ALIGN.CENTER
                        sp.font.name = font_name
                        sp.font.size = Pt(eff_pt)
                        sp.font.color.rgb = text_rgb
                        if s_i < step_n - 1:
                            arr = slide.shapes.add_shape(
                                MSO_SHAPE.RIGHT_ARROW,
                                Inches(sx + box_w + 0.04 * scale_w),
                                Inches(box_y + box_h / 2 - 0.12 * scale_h),
                                Inches(max(0.05 * scale_w, arrow_w - 0.08 * scale_w)),
                                Inches(0.24 * scale_h),
                            )
                            arr.fill.solid()
                            arr.fill.fore_color.rgb = accent_rgb
                            arr.line.fill.background()

            elif btype == "chart":
                kind = blk.get("chart_kind", "column")
                cats = blk.get("categories", [])
                series_list = blk.get("series", [])
                chart_data = CategoryChartData()
                chart_data.categories = cats
                for ser in series_list:
                    chart_data.add_series(ser.get("name", "系列"), ser.get("values", []))
                chart_type_map = {
                    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
                    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                    "line": XL_CHART_TYPE.LINE_MARKERS,
                    "pie": XL_CHART_TYPE.PIE,
                }
                xl_type = chart_type_map.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
                slide.shapes.add_chart(
                    xl_type,
                    Inches(left_x),
                    Inches(top_y),
                    Inches(col_w),
                    Inches(area_h - 0.2),
                    chart_data,
                )

            elif btype == "picture" and blk.get("image_path") and Path(blk["image_path"]).is_file():
                img_p = Path(blk["image_path"])
                px = _probe_image_size_px(img_p)
                if px and px[1] > 0:
                    ratio = px[0] / px[1]
                    target_w = col_w
                    target_h = target_w / ratio
                    if target_h > area_h:
                        target_h = area_h
                        target_w = target_h * ratio
                    slide.shapes.add_picture(str(img_p), Inches(left_x), Inches(top_y), width=Inches(target_w), height=Inches(target_h))
                else:
                    slide.shapes.add_picture(str(img_p), Inches(left_x), Inches(top_y), width=Inches(col_w))

        if spec.get("notes"):
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.text = spec["notes"]

    # Build or patch slides
    if mode == "patch" and base_pptx_str and Path(base_pptx_str).is_file():
        sld_id_lst = prs.slides._sldIdLst
        # 1. 原地替换指定既有页（其余页 100% 不动）
        for rep_spec in eff_blueprint.get("replace_slides", []):
            target_no = int(rep_spec.get("slide_no") or rep_spec.get("target_slide_no") or 0)
            if target_no < 1 or target_no > len(prs.slides):
                continue
            new_slide = prs.slides.add_slide(blank_layout)
            _populate_slide(new_slide, rep_spec)
            new_el = sld_id_lst[-1]
            old_el = sld_id_lst[target_no - 1]
            old_rid = old_el.attrib.get(REL_ID_ATTR, "")
            sld_id_lst.remove(new_el)
            old_el.addprevious(new_el)
            sld_id_lst.remove(old_el)
            if old_rid:
                try:
                    prs.part.drop_rel(old_rid)
                except Exception:
                    pass

        # 2. 在指定页位后插入新页
        insert_after = eff_blueprint.get("insert_after_slide", len(prs.slides))
        new_slides_specs = eff_blueprint.get("slides", [])
        initial_count = len(prs.slides)
        added_count = 0
        for s_spec in new_slides_specs:
            slide = prs.slides.add_slide(blank_layout)
            _populate_slide(slide, s_spec)
            added_count += 1
        if 0 <= insert_after < initial_count and added_count > 0:
            all_ids = list(sld_id_lst)
            moved = all_ids[initial_count:]
            kept = all_ids[:initial_count]
            new_order = kept[:insert_after] + moved + kept[insert_after:]
            for el in all_ids:
                sld_id_lst.remove(el)
            for el in new_order:
                sld_id_lst.append(el)
    else:
        for s_spec in eff_blueprint.get("slides", []):
            slide = prs.slides.add_slide(blank_layout)
            _populate_slide(slide, s_spec)

    # 继承原有隐形 doc_id 或创建新 doc_id
    doc_id = eff_blueprint.get("doc_id") or inherited_id or uuid.uuid4().hex[:12]
    existing_comments = prs.core_properties.comments or ""
    if "[ppt-style-id:" not in existing_comments:
        prs.core_properties.comments = (existing_comments + f" [ppt-style-id:{doc_id}]").strip()

    final_path, lock_incremented = _save_with_lock_fallback(prs, Path(out_path_str))
    snap_role = "ai_revise" if (mode == "patch" or inherited_id) else "ai_build"
    snap_res = snapshot_diff.save_snapshot(
        str(final_path),
        snapshot_root=snapshot_root,
        doc_id=doc_id,
        source_label=f"built_{mode}",
        role=snap_role,
    )

    return {
        "status": "BUILT",
        "mode": mode,
        "output_path": str(final_path),
        "lock_incremented": lock_incremented,
        "doc_id": doc_id,
        "rev_id": snap_res.get("rev_id"),
        "captured_prior_human_rev": captured_prior,
        "font_resolution": font_info,
        "slide_count": len(prs.slides),
        "snapshot_saved": snap_res.get("status") == "SAVED",
        "human_hint": (
            f"原文件正被办公软件占用，已自动为你另存为最新版《{final_path.name}》，可直接打开查看。"
            if lock_incremented
            else f"已生成可编辑文件《{final_path.name}》（共 {len(prs.slides)} 页），可直接打开修改并按 Ctrl+S 保存。"
        ),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PPT-style 分级可编辑 PPTX 构建器")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_b = sub.add_parser("build", help="从结构化 JSON 蓝图生成或局部增改 PPTX")
    p_b.add_argument("--blueprint", required=True, help="JSON 蓝图文件路径")
    p_b.add_argument("--out", required=True, help="输出 PPTX 路径")
    p_b.add_argument("--mode", choices=("preview", "trial", "full", "patch"), default="full")
    p_b.add_argument("--base-pptx", default=None, help="局部增改或续改模式下的基础 PPTX")
    p_b.add_argument("--snapshot-root", default=None, help="快照根目录（可选，便于隔离测试）")
    p_b.add_argument("--force-shrink", action="store_true", help="经用户确认后允许缩小字号")

    args = parser.parse_args(argv)
    if args.cmd == "build":
        bp = json.loads(Path(args.blueprint).read_text(encoding="utf-8"))
        res = build_pptx(
            bp,
            args.out,
            mode=args.mode,
            base_pptx_str=args.base_pptx,
            snapshot_root=args.snapshot_root,
            force_shrink=args.force_shrink,
        )
        print(json.dumps(res, ensure_ascii=False, indent=2))
        if res["status"] == "BUILT":
            return 0
        if res["status"] in ("NEEDS_OPTIONAL_INSTALL", "OVERFLOW_DETECTED"):
            return 3
        return 1
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
