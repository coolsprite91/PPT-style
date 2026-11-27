#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PPT-style 全链路隔离自动化测试套件（全部运行于临时目录，绝不触碰用户真实全局目录）。"""

from __future__ import annotations

import json
import re
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = SKILL_DIR / "scripts"
sys.path.insert(0, str(SKILL_DIR))
sys.path.insert(0, str(SCRIPTS_DIR))

import env_probe
import install as installer
import output_guard
import pptx_builder
import pptx_inspector
import snapshot_diff
import workspace_manager


def _make_sample_pptx(
    path: Path,
    slides_spec: list[dict],
    dark_bg: bool = False,
    font_size_pt: float = 14.0,
    x_in: float = 0.8,
    notes_list: list[str] | None = None,
    add_chart: bool = False,
    add_table: bool = False,
    overlap_boxes: bool = False,
) -> Path:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for idx, sp in enumerate(slides_spec):
        slide = prs.slides.add_slide(blank)
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(22, 34, 56) if dark_bg else RGBColor(255, 255, 255)

        tb_t = slide.shapes.add_textbox(Inches(x_in), Inches(0.45), Inches(11.5), Inches(0.8))
        p_t = tb_t.text_frame.paragraphs[0]
        p_t.text = sp["title"]
        p_t.font.name = "Microsoft YaHei"
        p_t.font.size = Pt(22)
        p_t.font.color.rgb = RGBColor(255, 255, 255) if dark_bg else RGBColor(31, 78, 121)

        tb_b = slide.shapes.add_textbox(Inches(x_in), Inches(1.6), Inches(11.0), Inches(3.5))
        p_b = tb_b.text_frame.paragraphs[0]
        p_b.text = sp["body"]
        p_b.font.name = "Microsoft YaHei"
        p_b.font.size = Pt(font_size_pt)
        p_b.font.color.rgb = RGBColor(235, 240, 245) if dark_bg else RGBColor(40, 40, 40)

        if overlap_boxes:
            tb_ov = slide.shapes.add_textbox(Inches(x_in + 0.2), Inches(1.8), Inches(9.0), Inches(3.0))
            p_ov = tb_ov.text_frame.paragraphs[0]
            p_ov.text = "重叠干扰框内容测试"
            p_ov.font.size = Pt(font_size_pt)

        if add_table and idx == 0:
            tbl = slide.shapes.add_table(2, 2, Inches(0.8), Inches(4.5), Inches(5.0), Inches(1.4)).table
            tbl.cell(0, 0).text = "区域"
            tbl.cell(0, 1).text = "完成度"
            tbl.cell(0, 0).fill.solid()
            tbl.cell(0, 0).fill.fore_color.rgb = RGBColor(31, 78, 121)
            tbl.cell(1, 0).text = "华南区"
            tbl.cell(1, 1).text = "达成 98%"

        if add_chart and idx == 0:
            cd = CategoryChartData()
            cd.categories = ["Q1", "Q2", "Q3"]
            cd.add_series("进度", (60, 78, 92))
            slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.5), Inches(4.2), Inches(5.0), Inches(2.5), cd)

        if notes_list and idx < len(notes_list) and notes_list[idx]:
            slide.notes_slide.notes_text_frame.text = notes_list[idx]

    prs.save(str(path))
    return path


class PptStyleFullSuiteTest(unittest.TestCase):
    def test_01_env_probe(self) -> None:
        info = env_probe.probe_environment()
        self.assertEqual(info["tier_read_plan_write_text"], "READY")
        self.assertIn(info["tier_build_pptx"], ("READY", "NEEDS_OPTIONAL_INSTALL"))
        self.assertTrue(env_probe._is_windows_apps_stub(r"C:\Users\u\AppData\Local\Microsoft\WindowsApps\python.exe"))
        self.assertFalse(env_probe._is_windows_apps_stub(sys.executable))

    def test_02_inspector_headers_theme_charts_notes_hygiene_and_folding(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            # 1. Magic header checks
            fake_ppt = tmp / "old_report.ppt"
            fake_ppt.write_bytes(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\x00" * 64)
            chk_ppt = pptx_inspector.check_file(str(fake_ppt))
            self.assertEqual(chk_ppt["status"], "NEEDS_USER_CONVERT")

            fake_dps = tmp / "wps_deck.dps"
            fake_dps.write_bytes(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\x00" * 64)
            self.assertEqual(pptx_inspector.check_file(str(fake_dps))["status"], "NEEDS_USER_CONVERT")

            fake_pdf = tmp / "brand_guide.pdf"
            fake_pdf.write_bytes(b"%PDF-1.7 test")
            self.assertEqual(pptx_inspector.check_file(str(fake_pdf))["status"], "VISUAL_REF_ONLY")

            # 2. Color scheme + lumMod + RGB CN translation
            hex_mod = pptx_inspector._apply_lum_mod_off("1F4E79", lum_mod=75000, lum_off=25000, tint=None, shade=None)
            self.assertEqual(len(hex_mod), 6)
            self.assertEqual(pptx_inspector.rgb_to_cn_name("1F4E79"), "蓝")
            self.assertEqual(pptx_inspector.rgb_to_cn_name("FFFFFF"), "纯白")

            # 3. Rough layout + chart + notes + iterative duplicate folding + bimodal detection
            d1 = _make_sample_pptx(
                tmp / "项目汇报-v1.pptx",
                [
                    {"title": "服务覆盖提升：客户交付周期缩短至2天", "body": "核心结论摘要。"},
                    {"title": "运营效率改善：团队周均处理提升30%", "body": "第二页摘要。"},
                ],
                dark_bg=True,
                font_size_pt=8.5,   # small font -> triggers hygiene
                x_in=0.05,          # edge overflow -> triggers hygiene
                notes_list=["各位领导好，这一页我们重点汇报交付周期缩短的关键动作与一线客户真实反馈。" * 2] * 2,
                add_chart=True,
                add_table=True,
                overlap_boxes=True,
            )
            d2 = _make_sample_pptx(
                tmp / "项目汇报-最终版.pptx",
                [
                    {"title": "服务覆盖提升：客户交付周期缩短至2天", "body": "核心结论摘要微调版。"},
                    {"title": "运营效率改善：团队周均处理提升30%", "body": "第二页摘要。"},
                ],
                dark_bg=True,
                font_size_pt=14.0,
                notes_list=["各位领导好，这一页我们重点汇报交付周期缩短的关键动作与一线客户真实反馈。" * 2] * 2,
            )
            d2_v1 = _make_sample_pptx(
                tmp / "项目汇报-草稿v1.pptx",
                [
                    {"title": "服务覆盖提升：客户交付周期缩短至2天", "body": "核心结论摘要初稿。"},
                    {"title": "运营效率改善：团队周均处理提升30%", "body": "第二页摘要。"},
                ],
                dark_bg=True,
                font_size_pt=14.0,
                notes_list=["各位领导好，这一页我们重点汇报交付周期缩短的关键动作与一线客户真实反馈。" * 2] * 2,
            )
            d3 = _make_sample_pptx(
                tmp / "日常工作周报.pptx",
                [
                    {"title": "本周推进事项总结：三项重点落地", "body": "详述文字内容。" * 28},
                    {"title": "下周风险清单：两项资源待协调", "body": "详述文字内容。" * 28},
                ],
                dark_bg=False,
                font_size_pt=13.0,
            )

            res = pptx_inspector.scan_corpus([
                {"path": str(d2_v1), "category": "full_personal"},
                {"path": str(d2), "category": "full_personal"},
                {"path": str(d3), "category": "full_personal"},
            ])
            summary = res["summary"]
            self.assertEqual(len(summary["duplicate_clusters"]), 1)
            self.assertEqual(summary["duplicate_clusters"][0]["recommended_primary"], "项目汇报-最终版.pptx")
            self.assertTrue(summary["bimodal_style_detected"])
            # Scan d1 alone for hygiene + chart + notes check
            res_d1 = pptx_inspector.scan_corpus([{"path": str(d1), "category": "full_personal"}])["summary"]
            self.assertTrue(res_d1["recommend_clean_grid"])
            self.assertIn("small_body_font", res_d1["hygiene_issues"])
            self.assertIn("edge_overflow", res_d1["hygiene_issues"])
            self.assertIn("text_box_overlap", res_d1["hygiene_issues"])
            self.assertTrue(res_d1["notes_dual_channel_recommended"])
            self.assertEqual(res_d1["carrier_counts"]["charts"].get("柱状图"), 1)

    def test_03_workspace_intake_dedup_partial_and_media_strip(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            home = tmp / "ppt_home"
            ws = tmp / "my_project"
            ws.mkdir()

            init_out = workspace_manager.init_workshop(str(ws), False, home)
            wdir = Path(init_out["workshop_dir"])
            self.assertTrue((wdir / "00-操作中心.md").is_file())

            # Create 8-slide partial collab deck and embed dummy >50MB-triggered mp4 check
            slides = [{"title": f"模块汇报第{i}页：进展完成", "body": f"第{i}页正文内容"} for i in range(1, 9)]
            collab = _make_sample_pptx(tmp / "部门月会合写.pptx", slides)

            # Inject fake video in zip to test >50MB media stripping path
            big_pptx = tmp / "含视频产品宣讲.pptx"
            with zipfile.ZipFile(collab, "r") as zin, zipfile.ZipFile(big_pptx, "w", compression=zipfile.ZIP_STORED) as zout:
                for item in zin.infolist():
                    zout.writestr(item, zin.read(item.filename))
                zout.writestr("ppt/media/demo_video.mp4", b"\x00" * (51 * 1024 * 1024))

            # Intake collab with 5 contiguous pages (2-6) -> module_l1_eligible == True
            in1 = workspace_manager.intake_file(wdir, str(collab), "partial_collab", "2-6", False, home)
            self.assertEqual(in1["status"], "ARCHIVED")
            self.assertTrue(in1["module_l1_eligible"])
            self.assertEqual(in1["total_slides"], 8)

            # Duplicate intake skipped
            dup_in = workspace_manager.intake_file(wdir, str(collab), "partial_collab", "2-6", False, home)
            self.assertEqual(dup_in["status"], "DUPLICATE_SKIPPED")

            # Intake heavy media file -> stripped archived copy is small!
            in_big = workspace_manager.intake_file(wdir, str(big_pptx), "full_personal", None, False, home)
            self.assertEqual(in_big["status"], "ARCHIVED")
            self.assertTrue(in_big["media_stripped"])
            self.assertLess(Path(in_big["path"]).stat().st_size, 1024 * 1024)

    def test_04_double_insurance_and_4A_checkpoint(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            # Case A: 4 decks with 100% conclusion titles -> NOT auto-confirmed (<5 independent decks!)
            home4 = tmp / "home4"
            ws4 = tmp / "ws4"
            wdir4 = Path(workspace_manager.init_workshop(str(ws4), False, home4)["workshop_dir"])
            for i in range(4):
                p = _make_sample_pptx(
                    tmp / f"deck4_{i}.pptx",
                    [{"title": f"结论汇报{i}：交付周期缩短20%", "body": f"独特内容分支{i * 97}"}],
                )
                workspace_manager.intake_file(wdir4, str(p), "full_personal", None, False, home4)
            workspace_manager.scan_workshop(wdir4, home4)
            plan4 = workspace_manager.plan_questions(wdir4, home4)
            self.assertEqual(plan4["auto_confirmed_count"], 0)

            # Case B: 5 decks with 100% conclusion titles -> auto-confirmed + <=5 per batch + 4A checkpoint + undo
            home5 = tmp / "home5"
            ws5 = tmp / "ws5"
            wdir5 = Path(workspace_manager.init_workshop(str(ws5), False, home5)["workshop_dir"])
            for i in range(5):
                p = _make_sample_pptx(
                    tmp / f"deck5_{i}.pptx",
                    [{"title": f"独立主题{i}：核心成本下降{10 + i}%", "body": f"完全不同的正文叙述段落{i}，包括渠道、客服、运维等方向{i * 33}"}],
                )
                workspace_manager.intake_file(wdir5, str(p), "full_personal", None, False, home5)
            workspace_manager.scan_workshop(wdir5, home5)
            plan5 = workspace_manager.plan_questions(wdir5, home5)
            self.assertEqual(plan5["auto_confirmed_count"], 1)

            b1 = workspace_manager.next_batch(wdir5)
            self.assertEqual(b1["status"], "BATCH_READY")
            self.assertLessEqual(b1["question_count"], 5)

            ans = {q["id"]: "A" for q in b1["questions"]}
            rec1 = workspace_manager.record_answer(wdir5, ans, home5)
            # Verify undo works
            undone = workspace_manager.undo_last(wdir5, home5)
            self.assertEqual(undone["status"], "UNDONE")
            b1_again = workspace_manager.next_batch(wdir5)
            self.assertEqual(b1_again["question_count"], b1["question_count"])

            # Answer through all core questions
            while True:
                nb = workspace_manager.next_batch(wdir5)
                if nb["status"] != "BATCH_READY" or nb["phase"] != "core":
                    break
                rec = workspace_manager.record_answer(wdir5, {q["id"]: "A" for q in nb["questions"]}, home5)
                if rec["status"] == "CHECKPOINT_CORE_DONE":
                    self.assertEqual(rec["remaining_optional_count"], 4)
                    self.assertEqual(rec["estimated_extra_rounds"], 1)
                    break

            # Choose ship_now (4A recommendation)
            ch = workspace_manager.choose_path(wdir5, "ship_now", home5)
            self.assertEqual(ch["deferred_optional_count"], 4)
            self.assertEqual(workspace_manager.next_batch(wdir5)["status"], "ALL_QUESTIONS_DONE")

    def test_05_profile_masks_swap_shell_rollback_sanitize_and_export(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            home = tmp / "home"
            workspace_manager.init_workshop(str(tmp / "ws"), False, home)

            class Args:
                profile = "default"
                mask = "external_pitch"
                label = "对外客户提案"
                layer = "L2"
                rule_text = "2026年8月某区域营收增长32.5%时先放核心数字"
                rule_no = 1
                out = str(tmp / "share_pkg")
                subaction = "split-mask"

            self.assertEqual(workspace_manager.handle_profile(Args, home)["status"], "MASK_CREATED")
            Args.subaction = "promote-rule"
            prom = workspace_manager.handle_profile(Args, home)
            self.assertEqual(prom["status"], "PROMOTED")
            self.assertNotIn("2026年8月", prom["sanitized_rule"])
            self.assertNotIn("32.5%", prom["sanitized_rule"])

            # Delete then rollback restores it
            Args.subaction = "delete-rule"
            self.assertEqual(workspace_manager.handle_profile(Args, home)["status"], "DELETED")
            Args.subaction = "rollback"
            self.assertEqual(workspace_manager.handle_profile(Args, home)["status"], "ROLLED_BACK")

            # Swap visual shell preserves L1/L2, resets L3
            Args.subaction = "swap-visual-shell"
            swapped = workspace_manager.handle_profile(Args, home)
            self.assertEqual(swapped["status"], "SHELL_SWAPPED")
            self.assertEqual(swapped["preserved_layers"], ["L1_planning", "L2_content"])

            # Export standalone sanitized share package
            Args.subaction = "export-share"
            exp = workspace_manager.handle_profile(Args, home)
            self.assertTrue((Path(exp["share_dir"]) / "bundled-profile" / "core_l2_content.json").is_file())
            self.assertTrue((Path(exp["share_dir"]) / "SKILL.md").is_file())

    def test_06_builder_modes_overflow_lock_fallback_and_patch(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            snap_root = tmp / "snaps"

            # 1. Missing python-pptx graceful fallback
            with mock.patch("importlib.util.find_spec", return_value=None):
                no_dep = pptx_builder.build_pptx({"slides": []}, str(tmp / "none.pptx"))
                self.assertEqual(no_dep["status"], "NEEDS_OPTIONAL_INSTALL")
                self.assertEqual(no_dep["fallback_mode"], "TEXT_DRAFT_ONLY")

            # 2. Overflow detection
            huge_bp = {
                "theme": {"body_size_pt": 18.0},
                "slides": [{
                    "title": "超载测试页",
                    "blocks": [{"type": "bullets", "text": "这是一段非常非常长的业务说明文字，用来故意撑爆文本框安全容量限制。\n" * 18}],
                }],
            }
            ov = pptx_builder.build_pptx(huge_bp, str(tmp / "ov.pptx"), snapshot_root=str(snap_root))
            self.assertEqual(ov["status"], "OVERFLOW_DETECTED")

            # 3. Trial 2-slide kit with table + chart + diagram + speaker notes
            trial_bp = {
                "theme": {"primary_hex": "1F4E79", "font_name": "Microsoft YaHei", "body_size_pt": 13.0},
                "slides": [
                    {
                        "title": "流程与数据对比：自动化率提升至85%",
                        "subtitle": "流程步骤清晰，关键指标可追溯",
                        "blocks": [
                            {"type": "diagram", "steps": ["步骤1：自动采集", "步骤2：智能归类", "步骤3：生成成稿"]},
                            {"type": "table", "headers": ["指标", "变更前", "变更后"], "rows": [["耗时", "120分钟", "15分钟"]]},
                        ],
                        "notes": "演讲备注：重点说明耗时从120分钟降到15分钟。",
                    },
                    {
                        "title": "季度趋势：持续稳步爬坡",
                        "blocks": [
                            {"type": "chart", "chart_kind": "column", "categories": ["Q1", "Q2", "Q3"], "series": [{"name": "覆盖率", "values": [40, 68, 85]}]}
                        ],
                    },
                ],
            }
            built = pptx_builder.build_pptx(trial_bp, str(tmp / "trial.pptx"), mode="trial", snapshot_root=str(snap_root))
            self.assertEqual(built["status"], "BUILT")
            self.assertEqual(built["slide_count"], 2)
            self.assertTrue(built["snapshot_saved"])

            # 4. Patch mode: insert 1 slide after slide 1, keep original slide 1 and original slide 2 untouched
            patch_bp = {
                "insert_after_slide": 1,
                "slides": [{"title": "局部新增第2页：补充专题说明", "blocks": [{"type": "bullets", "items": ["仅新增这一页，原第1页和原第2页（现第3页）纹丝不动"]}]}],
            }
            patched = pptx_builder.build_pptx(
                patch_bp,
                str(tmp / "patched.pptx"),
                mode="patch",
                base_pptx_str=built["output_path"],
                snapshot_root=str(snap_root),
            )
            self.assertEqual(patched["status"], "BUILT")
            self.assertEqual(patched["slide_count"], 3)
            sem_patched = snapshot_diff.extract_semantic_slides(Path(patched["output_path"]))
            self.assertIn("流程与数据对比", sem_patched[0].title)
            self.assertIn("局部新增第2页", sem_patched[1].title)
            self.assertIn("季度趋势", sem_patched[2].title)

            # 5. Simulated Windows PermissionError -> auto save -v2.pptx
            from pptx.presentation import Presentation as PresClass

            orig_save = PresClass.save
            call_count = {"n": 0}

            def flaky_save(self_prs, fname):
                call_count["n"] += 1
                if call_count["n"] == 1:
                    raise PermissionError("Simulated lock by PowerPoint")
                return orig_save(self_prs, fname)

            with mock.patch.object(PresClass, "save", flaky_save):
                locked_out = pptx_builder.build_pptx(
                    {"slides": [{"title": "占锁测试", "blocks": [{"type": "bullets", "items": ["一行内容"]}]}]},
                    str(tmp / "locked.pptx"),
                    snapshot_root=str(snap_root),
                )
            self.assertTrue(locked_out["lock_incremented"])
            self.assertTrue(locked_out["output_path"].endswith("locked-v2.pptx"))

    def test_07_snapshot_diff_cross_move_wps_noise_and_text_lock(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            snap_root = tmp / "snaps"
            bp = {
                "slides": [{
                    "title": "年度交付总结：达成率提升至96.5%",
                    "blocks": [
                        {"type": "bullets", "items": ["累计服务120家客户", "平均用时缩短30分钟"]},
                        {"type": "table", "headers": ["区域", "数值"], "rows": [["华东", "50万元"]]},
                    ],
                }]
            }
            out = pptx_builder.build_pptx(bp, str(tmp / "orig.pptx"), snapshot_root=str(snap_root))
            orig_p = Path(out["output_path"])

            # Move/rename to another path + simulate WPS <a:r> split & 0.5pt jitter -> 0 false positives!
            moved_p = tmp / "桌面" / "改名后的汇报.pptx"
            moved_p.parent.mkdir()
            with zipfile.ZipFile(orig_p, "r") as zin, zipfile.ZipFile(moved_p, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "ppt/slides/slide1.xml":
                        txt = data.decode("utf-8")
                        # Split text across multiple <a:r> runs (WPS behavior)
                        txt = txt.replace("累计服务120家客户", "累计服务</a:t></a:r><a:r><a:t>120家客户")
                        data = txt.encode("utf-8")
                    zout.writestr(item, data)

            no_diff = snapshot_diff.diff_against_snapshot(str(moved_p), snapshot_root=str(snap_root))
            self.assertEqual(no_diff["status"], "NO_CHANGES_DETECTED")

            # Verify dual-channel text lock (exact split passes, altered number fails, code/formula exempt passes)
            src_texts = [
                "年度交付总结：达成率提升至96.5%",
                "累计服务120家客户",
                "平均用时缩短30分钟",
                "区域",
                "数值",
                "华东",
                "50万元",
            ]
            lock_ok = snapshot_diff.verify_text_lock(src_texts, str(moved_p))
            self.assertTrue(lock_ok["ok"])

            lock_bad = snapshot_diff.verify_text_lock(src_texts + ["额外被删掉的一句"], str(moved_p))
            self.assertFalse(lock_bad["ok"])

            lock_num_bad = snapshot_diff.verify_text_lock(
                [t.replace("96.5%", "98.0%") for t in src_texts],
                str(moved_p),
            )
            self.assertFalse(lock_num_bad["numeric_channel_ok"])

            # Formula exemption works
            lock_exempt = snapshot_diff.verify_text_lock(
                src_texts + ["E = mc^2 (formula_v1)"],
                str(moved_p),
                exempt_blocks=["E = mc^2 (formula_v1)"],
            )
            self.assertTrue(lock_exempt["ok"])

            # Real edit: change title, increase font size, delete table -> detected & structural deletion alert
            from pptx import Presentation
            from pptx.util import Pt

            prs = Presentation(str(moved_p))
            s0 = prs.slides[0]
            s0.shapes[0].text_frame.paragraphs[0].text = "年度交付复盘：核心达成率突破96.5%"
            s0.shapes[2].text_frame.paragraphs[0].font.size = Pt(16)
            # Delete table graphicFrame
            for shp in list(s0.shapes):
                if getattr(shp, "has_table", False):
                    shp._element.getparent().remove(shp._element)
            prs.save(str(moved_p))

            real_diff = snapshot_diff.diff_against_snapshot(str(moved_p), snapshot_root=str(snap_root))
            self.assertEqual(real_diff["status"], "CHANGES_DETECTED")
            self.assertTrue(real_diff["structural_deletions"])
            self.assertEqual(real_diff["confirmation_checklist"][0]["item_no"], 0)

    def test_08_output_guard_and_installer_isolation(self) -> None:
        bad = output_guard.check_reply_text("请修改 #1F4E79 配色，或者运行 pptx_inspector.py 查看：\n1. 问题一？\nA. 是\nB. 否")
        self.assertFalse(bad["ok"])
        codes = {v["code"] for v in bad["violations"]}
        self.assertIn("HEX_COLOR_LEAK", codes)
        self.assertIn("INTERNAL_FILE_LEAK", codes)
        self.assertIn("MISSING_QUESTION_COUNT", codes)

        good = output_guard.check_reply_text("【进度 3/6】已识别你的深蓝主配色，共 2 个问题请确认：\n1. 选哪套？\nA. 第一套\nB. 第二套\n2. 要留白吗？\nA. 要\nB. 不要")
        self.assertTrue(good["ok"])

        # Verify lowercase machine-readable frontmatter name
        skill_md_head = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")[:200]
        self.assertIn("name: ppt-style", skill_md_head)

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            skills_dir = tmp / "global_skills"
            home_dir = tmp / "ppt_home"

            # Pre-seed legacy uppercase directory to ensure it gets cleaned into backups/
            (skills_dir / "PPT-style").mkdir(parents=True)
            (skills_dir / "PPT-style" / "old.txt").write_text("legacy", encoding="utf-8")

            r1 = installer.install_skill(skills_dir, home_dir)
            self.assertEqual(r1["status"], "INSTALLED")
            self.assertEqual(r1["skill_name"], "ppt-style")
            self.assertEqual(r1["version"], "V0.2.0")
            self.assertTrue((skills_dir / "ppt-style" / "SKILL.md").is_file())
            self.assertFalse((skills_dir / "ppt-style" / "old.txt").exists())
            self.assertEqual([p.name for p in skills_dir.iterdir()], ["ppt-style"])
            self.assertIsNotNone(r1["backup_path"])

            # Reinstall must back up to home_dir/backups/, NEVER leave backup inside skills_dir!
            r2 = installer.install_skill(skills_dir, home_dir)
            self.assertIsNotNone(r2["backup_path"])
            self.assertTrue(Path(r2["backup_path"]).is_dir())
            self.assertEqual([p.name for p in skills_dir.iterdir()], ["ppt-style"])

            r3 = installer.uninstall_skill(skills_dir, home_dir)
            self.assertEqual(r3["status"], "UNINSTALLED")
            self.assertFalse((skills_dir / "ppt-style").exists())

            # Multi-client auto-detection test in isolated fake user_home
            fake_user = tmp / "fake_user"
            (fake_user / ".trae-cn").mkdir(parents=True)
            (fake_user / ".claude").mkdir(parents=True)
            (fake_user / ".agents").mkdir(parents=True)
            r_multi = installer.install_skill(None, home_dir, user_home=fake_user)
            self.assertFalse(r_multi["fallback_used"])
            self.assertEqual(len(r_multi["installed_paths"]), 3)
            self.assertTrue((fake_user / ".trae-cn" / "skills" / "ppt-style" / "SKILL.md").is_file())
            self.assertTrue((fake_user / ".claude" / "skills" / "ppt-style" / "SKILL.md").is_file())
            self.assertTrue((fake_user / ".agents" / "skills" / "ppt-style" / "SKILL.md").is_file())

            st_multi = installer.check_status(None, user_home=fake_user)
            self.assertTrue(st_multi["installed"])
            self.assertEqual(len(st_multi["installations"]), 3)
            self.assertEqual({x["version"] for x in st_multi["installations"]}, {"V0.2.0"})

            u_multi = installer.uninstall_skill(None, home_dir, user_home=fake_user)
            self.assertEqual(len(u_multi["removed_from"]), 3)
            self.assertFalse(installer.check_status(None, user_home=fake_user)["installed"])

            # Fallback when zero client directories exist
            empty_user = tmp / "empty_user"
            empty_user.mkdir()
            r_fb = installer.install_skill(None, home_dir, user_home=empty_user)
            self.assertTrue(r_fb["fallback_used"])
            self.assertTrue((empty_user / ".agents" / "skills" / "ppt-style" / "SKILL.md").is_file())

    def test_09_v012_real_world_regressions(self) -> None:
        from pptx import Presentation
        from pptx.dml.color import RGBColor
        from pptx.util import Inches, Pt

        # 1. Sanitize-rule Chinese + decimal regex fix
        san = output_guard.sanitize_rule_text(
            "2018年6月某美妆活动带来32.5%曝光增长，正文统一用微软雅黑14pt",
            auto_abstract=True,
        )
        self.assertEqual(
            san["sanitized_text"],
            "[时间]某美妆活动带来[具体数值与单位]曝光增长，正文统一用微软雅黑14pt",
        )

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            home = tmp / "home"
            ws = tmp / "ws"
            ws.mkdir()

            # 2. Rejected .ppt returns exit code 2 and REJECTED status
            wdir = Path(workspace_manager.init_workshop(str(ws), False, home)["workshop_dir"])
            bad_ppt = tmp / "legacy.ppt"
            bad_ppt.write_bytes(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\x00" * 64)
            import contextlib, io
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                rc_ppt = workspace_manager.main([
                    "--home-dir", str(home),
                    "intake", "--workspace", str(ws), "--file", str(bad_ppt), "--category", "full_personal",
                ])
            self.assertEqual(rc_ppt, 2)
            ppt_resp = json.loads(buf.getvalue())
            self.assertEqual(ppt_resp["status"], "REJECTED")
            self.assertEqual(ppt_resp["reject_reason"], "NEEDS_USER_CONVERT")

            # 3. Image-only >50MB pptx -> has_heavy_media=True, has_av_to_strip=False
            base_deck = _make_sample_pptx(tmp / "base_clean.pptx", [{"title": "标题测试页面", "body": "正文内容不少于十个字"}])
            big_img_pptx = tmp / "big_images.pptx"
            with zipfile.ZipFile(base_deck, "r") as zin, zipfile.ZipFile(big_img_pptx, "w", compression=zipfile.ZIP_STORED) as zout:
                for item in zin.infolist():
                    zout.writestr(item, zin.read(item.filename))
                zout.writestr("ppt/media/poster.png", b"\x89PNG\r\n\x1a\n" + b"\x00" * (51 * 1024 * 1024))
            chk_big_img = pptx_inspector.check_file(str(big_img_pptx))
            self.assertTrue(chk_big_img["has_heavy_media"])
            self.assertFalse(chk_big_img["has_av_to_strip"])

            # 4. Inherited master solid bg + grpSp safe transform + 2x canvas font normalization + decorative 'C' filter
            prs2x = Presentation()
            prs2x.slide_width = Inches(26.666)
            prs2x.slide_height = Inches(15.0)
            s_2x = prs2x.slides.add_slide(prs2x.slide_layouts[6])
            tb_c = s_2x.shapes.add_textbox(Inches(1.2), Inches(0.5), Inches(1.5), Inches(1.0))
            tb_c.text_frame.paragraphs[0].text = "C"
            tb_real_title = s_2x.shapes.add_textbox(Inches(3.2), Inches(0.8), Inches(18.0), Inches(1.5))
            p_rt = tb_real_title.text_frame.paragraphs[0]
            p_rt.text = "双倍画布真实章节标题内容"
            p_rt.font.size = Pt(56.0)
            tb_2x_body = s_2x.shapes.add_textbox(Inches(2.0), Inches(4.0), Inches(20.0), Inches(4.0))
            p_2xb = tb_2x_body.text_frame.paragraphs[0]
            p_2xb.text = "双倍画布上的标准正文内容，折算回标准画布应为十六磅。"
            p_2xb.font.size = Pt(32.0)  # 32 / 2 = 16.0pt -> NOT bloated and NOT <10pt
            p2x_path = tmp / "double_canvas.pptx"
            prs2x.save(str(p2x_path))

            # Inject master yellow background + grouped inner shape whose local x=0 sits inside safe parent group x=2.0in
            patched_xml_deck = tmp / "double_master_grp.pptx"
            with zipfile.ZipFile(p2x_path, "r") as zin, zipfile.ZipFile(patched_xml_deck, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "ppt/slideMasters/slideMaster1.xml":
                        txt = data.decode("utf-8")
                        txt = re.sub(
                            r"<p:bg>.*?</p:bg>",
                            '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FCE199"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>',
                            txt,
                            flags=re.S,
                        )
                        if "<p:bg>" not in txt:
                            txt = txt.replace(
                                "<p:cSld>",
                                '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FCE199"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>',
                                1,
                            )
                        data = txt.encode("utf-8")
                    elif item.filename == "ppt/slides/slide1.xml":
                        txt = data.decode("utf-8")
                        grp_xml = (
                            '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="901" name="Grp1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                            '<p:grpSpPr><a:xfrm><a:off x="2743200" y="5486400"/><a:ext cx="5486400" cy="1828800"/>'
                            '<a:chOff x="0" y="0"/><a:chExt cx="5486400" cy="1828800"/></a:xfrm></p:grpSpPr>'
                            '<p:sp><p:nvSpPr><p:cNvPr id="902" name="Inner1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
                            '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm>'
                            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
                            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2800"/><a:t>组合内部安全框文字</a:t></a:r></a:p></p:txBody></p:sp>'
                            '</p:grpSp></p:spTree>'
                        )
                        txt = txt.replace("</p:spTree>", grp_xml)
                        data = txt.encode("utf-8")
                    zout.writestr(item, data)

            dec = pptx_inspector.inspect_deck(str(patched_xml_deck))
            s0 = dec["slides"][0]
            self.assertEqual(s0["bg_color_cn"], "浅黄")
            self.assertEqual(s0["bg_kind"], "solid")
            self.assertEqual(s0["title"], "双倍画布真实章节标题内容")
            self.assertIn(16.0, s0["body_sizes"])
            self.assertNotIn("edge_overflow", s0["hygiene_issues"])
            self.assertNotIn("text_box_overlap", s0["hygiene_issues"])

            # 5. Smart blank layout selection + zero ghost placeholders when slot 6 has placeholders
            tpl = Presentation()
            tpl.slide_width = Inches(13.333)
            tpl.slide_height = Inches(7.5)
            # Rename layout[4] vs layout[6] so slot 6 has placeholders and slot 4 is blank
            tpl.slide_layouts[6].name = "图片与标题"
            # Add a placeholder into layout[6] if empty, by inspecting: default layout[6] has 0 ph, so test layout swap via default layout[1] cloned/tested or pick non-zero layout
            # Better: make layout[6] actually hold a placeholder shape and keep layout[4] empty
            from pptx.oxml import parse_xml
            lay6_xml = tpl.slide_layouts[6]._element
            sp_tree_l6 = lay6_xml.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
            ph_xml = parse_xml(
                '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                '<p:nvSpPr><p:cNvPr id="77" name="GhostPh"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>'
                '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            )
            sp_tree_l6.append(ph_xml)
            tpl.slide_layouts[5].name = "空白"
            tpl.slides.add_slide(tpl.slide_layouts[0])
            tpl_p = tmp / "agency_tpl.pptx"
            tpl.save(str(tpl_p))

            patch_res = pptx_builder.build_pptx(
                {"insert_after_slide": 1, "slides": [{"title": "局部加页洁净测试", "blocks": [{"type": "bullets", "items": ["要点1"]}]}]},
                str(tmp / "agency_patched.pptx"),
                mode="patch",
                base_pptx_str=str(tpl_p),
                snapshot_root=str(tmp / "snaps"),
            )
            self.assertEqual(patch_res["status"], "BUILT")
            patched_prs = Presentation(patch_res["output_path"])
            new_s = patched_prs.slides[1]
            self.assertEqual(len([sh for sh in new_s.shapes if sh.is_placeholder]), 0)
            self.assertEqual(new_s.slide_layout.name, "空白")

            # 6. Diagram auto-shape deletion + table font size diff
            trial_res = pptx_builder.build_pptx(
                {
                    "slides": [
                        {
                            "title": "流程与表格手改测试",
                            "blocks": [
                                {"type": "diagram", "steps": ["步骤1", "步骤2", "步骤3"]},
                                {"type": "table", "headers": ["项", "值"], "rows": [["曝光", "100万"]]},
                            ],
                        }
                    ]
                },
                str(tmp / "diag_tbl.pptx"),
                mode="trial",
                snapshot_root=str(tmp / "snaps"),
            )
            ed_prs = Presentation(trial_res["output_path"])
            s_ed = ed_prs.slides[0]
            for shp in list(s_ed.shapes):
                if shp.shape_type == 1 and shp.name.startswith(("Rounded Rectangle", "Right Arrow")):
                    shp._element.getparent().remove(shp._element)
                elif shp.has_table:
                    for r in shp.table.cell(1, 1).text_frame.paragraphs[0].runs:
                        r.font.size = Pt(17.5)
            ed_p = tmp / "diag_tbl_edited.pptx"
            ed_prs.save(str(ed_p))
            diff_diag = snapshot_diff.diff_against_snapshot(str(ed_p), snapshot_root=str(tmp / "snaps"))
            self.assertEqual(diff_diag["status"], "CHANGES_DETECTED")
            self.assertTrue(diff_diag["structural_deletions"])
            self.assertIn("字号调整", {c["kind"] for c in diff_diag["changes"]})

            # 7. scan-dir multi-author + reference intake + auto-commit L1/L2/L3 rules
            bulk_dir = tmp / "bulk_samples"
            bulk_dir.mkdir()
            d_a = _make_sample_pptx(bulk_dir / "项目A-方案-李莉.pptx", [{"title": "A主题", "body": "内容A"}])
            d_b = _make_sample_pptx(bulk_dir / "项目B-方案-向重洁.pptx", [{"title": "B主题", "body": "内容B"}])
            sd_out = workspace_manager.scan_directory(str(bulk_dir))
            self.assertEqual(sd_out["status"], "DIR_TRIAGED")
            author_names = {a["name"] for a in sd_out["detected_authors"]}
            self.assertIn("李莉", author_names)
            self.assertIn("向重洁", author_names)

            in_ref = workspace_manager.intake_file(wdir, str(d_a), "full_personal", None, False, home, copy_mode="reference")
            self.assertEqual(in_ref["status"], "ARCHIVED")
            self.assertEqual(in_ref["storage_mode"], "reference")
            self.assertFalse((wdir / "01-原始代表作归档" / "完整个人代表作" / d_a.name).exists())
            workspace_manager.scan_workshop(wdir, home)
            workspace_manager.plan_questions(wdir, home)
            while True:
                nb = workspace_manager.next_batch(wdir)
                if nb["status"] != "BATCH_READY" or nb["phase"] != "core":
                    break
                workspace_manager.record_answer(wdir, {q["id"]: "A" for q in nb["questions"]}, home)
            workspace_manager.choose_path(wdir, "ship_now", home)
            l1_json = json.loads((home / "profiles" / "default" / "core_l1_planning.json").read_text(encoding="utf-8"))
            l2_json = json.loads((home / "profiles" / "default" / "core_l2_content.json").read_text(encoding="utf-8"))
            l3_json = json.loads((home / "profiles" / "default" / "masks" / "default.json").read_text(encoding="utf-8"))
            self.assertGreater(len(l1_json["rules"]), 0)
            self.assertGreater(len(l2_json["rules"]), 0)
            self.assertGreater(len(l3_json["rules"]), 0)

    def test_10_v020_compat_degradation_formats_and_gemini_symlink_install(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)

            # 1. Cross-platform font resolution & 3-sentence degradation notice
            font_map = env_probe.resolve_safe_font("Microsoft YaHei", installed_fonts=["苹方", "思源黑体", "Arial"])
            self.assertTrue(font_map["mapped"])
            self.assertEqual(font_map["resolved_font"], "苹方")

            with mock.patch("importlib.util.find_spec", return_value=None):
                no_pptx = pptx_builder.build_pptx({"slides": []}, str(tmp / "fallback.pptx"))
                self.assertEqual(no_pptx["status"], "NEEDS_OPTIONAL_INSTALL")
                dn = no_pptx["degradation_notice"]
                self.assertIn("原因：", dn["formatted_message"])
                self.assertIn("影响：", dn["formatted_message"])
                self.assertIn("请问你能接受先按这个降级效果继续吗", dn["formatted_message"])

            # 2. Keynote .key bundle directory, .odp, and suspected DRM .pptx
            key_dir = tmp / "产品发布会.key"
            key_dir.mkdir()
            chk_key = pptx_inspector.check_file(str(key_dir))
            self.assertEqual(chk_key["status"], "NEEDS_USER_CONVERT")
            self.assertEqual(chk_key["format_kind"], "keynote")

            odp_file = tmp / "汇报.odp"
            odp_file.write_bytes(b"PK\x03\x04dummy-odp")
            chk_odp = pptx_inspector.check_file(str(odp_file))
            self.assertEqual(chk_odp["status"], "NEEDS_USER_CONVERT")
            self.assertEqual(chk_odp["format_kind"], "odp")

            drm_pptx = tmp / "透明加密汇报.pptx"
            drm_pptx.write_bytes(b"\x99\x88\x77\x66ENCRYPTED_HEADER" + b"\xAB" * 2048)
            chk_drm = pptx_inspector.check_file(str(drm_pptx))
            self.assertEqual(chk_drm["status"], "SUSPECTED_DRM_ENCRYPTED")
            self.assertTrue(chk_drm["suspected_drm"])

            # 3. install.py detects ~/.gemini/config/skills/ and creates symlink when mirroring .agents
            fake_user = tmp / "fake_user"
            agents_skills = fake_user / ".agents" / "skills"
            gemini_cfg_skills = fake_user / ".gemini" / "config" / "skills"
            (agents_skills / "other-skill").mkdir(parents=True)
            gemini_cfg_skills.mkdir(parents=True)
            (gemini_cfg_skills / "other-skill").symlink_to("../../../.agents/skills/other-skill")

            res_inst = installer.install_skill(None, tmp / "home", user_home=fake_user)
            self.assertEqual(res_inst["status"], "INSTALLED")
            self.assertTrue((agents_skills / "ppt-style" / "SKILL.md").is_file())
            self.assertTrue((gemini_cfg_skills / "ppt-style").is_symlink())
            self.assertTrue((gemini_cfg_skills / "ppt-style" / "SKILL.md").is_file())

    def test_11_v020_multi_signal_triage_git_workshop_and_upload_copy(self) -> None:
        from pptx import Presentation
        from pptx.dml.color import RGBColor
        from pptx.util import Inches, Pt

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            home = tmp / "home"

            # 1. Green personal deck vs Yellow same-master multi-speaker + style shift vs Gray template
            green_p = _make_sample_pptx(
                tmp / "项目复盘个人版.pptx",
                [
                    {"title": "复盘结论：交付周期缩短30%", "body": "正常单人写作风格页。"},
                    {"title": "核心动作：自动化校准上线", "body": "保持统一字号与配色的第二页。"},
                ],
                add_table=True,
            )

            # Same master, 6 slides, 2 speakers + font/bg/density shift between first half and second half
            yellow_p = tmp / "跨组汇报汇编.pptx"
            prs_y = Presentation()
            prs_y.slide_width = Inches(13.333)
            prs_y.slide_height = Inches(7.5)
            blank = prs_y.slide_layouts[6]
            for i in range(6):
                sl = prs_y.slides.add_slide(blank)
                is_second_half = i >= 3
                fill = sl.background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(20, 30, 50) if is_second_half else RGBColor(255, 255, 255)
                tb = sl.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.0), Inches(4.5))
                p = tb.text_frame.paragraphs[0]
                speaker = "张敏" if i < 3 else "王磊"
                body = ("简短要点。" if not is_second_half else "大段详述分析内容补充说明。" * 14)
                p.text = f"汇报人：{speaker}\n{body}"
                p.font.name = "SimHei" if is_second_half else "Microsoft YaHei"
                p.font.size = Pt(14)
                p.font.color.rgb = RGBColor(255, 255, 255) if is_second_half else RGBColor(30, 30, 30)
            prs_y.save(str(yellow_p))

            gray_p = _make_sample_pptx(
                tmp / "空白模板壳.pptx",
                [
                    {"title": "点击添加标题", "body": "请输入正文占位文本"},
                    {"title": "TBD", "body": "点击添加内容"},
                ],
            )

            tr = workspace_manager.triage_paths([str(green_p), str(yellow_p), str(gray_p)])
            self.assertEqual(tr["status"], "TRIAGED")
            self.assertEqual(len(tr["green_likely_personal"]), 1)
            self.assertEqual(len(tr["yellow_suspected_collab"]), 1)
            self.assertEqual(len(tr["gray_template"]), 1)
            self.assertIn("一键去掉", tr["human_triage_prompt"])

            # 2. User-specified workshop inside a Git repo -> creates .gitignore + returns warning
            git_repo = tmp / "biz_repo"
            (git_repo / ".git").mkdir(parents=True)
            ws_res = workspace_manager.init_workshop(str(git_repo / "我的专属工坊"), False, home)
            self.assertEqual(ws_res["status"], "INITIALIZED")
            self.assertTrue(ws_res["inside_git_repo"])
            wdir = Path(ws_res["workshop_dir"])
            self.assertTrue((wdir / ".gitignore").is_file())
            self.assertIn("版本库", ws_res["git_warning_hint"])

            # 3. from_upload=True forces copy even when copy_mode="reference"
            in_up = workspace_manager.intake_file(
                wdir, str(green_p), "full_personal", None, False, home,
                copy_mode="reference", from_upload=True,
            )
            self.assertEqual(in_up["status"], "ARCHIVED")
            self.assertEqual(in_up["storage_mode"], "copy")
            self.assertTrue(in_up["forced_copy_from_upload"])
            self.assertTrue(Path(in_up["path"]).is_file())

    def test_12_v020_three_layer_assets_combo_preview_and_usage_guide(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            home = tmp / "home"
            wdir = Path(workspace_manager.init_workshop(str(tmp / "ws"), False, home)["workshop_dir"])

            # Feed two distinct scenarios and two distinct skins (light blue monthly + dark pitch)
            d_monthly_1 = _make_sample_pptx(
                tmp / "10月数字化月会进展.pptx",
                [
                    {"title": "10月核心进展：线上覆盖率达85%", "body": "重点工作稳步推进。"},
                    {"title": "大区明细数据：全部达标", "body": "详见表格对照。"},
                ],
                dark_bg=False,
                add_table=True,
            )
            d_monthly_2 = _make_sample_pptx(
                tmp / "11月数字化月会进展.pptx",
                [
                    {"title": "11月核心进展：用时减少35%", "body": "自动化报表全量切换。"},
                    {"title": "问题卡点与下一步：两项待协调", "body": "下周推进落地。"},
                ],
                dark_bg=False,
                add_table=True,
            )
            d_pitch = _make_sample_pptx(
                tmp / "对外客户提案宣讲.pptx",
                [
                    {"title": "客户痛点与方案价值", "body": "核心主张阐述。"},
                    {"title": "标杆案例效果展示", "body": "助力客户效率倍增。"},
                ],
                dark_bg=True,
                font_size_pt=9.0,  # trigger hygiene -> medium confidence adjustment suggestion
            )

            for p in (d_monthly_1, d_monthly_2, d_pitch):
                workspace_manager.intake_file(wdir, str(p), "full_personal", None, False, home)

            scan_out = workspace_manager.scan_workshop(wdir, home)
            summary = scan_out["summary"]
            self.assertGreaterEqual(len(summary["l1_scenarios"]), 2)
            self.assertGreaterEqual(len(summary["l2_frameworks"]), 2)
            self.assertEqual(len(summary["l3_visual_skins"]), 2)

            # Pure L3 preview mode uses extracted real master & default demo content
            skin_0 = summary["l3_visual_skins"][0]
            tpl_clean = tmp / "clean_master.pptx"
            ext_m = pptx_inspector.extract_clean_master_pptx(skin_0["sample_deck_path"], str(tpl_clean))
            self.assertEqual(ext_m["status"], "EXTRACTED")
            prev_out = pptx_builder.build_pptx(
                {"template_pptx": str(tpl_clean), "theme": skin_0["theme_tokens"]},
                str(tmp / "preview_l3.pptx"),
                mode="preview",
                snapshot_root=str(home / "snapshots"),
            )
            self.assertEqual(prev_out["status"], "BUILT")
            self.assertEqual(prev_out["slide_count"], 1)

            # Complete wizard -> syncs L1 scenarios, L2 frameworks, L3 skins & base templates
            workspace_manager.plan_questions(wdir, home)
            while True:
                nb = workspace_manager.next_batch(wdir)
                if nb["status"] != "BATCH_READY" or nb["phase"] != "core":
                    break
                workspace_manager.record_answer(wdir, {q["id"]: "A" for q in nb["questions"]}, home)
            workspace_manager.choose_path(wdir, "ship_now", home)

            # Manually add a scenario, adjust a medium-confidence framework, and bind/resolve combo
            class PArgs:
                profile = "default"
                mask = "skin_2"
                label = ""
                layer = "L2"
                rule_text = ""
                rule_no = 0
                out = ""
                scenarios = ""
                independent = True
                scenario_name = "年度述职答辩"
                scenario_id = ""
                skeleton = "成果总览 → 核心项目 → 明年规划"
                framework_id = ""
                decision = "adjust"
                note = "按两栏安全留白重排"
                scenario = "月会"
                subaction = "confirm-scenario"

            sc_res = workspace_manager.handle_profile(PArgs, home)
            self.assertEqual(sc_res["status"], "SCENARIO_CONFIRMED")

            PArgs.subaction = "bind-mask"
            bind_res = workspace_manager.handle_profile(PArgs, home)
            self.assertEqual(bind_res["binding_mode"], "independent")

            # Resolve default bound combo vs mixed combo
            PArgs.subaction = "resolve-combo"
            PArgs.mask = None
            combo_def = workspace_manager.handle_profile(PArgs, home)
            self.assertEqual(combo_def["status"], "COMBO_RESOLVED")
            self.assertEqual(combo_def["combo_mode"], "scenario_default")
            self.assertTrue(Path(combo_def["l3_mask"]["template_pptx"]).is_file())

            PArgs.mask = "skin_2"
            combo_mix = workspace_manager.handle_profile(PArgs, home)
            self.assertEqual(combo_mix["combo_mode"], "manual_mix")
            self.assertEqual(combo_mix["l3_mask"]["mask_id"], "skin_2")

            # Fixed 5-section usage guide passes output_guard cleanly
            ug = workspace_manager.generate_usage_guide(home, wdir=wdir)
            self.assertEqual(ug["status"], "OK")
            self.assertTrue(ug["guard_ok"])
            for sec_hdr in (
                "【你的 PPT 风格资产速览】",
                "【按你的场景直接发起】",
                "【日常 5 档怎么开口（按需点菜）】",
                "【自由混搭与临时借参考】",
                "【手改进化、反悔与随时查口令】",
            ):
                self.assertIn(sec_hdr, ug["guide_text"])

    def test_13_v020_multi_round_human_ai_timeline_and_in_place_replace(self) -> None:
        from pptx import Presentation

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            snap_root = tmp / "snaps"
            work_pptx = tmp / "季度汇报.pptx"

            # R1: AI builds 2 slides
            r1 = pptx_builder.build_pptx(
                {
                    "slides": [
                        {"title": "第一页：整体进展", "blocks": [{"type": "bullets", "items": ["完成基础搭建"]}]},
                        {"title": "第二页：待AI优化", "blocks": [{"type": "bullets", "items": ["初版要点"]}]},
                    ]
                },
                str(work_pptx),
                mode="full",
                snapshot_root=str(snap_root),
            )
            self.assertEqual(r1["rev_id"], "R1")
            doc_id = r1["doc_id"]

            # Human Edit 1: user edits slide 1 title on disk
            prs1 = Presentation(str(work_pptx))
            prs1.slides[0].shapes[0].text_frame.paragraphs[0].text = "第一页：Q3整体进展超预期"
            prs1.save(str(work_pptx))

            # AI Revise (patch replace slide 2 in-place): inherits doc_id, captures H1 as R2, saves AI R3, leaves slide 1 untouched!
            r3 = pptx_builder.build_pptx(
                {
                    "replace_slides": [
                        {"slide_no": 2, "title": "第二页：AI重写深化方案", "blocks": [{"type": "bullets", "items": ["补充细节数据"]}]}
                    ]
                },
                str(work_pptx),
                mode="patch",
                base_pptx_str=str(work_pptx),
                snapshot_root=str(snap_root),
            )
            self.assertEqual(r3["status"], "BUILT")
            self.assertEqual(r3["doc_id"], doc_id)
            self.assertTrue(r3["captured_prior_human_rev"]["captured"])
            self.assertEqual(r3["captured_prior_human_rev"]["rev_id"], "R2")
            self.assertEqual(r3["rev_id"], "R3")
            self.assertEqual(r3["slide_count"], 2)

            sem_r3 = snapshot_diff.extract_semantic_slides(work_pptx)
            self.assertEqual(sem_r3[0].title, "第一页：Q3整体进展超预期")
            self.assertEqual(sem_r3[1].title, "第二页：AI重写深化方案")

            # Human Edit 2: user edits slide 2 body on disk
            prs2 = Presentation(str(work_pptx))
            prs2.slides[1].shapes[2].text_frame.paragraphs[0].text = "补充细节数据：人工核实达成92%"
            prs2.save(str(work_pptx))

            # Latest diff sees ONLY R3 -> H2 (slide 2 body edit), and mark_reviewed prevents re-reporting it!
            diff_latest = snapshot_diff.diff_against_snapshot(
                str(work_pptx),
                snapshot_root=str(snap_root),
                mode="latest",
                mark_reviewed_after=True,
            )
            self.assertEqual(diff_latest["status"], "CHANGES_DETECTED")
            self.assertEqual(diff_latest["base_rev_id"], "R3")
            self.assertEqual(len(diff_latest["changes"]), 1)
            self.assertIn("第 2 页", diff_latest["changes"][0]["slide_ref"])

            diff_after_mark = snapshot_diff.diff_against_snapshot(
                str(work_pptx),
                snapshot_root=str(snap_root),
                mode="latest",
            )
            self.assertEqual(diff_after_mark["status"], "NO_CHANGES_DETECTED")

            # Cumulative human diff summarizes R1->R2 (slide 1) and R3->R4 (slide 2), excluding AI R2->R3!
            diff_cum = snapshot_diff.diff_against_snapshot(
                str(work_pptx),
                snapshot_root=str(snap_root),
                mode="cumulative_human",
            )
            self.assertEqual(diff_cum["status"], "CHANGES_DETECTED")
            self.assertEqual(diff_cum["human_rounds_count"], 2)
            self.assertEqual(len(diff_cum["excluded_ai_revisions"]), 2)
            slide_refs = " | ".join(c["slide_ref"] for c in diff_cum["changes"])
            self.assertIn("第 1 页", slide_refs)
            self.assertIn("第 2 页", slide_refs)


if __name__ == "__main__":
    unittest.main(verbosity=2)
