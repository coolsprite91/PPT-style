#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""output_guard.py — 全链路对用户输出守卫与规则库脱敏核查器（零第三方依赖）。

子命令:
  check-reply    检查即将发给用户的回复是否泄露色号/内部文件名/XML标签、超5题或漏报题数
  sanitize-rule  校验规则条目是否混入未抽象的真实业务数字/日期，并支持自动脱敏抽象
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


HEX_COLOR_RE = re.compile(r"(?<![A-Za-z0-9&])#[0-9A-Fa-f]{6}\b")
RGB_FUNC_RE = re.compile(r"\brgb\s*\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}\s*\)", re.IGNORECASE)
INTERNAL_FILE_RE = re.compile(
    r"(?:env_probe|pptx_inspector|workspace_manager|pptx_builder|snapshot_diff|output_guard|install)\.py"
    r"|\.(?:archive-manifest|wizard-state)\.json\b"
    r"|registry\.json\b|core_l[12]_[a-z]+\.json\b|\.snapshots\b|core_properties\b"
)
XML_LEAK_RE = re.compile(r"(?:<a:[a-zA-Z]+|<p:[a-zA-Z]+|\bschemeClr\b|\blumMod\b|\blumOff\b|\bclrMap\b|\bslideMaster\b|\bslideLayout\b)")

QUESTION_LINE_RE = re.compile(r"(?m)^\s*(?:\*\*)?([1-9]\d?)\s*[\.、\)](?!\d)")
OPTION_MARKER_RE = re.compile(r"(?m)(?:^|\s|\*\*)(?:A|B|C)\s*[\.、\)]")
QUESTION_COUNT_RE = re.compile(
    r"(?:共\s*\d+\s*[个道项轮]|还有\s*\d+\s*[个道项]|\d+\s*[个道项]\s*(?:问题|确认项|待补充|小习惯))"
)

CONCRETE_DATE_RE = re.compile(
    r"(?<![0-9])(?:19|20)\d{2}\s*(?:"
    r"年\s*\d{1,2}\s*月(?:\s*\d{1,2}\s*[日号]?)?"
    r"|年"
    r"|[\-/\.]\s*\d{1,2}(?:[\-/\.]\s*\d{1,2})?"
    r")"
)
CONCRETE_METRIC_RE = re.compile(
    r"(?<![0-9A-Za-z\.])\d+(?:,\d{3})*(?:\.\d+)?\s*(?:%|‰|个百分点|万元|亿元|万美元|元|万家|万户|家|台|人|小时|分钟|秒)(?![A-Za-z])"
)


def check_reply_text(text: str, expect_questions: int | None = None) -> dict[str, Any]:
    """机械校验拟发送给用户的文本是否符合 C 端极简规范。"""
    violations: list[dict[str, str]] = []

    hex_hits = HEX_COLOR_RE.findall(text) + RGB_FUNC_RE.findall(text)
    if hex_hits:
        violations.append({
            "code": "HEX_COLOR_LEAK",
            "detail": f"检测到十六进制或RGB色号泄露：{', '.join(hex_hits[:3])}（请改用“深蓝/浅灰”等自然语言）",
        })

    file_hits = INTERNAL_FILE_RE.findall(text)
    if file_hits:
        violations.append({
            "code": "INTERNAL_FILE_LEAK",
            "detail": f"检测到内部脚本或配置文件名泄露：{', '.join(dict.fromkeys(file_hits))}",
        })

    xml_hits = XML_LEAK_RE.findall(text)
    if xml_hits:
        violations.append({
            "code": "XML_JARGON_LEAK",
            "detail": f"检测到内部技术标签泄露：{', '.join(dict.fromkeys(xml_hits))}",
        })

    # Extract distinct question numbers only when options or question phrasing exist
    q_nums = [int(m.group(1)) for m in QUESTION_LINE_RE.finditer(text)]
    has_question_context = bool(OPTION_MARKER_RE.search(text)) or ("？" in text) or (expect_questions is not None)
    distinct_q = sorted(set(q_nums)) if has_question_context else []

    if distinct_q and max(distinct_q) > 5 and len(distinct_q) > 5:
        violations.append({
            "code": "TOO_MANY_QUESTIONS",
            "detail": f"单轮问题数达到 {len(distinct_q)} 道，超过每轮最多 5 题的硬上限",
        })

    if expect_questions is not None and expect_questions > 5:
        violations.append({
            "code": "TOO_MANY_QUESTIONS",
            "detail": f"指定问题数 {expect_questions} 超过单轮 5 题上限",
        })

    if (len(distinct_q) >= 1 and bool(OPTION_MARKER_RE.search(text))) or (expect_questions and expect_questions > 0):
        # Check first 450 characters or first 6 lines for explicit total question count
        head_text = "\n".join(text.splitlines()[:8])
        if not QUESTION_COUNT_RE.search(head_text):
            violations.append({
                "code": "MISSING_QUESTION_COUNT",
                "detail": "包含待回答问题，但开头未显式报明精确的问题总数（如“共 N 个问题”）",
            })

    return {
        "ok": len(violations) == 0,
        "detected_question_count": len(distinct_q),
        "violations": violations,
    }


def sanitize_rule_text(text: str, auto_abstract: bool = False) -> dict[str, Any]:
    """检查即将入库规则库或导出分享的条目是否包含未抽象的真实业务日期或具体数字。"""
    findings: list[str] = []
    sanitized = text

    dates = CONCRETE_DATE_RE.findall(text)
    if dates:
        findings.append(f"包含具体业务日期：{', '.join(dates[:3])}")
        if auto_abstract:
            sanitized = CONCRETE_DATE_RE.sub("[时间]", sanitized)

    metrics = CONCRETE_METRIC_RE.findall(text)
    # Allow benign typography sizes like pt/磅/页/轮/条/题/列/行/栏
    if metrics:
        findings.append(f"包含具体业务数字或金额：{', '.join(metrics[:3])}")
        if auto_abstract:
            sanitized = CONCRETE_METRIC_RE.sub("[具体数值与单位]", sanitized)

    ok_after = True if (auto_abstract or not findings) else False
    return {
        "ok": ok_after,
        "had_raw_specifics": bool(findings),
        "findings": findings,
        "sanitized_text": sanitized,
    }


def _read_text_arg(text_arg: str | None, file_arg: str | None) -> str:
    if text_arg is not None:
        return text_arg
    if file_arg is not None:
        return Path(file_arg).read_text(encoding="utf-8")
    return sys.stdin.read()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PPT-style 对用户输出守卫与规则脱敏校验工具")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_check = sub.add_parser("check-reply", help="检查回复文本是否违规")
    p_check.add_argument("--text", help="待检查文本")
    p_check.add_argument("--file", help="待检查文本文件路径")
    p_check.add_argument("--expect-questions", type=int, default=None, help="预期问题总数")

    p_san = sub.add_parser("sanitize-rule", help="检查并脱敏规则文本")
    p_san.add_argument("--text", help="规则文本")
    p_san.add_argument("--file", help="规则文本文件")
    p_san.add_argument("--auto-abstract", action="store_true", help="自动替换为抽象占位符")

    args = parser.parse_args(argv)
    if args.cmd == "check-reply":
        raw = _read_text_arg(args.text, args.file)
        res = check_reply_text(raw, expect_questions=args.expect_questions)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0 if res["ok"] else 2

    if args.cmd == "sanitize-rule":
        raw = _read_text_arg(args.text, args.file)
        res = sanitize_rule_text(raw, auto_abstract=args.auto_abstract)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return 0 if res["ok"] else 2

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
