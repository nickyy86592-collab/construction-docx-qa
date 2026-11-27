#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按冻结稿的 Microsoft Word 实排物理页逐项核查评审点与非封面页数。"""
from __future__ import annotations

import argparse
from bisect import bisect_left
from datetime import datetime
import hashlib
import json
from pathlib import Path
import sys


# 数值政策的唯一来源。EXPECTED_BODY_PAGES 在本项目表示非封面页总上限。
EXPECTED_BODY_PAGES = 100
PAGE_TOLERANCE = 0
FRONT_MATTER_PAGES = 1
COVER_PAGES = 1
REVIEW_POINT_COUNT = 10
REVIEW_POINT_PAGE_LIMIT = 10
# 编制前由 AI 按有效招标顺序填写全部十个最终正文一级章名，包含实际使用的编号。
REVIEW_POINT_HEADINGS = []

WD_STATISTIC_PAGES = 2
WD_ACTIVE_END_PAGE_NUMBER = 3
WD_WITHIN_TABLE = 12
LAYOUT = " \t\r\n\v\f\x00\x0e\u00a0\u200b"


def normalized(text):
    return "".join(text.split())


def rule_snapshot():
    headings = REVIEW_POINT_HEADINGS
    if (type(REVIEW_POINT_COUNT) is not int or REVIEW_POINT_COUNT < 1
            or len(headings) != REVIEW_POINT_COUNT
            or not all(isinstance(h, str) and normalized(h) for h in headings)
            or len({normalized(h) for h in headings}) != REVIEW_POINT_COUNT):
        raise ValueError("须先按有效招标原文登记全部评审点的不同一级章名")
    if (not all(type(n) is int for n in (
            EXPECTED_BODY_PAGES, PAGE_TOLERANCE, FRONT_MATTER_PAGES, COVER_PAGES, REVIEW_POINT_PAGE_LIMIT))
            or PAGE_TOLERANCE != 0 or not 0 <= COVER_PAGES <= FRONT_MATTER_PAGES
            or REVIEW_POINT_PAGE_LIMIT < 1 or EXPECTED_BODY_PAGES <= FRONT_MATTER_PAGES - COVER_PAGES):
        raise ValueError("封面、前置页及逐项上限配置无效")
    body_limit = EXPECTED_BODY_PAGES - (FRONT_MATTER_PAGES - COVER_PAGES)
    return {
        "body_pages_target": EXPECTED_BODY_PAGES,
        "page_tolerance": PAGE_TOLERANCE,
        "front_matter_pages_excluded": FRONT_MATTER_PAGES,
        "body_pages_allowed_range": [1, body_limit],
        "total_pages_allowed_range": [FRONT_MATTER_PAGES + 1, EXPECTED_BODY_PAGES + COVER_PAGES],
        "non_cover_page_limit": EXPECTED_BODY_PAGES,
        "cover_pages_excluded": COVER_PAGES,
        "review_point_count": REVIEW_POINT_COUNT,
        "review_point_page_limit": REVIEW_POINT_PAGE_LIMIT,
        "review_point_headings": headings,
    }


def page_at(doc, position):
    page = int(doc.Range(position, position).Information(WD_ACTIVE_END_PAGE_NUMBER))
    if page < 1:
        raise RuntimeError("Word 无法确定正文位置的物理页")
    return page


def paragraph_ranges(doc):
    """Bound ranges by the next paragraph when Word treats a native field as one large range."""
    items = iter(doc.Paragraphs)
    current = next(items, None)
    while current is not None:
        following = next(items, None)
        rng = current.Range
        if following is not None:
            next_start = int(following.Range.Start)
            if int(rng.Start) < next_start < int(rng.End):
                rng = doc.Range(int(rng.Start), next_start)
        yield current, rng
        current = following


def measure_points(doc, rule, reference=False):
    """Locate each point's heading and farthest occupied physical page."""
    front = rule["front_matter_pages_excluded"]
    planned = {normalized(title): title for title in rule["review_point_headings"]}
    floating = sorted(int(shape.Anchor.Start) for shape in doc.Shapes if int(shape.Anchor.StoryType) == 1)
    rows, errors = [], []
    for paragraph, rng in paragraph_ranges(doc):
        text = str(rng.Text or "")
        visible = text.strip(LAYOUT + "\x07")
        start, end = int(rng.Start), int(rng.End)
        anchors = floating[bisect_left(floating, start):bisect_left(floating, end)]
        if not visible and "\x07" not in text and not anchors:
            continue
        tail = len(text) - len(text.rstrip(LAYOUT))
        last = max(start, end - tail - 1)
        positions = [last, *anchors]
        end_page = max(page_at(doc, pos) for pos in set(positions))
        if int(paragraph.OutlineLevel) == 1 and visible and not rng.Information(WD_WITHIN_TABLE):
            first = min(start + len(text) - len(text.lstrip(LAYOUT)), last)
            start_page = page_at(doc, first)
            if start_page <= front:
                if normalized(visible) in planned or end_page > front:
                    errors.append(f"物理第{start_page}页的评审点标题进入封面或目录")
                continue
            heading = visible
            if normalized(heading) not in planned:
                list_number = str(rng.ListFormat.ListString or "").strip()
                if list_number:
                    heading = list_number + " " + heading
            rows.append({"heading": planned.get(normalized(heading), heading),
                         "start_page": start_page, "end_page": end_page, "has_content": False})
        elif rows:
            # In a split table row the visually last cell need not be the last Word paragraph.
            rows[-1]["end_page"] = max(rows[-1]["end_page"], end_page)
            rows[-1]["has_content"] |= bool(visible or anchors)
        elif end_page > front:
            errors.append("首个评审点前存在未归属的正文内容")
    limit = rule["review_point_page_limit"]
    for index, row in enumerate(rows, 1):
        pages = row["end_page"] - row["start_page"] + 1
        has_content = row["has_content"]
        row.update(index=index, pages=pages,
                   status="PASS" if has_content and (reference or 1 <= pages <= limit) else "FAIL")
        if not has_content:
            errors.append(f"第{index}项 {row['heading']} 缺少正文、图或表")
        if not reference and not 1 <= pages <= limit:
            errors.append(f"第{index}项 {row['heading']} 实排 {pages} 页，上限 {limit} 页")
    return rows, list(dict.fromkeys(errors))


def build_report(doc, word_version, document, document_hash, rule, stage=None, reference=False):
    doc.Repaginate()
    total = int(doc.ComputeStatistics(WD_STATISTIC_PAGES))
    rows, failures = measure_points(doc, rule, reference)
    front = rule["front_matter_pages_excluded"]
    cover = rule["cover_pages_excluded"]
    non_cover = total - cover
    headings = rule["review_point_headings"]
    expected = headings[:len(rows)] if stage is not None else headings
    if not rows or [row["heading"] for row in rows] != expected:
        failures.append("正文评审点标题缺失、重复、顺序错误或与项目登记不符")
    if rows and rows[0]["start_page"] != front + 1:
        failures.append("首个评审点与登记的正文起始物理页不符")
    if any(not front < row["start_page"] <= row["end_page"] <= total for row in rows):
        failures.append("逐项报告存在无效的物理页区间")
    if any(previous["end_page"] > current["start_page"] for previous, current in zip(rows, rows[1:])):
        failures.append("相邻评审点物理页范围倒序")
    if not total > front or (not reference and non_cover > rule["non_cover_page_limit"]):
        failures.append(f"扣除封面实排 {non_cover} 页，上限 {rule['non_cover_page_limit']} 页")
    report = {
        "audit_type": "project_page_count_gate" if stage is None else "word_stage_measurement",
        "page_source": "microsoft_word",
        "reference_only": reference,
        "document": str(document.resolve()),
        "document_sha256": document_hash,
        "page_script_sha256": sha256(Path(__file__)),
        "audited_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "status": ("FAIL" if failures else "PASS") if stage is None else "MEASURED",
        "point_limit_status": "FAIL" if failures else "PASS",
        "rule": rule,
        "actual": {
            "word_total_pages": total,
            "body_pages": max(0, total - front),
            "non_cover_pages": non_cover,
            "excluded_cover_pages": cover,
            "excluded_front_matter_pages": front,
            "review_points": rows,
            "word_version": str(word_version),
        },
        "failures": failures,
    }
    if stage is not None:
        report.update(stage_percent=stage, measurement_role="checkpoint",
                      project_rule={"target_body_pages": rule["body_pages_target"],
                                    "page_tolerance": rule["page_tolerance"],
                                    "front_matter_pages_excluded": front})
        report["actual"]["word_body_pages"] = max(0, total - front)
        if rows:
            start = int(doc.GoTo(What=1, Which=1, Count=front + 1).Start)
            body = str(doc.Range(start, int(doc.Content.End)).Text or "").replace("\x07", "")
            report["actual"]["word_body_characters"] = len(normalized(body))
    return report


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def audit(document, report_path, stage=None, reference=False):
    rule = rule_snapshot()
    if sys.platform != "win32":
        raise RuntimeError("本脚本需要 Windows Microsoft Word 实排")
    if not document.is_file() or document.suffix.lower() != ".docx":
        raise ValueError("须提供存在的 DOCX 冻结稿")
    import win32com.client

    document_hash = sha256(document)
    word = win32com.client.DispatchEx("Word.Application")
    doc = None
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(document.resolve()), ReadOnly=True, AddToRecentFiles=False)
        report = build_report(doc, word.Version, document, document_hash, rule, stage, reference)
        if sha256(document) != document_hash:
            raise RuntimeError("测量期间 DOCX 发生变化")
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        finally:
            word.Quit()

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"PAGE_AUDIT_STATUS={report['status']}\nPOINT_LIMIT_STATUS={report['point_limit_status']}")
    for point in report["actual"]["review_points"]:
        print(f"{point['heading']} | 物理页 {point['start_page']}–{point['end_page']} | {point['pages']} 页 | {point['status']}")
    for failure in report["failures"]:
        print("- " + failure)
    print("REPORT=" + str(report_path))
    return 2 if report["status"] == "FAIL" else 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("document", type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--stage", type=int, choices=(40, 75))
    parser.add_argument("--reference", action="store_true")
    args = parser.parse_args()
    report = args.report or args.document.with_name(args.document.stem + "_页数硬审报告.json")
    try:
        return audit(args.document, report, args.stage, args.reference)
    except Exception as exc:
        print("PAGE_AUDIT_STATUS=ERROR\n" + str(exc), file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
