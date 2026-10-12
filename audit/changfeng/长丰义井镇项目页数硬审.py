#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本项目页数硬审脚本
项目要求：
- 施工组织设计正文目标：50页，允许±1页
- 不含封面和目录
- 当前配套主硬审规则已规定：封面 + 一级目录仅占第1页，正文从第2页开始
- 因此：正文合格范围为49–51页，对应Word总页数合格范围为50–52页

运行环境：
    Windows
    Microsoft Word 已安装
    Python 3.10+
    pip install pywin32

用法：
    python "本项目页数硬审.py" "施工组织设计.docx"

退出码：
    0 = PASS
    2 = FAIL
    3 = ERROR
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime
from pathlib import Path

WD_STATISTIC_PAGES = 2

# -----------------------------
# 本项目固定页数规则
# -----------------------------
EXPECTED_BODY_PAGES = 50
PAGE_TOLERANCE = 1
FRONT_MATTER_PAGES = 1   # 第1页：封面 + 一级目录
MIN_BODY_PAGES = EXPECTED_BODY_PAGES - PAGE_TOLERANCE
MAX_BODY_PAGES = EXPECTED_BODY_PAGES + PAGE_TOLERANCE
MIN_TOTAL_PAGES = MIN_BODY_PAGES + FRONT_MATTER_PAGES
MAX_TOTAL_PAGES = MAX_BODY_PAGES + FRONT_MATTER_PAGES


class AuditError(RuntimeError):
    pass


def audit(document: Path, report_path: Path, stage: int | None = None) -> int:
    if sys.platform != "win32":
        raise AuditError(
            "该脚本必须在 Windows 环境运行，因为页数必须由本机 Microsoft Word 实际重新分页后判定。"
        )

    if not document.exists():
        raise AuditError(f"文件不存在：{document}")

    if document.suffix.lower() not in (".docx", ".docm", ".doc"):
        raise AuditError("仅支持 Microsoft Word 文档（.docx/.docm/.doc）。")

    try:
        import win32com.client  # type: ignore
    except ImportError as e:
        raise AuditError("缺少 pywin32，请先运行：pip install pywin32") from e

    word = None
    doc = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(
            str(document.resolve()),
            ReadOnly=True,
            AddToRecentFiles=False
        )

        # 使用 Microsoft Word 自身排版引擎重新分页
        doc.Repaginate()
        total_pages = int(doc.ComputeStatistics(WD_STATISTIC_PAGES))

        if total_pages <= FRONT_MATTER_PAGES:
            body_pages = 0
        else:
            body_pages = total_pages - FRONT_MATTER_PAGES

        passed = (
            MIN_BODY_PAGES <= body_pages <= MAX_BODY_PAGES
            and MIN_TOTAL_PAGES <= total_pages <= MAX_TOTAL_PAGES
        )

        status = "PASS" if passed else "FAIL"

        report = {
            "audit_type": "project_page_count_gate",
            "page_source": "microsoft_word",
            "document_sha256": hashlib.sha256(document.read_bytes()).hexdigest(),
            "audited_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "document": str(document.resolve()),
            "status": status,
            "rule": {
                "body_pages_target": EXPECTED_BODY_PAGES,
                "page_tolerance": PAGE_TOLERANCE,
                "body_pages_allowed_range": [MIN_BODY_PAGES, MAX_BODY_PAGES],
                "front_matter_pages_excluded": FRONT_MATTER_PAGES,
                "front_matter_definition": "第1页为封面 + 一级目录",
                "body_start_page": 2,
                "total_pages_allowed_range": [MIN_TOTAL_PAGES, MAX_TOTAL_PAGES]
            },
            "actual": {
                "word_total_pages": total_pages,
                "excluded_front_matter_pages": FRONT_MATTER_PAGES,
                "body_pages": body_pages
            }
        }

        if stage is not None:
            report.update({"audit_type":"word_stage_measurement","status":"MEASURED","stage_percent":stage,
                "project_rule":{"target_body_pages":EXPECTED_BODY_PAGES,"page_tolerance":PAGE_TOLERANCE,"front_matter_pages_excluded":FRONT_MATTER_PAGES}})
            report["actual"]["word_body_pages"]=body_pages
            report["actual"]["word_version"]=str(word.Version)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

        if stage is not None:
            print(f"WORD_STAGE_STATUS=MEASURED\nSTAGE={stage}\nBODY_PAGES={body_pages}\nREPORT={report_path}")
            return 0

        print("\n=== 本项目页数硬审结果 ===")
        print(f"Microsoft Word 实际总页数：{total_pages}")
        print(f"排除封面及目录页数：{FRONT_MATTER_PAGES}")
        print(f"实际正文页数：{body_pages}")
        print(f"目标正文页数：{EXPECTED_BODY_PAGES}（允许±{PAGE_TOLERANCE}页）")
        print(f"正文合格范围：{MIN_BODY_PAGES}–{MAX_BODY_PAGES}页")
        print(f"总页数合格范围：{MIN_TOTAL_PAGES}–{MAX_TOTAL_PAGES}页")
        print("--------------------------------")
        print(f"PAGE_AUDIT_STATUS={status}")
        print(f"REPORT={report_path}")

        if not passed:
            print("FINAL_DELIVERY_BLOCKED")
            return 2

        return 0

    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="本项目页数硬审：正文目标50页，允许±1页（不含封面和目录）"
    )
    parser.add_argument("--stage",type=int,choices=(40,75),help="只测阶段页数，不按最终49–51页判定")
    parser.add_argument("document", type=Path, help="待审 Word 文件路径")
    parser.add_argument(
        "--report",
        type=Path,
        default=None,
        help="审核报告 JSON 输出路径（可选）"
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report = args.report or args.document.with_name(
        args.document.stem + "_页数硬审报告.json"
    )

    try:
        return audit(args.document, report, args.stage)
    except AuditError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 3
    except Exception as e:
        print(f"ERROR: 未预期错误：{e}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
