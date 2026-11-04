#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
巢湖项目页数硬审脚本
项目要求：
- 施工组织设计全篇共50页，不设页数浮动
- 封面和一级目录共占第1页，按1页前置页扣除
- 正文从物理总第2页开始，共49页；该页显示正文页码1，后续顺序编号至49
- 封面和一级目录所在的物理第1页不显示页码；正文页码使用阿拉伯数字
- 合格范围：正文49–49页；Word总页数50–50页
- 页数以Windows Microsoft Word实际打开并重新分页后的结果为准
- 本脚本核验Word实排总页数及扣除前置页后的正文页数，不修改被审文档
- 封面、一级目录实际位置、正文起始位置、前置页不显示页码及正文从1编号，由通用Word格式硬审核验
- 报告rule中的页码字段是项目要求，不代表本脚本已完成页码显示核验

运行环境：
    Windows
    Microsoft Word 已安装
    Python 3.10+
    pip install pywin32

用法：
    python "巢湖项目页数硬审.py" "施工组织设计.docx"

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
EXPECTED_BODY_PAGES = 49
PAGE_TOLERANCE = 0
FRONT_MATTER_PAGES = 1   # 固定：封面和一级目录共占第1页；正文从第2页开始
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
            "project_name": "巢湖项目",
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
                "front_matter_definition": f"封面和一级目录合计{FRONT_MATTER_PAGES}页，占第1页",
                "body_start_page": FRONT_MATTER_PAGES + 1,
                "front_matter_page_numbers_visible": False,
                "body_page_number_start": 1,
                "body_page_number_format": "arabic",
                "page_numbering_verifier": "generic_word_format_audit",
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

        print("\n=== 巢湖项目页数硬审结果 ===")
        print(f"Microsoft Word 实际总页数：{total_pages}")
        print(f"排除封面和一级目录页数：{FRONT_MATTER_PAGES}")
        print(f"实际正文页数：{body_pages}")
        print("页数要求：全篇50页；封面和一级目录共占物理第1页且不显示页码；正文从物理第2页开始，共49页，显示页码1—49（不设浮动）")
        print("核验分工：本脚本核验实排页数；封面目录位置及页码显示由通用Word格式硬审核验。")
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
        description="巢湖项目页数硬审：全篇50页；封面和一级目录共占物理第1页且不显示页码；正文从物理第2页开始共49页，正文页码从1开始（不设浮动）。本脚本核验页数，页码显示由通用Word格式硬审核验。"
    )
    parser.add_argument("--stage",type=int,choices=(40,75),help="只测阶段页数，不按最终全篇50页、正文49页判定")
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
