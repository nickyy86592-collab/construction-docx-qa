#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
长丰义井镇施工组织设计：Microsoft Word 阶段页数实测

用途：
- 仅用于正文编制约 40% / 75% 时的阶段校准。
- 直接调用 Windows Microsoft Word 重新分页，返回真实 Word 页数。
- 不执行最终 50±1 页硬审，不因阶段稿尚未达到 50 页而 FAIL。
- 目标页数、容差、前置页数唯一读取自同目录 page_gate.py，避免重复维护。

退出码：
    0 = MEASURED（已成功取得 Microsoft Word 实测页数）
    3 = ERROR
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import sys
from datetime import datetime
from pathlib import Path

WD_STATISTIC_PAGES = 2

class MeasureError(RuntimeError):
    pass

def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def load_project_page_rules():
    gate_path = Path(__file__).with_name("page_gate.py")
    if not gate_path.exists():
        raise MeasureError(f"缺少项目页数硬审脚本：{gate_path}")

    spec = importlib.util.spec_from_file_location("project_page_gate", gate_path)
    if spec is None or spec.loader is None:
        raise MeasureError("无法加载项目页数硬审脚本。")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    required = ("EXPECTED_BODY_PAGES", "PAGE_TOLERANCE", "FRONT_MATTER_PAGES")
    missing = [x for x in required if not hasattr(module, x)]
    if missing:
        raise MeasureError("项目页数硬审脚本缺少必要常量：" + ", ".join(missing))

    return (
        int(module.EXPECTED_BODY_PAGES),
        int(module.PAGE_TOLERANCE),
        int(module.FRONT_MATTER_PAGES),
    )

def measure(document: Path, stage_percent: int, report_path: Path, text_report_path: Path) -> int:
    if sys.platform != "win32":
        raise MeasureError("阶段实测必须在 Windows + Microsoft Word 环境运行。")

    if not document.exists():
        raise MeasureError(f"阶段稿不存在：{document}")

    if document.suffix.lower() not in (".docx", ".docm", ".doc"):
        raise MeasureError("仅支持 Microsoft Word 文档（.docx/.docm/.doc）。")

    if not (1 <= stage_percent <= 99):
        raise MeasureError("stage_percent 必须为 1–99 的整数。")

    target_body, tolerance, front_pages = load_project_page_rules()

    try:
        import win32com.client  # type: ignore
    except ImportError as e:
        raise MeasureError("缺少 pywin32。") from e

    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(
            str(document.resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        doc.Repaginate()

        total_pages = int(doc.ComputeStatistics(WD_STATISTIC_PAGES))
        body_pages = max(total_pages - front_pages, 0)
        ratio = stage_percent / 100.0
        expected_at_stage = target_body * ratio
        rough_projected_final = body_pages / ratio if ratio > 0 else None

        min_final = target_body - tolerance
        max_final = target_body + tolerance
        if rough_projected_final is None:
            advisory = "UNKNOWN"
        elif rough_projected_final < min_final:
            advisory = "EXPAND_LATER_CONTENT"
        elif rough_projected_final > max_final:
            advisory = "COMPRESS_LATER_CONTENT"
        else:
            advisory = "ON_TRACK"

        try:
            word_version = str(word.Version)
        except Exception:
            word_version = None

        report = {
            "measurement_type": "microsoft_word_stage_page_measurement",
            "measured_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "page_source": "microsoft_word",
            "document": str(document.resolve()),
            "document_sha256": sha256_file(document),
            "status": "MEASURED",
            "stage_percent": stage_percent,
            "project_rule_source": str(Path(__file__).with_name("page_gate.py").resolve()),
            "project_rule": {
                "target_body_pages": target_body,
                "page_tolerance": tolerance,
                "front_matter_pages_excluded": front_pages,
                "allowed_final_body_pages": [target_body - tolerance, target_body + tolerance],
            },
            "actual": {
                "word_total_pages": total_pages,
                "word_body_pages": body_pages,
                "word_version": word_version,
            },
            "planning_reference": {
                "linear_expected_body_pages_at_stage": round(expected_at_stage, 2),
                "rough_linear_projected_final_body_pages": round(rough_projected_final, 2) if rough_projected_final is not None else None,
                "advisory": advisory,
                "note": "阶段投影仅作篇幅校准参考；正式调整仍应结合评分权重、章节计划和图表占用，不得机械按线性比例凑页。",
            },
        }

        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

        lines = [
            "=== Microsoft Word 阶段页数实测 ===",
            f"阶段：约 {stage_percent}%",
            f"页数数据源：Microsoft Word {word_version or ''}".rstrip(),
            f"Word 实际总页数：{total_pages}",
            f"排除前置页：{front_pages}",
            f"Word 实际正文页数：{body_pages}",
            f"项目最终目标：正文 {target_body} 页，允许 ±{tolerance} 页",
            f"阶段线性参考值：{expected_at_stage:.2f} 页",
            f"粗略线性投影最终正文：{rough_projected_final:.2f} 页",
            f"建议方向：{advisory}",
            "说明：阶段投影不是最终硬审结果，不得机械凑页。",
        ]
        text_report_path.parent.mkdir(parents=True, exist_ok=True)
        text_report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

        print("\n".join(lines))
        print("WORD_STAGE_MEASURE_STATUS=MEASURED")
        print("PAGE_SOURCE=MICROSOFT_WORD")
        print(f"WORD_BODY_PAGES={body_pages}")
        print(f"REPORT={report_path}")
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
    parser = argparse.ArgumentParser(description="Microsoft Word 阶段页数实测（不执行最终页数硬审）")
    parser.add_argument("document", type=Path, help="阶段稿 Word 文件")
    parser.add_argument("--stage-percent", type=int, required=True, help="当前编制阶段百分比，例如 40 或 75")
    parser.add_argument("--report", type=Path, required=True, help="JSON 报告输出路径")
    parser.add_argument("--text-report", type=Path, required=True, help="文本报告输出路径")
    return parser.parse_args()

def main() -> int:
    args = parse_args()
    try:
        return measure(args.document, args.stage_percent, args.report, args.text_report)
    except MeasureError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 3
    except Exception as e:
        print(f"ERROR: 未预期错误：{e}", file=sys.stderr)
        return 3

if __name__ == "__main__":
    raise SystemExit(main())
