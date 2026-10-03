#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用 Word 正式交付硬审脚本（Windows + Microsoft Word）

用途：
- 不绑定某一个编制 Skill，可被不同文档生产 Skill 复用。
- 使用 Windows Microsoft Word 实际排版结果定位首页、正文起始页、空白页等版式信息。
- 不再审核“总页数/正文页数目标”；页数要求由独立页数硬审脚本负责。
- 对首页一级目录、正文三级标题、连续排版、必要图件、
  实景图占位、字体规则、Word 导航窗格、全篇竖向页面，以及图片迁移导致的大块页底留白等进行硬审。

运行环境：
    Windows
    Microsoft Word 已安装
    Python 3.10+
    pip install pywin32

用法：
    python word_delivery_gate.py "施工组织设计.docx"
    python word_delivery_gate.py "施工组织设计.docx" --rules "audit_rules.json"
    python word_delivery_gate.py "施工组织设计.docx" --rules "audit_rules.json" --report "audit_report.json"

退出码：
    0 = PASS（硬性规则全部通过）
    2 = FAIL（存在硬性失败项）
    3 = ERROR（脚本/Word 环境错误）

说明：
1. 与招标文件相关的“一级章节、字体、必备图件、实景图占位”等项目参数，可由编制 Skill
   在读取招标文件后自动生成临时规则并传入；用户无需手工编辑 JSON。
2. 对项目敏感参数采用“未知(None) / 明确无要求(空列表或空字典)”两种状态。
   若参数仍为未知，视为项目规则尚未提取完成，将直接阻止正式交付，避免假 PASS。
3. “是否为 AI 生成的实景图”无法仅靠 DOCX 文件可靠鉴别。脚本能硬性检查：
   当规则明确某处暂无真实素材时，必须存在规范实景图占位；不能机器证明图片来源真实性。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


# -----------------------------
# Word 常量（避免依赖 makepy）
# -----------------------------
WD_STATISTIC_PAGES = 2
WD_ACTIVE_END_PAGE_NUMBER = 3
WD_VERTICAL_POSITION_RELATIVE_TO_PAGE = 6
WD_GOTO_PAGE = 1
WD_GOTO_ABSOLUTE = 1
WD_OUTLINE_BODY_TEXT = 10
WD_SECTION_CONTINUOUS = 0
WD_SECTION_NEW_COLUMN = 1
WD_SECTION_NEW_PAGE = 2
WD_SECTION_EVEN_PAGE = 3
WD_SECTION_ODD_PAGE = 4
WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1


DEFAULT_RULES: Dict[str, Any] = {
    # 页数目标已剥离到独立页数硬审脚本；本脚本仅保留页面定位所需的 front/body 参数。
    "front_page": 1,
    "body_start_page": 2,
    "allow_front_matter_break": True,
    "hard_fail_on_mechanical_structure": True,
    "min_chapters_for_mechanical_check": 3,
    # None = 尚未从招标文件/项目资料提取；[] / {} = 已确认“没有额外硬性要求”
    "required_level1_titles": None,
    "required_figures": None,
    "required_real_photo_slots": None,
    "font_rules": None,
    "low_content_page_char_threshold": 30,
    # 通用版式待审阈值；不是招标页数参数。1/3仅用于触发图片迁移留白复核。
    "bottom_whitespace_review_ratio": 1 / 3,
    "next_page_visual_top_ratio": 0.30,
    "next_page_pre_visual_char_limit": 140,
}

# 三级标题格式
RE_L1 = re.compile(r"^第([一二三四五六七八九十百零〇两\d]+)章(?:[、：:．.\s]+)?(.+)?$")
RE_L2 = re.compile(r"^第([一二三四五六七八九十百零〇两\d]+)节(?:[、：:．.\s]+)?(.+)?$")
RE_L3 = re.compile(r"^([一二三四五六七八九十百零〇两\d]+)、\s*(.+)?$")
RE_REAL_PHOTO_PLACEHOLDER = re.compile(r"【实景图占位：[^】]+】")


@dataclass
class CheckResult:
    name: str
    status: str  # PASS / FAIL / WARN / SKIP
    message: str
    details: Any = None


class AuditError(RuntimeError):
    pass


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def clean_text(text: str) -> str:
    if not text:
        return ""
    return text.replace("\r", "").replace("\x07", "").replace("\x0b", "").replace("\x0c", "").strip()


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", clean_text(text)).strip()


def normalize_toc_title(text: str) -> str:
    """去掉目录点线和末尾页码，用于首页目录 vs 正文一级标题比对。"""
    t = normalize_spaces(text)
    # 例如：第一章 工程概况........2 / 第一章 工程概况  2
    t = re.sub(r"[\.·…．。]{2,}\s*\d+\s*$", "", t)
    t = re.sub(r"\s+\d+\s*$", "", t)
    return t.strip()


def load_rules(path: Optional[Path]) -> Dict[str, Any]:
    rules = json.loads(json.dumps(DEFAULT_RULES, ensure_ascii=False))
    if path:
        with path.open("r", encoding="utf-8") as f:
            incoming = json.load(f)
        if not isinstance(incoming, dict):
            raise AuditError("规则文件顶层必须是 JSON 对象。")
        rules.update(incoming)
    return rules


def chinese_num_to_int(s: str) -> Optional[int]:
    s = s.strip()
    if not s:
        return None
    if s.isdigit():
        return int(s)

    s = s.replace("〇", "零").replace("两", "二")
    digit = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    unit = {"十": 10, "百": 100, "千": 1000}

    total = 0
    section = 0
    number = 0
    seen = False
    for ch in s:
        if ch in digit:
            number = digit[ch]
            seen = True
        elif ch in unit:
            seen = True
            u = unit[ch]
            if number == 0:
                number = 1
            section += number * u
            number = 0
        else:
            return None
    total += section + number
    return total if seen else None


def page_of_range(rng) -> int:
    try:
        return int(rng.Information(WD_ACTIVE_END_PAGE_NUMBER))
    except Exception:
        return -1


def get_page_range(doc, page_num: int, page_count: int):
    start = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=page_num).Start
    if page_num < page_count:
        end = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=page_num + 1).Start
    else:
        end = doc.Content.End
    return doc.Range(start, end)


def iter_find_ranges(doc, find_text: str) -> List[Tuple[int, int, int]]:
    """返回 (start, end, page) 列表。Word Find 特殊字符：^m 手工分页；^b 分节符。"""
    hits: List[Tuple[int, int, int]] = []
    r = doc.Content.Duplicate
    while True:
        finder = r.Find
        finder.ClearFormatting()
        found = finder.Execute(FindText=find_text, Forward=True, Wrap=0)
        if not found:
            break
        hits.append((int(r.Start), int(r.End), page_of_range(r)))
        new_start = int(r.End)
        if new_start >= int(doc.Content.End):
            break
        r.SetRange(new_start, doc.Content.End)
    return hits


def get_paragraph_records(doc) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    for i in range(1, int(doc.Paragraphs.Count) + 1):
        p = doc.Paragraphs(i)
        text = normalize_spaces(p.Range.Text)
        if not text:
            continue
        try:
            outline = int(p.OutlineLevel)
        except Exception:
            outline = WD_OUTLINE_BODY_TEXT
        try:
            page_break_before = bool(p.Format.PageBreakBefore)
        except Exception:
            page_break_before = False
        rec = {
            "index": i,
            "text": text,
            "page": page_of_range(p.Range),
            "outline": outline,
            "page_break_before": page_break_before,
            "start": int(p.Range.Start),
            "end": int(p.Range.End),
        }
        records.append(rec)
    return records


def classify_heading(text: str) -> Tuple[int, Optional[int]]:
    for level, regex in ((1, RE_L1), (2, RE_L2), (3, RE_L3)):
        m = regex.match(text)
        if m:
            return level, chinese_num_to_int(m.group(1))
    return 0, None


def extract_heading_structure(records: List[Dict[str, Any]], body_start_page: int) -> Tuple[List[Dict[str, Any]], List[CheckResult]]:
    results: List[CheckResult] = []
    headings: List[Dict[str, Any]] = []
    bad_outline: List[Dict[str, Any]] = []
    bad_syntax: List[Dict[str, Any]] = []

    for rec in records:
        if rec["page"] < body_start_page:
            continue
        textual_level, number = classify_heading(rec["text"])
        outline_level = rec["outline"] if rec["outline"] in (1, 2, 3) else 0

        if textual_level:
            x = dict(rec)
            x["level"] = textual_level
            x["number"] = number
            headings.append(x)
            if outline_level != textual_level:
                bad_outline.append({"text": rec["text"], "page": rec["page"], "expected_outline": textual_level, "actual_outline": rec["outline"]})
        elif outline_level in (1, 2, 3):
            bad_syntax.append({"text": rec["text"], "page": rec["page"], "outline": outline_level})

    if bad_outline:
        results.append(CheckResult("导航窗格_标题大纲级别", "FAIL", "存在标题文字格式正确但 Word 大纲级别错误，导航窗格会异常。", bad_outline))
    else:
        results.append(CheckResult("导航窗格_标题大纲级别", "PASS", "三级标题均使用对应 Word 大纲级别。"))

    if bad_syntax:
        results.append(CheckResult("标题样式_语法一致性", "FAIL", "存在已设置为 1–3 级大纲、但文字不符合“章/节/一、二……”格式的标题。", bad_syntax))
    else:
        results.append(CheckResult("标题样式_语法一致性", "PASS", "Word 1–3 级大纲标题均符合规定文字格式。"))

    return headings, results


def check_front_page_and_body(doc, records: List[Dict[str, Any]], headings: List[Dict[str, Any]], rules: Dict[str, Any], page_count: int) -> List[CheckResult]:
    out: List[CheckResult] = []
    front = int(rules["front_page"])
    body_start = int(rules["body_start_page"])

    # 正文一级标题必须从第2页开始；且该页必须有实际正文内容承接。
    l1 = [h for h in headings if h["level"] == 1]
    if not l1:
        out.append(CheckResult("正文起始页", "FAIL", "未检测到正文一级标题。"))
    else:
        first = l1[0]
        first_page = first["page"]
        page_start_records_after = [
            r for r in records
            if r["page"] == body_start and r["start"] > first["start"] and normalize_spaces(r["text"])
        ]
        try:
            start_page_range = get_page_range(doc, body_start, page_count)
            start_page_has_object = count_objects_on_page(doc, start_page_range) > 0
        except Exception:
            start_page_has_object = False
        has_follow_content = bool(page_start_records_after) or start_page_has_object
        status = "PASS" if first_page == body_start and has_follow_content else "FAIL"
        msg = (
            f"首个正文一级标题位于第 {first_page} 页，且第 {body_start} 页存在后续正文/图表内容。"
            if status == "PASS"
            else f"正文起始不符合要求：首个一级标题在第 {first_page} 页；要求第 {body_start} 页开始且该页应有实际正文内容。"
        )
        out.append(CheckResult("正文起始页", status, msg, {
            "first_heading_page": first_page,
            "has_follow_content_on_start_page": has_follow_content,
        }))

    # 首页目录项不得使用 Heading 1–3，否则会污染导航窗格并被误判为正文标题。
    front_outline = [
        {"text": r["text"], "outline": r["outline"]}
        for r in records if r["page"] == front and r["outline"] in (1, 2, 3)
    ]
    if front_outline:
        out.append(CheckResult(
            "首页内容_目录不得进入导航窗格", "FAIL",
            "首页封面/一级目录存在 Heading 1–3 大纲级别；首页目录应为普通目录文本，正文标题才进入导航窗格。",
            front_outline,
        ))
    else:
        out.append(CheckResult("首页内容_目录不得进入导航窗格", "PASS", "首页未发现 Heading 1–3 大纲标题。"))

    front_range = get_page_range(doc, front, page_count)
    front_text = clean_text(front_range.Text.replace("\r", "\n"))
    front_lines = [normalize_spaces(x) for x in re.split(r"[\r\n]+", front_text) if normalize_spaces(x)]

    # 首页不得出现二级、三级标题格式内容。
    bad_lower = [x for x in front_lines if RE_L2.match(x) or RE_L3.match(x)]
    if bad_lower:
        out.append(CheckResult("首页内容_仅封面及一级目录", "FAIL", "首页出现二级或三级标题格式内容；首页仅允许封面信息及一级目录。", bad_lower))
    else:
        out.append(CheckResult("首页内容_仅封面及一级目录", "PASS", "首页未发现二级/三级标题格式内容。"))

    # 首页一级目录与正文一级标题名称、数量、顺序必须完全一致。
    toc_candidates: List[Tuple[int, str]] = []
    for idx, line in enumerate(front_lines):
        n = normalize_toc_title(line)
        if RE_L1.match(n):
            if not toc_candidates or toc_candidates[-1][1] != n:
                toc_candidates.append((idx, n))
    page1_toc = [x[1] for x in toc_candidates]
    body_l1 = [normalize_toc_title(h["text"]) for h in l1]

    # 首页必须同时存在有效封面信息，不能只有一级目录。
    # “有效封面信息”按可程序化边界定义为：首个一级目录项之前至少存在一行
    # 非空、非纯页码/装饰符号、且不是一级目录格式的文字。
    cover_lines: List[str] = []
    if toc_candidates:
        first_toc_idx = toc_candidates[0][0]
        for line in front_lines[:first_toc_idx]:
            n = normalize_toc_title(line)
            if not n:
                continue
            if RE_L1.match(n):
                continue
            if re.fullmatch(r"[\d\s\.·…．。_—-]+", n):
                continue
            cover_lines.append(line)
    if cover_lines:
        out.append(CheckResult(
            "首页封面信息", "PASS",
            "首页一级目录之前检测到有效封面信息。",
            cover_lines[:20],
        ))
    else:
        out.append(CheckResult(
            "首页封面信息", "FAIL",
            "首页未检测到有效封面信息，或一级目录从首页首项直接开始；第1页必须同时包含封面及一级目录。",
        ))

    if page1_toc == body_l1 and body_l1:
        out.append(CheckResult("首页一级目录_与正文一致", "PASS", "首页一级目录名称、数量及顺序与正文一级标题完全一致。", body_l1))
    else:
        out.append(CheckResult("首页一级目录_与正文一致", "FAIL", "首页一级目录与正文一级标题存在缺项、错序、数量或名称不一致。", {
            "page1_toc": page1_toc,
            "body_level1": body_l1,
        }))

    # 可程序化的首页边界检查：一级目录开始后，不得再混入普通正文。
    # 目录之前的普通文字视作封面信息；其语义无法仅凭 DOCX 可靠区分，因此不做猜测性误杀。
    suspicious_after_toc: List[str] = []
    if toc_candidates:
        first_toc_idx = toc_candidates[0][0]
        toc_line_indexes = {x[0] for x in toc_candidates}
        for idx in range(first_toc_idx, len(front_lines)):
            if idx in toc_line_indexes:
                continue
            line = front_lines[idx]
            if re.fullmatch(r"[\d\s\.·…．。-]+", line):
                continue
            suspicious_after_toc.append(line)
    if suspicious_after_toc:
        out.append(CheckResult("首页目录后正文混入", "FAIL", "首页一级目录开始后出现非一级目录的普通正文/其他内容。", suspicious_after_toc))
    else:
        out.append(CheckResult("首页目录后正文混入", "PASS", "首页一级目录区域未发现普通正文混入。"))

    return out

def check_project_parameter_completeness(rules: Dict[str, Any]) -> CheckResult:
    required_keys = {
        "required_level1_titles": "招标文件一级章节要求",
        "required_figures": "招标文件/评分要求中的必要图件",
        "required_real_photo_slots": "实景图素材/占位需求",
        "font_rules": "招标文件字体字号要求",
    }
    missing = [label for key, label in required_keys.items() if rules.get(key) is None]
    if missing:
        return CheckResult(
            "项目硬审参数完整性",
            "FAIL",
            "项目敏感规则尚未提取完整，禁止在未知条件下判定正式交付。编制 Skill 应先读取招标文件/项目资料并自动传入参数；用户无需手工维护 JSON。",
            missing,
        )
    return CheckResult("项目硬审参数完整性", "PASS", "项目敏感规则已明确；空列表/空字典表示已确认无额外硬性要求。")


def check_required_level1(headings: List[Dict[str, Any]], rules: Dict[str, Any]) -> CheckResult:
    raw = rules.get("required_level1_titles")
    if raw is None:
        return CheckResult("招标文件一级章节", "FAIL", "尚未取得招标文件对一级章节的要求，无法完成硬审。")
    required = [normalize_spaces(x) for x in raw if normalize_spaces(x)]
    actual = [normalize_toc_title(h["text"]) for h in headings if h["level"] == 1]
    if not required:
        return CheckResult("招标文件一级章节", "PASS", "已确认招标文件未规定固定一级章节清单；不执行名称逐项比对。", {"actual": actual})
    status = "PASS" if actual == required else "FAIL"
    return CheckResult("招标文件一级章节", status, "正文一级章节与招标文件清单完全一致。" if status == "PASS" else "正文一级章节与招标文件规定不一致。", {"required": required, "actual": actual})

def check_heading_level_presence(headings: List[Dict[str, Any]]) -> CheckResult:
    counts = {1: 0, 2: 0, 3: 0}
    for h in headings:
        level = h.get("level")
        if level in counts:
            counts[level] += 1
    missing = [level for level, count in counts.items() if count == 0]
    if missing:
        names = {1: "一级标题（第一章、第二章……）", 2: "二级标题（第一节、第二节……）", 3: "三级标题（一、二、三……）"}
        return CheckResult(
            "三级标题体系_层级完整性",
            "FAIL",
            "正文未实际形成完整的三级标题体系，缺少：" + "、".join(names[x] for x in missing) + "。",
            {"counts": counts, "missing_levels": missing},
        )
    return CheckResult(
        "三级标题体系_层级完整性",
        "PASS",
        "正文已实际检测到一级、二级、三级三种标题层级。",
        {"counts": counts},
    )


def check_heading_sequence(headings: List[Dict[str, Any]]) -> CheckResult:
    errors: List[Dict[str, Any]] = []
    expected_l1 = 1
    expected_l2 = 1
    expected_l3 = 1
    current_l1 = None
    current_l2 = None

    for h in headings:
        level = h["level"]
        num = h["number"]
        if num is None:
            errors.append({"text": h["text"], "page": h["page"], "reason": "标题编号无法解析"})
            continue

        if level == 1:
            if num != expected_l1:
                errors.append({"text": h["text"], "page": h["page"], "expected": expected_l1, "actual": num})
            current_l1 = num
            current_l2 = None
            expected_l1 = num + 1
            expected_l2 = 1
            expected_l3 = 1
        elif level == 2:
            if current_l1 is None:
                errors.append({"text": h["text"], "page": h["page"], "reason": "二级标题出现在一级标题之前"})
            if num != expected_l2:
                errors.append({"text": h["text"], "page": h["page"], "expected": expected_l2, "actual": num})
            current_l2 = num
            expected_l2 = num + 1
            expected_l3 = 1
        elif level == 3:
            if current_l1 is None or current_l2 is None:
                errors.append({"text": h["text"], "page": h["page"], "reason": "三级标题未处于有效二级标题之下"})
            if num != expected_l3:
                errors.append({"text": h["text"], "page": h["page"], "expected": expected_l3, "actual": num})
            expected_l3 = num + 1

    return CheckResult("三级标题编号连续性", "FAIL" if errors else "PASS", "发现标题跳号、重号或层级异常。" if errors else "三级标题编号连续、层级关系正常。", errors or None)


def build_structure_counts(headings: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    chapters: List[Dict[str, Any]] = []
    current_ch = None
    current_sec = None
    for h in headings:
        if h["level"] == 1:
            current_ch = {"title": h["text"], "sections": []}
            chapters.append(current_ch)
            current_sec = None
        elif h["level"] == 2 and current_ch is not None:
            current_sec = {"title": h["text"], "level3_count": 0}
            current_ch["sections"].append(current_sec)
        elif h["level"] == 3 and current_sec is not None:
            current_sec["level3_count"] += 1
    return chapters


def check_mechanical_structure(headings: List[Dict[str, Any]], rules: Dict[str, Any]) -> CheckResult:
    chapters = build_structure_counts(headings)
    min_ch = int(rules.get("min_chapters_for_mechanical_check", 3))
    if len(chapters) < min_ch:
        return CheckResult("机械化固定目录结构", "SKIP", f"一级章节少于 {min_ch}，不启用固定 N×N 结构拦截。")

    l2_counts = [len(ch["sections"]) for ch in chapters]
    all_l2_same = bool(l2_counts) and len(set(l2_counts)) == 1 and l2_counts[0] > 0
    all_l3_counts = [sec["level3_count"] for ch in chapters for sec in ch["sections"]]
    all_l3_same = bool(all_l3_counts) and len(set(all_l3_counts)) == 1 and all_l3_counts[0] > 0

    if all_l2_same and all_l3_same:
        detail = {
            "chapter_count": len(chapters),
            "level2_per_chapter": l2_counts[0],
            "level3_per_section": all_l3_counts[0],
        }
        return CheckResult("机械化固定目录结构", "FAIL", "检测到全篇明显机械化固定 N×N 目录结构。", detail)
    return CheckResult("机械化固定目录结构", "PASS", "未发现全篇统一固定 N×N 的明显机械化目录模式。", {"level2_counts": l2_counts})


def _section_layout_signature(section) -> Tuple[Any, Any, Any]:
    try:
        ps = section.PageSetup
        return (int(ps.Orientation), round(float(ps.PageWidth), 1), round(float(ps.PageHeight), 1))
    except Exception:
        return (None, None, None)


def _first_nonempty_paragraph_in_section(section) -> Optional[Dict[str, Any]]:
    try:
        for i in range(1, int(section.Range.Paragraphs.Count) + 1):
            p = section.Range.Paragraphs(i)
            text = normalize_spaces(p.Range.Text)
            if text:
                return {"text": text, "page": page_of_range(p.Range), "start": int(p.Range.Start)}
    except Exception:
        pass
    return None


def check_pagination_controls(doc, records: List[Dict[str, Any]], headings: List[Dict[str, Any]], rules: Dict[str, Any]) -> List[CheckResult]:
    out: List[CheckResult] = []
    body_start_page = int(rules["body_start_page"])

    # 标题段前分页：硬失败。
    bad_heading_pb = [
        {"text": h["text"], "page": h["page"]}
        for h in headings if h.get("page_break_before")
    ]
    out.append(CheckResult(
        "标题强制分页_PageBreakBefore",
        "FAIL" if bad_heading_pb else "PASS",
        "存在标题设置‘段前分页/另起一页’。" if bad_heading_pb else "未发现标题段前强制分页。",
        bad_heading_pb or None,
    ))

    # 正文内部手工分页符：硬失败。
    manual_breaks = iter_find_ranges(doc, "^m")
    body_manual = [x for x in manual_breaks if x[2] >= body_start_page]
    out.append(CheckResult(
        "正文手工分页符",
        "FAIL" if body_manual else "PASS",
        "正文中存在手工分页符。" if body_manual else "正文中未发现手工分页符。",
        [{"page": x[2], "start": x[0]} for x in body_manual] or None,
    ))

    # 分节符不能一刀切：有些文档会因页眉页脚、页码等使用分节。
    # 但本规则明确禁止任何横向页面，因此横向版式由独立“页面方向”硬审直接 FAIL。
    # 此处仅对“同版式 + 强制新页 + 新节第一段就是标题”的情形判硬失败；其他正文强制新页分节给 WARN。
    heading_forced: List[Dict[str, Any]] = []
    layout_sections: List[Dict[str, Any]] = []
    section_count = int(doc.Sections.Count)
    for idx in range(2, section_count + 1):
        sec = doc.Sections(idx)
        first = _first_nonempty_paragraph_in_section(sec)
        if not first or first["page"] < body_start_page:
            continue
        try:
            start_type = int(sec.PageSetup.SectionStart)
        except Exception:
            start_type = None
        if start_type not in (WD_SECTION_NEW_PAGE, WD_SECTION_EVEN_PAGE, WD_SECTION_ODD_PAGE):
            continue
        prev = doc.Sections(idx - 1)
        layout_changed = _section_layout_signature(prev) != _section_layout_signature(sec)
        heading_level, _ = classify_heading(first["text"])
        detail = {
            "section": idx,
            "page": first["page"],
            "first_text": first["text"],
            "section_start": start_type,
            "layout_changed": layout_changed,
        }
        if heading_level in (1, 2, 3) and not layout_changed:
            heading_forced.append(detail)
        else:
            layout_sections.append(detail)

    out.append(CheckResult(
        "标题相关强制新页分节",
        "FAIL" if heading_forced else "PASS",
        "发现同版式下通过新页分节使标题另起一页。" if heading_forced else "未发现利用新页分节强制标题另起一页。",
        heading_forced or None,
    ))
    if layout_sections:
        out.append(CheckResult(
            "正文其他强制新页分节",
            "WARN",
            "正文存在其他强制新页分节；虽然未直接认定为标题强制分页，但仍建议确认其必要性。页面横向与否由独立页面方向硬审判定。",
            layout_sections,
        ))
    else:
        out.append(CheckResult("正文其他强制新页分节", "PASS", "未发现需人工确认的正文强制新页分节。"))
    return out

def _section_page_span(doc, section) -> Tuple[int, int]:
    """返回分节大致起止页码，用于报告定位。"""
    try:
        start_pos = int(section.Range.Start)
        end_pos = max(start_pos, int(section.Range.End) - 1)
        start_rng = doc.Range(start_pos, min(start_pos + 1, int(doc.Content.End)))
        end_rng = doc.Range(end_pos, min(end_pos + 1, int(doc.Content.End)))
        return page_of_range(start_rng), page_of_range(end_rng)
    except Exception:
        return -1, -1


def check_page_orientation(doc, rules: Dict[str, Any]) -> CheckResult:
    """硬审：整份正式 Word 必须全程竖向，任何图表不得通过横向页面承载。"""
    bad: List[Dict[str, Any]] = []
    for idx in range(1, int(doc.Sections.Count) + 1):
        sec = doc.Sections(idx)
        try:
            ps = sec.PageSetup
            orientation = int(ps.Orientation)
            width = float(ps.PageWidth)
            height = float(ps.PageHeight)
        except Exception as e:
            bad.append({"section": idx, "reason": f"无法读取页面方向：{e}"})
            continue

        # Word: 0=Portrait, 1=Landscape。再用宽高关系做冗余校验。
        is_landscape = orientation == WD_ORIENT_LANDSCAPE or width > height
        if is_landscape:
            start_page, end_page = _section_page_span(doc, sec)
            bad.append({
                "section": idx,
                "start_page": start_page,
                "end_page": end_page,
                "orientation": orientation,
                "page_width_pt": round(width, 2),
                "page_height_pt": round(height, 2),
                "reason": "检测到横向页面/横向分节",
            })

    if bad:
        return CheckResult(
            "页面方向_全篇竖版",
            "FAIL",
            "发现横向页面。正式 Word 全篇必须保持竖向；横道图、平面图等过宽内容应缩放、重排或拆分，不得通过横向页面展示。",
            bad,
        )
    return CheckResult(
        "页面方向_全篇竖版",
        "PASS",
        "全部分节均为竖向页面，未发现横向页面。",
    )


def count_objects_on_page(doc, page_range) -> int:
    count = 0
    try:
        count += int(page_range.InlineShapes.Count)
    except Exception:
        pass
    try:
        count += int(page_range.Tables.Count)
    except Exception:
        pass
    try:
        for i in range(1, int(doc.Shapes.Count) + 1):
            shp = doc.Shapes(i)
            try:
                if int(shp.Anchor.Start) >= int(page_range.Start) and int(shp.Anchor.Start) < int(page_range.End):
                    count += 1
            except Exception:
                pass
    except Exception:
        pass
    return count


def check_blank_and_low_content_pages(doc, page_count: int, rules: Dict[str, Any]) -> List[CheckResult]:
    """基础检查：整页空白与极低文字量页面。不能替代真实几何留白检查。"""
    blank: List[int] = []
    low: List[Dict[str, Any]] = []
    threshold = int(rules.get("low_content_page_char_threshold", 30))

    for p in range(1, page_count + 1):
        r = get_page_range(doc, p, page_count)
        text = clean_text(r.Text)
        compact = re.sub(r"\s+", "", text)
        obj_count = count_objects_on_page(doc, r)
        if len(compact) == 0 and obj_count == 0:
            blank.append(p)
        elif p not in (1, page_count) and len(compact) < threshold and obj_count == 0:
            low.append({"document_page": p, "char_count": len(compact)})

    return [
        CheckResult("空白页", "FAIL" if blank else "PASS", "发现无内容空白页。" if blank else "未发现无内容空白页。", blank or None),
        CheckResult("低内容页_辅助线索", "WARN" if low else "PASS", "存在正文页文字量异常偏低；该项仅作辅助线索，仍需结合实际页面几何判断。" if low else "未发现明显低内容页。", low or None),
    ]


def _safe_float(value) -> Optional[float]:
    try:
        x = float(value)
        if x <= -9990000 or x != x:
            return None
        return x
    except Exception:
        return None


def _range_vertical_y(doc, pos: int) -> Optional[float]:
    try:
        end_limit = max(1, int(doc.Content.End))
        pos = max(0, min(int(pos), end_limit - 1))
        r = doc.Range(pos, min(pos + 1, end_limit))
        y = _safe_float(r.Information(WD_VERTICAL_POSITION_RELATIVE_TO_PAGE))
        return y if y is not None and y >= 0 else None
    except Exception:
        return None


def _page_body_box(page_range) -> Optional[Dict[str, float]]:
    try:
        ps = page_range.Sections(1).PageSetup
        page_height = _safe_float(ps.PageHeight)
        top_margin = _safe_float(ps.TopMargin)
        bottom_margin = _safe_float(ps.BottomMargin)
        if page_height is None or top_margin is None or bottom_margin is None:
            return None
        body_top = top_margin
        body_bottom = page_height - bottom_margin
        usable = body_bottom - body_top
        if usable <= 0:
            return None
        return {
            "body_top_pt": body_top,
            "body_bottom_pt": body_bottom,
            "usable_height_pt": usable,
        }
    except Exception:
        return None


def _visuals_on_page(doc, page_range, body_box: Optional[Dict[str, float]]) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    page_start, page_end = int(page_range.Start), int(page_range.End)

    try:
        count = int(page_range.InlineShapes.Count)
    except Exception:
        count = 0
    for i in range(1, count + 1):
        try:
            shp = page_range.InlineShapes(i)
            rng = shp.Range
            top = _range_vertical_y(doc, int(rng.Start))
            height = _safe_float(shp.Height)
            width = _safe_float(shp.Width)
            items.append({
                "kind": "InlineShape",
                "anchor_start": int(rng.Start),
                "top_pt": round(top, 2) if top is not None else None,
                "height_pt": round(height, 2) if height is not None else None,
                "width_pt": round(width, 2) if width is not None else None,
                "bottom_pt": round(top + height, 2) if top is not None and height is not None else None,
            })
        except Exception:
            pass

    try:
        shape_count = int(doc.Shapes.Count)
    except Exception:
        shape_count = 0
    for i in range(1, shape_count + 1):
        try:
            shp = doc.Shapes(i)
            a = int(shp.Anchor.Start)
            if not (page_start <= a < page_end):
                continue
            anchor_y = _range_vertical_y(doc, a)
            raw_top = _safe_float(shp.Top)
            height = _safe_float(shp.Height)
            width = _safe_float(shp.Width)
            try:
                rel = int(shp.RelativeVerticalPosition)
            except Exception:
                rel = None
            top = None
            if raw_top is not None:
                if rel == 1:
                    top = raw_top
                elif rel == 0 and body_box is not None:
                    top = body_box["body_top_pt"] + raw_top
                elif anchor_y is not None:
                    top = anchor_y + raw_top
            items.append({
                "kind": "Shape",
                "anchor_start": a,
                "top_pt": round(top, 2) if top is not None else None,
                "height_pt": round(height, 2) if height is not None else None,
                "width_pt": round(width, 2) if width is not None else None,
                "bottom_pt": round(top + height, 2) if top is not None and height is not None else None,
                "relative_vertical_position": rel,
            })
        except Exception:
            pass
    return items


def _page_content_bottom(doc, page_range, visuals: List[Dict[str, Any]]) -> Optional[float]:
    bottoms: List[float] = []
    page_start, page_end = int(page_range.Start), int(page_range.End)
    try:
        pcount = int(page_range.Paragraphs.Count)
    except Exception:
        pcount = 0
    for i in range(1, pcount + 1):
        try:
            p = page_range.Paragraphs(i)
            start = max(page_start, int(p.Range.Start))
            end = min(page_end - 1, max(start, int(p.Range.End) - 1))
            if end < start:
                continue
            seg = doc.Range(start, min(end + 1, int(doc.Content.End)))
            if not re.sub(r"\s+", "", clean_text(seg.Text)):
                continue
            y = _range_vertical_y(doc, end)
            if y is None:
                continue
            try:
                line = _safe_float(p.Format.LineSpacing)
            except Exception:
                line = None
            line = line if line is not None and 8 <= line <= 80 else 18.0
            bottoms.append(y + min(max(line * 0.85, 10.0), 36.0))
        except Exception:
            pass
    for item in visuals:
        b = _safe_float(item.get("bottom_pt"))
        if b is not None:
            bottoms.append(b)
    return max(bottoms) if bottoms else None


def _body_page_number(document_page: int, body_start_page: int) -> Optional[int]:
    return document_page - body_start_page + 1 if document_page >= body_start_page else None


def check_image_migration_whitespace(doc, page_count: int, rules: Dict[str, Any]) -> Tuple[List[CheckResult], List[int]]:
    body_start = int(rules.get("body_start_page", 2))
    threshold = float(rules.get("bottom_whitespace_review_ratio", 1 / 3))
    visual_top_limit = float(rules.get("next_page_visual_top_ratio", 0.30))
    pre_char_limit = int(rules.get("next_page_pre_visual_char_limit", 140))
    confirmed: List[Dict[str, Any]] = []
    review: List[Dict[str, Any]] = []
    incomplete: List[Dict[str, Any]] = []
    evidence_pages = set()
    saw_visual = False

    for p in range(max(body_start, 1), page_count):
        prev_r = get_page_range(doc, p, page_count)
        next_r = get_page_range(doc, p + 1, page_count)
        prev_box = _page_body_box(prev_r)
        next_box = _page_body_box(next_r)
        prev_visuals = _visuals_on_page(doc, prev_r, prev_box)
        next_visuals = _visuals_on_page(doc, next_r, next_box)
        if not next_visuals:
            continue
        saw_visual = True
        if prev_box is None or next_box is None:
            incomplete.append({"document_page": p, "reason": "无法读取页面正文区域几何"})
            continue
        content_bottom = _page_content_bottom(doc, prev_r, prev_visuals)
        if content_bottom is None:
            incomplete.append({"document_page": p, "reason": "无法取得该页最后实际内容纵向位置"})
            continue
        blank_height = max(0.0, prev_box["body_bottom_pt"] - content_bottom)
        blank_ratio = blank_height / prev_box["usable_height_pt"]
        if blank_ratio < threshold:
            continue

        geom = [x for x in next_visuals if _safe_float(x.get("top_pt")) is not None]
        if not geom:
            incomplete.append({"document_page": p + 1, "reason": "检测到下一页图片，但无法取得图片纵坐标"})
            continue
        first_visual = min(geom, key=lambda x: float(x["top_pt"]))
        visual_top = float(first_visual["top_pt"])
        top_ratio = max(0.0, (visual_top - next_box["body_top_pt"]) / next_box["usable_height_pt"])
        anchor_start = int(first_visual.get("anchor_start") or int(next_r.Start))
        try:
            pre_text = clean_text(doc.Range(int(next_r.Start), min(max(anchor_start, int(next_r.Start)), int(next_r.End))).Text)
            pre_chars = len(re.sub(r"\s+", "", pre_text))
        except Exception:
            pre_chars = None

        detail = {
            "previous_document_page": p,
            "previous_body_page": _body_page_number(p, body_start),
            "next_document_page": p + 1,
            "next_body_page": _body_page_number(p + 1, body_start),
            "bottom_blank_height_pt": round(blank_height, 2),
            "bottom_blank_ratio": round(blank_ratio, 4),
            "review_threshold_ratio": round(threshold, 4),
            "next_visual_top_ratio_of_body": round(top_ratio, 4),
            "pre_visual_char_count": pre_chars,
            "visual": first_visual,
        }
        if top_ratio <= visual_top_limit and pre_chars is not None and pre_chars <= pre_char_limit:
            detail["classification"] = "CONFIRMED_IMAGE_MIGRATION_RISK"
            detail["reason"] = "上一页大块页底留白后，下一页靠近正文顶部即出现图片，且图片前文字很少。"
            confirmed.append(detail)
        else:
            detail["classification"] = "REVIEW_REQUIRED"
            detail["reason"] = "页底留白超过待审阈值，但自动证据不足以直接判定图片迁移，需要人工复核。"
            review.append(detail)
        evidence_pages.update((p, p + 1))

    return [
        CheckResult(
            "彩图前页大块留白_图片迁移",
            "FAIL" if confirmed else "PASS",
            "发现高置信度图片迁移型大块页底留白；必须调整图片尺寸、插入位置或段落关系后重新 Word 渲染与复验。" if confirmed else "未发现高置信度图片迁移型大块页底留白。",
            confirmed or None,
        ),
        CheckResult(
            "大块页底留白_待复核",
            "WARN" if review else "PASS",
            "发现超过通用待审阈值的页底留白，但证据不足以自动确认；请结合相邻页 Word 版式复核。" if review else "未发现需要人工复核的图片相邻大块页底留白。",
            review or None,
        ),
        CheckResult(
            "彩图留白检测完整性",
            "WARN" if incomplete else "PASS",
            "部分图片页面无法取得足够实际几何信息，留白检测未完成；不得解释为已检查通过。" if incomplete else ("已使用 Microsoft Word 实际分页与页面位置数据完成彩图相邻留白检查。" if saw_visual else "正文未检测到可分析图片对象。"),
            incomplete or None,
        ),
    ], sorted(evidence_pages)


def export_whitespace_evidence(doc, pages: List[int], report_path: Path, evidence_dir: Optional[Path]) -> CheckResult:
    if not pages:
        return CheckResult("彩图留白证据", "PASS", "未触发彩图相邻留白异常，无需生成相邻页证据截图。")

    target = evidence_dir or report_path.with_name(report_path.stem + "_evidence")
    target.mkdir(parents=True, exist_ok=True)
    pdf_path = target / "word_layout_evidence.pdf"
    errors: List[str] = []
    png_files: List[str] = []
    try:
        doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
    except Exception as e:
        errors.append("Word PDF导出失败：" + str(e))

    if pdf_path.exists():
        try:
            import fitz  # type: ignore
            pdf = fitz.open(str(pdf_path))
            try:
                for p in pages:
                    if 1 <= p <= len(pdf):
                        pix = pdf[p - 1].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                        fn = target / f"word_page_{p:03d}.png"
                        pix.save(str(fn))
                        png_files.append(str(fn))
            finally:
                pdf.close()
        except Exception as e:
            errors.append("逐页PNG截图不可用：" + str(e))

    manifest = {
        "layout_engine": "Microsoft Word",
        "document_pages": pages,
        "word_pdf": str(pdf_path) if pdf_path.exists() else None,
        "png_screenshots": png_files,
        "errors": errors,
        "note": "1/3为空白待审初始阈值，不是招标硬性要求。",
    }
    (target / "evidence_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    if png_files:
        return CheckResult("彩图留白证据", "PASS", "已生成异常页及相邻图片页的 Word 页面 PNG 证据，并保留 Word 导出 PDF。", manifest)
    if pdf_path.exists():
        return CheckResult("彩图留白证据", "WARN", "已生成 Microsoft Word 版式 PDF；当前环境未生成逐页 PNG，请按报告页码在 PDF 中复核。", manifest)
    return CheckResult("彩图留白证据", "WARN", "异常页证据导出未完成；不得把证据项解释为已完成复核。", manifest)


def doc_full_text(doc) -> str:
    try:
        return clean_text(doc.Content.Text)
    except Exception:
        return ""


def find_text_range(doc, keyword: str):
    r = doc.Content.Duplicate
    finder = r.Find
    finder.ClearFormatting()
    found = finder.Execute(FindText=keyword, Forward=True, Wrap=0)
    return r if found else None


def has_object_near_range(doc, rng, window: int = 600) -> bool:
    start = max(0, int(rng.Start) - window)
    end = min(int(doc.Content.End), int(rng.End) + window)
    around = doc.Range(start, end)
    try:
        if int(around.InlineShapes.Count) > 0 or int(around.Tables.Count) > 0:
            return True
    except Exception:
        pass
    try:
        for i in range(1, int(doc.Shapes.Count) + 1):
            shp = doc.Shapes(i)
            a = int(shp.Anchor.Start)
            if start <= a <= end:
                return True
    except Exception:
        pass
    return False


def check_required_figures(doc, rules: Dict[str, Any]) -> List[CheckResult]:
    out: List[CheckResult] = []
    raw = rules.get("required_figures")
    if raw is None:
        return [CheckResult("招标要求必要图件", "FAIL", "尚未取得本项目必备图件要求，无法完成硬审。")]
    items = raw or []
    if not items:
        return [CheckResult("招标要求必要图件", "PASS", "已确认本项目没有额外指定的必备图件清单。")]

    for item in items:
        name = str(item.get("name", "未命名图件"))
        keywords = item.get("keywords") or [name]
        hard = bool(item.get("hard", False))
        found_kw = None
        rng = None
        for kw in keywords:
            rng = find_text_range(doc, str(kw))
            if rng is not None:
                found_kw = str(kw)
                break
        if rng is None:
            out.append(CheckResult(f"必要图件_{name}", "FAIL" if hard else "WARN", f"未检测到“{name}”相关标题/关键词；请补充或确认。", {"keywords": keywords}))
            continue
        has_obj = has_object_near_range(doc, rng)
        if has_obj:
            out.append(CheckResult(f"必要图件_{name}", "PASS", f"检测到“{found_kw}”且附近存在图形/表格对象。"))
        else:
            out.append(CheckResult(f"必要图件_{name}", "FAIL" if hard else "WARN", f"检测到“{found_kw}”，但附近未检测到图形/表格对象；可能只有文字标题。"))
    return out


def check_real_photo_slots(doc, rules: Dict[str, Any]) -> List[CheckResult]:
    raw = rules.get("required_real_photo_slots")
    if raw is None:
        return [CheckResult("实景图占位", "FAIL", "尚未完成实景图素材/占位需求判断，无法完成硬审。")]
    slots = raw or []
    if not slots:
        return [
            CheckResult("实景图占位", "PASS", "已确认当前文档不存在‘缺真实素材但必须预留’的实景图项。"),
            CheckResult("实景图_AI来源真实性", "WARN", "脚本不能仅凭 DOCX 可靠鉴别既有图片是否由 AI 生成；图片真实性仍需人工或素材来源记录确认。"),
        ]
    full = doc_full_text(doc)
    out: List[CheckResult] = []
    for item in slots:
        name = str(item.get("name", "实景图"))
        keywords = [str(x) for x in item.get("keywords", [])]
        # 必须出现规范占位，且占位内至少命中名称或关键词之一
        placeholders = RE_REAL_PHOTO_PLACEHOLDER.findall(full)
        matched = None
        for ph in placeholders:
            if name in ph or any(k in ph for k in keywords):
                matched = ph
                break
        if matched:
            out.append(CheckResult(f"实景图占位_{name}", "PASS", "已检测到规范实景图占位。", matched))
        else:
            out.append(CheckResult(f"实景图占位_{name}", "FAIL", "规则声明该处暂无真实素材，但未检测到规范占位。", {"expected_example": f"【实景图占位：拟插入{name}，建议尺寸××】"}))
    out.append(CheckResult("实景图_AI来源真实性", "WARN", "脚本不能仅凭 DOCX 可靠鉴别既有图片是否由 AI 生成；图片真实性仍需人工或素材来源记录确认。"))
    return out


def font_value_equal(actual: Any, expected: Any, numeric: bool = False) -> bool:
    if expected is None:
        return True
    if actual is None:
        return False
    if numeric:
        try:
            return abs(float(actual) - float(expected)) < 0.15
        except Exception:
            return False
    return str(actual).strip().lower() == str(expected).strip().lower()


def inspect_range_font(rng) -> Dict[str, Any]:
    try:
        if int(rng.End) > int(rng.Start) and clean_text(rng.Text):
            rr = rng.Duplicate
            if int(rr.End) > int(rr.Start):
                rr.End = max(int(rr.Start), int(rr.End) - 1)  # 去掉段落标记
        else:
            rr = rng
        font = rr.Font
        return {
            "name_east_asia": None if font.NameFarEast in (None, 9999999) else str(font.NameFarEast),
            "name": None if font.Name in (None, 9999999) else str(font.Name),
            "size_pt": None if font.Size in (None, 9999999) else float(font.Size),
            "bold": None if font.Bold in (None, 9999999) else bool(font.Bold),
        }
    except Exception:
        return {"name_east_asia": None, "name": None, "size_pt": None, "bold": None}


def check_fonts(doc, records: List[Dict[str, Any]], rules: Dict[str, Any]) -> CheckResult:
    raw = rules.get("font_rules")
    if raw is None:
        return CheckResult("招标文件字体规则", "FAIL", "尚未取得招标文件字体/字号要求，无法完成硬审。")
    fr = raw or {}
    if not fr:
        return CheckResult("招标文件字体规则", "PASS", "已确认招标文件没有额外字体/字号硬性规定。")

    errors: List[Dict[str, Any]] = []
    checked = 0
    # 重新按 Word Paragraph 索引取 Range，确保能读取有效 Font
    for rec in records:
        if rec["page"] < int(rules["body_start_page"]):
            continue
        level, _ = classify_heading(rec["text"])
        key = f"heading{level}" if level else "body"
        expected = fr.get(key)
        if not expected:
            continue
        checked += 1
        p = doc.Paragraphs(rec["index"])
        actual = inspect_range_font(p.Range)
        mismatch = {}
        for k in ("name_east_asia", "name", "size_pt", "bold"):
            if k in expected and expected[k] is not None:
                ok = font_value_equal(actual.get(k), expected[k], numeric=(k == "size_pt"))
                if not ok:
                    mismatch[k] = {"expected": expected[k], "actual": actual.get(k)}
        if mismatch:
            errors.append({"text": rec["text"][:80], "page": rec["page"], "type": key, "mismatch": mismatch})

    if checked == 0:
        return CheckResult("招标文件字体规则", "FAIL", "已提供字体规则，但没有匹配到任何可检查段落，无法证明字体要求已被执行，禁止正式交付。")
    return CheckResult("招标文件字体规则", "FAIL" if errors else "PASS", "发现字体/字号与招标文件参数不一致。" if errors else f"已按配置检查 {checked} 个段落，字体规则通过。", errors or {"checked_paragraphs": checked})


def summarize(results: Iterable[CheckResult]) -> Tuple[int, int, int, int]:
    counts = {"PASS": 0, "FAIL": 0, "WARN": 0, "SKIP": 0}
    for r in results:
        counts[r.status] = counts.get(r.status, 0) + 1
    return counts["PASS"], counts["FAIL"], counts["WARN"], counts["SKIP"]


def print_console(results: List[CheckResult], overall: str, report_path: Path) -> None:
    print("\n=== Word 正式交付硬审结果 ===")
    for r in results:
        print(f"[{r.status:<4}] {r.name}: {r.message}")
    p, f, w, s = summarize(results)
    print("--------------------------------")
    print(f"PASS={p}  FAIL={f}  WARN={w}  SKIP={s}")
    print(f"DELIVERY_STATUS={overall}")
    print(f"REPORT={report_path}")
    if overall != "PASS":
        print("FINAL_DELIVERY_BLOCKED")


def audit(docx_path: Path, rules_path: Optional[Path], report_path: Path, evidence_dir: Optional[Path] = None) -> int:
    if sys.platform != "win32":
        raise AuditError("该脚本要求在 Windows 上运行，因为页数必须由本机 Microsoft Word 实际重新分页后判定。")
    if not docx_path.exists():
        raise AuditError(f"文件不存在：{docx_path}")
    if docx_path.suffix.lower() not in (".docx", ".docm", ".doc"):
        raise AuditError("仅支持 Microsoft Word 文档（.docx/.docm/.doc）。")

    try:
        import win32com.client  # type: ignore
    except ImportError as e:
        raise AuditError("缺少 pywin32。请先运行：pip install pywin32") from e

    rules = load_rules(rules_path)
    word = None
    doc = None
    results: List[CheckResult] = []
    started = datetime.now().astimezone().isoformat(timespec="seconds")

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True, AddToRecentFiles=False)
        doc.Repaginate()

        # 总页数目标已从本脚本剥离；这里仅取得 Word 实际页数，供首页、正文起始页、空白页等版式检查定位使用。
        page_count = int(doc.ComputeStatistics(WD_STATISTIC_PAGES))

        records = get_paragraph_records(doc)
        results.append(check_project_parameter_completeness(rules))
        headings, heading_checks = extract_heading_structure(records, int(rules["body_start_page"]))
        results.extend(heading_checks)
        results.extend(check_front_page_and_body(doc, records, headings, rules, page_count))
        results.append(check_heading_level_presence(headings))
        results.append(check_required_level1(headings, rules))
        results.append(check_heading_sequence(headings))
        if rules.get("hard_fail_on_mechanical_structure", True):
            results.append(check_mechanical_structure(headings, rules))
        results.extend(check_pagination_controls(doc, records, headings, rules))
        results.append(check_page_orientation(doc, rules))
        results.extend(check_blank_and_low_content_pages(doc, page_count, rules))
        whitespace_checks, whitespace_evidence_pages = check_image_migration_whitespace(doc, page_count, rules)
        results.extend(whitespace_checks)
        results.append(export_whitespace_evidence(doc, whitespace_evidence_pages, report_path, evidence_dir))
        results.extend(check_required_figures(doc, rules))
        results.extend(check_real_photo_slots(doc, rules))
        results.append(check_fonts(doc, records, rules))

        hard_failures = [r for r in results if r.status == "FAIL"]
        overall = "PASS" if not hard_failures else "FAIL"

        report = {
            "audit_version": "1.4.0",
            "layout_engine": "Microsoft Word",
            "document": str(docx_path.resolve()),
            "document_sha256": sha256_file(docx_path),
            "audited_at": started,
            "delivery_status": overall,
            "hard_failure_count": len(hard_failures),
            "rules": rules,
            "checks": [asdict(r) for r in results],
            "notes": [
                "本脚本不审核总页数/正文页数目标；页数要求由独立页数硬审脚本负责。",
                "WARN/SKIP 不等于硬失败；只有 FAIL 会阻止正式交付。项目敏感参数未知、或已提供字体规则但未实际检查到任何段落时，会主动产生 FAIL，避免假 PASS。",
                "正式版全篇必须保持竖向页面；任何横道图、平面图等均不得通过横向页面展示。",
                "脚本不能单凭 DOCX 可靠鉴别既有实景图片是否由 AI 生成。",
                "彩图留白检测基于 Microsoft Word 实际分页后的页面几何；页底留白超过正文可用高度1/3仅为通用待审初始阈值，不是招标硬性要求。",
                "高置信度上一页大留白+下一页靠近页首出现图片判为FAIL；证据不足时为WARN并要求复核。",
            ],
        }
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print_console(results, overall, report_path)
        return 0 if overall == "PASS" else 2

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
    p = argparse.ArgumentParser(description="通用 Word 正式交付硬审脚本（不含页数目标审核）")
    p.add_argument("document", type=Path, help="待审 Word 文件路径")
    p.add_argument("--rules", type=Path, default=None, help="项目规则 JSON（可选）")
    p.add_argument("--report", type=Path, default=None, help="审查报告 JSON 输出路径（可选）")
    p.add_argument("--evidence-dir", type=Path, default=None, help="彩图留白异常页证据输出目录（可选）")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    report = args.report or args.document.with_name(args.document.stem + "_audit_report.json")
    try:
        return audit(args.document, args.rules, report, args.evidence_dir)
    except AuditError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 3
    except Exception as e:
        print(f"ERROR: 未预期错误：{e}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
