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
import importlib.util
import json
import re
import sys
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


# The batch workflow ships audit/common with this script. Use the same font
# resolution and Word checks as the regular workflow.
common_dir = Path(__file__).resolve().parents[1] / "common"
sys.path.insert(0, str(common_dir))
spec = importlib.util.spec_from_file_location(
    "_batch_common_word_audit", common_dir / "Word正式交付硬审（不含页数）.py")
common_audit = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = common_audit
spec.loader.exec_module(common_audit)


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
    "photo_delivery_stage": "draft",  # draft=允许占位；final=未替换的占位须阻断
    "font_rules": None,
    "low_content_page_char_threshold": 30,
    # 以下为通用版式待审阈值，不是招标页数参数。1/3 仅用于触发图片迁移留白复核。
    "bottom_whitespace_review_ratio": 1 / 3,  # 高置信度图片迁移 FAIL 的起始判断线
    "bottom_whitespace_optimization_ratio": 0.20,  # 更早触发人工优化 WARN，非招标硬要求
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


def load_rules(path: Optional[Path], page_dir=None) -> Dict[str, Any]:
    rules = json.loads(json.dumps(DEFAULT_RULES, ensure_ascii=False))
    if path:
        with path.open("r", encoding="utf-8") as f:
            incoming = json.load(f)
        if not isinstance(incoming, dict):
            raise AuditError("规则文件顶层必须是 JSON 对象。")
        rules.update(incoming)
    try:
        rules.update(common_audit.effective_font_settings(rules))
    except common_audit.AuditError as exc:
        raise AuditError(str(exc)) from exc
    # Location only: page target remains exclusively in the project page script.
    import ast
    here=Path(page_dir or Path(__file__).resolve().parent)
    if not here.is_dir():
        raise AuditError('页数脚本目录不存在')
    candidates=[]
    for p in here.glob('*页数硬审*.py'):
        values={}
        for node in ast.parse(p.read_text(encoding='utf-8-sig')).body:
            if isinstance(node,ast.Assign) and len(node.targets)==1 and isinstance(node.targets[0],ast.Name) and node.targets[0].id in ('EXPECTED_BODY_PAGES','PAGE_TOLERANCE','FRONT_MATTER_PAGES'):
                values[node.targets[0].id]=ast.literal_eval(node.value)
        if 'EXPECTED_BODY_PAGES' in values:
            candidates.append(values)
    if len(candidates)>1:
        raise AuditError('页数来源不唯一')
    if candidates:
        front=candidates[0].get('FRONT_MATTER_PAGES',0)
        if type(front) is not int or front<0:
            raise AuditError('封面目录排除页数必须为非负整数')
        if path and 'body_start_page' in incoming and incoming['body_start_page']!=front+1:
            raise AuditError('格式规则正文起始页与当前页数脚本不一致，先解决冲突')
        rules['body_start_page']=front+1
        rules['front_page']=1
    elif page_dir is not None:
        raise AuditError('指定目录未找到项目页数脚本')
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


_physical_page_starts = {}

def get_page_range(doc, page_num: int, page_count: int):
    # GoTo(Page, Absolute, Count) can follow restarted display numbering.
    # Audit ranges must use physical document pages, verified with Information(3).
    content_end = int(doc.Content.End)
    key = (str(doc.FullName), content_end, page_count)
    starts = _physical_page_starts.setdefault(key, {1: 0, page_count + 1: content_end})
    def physical_start(target):
        if target in starts:
            return starts[target]
        candidate = int(doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=target).Start)
        at_page = page_of_range(doc.Range(candidate, candidate))
        before_page = page_of_range(doc.Range(max(0, candidate - 1), max(0, candidate - 1)))
        if at_page == target and before_page < target:
            starts[target] = candidate
            return candidate
        low, high = 0, content_end
        while low < high:
            middle = (low + high) // 2
            actual = page_of_range(doc.Range(middle, middle))
            if actual < 1:
                raise AuditError('无法取得物理页码，不能构造页面审核范围')
            if actual < target:
                low = middle + 1
            else:
                high = middle
        if page_of_range(doc.Range(low, low)) != target:
            raise AuditError(f'无法定位物理第{target}页')
        starts[target] = low
        return low
    return doc.Range(physical_start(page_num), physical_start(page_num + 1))


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

    if body_start<=front:
        out.append(CheckResult("前置封面目录", "SKIP", "当前项目页数脚本配置无前置扣除页；需按项目要求复核。"))
        return out
    front_text = "\n".join(clean_text(get_page_range(doc, n, page_count).Text) for n in range(front,body_start))
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
            "首页未检测到有效封面信息，或一级目录从首页首项直接开始；前置扣除页范围内须包含封面及一级目录。",
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
        "font_rules": "项目执行字体字号规则",
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
        # Only the first cover/contents-to-body boundary is permitted.
        # All later same-layout forced heading sections retain the hard failure.
        if (idx == 2 and rules.get("allow_front_matter_break", False)
                and body_start_page > int(rules["front_page"])
                and first["page"] == body_start_page):
            out.append(CheckResult("前置页与正文分隔", "PASS",
                "按既定规则允许封面目录与正文之间唯一的新页分节；不放宽正文内部标题分页检查。", detail))
            continue
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
    """读取 Word 实际排版后某字符位置相对页面顶端的纵坐标（pt）。"""
    try:
        end_limit = max(1, int(doc.Content.End))
        pos = max(0, min(int(pos), end_limit - 1))
        r = doc.Range(pos, min(pos + 1, end_limit))
        y = _safe_float(r.Information(WD_VERTICAL_POSITION_RELATIVE_TO_PAGE))
        return y if y is not None and y >= 0 else None
    except Exception:
        return None


def _page_body_box(page_range) -> Optional[Dict[str, float]]:
    """返回页面正文可用区域几何；排除页眉、页脚和页边距。"""
    try:
        sec = page_range.Sections(1)
        ps = sec.PageSetup
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
            "page_height_pt": page_height,
            "body_top_pt": body_top,
            "body_bottom_pt": body_bottom,
            "usable_height_pt": usable,
        }
    except Exception:
        return None


def _paragraph_visual_bottom(doc, p, page_start: int, page_end: int) -> Optional[float]:
    try:
        start = max(page_start, int(p.Range.Start))
        end = min(page_end - 1, max(start, int(p.Range.End) - 1))
        if end < start:
            return None
        seg = doc.Range(start, min(end + 1, int(doc.Content.End)))
        if not re.sub(r"\s+", "", clean_text(seg.Text)):
            return None
        y = _range_vertical_y(doc, end)
        if y is None:
            return None
        try:
            line = _safe_float(p.Format.LineSpacing)
        except Exception:
            line = None
        # 只用于估算最后一行实际占用的底边；避免把基线当成内容底部。
        line = line if line is not None and 8 <= line <= 80 else 18.0
        return y + min(max(line * 0.85, 10.0), 36.0)
    except Exception:
        return None


def _visuals_on_page(doc, page_range, body_box: Optional[Dict[str, float]]) -> List[Dict[str, Any]]:
    """收集实际落在页面范围内的 InlineShape / Shape 几何。生成图通常为 InlineShape。"""
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
            name = None
            for attr in ("AlternativeText", "Title"):
                try:
                    val = str(getattr(shp, attr) or "").strip()
                    if val:
                        name = val
                        break
                except Exception:
                    pass
            items.append({
                "kind": "InlineShape",
                "anchor_start": int(rng.Start),
                "top_pt": round(top, 2) if top is not None else None,
                "height_pt": round(height, 2) if height is not None else None,
                "width_pt": round(width, 2) if width is not None else None,
                "bottom_pt": round(top + height, 2) if top is not None and height is not None else None,
                "label": name,
            })
        except Exception:
            continue

    # 一格边框表格形式的真实照片占位框也属于图文排布对象。
    # 不把普通数据表一概当成图，避免误报图片迁移。
    try:
        tcount = int(page_range.Tables.Count)
    except Exception:
        tcount = 0
    for i in range(1, tcount + 1):
        try:
            table = page_range.Tables(i)
            tr = table.Range
            label = clean_text(tr.Text)
            if "实景图占位" not in label and "需插入现场照片" not in label:
                continue
            start = int(tr.Start)
            if not (page_start <= start < page_end):
                continue
            top = _range_vertical_y(doc, start)
            bottom = _range_vertical_y(doc, max(start, int(tr.End) - 2))
            items.append({
                "kind": "PhotoPlaceholderTable",
                "anchor_start": start,
                "top_pt": round(top, 2) if top is not None else None,
                "bottom_pt": round(bottom, 2) if bottom is not None and top is not None and bottom >= top else None,
                "label": label[:100],
            })
        except Exception:
            continue

    # 浮动 Shape：以 Anchor 是否位于当前页为初筛，再按相对定位方式估算实际纵坐标.
    try:
        shape_count = int(doc.Shapes.Count)
    except Exception:
        shape_count = 0
    for i in range(1, shape_count + 1):
        try:
            shp = doc.Shapes(i)
            anchor = shp.Anchor
            a = int(anchor.Start)
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
                if rel == 1:  # wdRelativeVerticalPositionPage
                    top = raw_top
                elif rel == 0 and body_box is not None:  # margin
                    top = body_box["body_top_pt"] + raw_top
                elif anchor_y is not None:
                    top = anchor_y + raw_top
            label = None
            for attr in ("AlternativeText", "Title", "Name"):
                try:
                    val = str(getattr(shp, attr) or "").strip()
                    if val:
                        label = val
                        break
                except Exception:
                    pass
            items.append({
                "kind": "Shape",
                "anchor_start": a,
                "top_pt": round(top, 2) if top is not None else None,
                "height_pt": round(height, 2) if height is not None else None,
                "width_pt": round(width, 2) if width is not None else None,
                "bottom_pt": round(top + height, 2) if top is not None and height is not None else None,
                "relative_vertical_position": rel,
                "label": label,
            })
        except Exception:
            continue

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
            b = _paragraph_visual_bottom(doc, page_range.Paragraphs(i), page_start, page_end)
            if b is not None:
                bottoms.append(float(b))
        except Exception:
            continue
    for item in visuals:
        b = _safe_float(item.get("bottom_pt"))
        if b is not None:
            bottoms.append(b)
    return max(bottoms) if bottoms else None


def _body_page_number(document_page: int, body_start_page: int) -> Optional[int]:
    return document_page - body_start_page + 1 if document_page >= body_start_page else None


def check_image_migration_whitespace(doc, page_count: int, rules: Dict[str, Any]) -> Tuple[List[CheckResult], List[int]]:
    """
    重点识别：上一页出现大块页底留白，而下一页靠近页首出现图片。

    1/3 仅为高置信度迁移异常的初始判断线；0.20 用于更早的人工优化提示。
    图片前文字很少”等组合证据时，才将其视为高置信度图片迁移留白并 FAIL。
    单纯大留白只给 WARN，避免误伤文末、封面目录或其他正常布局。
    """
    body_start = int(rules.get("body_start_page", 2))
    blank_ratio_threshold = float(rules.get("bottom_whitespace_review_ratio", 1 / 3))
    optimize_ratio = float(rules.get("bottom_whitespace_optimization_ratio", 0.20))
    if not 0 < optimize_ratio <= blank_ratio_threshold < 1:
        raise AuditError("留白阈值无效：优化提示阈值必须大于0且不高于迁移硬审阈值")
    next_visual_top_ratio = float(rules.get("next_page_visual_top_ratio", 0.30))
    pre_visual_char_limit = int(rules.get("next_page_pre_visual_char_limit", 140))

    confirmed: List[Dict[str, Any]] = []
    review: List[Dict[str, Any]] = []
    incomplete: List[Dict[str, Any]] = []
    evidence_pages: set[int] = set()
    saw_visual = False

    # 仅检查正文页与其下一页；首页、最后一页不作为该模式的前页。
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
            incomplete.append({
                "document_page": p,
                "body_page": _body_page_number(p, body_start),
                "reason": "无法读取页面正文区域几何",
            })
            continue

        content_bottom = _page_content_bottom(doc, prev_r, prev_visuals)
        if content_bottom is None:
            incomplete.append({
                "document_page": p,
                "body_page": _body_page_number(p, body_start),
                "reason": "无法取得该页最后实际内容的纵向位置",
            })
            continue

        body_bottom = prev_box["body_bottom_pt"]
        usable = prev_box["usable_height_pt"]
        blank_height = max(0.0, body_bottom - content_bottom)
        blank_ratio = blank_height / usable if usable > 0 else 0.0
        if blank_ratio < optimize_ratio:
            continue

        geometrical_visuals = [x for x in next_visuals if _safe_float(x.get("top_pt")) is not None]
        if not geometrical_visuals:
            incomplete.append({
                "document_page": p + 1,
                "body_page": _body_page_number(p + 1, body_start),
                "reason": "检测到下一页图片，但无法取得图片实际纵坐标",
            })
            continue

        first_visual = min(geometrical_visuals, key=lambda x: float(x["top_pt"]))
        visual_top = float(first_visual["top_pt"])
        top_ratio = (visual_top - next_box["body_top_pt"]) / next_box["usable_height_pt"]
        top_ratio = max(0.0, top_ratio)

        anchor_start = int(first_visual.get("anchor_start") or int(next_r.Start))
        pre_end = min(max(anchor_start, int(next_r.Start)), int(next_r.End))
        try:
            pre_text = clean_text(doc.Range(int(next_r.Start), pre_end).Text)
            pre_chars = len(re.sub(r"\s+", "", pre_text))
        except Exception:
            pre_chars = None

        detail = {
            "previous_document_page": p,
            "previous_body_page": _body_page_number(p, body_start),
            "next_document_page": p + 1,
            "next_body_page": _body_page_number(p + 1, body_start),
            "page_numbering_note": "document_page=Microsoft Word文档总页码；body_page=从正文起始页计的页码",
            "bottom_blank_height_pt": round(blank_height, 2),
            "bottom_blank_ratio": round(blank_ratio, 4),
            "review_threshold_ratio": round(blank_ratio_threshold, 4),
            "optimization_threshold_ratio": round(optimize_ratio, 4),
            "next_visual_top_pt": round(visual_top, 2),
            "next_visual_top_ratio_of_body": round(top_ratio, 4),
            "pre_visual_char_count": pre_chars,
            "visual": first_visual,
            "trigger": "上一页页底大留白 + 下一页存在图片",
        }

        near_top = top_ratio <= next_visual_top_ratio
        little_text_before = pre_chars is not None and pre_chars <= pre_visual_char_limit
        if near_top and little_text_before and blank_ratio >= blank_ratio_threshold:
            detail["classification"] = "CONFIRMED_IMAGE_MIGRATION_RISK"
            detail["reason"] = "大块页底留白后，下一页靠近正文顶部即出现图片，且图片前文字很少；高置信度符合图片整体被推至下一页的版式问题。"
            confirmed.append(detail)
        else:
            detail["classification"] = "OPTIMIZATION_REVIEW" if blank_ratio < blank_ratio_threshold else "REVIEW_REQUIRED"
            detail["reason"] = "存在可优化的图前页底留白，尚未达到高置信度迁移型硬审条件；需看 Word 相邻页决定是否调整。"
            review.append(detail)
        evidence_pages.update((p, p + 1))

    checks: List[CheckResult] = []
    checks.append(CheckResult(
        "彩图前页大块留白_图片迁移",
        "FAIL" if confirmed else "PASS",
        "发现高置信度‘图片整体移到下一页并在前一页留下大块页底留白’问题；必须调整图片尺寸、插入位置或相关段落后重新 Word 渲染与复验。" if confirmed else "未发现高置信度图片迁移型大块页底留白。",
        confirmed or None,
    ))
    checks.append(CheckResult(
        "大块页底留白_待复核",
        "WARN" if review else "PASS",
        "检测到可进一步压缩的图前页底留白（默认20%起提示），请按页码结合Word渲染结果优化；这不是新的招标硬性阈值。" if review else "未发现需要人工优化的图前页底留白。",
        review or None,
    ))
    checks.append(CheckResult(
        "彩图留白检测完整性",
        "WARN" if incomplete else "PASS",
        "部分包含图片的页面无法取得足够实际几何信息，留白检测未完成；不得把该项解释为已检查通过。" if incomplete else ("已使用 Microsoft Word 实际分页与页面位置数据完成彩图相邻留白检查。" if saw_visual else "正文未检测到可分析图片对象；无图片迁移型留白可检查。"),
        incomplete or None,
    ))
    return checks, sorted(evidence_pages)


def export_whitespace_evidence(doc, pages: List[int], report_path: Path, evidence_dir: Optional[Path]) -> CheckResult:
    """对异常页生成 Word 版式证据：优先 PNG，其次 EMF；始终尝试导出 Word PDF 供复核。"""
    if not pages:
        return CheckResult("彩图留白证据", "PASS", "未触发彩图相邻留白异常，无需生成相邻页证据截图。")

    target = evidence_dir or report_path.with_name(report_path.stem + "_evidence")
    target.mkdir(parents=True, exist_ok=True)
    pdf_path = target / "word_layout_evidence.pdf"
    errors: List[str] = []
    png_files: List[str] = []
    emf_files: List[str] = []

    try:
        doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17)  # wdExportFormatPDF
    except Exception as e:
        errors.append(f"Word PDF导出失败：{e}")

    # 若环境已有 PyMuPDF，则从“Microsoft Word 导出的 PDF”精确截取异常页为 PNG。
    if pdf_path.exists():
        try:
            import fitz  # type: ignore
            pdf = fitz.open(str(pdf_path))
            try:
                for p in pages:
                    if not (1 <= p <= len(pdf)):
                        continue
                    pix = pdf[p - 1].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                    fn = target / f"word_page_{p:03d}.png"
                    pix.save(str(fn))
                    png_files.append(str(fn))
            finally:
                pdf.close()
        except Exception as e:
            errors.append(f"PNG截图不可用：{e}")

    # 不依赖第三方库的 Word 原生页面图像备份（EMF）。不同 Word 版本可能不支持该接口。
    if not png_files:
        try:
            window = doc.ActiveWindow
            try:
                window.View.Type = 3  # wdPrintView
            except Exception:
                pass
            doc.Repaginate()
            page_collection = window.Panes(1).Pages
            for p in pages:
                try:
                    bits = page_collection.Item(int(p)).EnhMetaFileBits
                    raw = bytes(bytearray(bits))
                    fn = target / f"word_page_{p:03d}.emf"
                    fn.write_bytes(raw)
                    emf_files.append(str(fn))
                except Exception as e:
                    errors.append(f"第{p}页EMF截图失败：{e}")
        except Exception as e:
            errors.append(f"Word原生页面截图接口不可用：{e}")

    manifest = {
        "layout_engine": "Microsoft Word",
        "document_pages": pages,
        "word_pdf": str(pdf_path) if pdf_path.exists() else None,
        "png_screenshots": png_files,
        "emf_screenshots": emf_files,
        "errors": errors,
        "note": "Word相邻页证据；20%为可优化留白提示，1/3为图片迁移异常起始判断线，均非招标硬要求。",
    }
    (target / "evidence_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    if png_files or emf_files:
        return CheckResult("彩图留白证据", "PASS", "已生成异常页及相邻图片页的 Microsoft Word 页面证据图，并保留 Word 导出 PDF。", manifest)
    if pdf_path.exists():
        return CheckResult("彩图留白证据", "WARN", "已生成 Microsoft Word 版式 PDF，但当前运行环境未生成逐页 PNG/EMF 截图；请按报告页码在 PDF 中复核。", manifest)
    return CheckResult("彩图留白证据", "WARN", "异常页证据导出未完成；已在报告中保留页码与几何触发依据，不得把证据项解释为已完成截图复核。", manifest)


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


def has_object_near_range(doc, rng, window: int = 600, include_tables: bool = True) -> bool:
    start = max(0, int(rng.Start) - window)
    end = min(int(doc.Content.End), int(rng.End) + window)
    around = doc.Range(start, end)
    try:
        if int(around.InlineShapes.Count) > 0 or (include_tables and int(around.Tables.Count) > 0):
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


def _placeholder_has_visible_frame(doc, placeholder: str) -> bool:
    """识别表格或文本形状内的占位内容；纯正文文字不冒充占位框。"""
    try:
        for i in range(1, int(doc.Tables.Count) + 1):
            if placeholder in str(doc.Tables(i).Range.Text):
                return True
    except Exception:
        pass
    try:
        for i in range(1, int(doc.Shapes.Count) + 1):
            shape = doc.Shapes(i)
            if placeholder in str(shape.TextFrame.TextRange.Text):
                return True
    except Exception:
        pass
    return False


def check_real_photo_slots(doc, rules: Dict[str, Any]) -> List[CheckResult]:
    """草稿允许标注真实照片占位；最终版必须有真实照片的来源记录并在图注旁确有图片。"""
    raw = rules.get("required_real_photo_slots")
    if raw is None:
        return [CheckResult("实景图占位", "FAIL", "尚未确认本项目的真实照片需求，禁止将格式审查标记为PASS。")]
    if not isinstance(raw, list):
        return [CheckResult("实景图占位", "FAIL", "required_real_photo_slots 必须是数组。")]
    stage = str(rules.get("photo_delivery_stage", "draft")).strip().lower()
    if stage not in ("draft", "final"):
        return [CheckResult("实景图交付阶段", "FAIL", "photo_delivery_stage 只能为 draft 或 final。")]
    full = doc_full_text(doc)
    slots_found = RE_REAL_PHOTO_PLACEHOLDER.findall(full)
    out: List[CheckResult] = []
    if stage == "final" and slots_found:
        out.append(CheckResult("正式稿_遗留实景图占位", "FAIL", "正式提交稿仍存在未替换的实景图占位；须补齐真实照片，不得自行删框规避。", slots_found))
    elif stage == "draft":
        out.append(CheckResult("草稿_实景图占位允许", "PASS", "当前为占位编制阶段；照片须在正式投标前替换；不得自行去除占位。"))

    for item in raw:
        if not isinstance(item, dict):
            out.append(CheckResult("实景图配置项", "FAIL", "实景图需求条目必须是对象。"))
            continue
        name = str(item.get("name", "")).strip()
        if not name:
            out.append(CheckResult("实景图配置项", "FAIL", "实景图需求缺少 name。"))
            continue
        keywords = [str(k) for k in item.get("keywords", []) if str(k).strip()]
        placeholders = [x for x in slots_found if name in x or any(k in x for k in keywords)]
        if stage == "draft":
            # 尚未有现场照片，可通过明确的占位框进行版面预算。
            if placeholders:
                framed = any(_placeholder_has_visible_frame(doc, text) for text in placeholders)
                out.append(CheckResult(
                    f"实景图占位_{name}", "PASS" if framed else "FAIL",
                    "检测到带框的照片占位内容，图号另作连续性校验。" if framed else "仅找到照片占位文字，未找到表格/形状占位框，不能视为已预留图框。",
                    {"placeholder": placeholders[0], "framed": framed},
                ))
            elif item.get("replaced") and item.get("source_record"):
                out.append(CheckResult(f"实景图占位_{name}", "PASS", "该处已经有来源记录的实景图，可跳过占位。"))
            else:
                out.append(CheckResult(f"实景图占位_{name}", "FAIL", "草稿缺少约定的实景图占位或来源明确的实景照片。"))
            continue

        # 已进入第二阶段的正式稿：旧占位文本不能直接冒充实景照片。
        if placeholders:
            continue  # 上方已有全局 FAIL，避免重复报告。
        caption = find_text_range(doc, name)
        has_photo = bool(caption and has_object_near_range(doc, caption, window=450, include_tables=False))
        source_record = str(item.get("source_record", "")).strip()
        if has_photo and source_record and bool(item.get("replaced")):
            out.append(CheckResult(f"正式实景图_{name}", "PASS", "检测到图名附近的图形对象及用户提供的素材来源记录；真实拍摄来源仍需人工核验。"))
        else:
            out.append(CheckResult(f"正式实景图_{name}", "FAIL", "正式稿缺少可定位图片、实景图替换标记或真实素材来源记录。",
                                   {"nearby_image": has_photo, "source_record_present": bool(source_record)}))
    out.append(CheckResult("实景图_AI真实性", "WARN", "脚本不能从 DOCX 自动证明图片真实拍摄或验证来源，仍须核对原始照片/来源记录。"))
    return out


def check_figure_caption_numbering(records: List[Dict[str, Any]], body_start: int) -> CheckResult:
    """检查图题显式显示的图1、图2……是否随正文连续、唯一、非图X。"""
    captions = []
    faults = []
    for rec in records:
        if int(rec.get("page") or 0) < body_start:
            continue
        content = str(rec.get("text") or "").strip()
        if re.match(r"^图\s*[XxＸ]\b", content):
            faults.append({"page": rec.get("page"), "text": content[:80], "reason": "存在未确定编号图X"})
            continue
        m = re.match(r"^图\s*(\d+)\s*[ 　:：．.、-]+\s*(.+)$", content)
        if not m:
            continue
        captions.append({"number": int(m.group(1)), "text": content[:100], "page": rec.get("page")})
    for i, cap in enumerate(captions, 1):
        if cap["number"] != i:
            faults.append({"page": cap["page"], "text": cap["text"], "expected": i, "actual": cap["number"]})
    if not captions:
        return CheckResult("图号连续性", "WARN", "未识别到采用‘图1 图名’格式的图题，请核查 Word SEQ 域和实际图题编号。")
    return CheckResult("图号连续性", "FAIL" if faults else "PASS", "图号存在跳号、重号或未定数字。" if faults else f"识别到{len(captions)}条图题且编号连续。", faults or {"count": len(captions)})


def check_fonts(doc, records: List[Dict[str, Any]], rules: Dict[str, Any]) -> CheckResult:
    checked = common_audit.check_fonts(doc, records, rules)
    return CheckResult(checked.name, checked.status, checked.message, checked.details)


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


def audit(docx_path: Path, rules_path: Optional[Path], report_path: Path, evidence_dir: Optional[Path] = None, page_dir=None) -> int:
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

    rules = load_rules(rules_path,page_dir)
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
        results.extend(portable_checks(docx_path, rules))
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
        results.append(check_figure_caption_numbering(records, int(rules["body_start_page"])))
        results.extend(check_real_photo_slots(doc, rules))
        results.append(check_fonts(doc, records, rules))

        hard_failures = [r for r in results if r.status == "FAIL"]
        overall = "PASS" if not hard_failures else "FAIL"

        report = {
            "audit_version": "1.6.0",
            "rules_sha256": rules_digest(rules),
            "photo_delivery_stage": rules.get("photo_delivery_stage"),
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
                "彩图留白检测基于 Microsoft Word 实际分页后的页面几何；默认20%起提示可优化的图前页底留白，1/3作为高置信度迁移问题的初始判断线，均非招标硬性要求。",
                "正式稿 real photo 检查依靠占位残留、图旁对象及来源记录；机器不能证明照片来源真实性。",
                "高置信度‘上一页大留白 + 下一页靠近页首出现图片’判为 FAIL；仅有大留白但证据不足时为 WARN，需结合相邻页证据复核。",
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


# Portable checks inspect the embedded DOCX bytes, not an external image filename.
FIGURE_NAMES = ('施工功能分区示意图', '施工进度横道图', '施工进度网络图')
EDITORIAL_RE = re.compile(r'本[章节段](?:应|需|可)(?:写|补充|描述|引用)|(?:后续|此处|这里)(?:补写|扩写|插入正文)|得分点|(?:引用|调用|根据)知识库|Skill|SKILL|(?:编制|审核|硬审|页数|门禁)脚本|待完善正文|编制时应|本段可用')


def rules_digest(rules):
    return hashlib.sha256(json.dumps(rules, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def _docx_blocks(path):
    import zipfile
    import xml.etree.ElementTree as ET
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    blocks = []
    with zipfile.ZipFile(path) as z:
        rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
        targets = {r.get('Id'): r.get('Target') for r in rels if r.get('TargetMode') != 'External'}
        body = ET.fromstring(z.read('word/document.xml')).find('w:body', ns)
        for el in body:
            if el.tag.rsplit('}', 1)[-1] not in ('p', 'tbl'):
                continue
            text = ''.join(x.text or '' for x in el.findall('.//w:t', ns))
            images = []
            for blip in el.findall('.//a:blip', ns):
                rid = blip.get('{'+ns['r']+'}embed')
                target = targets.get(rid)
                if target:
                    import posixpath
                    member = posixpath.normpath('word/'+target) if not target.startswith('/') else target.lstrip('/')
                    images.append({'member': member, 'bytes': z.read(member)})
            borders = el.findall('.//w:tblBorders/*', ns) + el.findall('.//w:tcBorders/*', ns)
            framed = any(b.get('{'+ns['w']+'}val') not in (None, 'nil', 'none') for b in borders)
            heights = [float(x.get('{'+ns['w']+'}val', '0'))/20 for x in el.findall('.//w:trHeight', ns)]
            blocks.append({'text': text, 'images': images, 'framed': framed, 'height_pt': max(heights, default=0)})
    return blocks


def _caption_images(blocks, index):
    own = blocks[index]['images']
    if own:
        return own
    # Captions must be adjacent to their own drawing, not merely somewhere near any object.
    for step in (-1, 1):
        for offset in (1, 2):
            j = index + step*offset
            if not 0 <= j < len(blocks):
                break
            b = blocks[j]
            if re.match(r'^图\s*\d+\s', b['text']) or RE_L1.match(b['text']) or RE_L2.match(b['text']):
                break
            if b['images']:
                return b['images']
            if b['text'].strip():
                break
    return []


def _color_metrics(raw):
    from PIL import Image
    from io import BytesIO
    im = Image.open(BytesIO(raw)).convert('RGBA')
    width, height = im.size
    im.thumbnail((600, 900))
    bg = Image.new('RGBA', im.size, 'white')
    bg.alpha_composite(im)
    pixels = list(bg.convert('RGB').getdata())
    colored = 0
    foreground = 0
    buckets = {}
    for r, g, b in pixels:
        if min(r, g, b) < 235:
            foreground += 1
        if max(r, g, b)-min(r, g, b) >= 32 and min(r, g, b) < 230:
            import colorsys
            hue = int(colorsys.rgb_to_hsv(r/255, g/255, b/255)[0]*12)%12
            buckets[hue] = buckets.get(hue, 0)+1
            colored += 1
    substantive = [k for k,v in buckets.items() if v >= max(30, len(pixels)*.002)]
    return {'width': width, 'height': height, 'color_ratio': colored/max(1,len(pixels)),
            'foreground_color_ratio': colored/max(1,foreground), 'hue_groups': len(substantive)}


def portable_checks(document, rules):
    blocks = _docx_blocks(document)
    out = []
    configured = rules.get('required_figures')
    if not isinstance(configured, list):
        configured = []
    caption_ids = []
    used = set()
    for name in FIGURE_NAMES:
        item = next((x for x in configured if x.get('name') == name), None)
        if not item or item.get('hard') is not True or item.get('color_required') is not True:
            out.append(CheckResult('彩色图规则_'+name, 'FAIL', '必备图必须配置hard=true及color_required=true。'))
        indices = [i for i,b in enumerate(blocks) if re.match(r'^图\s*\d+\s', b['text']) and name in b['text']]
        if len(indices) != 1:
            out.append(CheckResult('嵌入彩图_'+name, 'FAIL', '须有唯一的正式图题和实际嵌入图。', {'captions': len(indices)}))
            continue
        pics = _caption_images(blocks, indices[0])
        if len(pics) != 1:
            out.append(CheckResult('嵌入彩图_'+name, 'FAIL', '图题旁须有一张可定位的独立嵌入图片；复杂矢量对象先导出PNG再嵌入。'))
            continue
        pic = pics[0]
        digest = hashlib.sha256(pic['bytes']).hexdigest()
        if digest in used:
            out.append(CheckResult('三图独立_'+name, 'FAIL', '三种图不能共用同一图片。'))
        used.add(digest)
        try:
            metrics = _color_metrics(pic['bytes'])
            color_ok = metrics['color_ratio'] >= .01 and metrics['foreground_color_ratio'] >= .04 and metrics['hue_groups'] >= 2
            portrait_ok = metrics['width'] <= metrics['height']
            out.append(CheckResult('彩图配色_'+name, 'PASS' if color_ok else 'FAIL',
                '嵌入图片符合基础彩色比例及多色系检测；主体视觉仍须人工核对。' if color_ok else '黑灰图、彩色比例过低或色系不足；需重绘主体配色。', {**metrics,'image_sha256':digest}))
            out.append(CheckResult('图件竖向_'+name, 'PASS' if portrait_ok else 'FAIL', '图件须按竖向画布组织。', metrics))
        except Exception as e:
            out.append(CheckResult('彩图配色_'+name, 'FAIL', '无法检测嵌入图片颜色，不能当作通过：'+str(e)))
    for b in blocks:
        m = re.match(r'^图\s*(\d+)\s', b['text'])
        if m:
            caption_ids.append(int(m[1]))
    out.append(CheckResult('静态图号连续性', 'PASS' if caption_ids and caption_ids == list(range(1,len(caption_ids)+1)) else 'FAIL', '图题缓存须从图1连续显示。',caption_ids))
    photo_slots = rules.get('required_real_photo_slots')
    if not isinstance(photo_slots, list) or not photo_slots:
        out.append(CheckResult('实景图必有', 'FAIL', '实景图清单必须至少一项，缺素材必须占位。'))
    else:
        stage = rules.get('photo_delivery_stage')
        if stage not in ('draft','final'):
            out.append(CheckResult('照片交付阶段','FAIL','必须明确draft/final。'))
        if stage == 'final' and any('需插入现场照片' in b['text'] or RE_REAL_PHOTO_PLACEHOLDER.search(b['text']) for b in blocks):
            out.append(CheckResult('正式稿占位残留','FAIL','正式稿不得保留任何照片占位。'))
        for item in photo_slots:
            name = str(item.get('name','')).strip()
            if not name:
                out.append(CheckResult('照片配置','FAIL','照片主题名不能为空。'))
                continue
            placeholders = [b for b in blocks if ('【实景图占位：'+name+'】') in b['text']]
            if stage == 'draft' and placeholders:
                valid = len(placeholders)==1 and placeholders[0]['framed'] and placeholders[0]['height_pt'] >= 50 and '需插入现场照片' in placeholders[0]['text']
                out.append(CheckResult('照片占位_'+name,'PASS' if valid else 'FAIL','占位须唯一、有可见边框、至少50pt预留高度并标注需插入现场照片。'))
            else:
                indices = [i for i,b in enumerate(blocks) if re.match(r'^图\s*\d+\s',b['text']) and name in b['text']]
                pics = _caption_images(blocks,indices[0]) if len(indices)==1 else []
                expected_hash = item.get('image_sha256')
                actual_hashes = [hashlib.sha256(p['bytes']).hexdigest() for p in pics]
                valid = bool(item.get('replaced') is True and item.get('source_record') and item.get('source_verified') is True and expected_hash in actual_hashes)
                out.append(CheckResult('真实照片_'+name,'PASS' if valid else 'FAIL','实景图片须邻接图题、绑定嵌入图片哈希、有来源并完成真实性人工核验。'))
    editorial = [{'block': i+1,'text': b['text'][:180]} for i,b in enumerate(blocks) if EDITORIAL_RE.search(b['text'])]
    out.append(CheckResult('正文内部编制语言','FAIL' if editorial else 'PASS','发现内部编制说明，须改为施工执行语言。' if editorial else '未发现已知内部编制关键词；仍需逐段语义复核。',editorial))
    return out


def static_audit(document, rules_path, report_path,page_dir=None):
    rules = load_rules(rules_path,page_dir)
    results = portable_checks(document, rules)
    overall = 'FAIL' if any(x.status=='FAIL' for x in results) else 'PASS'
    report = {'audit_version':'1.6.0', 'audit_type':'embedded_docx_static', 'document_sha256':sha256_file(document),
              'rules_sha256':rules_digest(rules), 'photo_delivery_stage':rules.get('photo_delivery_stage'),
              'delivery_status':overall, 'checks':[asdict(x) for x in results],
              'notes':['静态核验不使用Word，不能证明实际页数、版式、照片真实性或技术参数正确。']}
    report_path.parent.mkdir(parents=True,exist_ok=True)
    report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print_console(results,overall,report_path)
    return 0 if overall=='PASS' else 2


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="通用 Word 正式交付硬审脚本（不含页数目标审核）")
    p.add_argument("document", type=Path, help="待审 Word 文件路径")
    p.add_argument("--page-dir",type=Path,default=None,help="当前唯一项目页数脚本目录，自动派生正文页定位")
    p.add_argument("--static-only", action="store_true", help="跨平台静态检测嵌入彩图/照片框/图号/内部语言，不代表Word版式验收")
    p.add_argument("--rules", type=Path, default=None, help="项目规则 JSON（可选）")
    p.add_argument("--report", type=Path, default=None, help="审查报告 JSON 输出路径（可选）")
    p.add_argument("--evidence-dir", type=Path, default=None, help="彩图留白异常页证据输出目录（可选；默认与报告同目录）")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    report = args.report or args.document.with_name(args.document.stem + "_audit_report.json")
    try:
        if args.static_only:
            return static_audit(args.document, args.rules, report,args.page_dir)
        return audit(args.document, args.rules, report, args.evidence_dir,args.page_dir)
    except AuditError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 3
    except Exception as e:
        print(f"ERROR: 未预期错误：{e}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
