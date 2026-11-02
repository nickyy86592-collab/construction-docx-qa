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
    "front_matter": {"cover_required": True, "toc_levels": [1]},
    "heading_patterns": None,  # 招标另定标题体系时，由AI提取1/2/3级正则；不改变大纲级别
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
    t = re.sub(r"(?:[\.·…．。]{2,}\s*|\s+)(?:\d+|[ivxlcdmIVXLCDM]+)\s*$", "", t)
    return t.strip()


def load_rules(path: Optional[Path], page_dir=None) -> Dict[str, Any]:
    rules = json.loads(json.dumps(DEFAULT_RULES, ensure_ascii=False))
    incoming = {}
    if path:
        with path.open("r", encoding="utf-8") as f:
            incoming = json.load(f)
        if not isinstance(incoming, dict):
            raise AuditError("规则文件顶层必须是 JSON 对象。")
        rules.update(incoming)
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
    front = rules['body_start_page'] - 1
    if type(rules['body_start_page']) is not int or front < 0:
        raise AuditError('正文起始页必须为正整数')
    fm = rules.get('front_matter')
    if not isinstance(fm, dict):
        raise AuditError('front_matter必须为对象')
    fm = {"cover_required": True, "toc_levels": [1], **fm}
    if type(fm['cover_required']) is not bool:
        raise AuditError('cover_required必须为布尔值')
    levels = fm['toc_levels']
    if not isinstance(levels, list) or any(type(v) is not int or v not in (1,2,3) for v in levels) or levels != sorted(set(levels)):
        raise AuditError('toc_levels须为递增且不重复的1至3级目录列表；[]表示有依据地无目录')
    if front == 0 and (fm['cover_required'] or levels):
        raise AuditError('无前置页时须按招标依据关闭封面及目录要求')
    for key in ('cover_page_range', 'toc_page_range'):
        span = fm.get(key, [1, front] if front else [])
        if span and not (isinstance(span,list) and len(span)==2 and all(type(v) is int for v in span) and 1<=span[0]<=span[1]<=front):
            raise AuditError(key+'须落在项目脚本提供的前置页范围')
        fm[key] = span
    keywords = fm.get('cover_keywords', [])
    if not isinstance(keywords,list) or any(not isinstance(v,str) or not v.strip() for v in keywords):
        raise AuditError('cover_keywords须为非空字符串数组')
    rules['front_matter'] = fm
    table_layout(rules)  # 同一运行期规则中校验表格配置
    toc_layout(rules)  # 校验运行期招标例外；不增加独立配置文件
    patterns = rules.get('heading_patterns')
    if patterns is not None:
        if not isinstance(patterns,dict) or set(patterns)!= {'1','2','3'}:
            raise AuditError('heading_patterns须明确1/2/3三级模式')
        for pattern in patterns.values():
            try:
                if re.compile(pattern).groups < 1:
                    raise AuditError('标题模式第一捕获组须为可解析编号')
            except (re.error, TypeError) as exc:
                raise AuditError('标题模式无效：'+str(exc)) from exc
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


def paragraph_text_range(doc, index):
    """Keep native TOC's first paragraph inside its visible result, not the outer field."""
    original=doc.Paragraphs(index).Range
    start,end=int(original.Start),int(original.End)
    try:
        fields=original.Fields
        for n in range(1,int(fields.Count)+1):
            field=fields(n)
            if int(field.Type)==13:  # wdFieldTOC
                start=max(start,int(field.Result.Start))
                break
    except Exception:pass
    try:
        following=int(doc.Paragraphs(index+1).Range.Start)
        if start<following<end:end=following
    except Exception:pass
    if start!=int(original.Start) or end!=int(original.End):return doc.Range(start,end)
    return original.Duplicate


def get_paragraph_records(doc) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    for i in range(1, int(doc.Paragraphs.Count) + 1):
        p = doc.Paragraphs(i)
        text_range = paragraph_text_range(doc, i)
        text = normalize_spaces(text_range.Text)
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
            "page": page_of_range(text_range),
            "outline": outline,
            "page_break_before": page_break_before,
            "start": int(text_range.Start),
            "end": int(text_range.End),
        }
        records.append(rec)
    return records


def classify_heading(text: str, rules=None) -> Tuple[int, Optional[int]]:
    patterns = (rules or {}).get('heading_patterns')
    regexes = [(n, re.compile(patterns[str(n)])) for n in (1,2,3)] if patterns else ((1, RE_L1), (2, RE_L2), (3, RE_L3))
    for level, regex in regexes:
        m = regex.match(text)
        if m:
            return level, chinese_num_to_int(m.group(1))
    return 0, None


def extract_heading_structure(records: List[Dict[str, Any]], body_start_page: int, rules=None) -> Tuple[List[Dict[str, Any]], List[CheckResult]]:
    results: List[CheckResult] = []
    headings: List[Dict[str, Any]] = []
    bad_outline: List[Dict[str, Any]] = []
    bad_syntax: List[Dict[str, Any]] = []

    for rec in records:
        if rec["page"] < body_start_page:
            continue
        textual_level, number = classify_heading(rec["text"], rules)
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

    fm = rules.get('front_matter') or {"cover_required": True, "toc_levels": [1]}
    prefix = [r for r in records if 1 <= r['page'] < body_start]
    polluted = [r for r in prefix if r['outline'] in (1,2,3)]
    out.append(CheckResult('前置页_导航隔离', 'FAIL' if polluted else 'PASS',
        '封面及目录不得使用正文标题大纲级别。', polluted or None))
    if body_start == 1:
        out.append(CheckResult('前置封面目录', 'PASS', '按当前项目规则，无独立前置页。'))
        return out
    def in_span(rec, key):
        span = fm.get(key) or [1,body_start-1]
        return span[0] <= rec['page'] <= span[1]
    toc_records = [r for r in prefix if in_span(r,'toc_page_range')]
    levels = fm.get('toc_levels',[1])
    actual, extra, trailing = [], [], []
    started = False
    for rec in toc_records:
        text = normalize_toc_title(rec['text']) if toc_layout(rules)['show_page_numbers'] else normalize_spaces(rec['text'])
        level, _ = classify_heading(text, rules)
        if level:
            started = True
            if level in levels:
                actual.append((level,text))
            else:
                extra.append(rec['text'])
        elif started and levels and not re.fullmatch(r'[\d\s.·…．。—-]+',text) and text not in ('目录','目 录'):
            trailing.append(rec['text'])
    expected = [(h['level'],normalize_spaces(h['text'])) for h in headings if h['level'] in levels]
    out.append(CheckResult('前置目录_与正文导航一致', 'PASS' if actual==expected and not extra else 'FAIL',
        '按项目目录层级比对正文标题名称、数量、顺序及Word导航。',
        {'toc':actual,'body':expected,'unexpected_levels':extra}))
    out.append(CheckResult('前置目录_正文混入', 'FAIL' if trailing else 'PASS',
        '目录项开始后不得混入普通正文。', trailing or None))
    if fm.get('cover_required',True):
        first_toc_start = min((r['start'] for r in toc_records if classify_heading(normalize_toc_title(r['text']),rules)[0]),default=float('inf'))
        cover = [r['text'] for r in prefix if in_span(r,'cover_page_range') and r['start'] < first_toc_start
                 and not classify_heading(normalize_toc_title(r['text']),rules)[0]
                 and re.sub(r'\s+','',r['text']) not in ('目录',)
                 and not re.fullmatch(r'[\d\s.·…．。_—-]+',r['text'])]
        keywords = fm.get('cover_keywords',[])
        missing = [k for k in keywords if not any(k in t for t in cover)]
        out.append(CheckResult('前置封面信息','PASS' if cover and not missing else 'FAIL',
            '按项目封面范围检查有效文字及明确要求的封面字段；具体语义由AI复核。',
            {'cover_lines':cover[:20],'missing_keywords':missing}))
    return out

# TOC format checks share the existing static/Word audit, never modify the document.
TOC_LAYOUT_DEFAULT = {'show_page_numbers': True, 'right_align': True, 'leader': 'dots', 'right_inset_pt': 0}
TOC_PAGE_RE = re.compile(r'(\d+|[ivxlcdmIVXLCDM]+)\s*$')


def toc_layout(rules):
    raw = (rules.get('front_matter') or {}).get('toc_layout') or {}
    if not isinstance(raw, dict):
        raise AuditError('toc_layout必须为对象')
    cfg = {**TOC_LAYOUT_DEFAULT, **raw}
    if any(type(cfg[k]) is not bool for k in ('show_page_numbers','right_align')) or cfg['leader'] not in ('dots','none','dashes','lines'):
        raise AuditError('目录页码/右对齐开关或引导符无效')
    inset = cfg['right_inset_pt']
    if isinstance(inset,bool) or not isinstance(inset,(float,int)) or not 0 <= inset < 1000:
        raise AuditError('目录右侧内缩必须为有限非负磅值')
    if any(cfg[k] != v for k,v in TOC_LAYOUT_DEFAULT.items()) and not str(cfg.get('source_ref','')).strip():
        raise AuditError('偏离默认目录格式须引用招标/有效补遗格式依据source_ref')
    return cfg


def check_toc_layout_static(document, rules):
    import zipfile
    import xml.etree.ElementTree as ET
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};W='{'+ns['w']+'}'
    cfg=toc_layout(rules); levels=(rules.get('front_matter') or {}).get('toc_levels',[1])
    if not levels:return [CheckResult('目录点线与右对齐_静态','PASS','当前有效规则无目录要求。')]
    failures=[];warnings=[];checked=0
    with zipfile.ZipFile(document) as z:
        body=ET.fromstring(z.read('word/document.xml')).find('w:body',ns)
        styles=ET.fromstring(z.read('word/styles.xml')) if 'word/styles.xml' in z.namelist() else ET.Element('styles')
        style_map={x.get(W+'styleId'):x for x in styles.findall('w:style',ns)}
        defaults=styles.find('w:docDefaults/w:pPrDefault/w:pPr',ns)
        def effective(p):
            props=[];seen=set();pr=p.find('w:pPr',ns);st=pr.find('w:pStyle',ns) if pr is not None else None
            sid=st.get(W+'val') if st is not None else next((key for key,node in style_map.items() if node.get(W+'type')=='paragraph' and node.get(W+'default') in ('1','true')),None)
            while sid and sid not in seen:
                seen.add(sid);node=style_map.get(sid)
                if node is None:break
                props.insert(0,node.find('w:pPr',ns));parent=node.find('w:basedOn',ns);sid=parent.get(W+'val') if parent is not None else None
            props=[defaults]+props+[pr];outline=10;tabs={}
            for item in props:
                if item is None:continue
                ol=item.find('w:outlineLvl',ns)
                if ol is not None:outline=int(ol.get(W+'val'))+1
                for tab in item.findall('w:tabs/w:tab',ns):
                    pos=int(tab.get(W+'pos','0'))
                    if tab.get(W+'val')=='clear':tabs.pop(pos,None)
                    else:tabs[pos]=(tab.get(W+'val'),{'dot':'dots','hyphen':'dashes','underscore':'lines'}.get(tab.get(W+'leader'),tab.get(W+'leader','none')))
            return outline,tabs
        paras=list(body.iter(W+'p'));all_text=lambda p:''.join(('\t' if x.tag==W+'tab' else x.text or '') for x in p.iter() if x.tag in (W+'t',W+'tab'))
        first_body=next((i for i,p in enumerate(paras) if effective(p)[0] in (1,2,3)),None)
        if first_body is None:
            return [CheckResult('目录点线与右对齐_静态','WARN','静态无法可靠定位正文标题边界，须用Word页位及AI渲染复核，不代表目录合格。')]
        # A section's sectPr belongs to the preceding content; use the next section end.
        sections=[(i,sec) for i,p in enumerate(paras) for sec in p.findall('w:pPr/w:sectPr',ns)]
        if body.find('w:sectPr',ns) is not None:sections.append((len(paras),body.find('w:sectPr',ns)))
        for i,p in enumerate(paras[:first_body]):
            text=all_text(p).strip();level,_=classify_heading(normalize_toc_title(text),rules)
            if level not in levels:continue
            checked+=1;detail={'paragraph':i+1,'text':text[:180]}
            if not cfg['show_page_numbers']:
                if re.search(r'\t\s*(?:\d+|[ivxlcdmIVXLCDM]+)\s*$',text):
                    failures.append({**detail,'reason':'有效项目规则要求不显示目录页码，实际仍有页码'})
                elif TOC_PAGE_RE.search(text):
                    warnings.append({**detail,'reason':'无页码目录末尾疑似页码，须核实是否属于标题'})
                continue
            m=TOC_PAGE_RE.search(text)
            if not m:
                warnings.append({**detail,'reason':'目录项页码无法可靠识别'});continue
            if '\t' not in text or not re.search(r'\t\s*'+re.escape(m[1])+r'\s*$',text):
                failures.append({**detail,'reason':'标题与页码之间缺真实制表符，不能用空格/手打点线替代'});continue
            if re.search(r'(?:\.{3,}|·{3,}|…{2,}|．{3,})\s*\t',text):
                failures.append({**detail,'reason':'页码制表符前仍有手工拼接点线'});continue
            _,tabs=effective(p);expected_alignment='right' if cfg['right_align'] else None
            right_tabs=[(pos,pair) for pos,pair in tabs.items() if pair[0]==expected_alignment] if expected_alignment else list(tabs.items())
            if not right_tabs:
                failures.append({**detail,'reason':'未配置有效右对齐制表位' if expected_alignment else '缺少实际制表位'});continue
            sec=next((s for end,s in sections if end>=i),None)
            sz=sec.find('w:pgSz',ns) if sec is not None else None;mar=sec.find('w:pgMar',ns) if sec is not None else None
            cols=sec.find('w:cols',ns) if sec is not None else None
            if sz is None or mar is None or (cols is not None and int(cols.get(W+'num','1'))>1):
                warnings.append({**detail,'reason':'版心/分栏信息不足，须Word实测制表位落点'});continue
            if int(mar.get(W+'gutter','0'))>0:
                warnings.append({**detail,'reason':'装订线布局须用Word实际版心核实制表位，静态不猜测装订边'});continue
            try:width=(int(sz.get(W+'w'))-int(mar.get(W+'left'))-int(mar.get(W+'right'))-int(mar.get(W+'gutter','0')))/20
            except (TypeError,ValueError):
                warnings.append({**detail,'reason':'版心尺寸不可可靠读取'});continue
            target=width-cfg['right_inset_pt']
            if target<=0:failures.append({**detail,'reason':'目录页码目标位置不在版心内'});continue
            matches=[(pos,pair) for pos,pair in right_tabs if not cfg['right_align'] or abs(pos/20-target)<=2]
            if not matches:
                failures.append({**detail,'reason':'页码右制表位未落在统一版心右边界','target_pt':target,'tabs':right_tabs});continue
            if not any(pair[1]==cfg['leader'] for _,pair in matches):
                failures.append({**detail,'reason':'目录引导符不符合有效项目规则','expected':cfg['leader'],'tabs':matches})
            if len(tabs)>1:
                warnings.append({**detail,'reason':'存在多个自定义制表位，须用Word核实末个页码实际落点'})
        if not checked:warnings.append({'reason':'前置区未静态识别到有效目录项，须Word定位；静态检查不证明页位或目录完整性'})
    return [CheckResult('目录点线与右对齐_静态','FAIL' if failures else 'PASS','检查真实制表符、继承后的制表位/引导符及版心右边界；不替代Word视觉复核。',failures or {'checked':checked}),
            CheckResult('目录点线与右对齐_待复核','WARN' if warnings else 'PASS','目录疑点须按当前文件逐项复核。' if warnings else '未发现静态目录格式疑点。',warnings or None)]


def check_toc_layout_word(doc, records, headings, rules):
    cfg=toc_layout(rules);fm=rules.get('front_matter') or {};levels=fm.get('toc_levels',[1]);body_start=rules['body_start_page']
    if not levels:return [CheckResult('目录页码对齐_Word','PASS','当前有效规则无目录要求。')]
    span=fm.get('toc_page_range') or [1,body_start-1]
    entries=[r for r in records if span[0]<=r['page']<=span[1] and classify_heading(normalize_toc_title(r['text']),rules)[0] in levels]
    expected=[h for h in headings if h['level'] in levels];failures=[];warnings=[];measured=[]
    try:doc.ActiveWindow.View.Type=3  # 仅切换视图测量，不修改或保存正文
    except Exception:pass
    for i,rec in enumerate(entries):
        detail={'paragraph':rec['index'],'page':rec['page'],'text':rec['text'][:160]}
        try:
            p=doc.Paragraphs(rec['index']);entry=paragraph_text_range(doc,rec['index']);raw=clean_text(entry.Text);m=TOC_PAGE_RE.search(raw)
            if not cfg['show_page_numbers']:
                if re.search(r'\t\s*(?:\d+|[ivxlcdmIVXLCDM]+)\s*$',raw):
                    failures.append({**detail,'reason':'有效项目规则要求不显示目录页码，实际仍有页码'})
                elif TOC_PAGE_RE.search(raw):
                    warnings.append({**detail,'reason':'无页码目录末尾疑似页码，须核实是否属于标题'})
                continue
            if not m:
                warnings.append({**detail,'reason':'目录页码无法可靠识别，须核对实际页码格式'});continue
            if not re.search(r'\t\s*'+re.escape(m[1])+r'\s*$',raw):
                failures.append({**detail,'reason':'页码前缺真实制表符'});continue
            token=m[1];tabs=[]
            for n in range(1,int(p.Format.TabStops.Count)+1):
                tab=p.Format.TabStops(n);tabs.append((float(tab.Position),int(tab.Alignment),int(tab.Leader)))
            ps=p.Range.Sections(1).PageSetup
            if int(ps.TextColumns.Count)>1:
                warnings.append({**detail,'reason':'多栏目录需按实际栏边界人工复核'});continue
            if float(ps.Gutter)>0 or bool(getattr(ps,'MirrorMargins',False)):
                warnings.append({**detail,'reason':'装订线或对称页边布局须结合实际页面核实右边界'});continue
            target=float(ps.PageWidth)-float(ps.LeftMargin)-float(ps.RightMargin)-float(ps.Gutter)-cfg['right_inset_pt']
            matched=[t for t in tabs if (not cfg['right_align'] or t[1]==2 and abs(t[0]-target)<=2)]
            leader={'none':0,'dots':1,'dashes':2,'lines':3}[cfg['leader']]
            if not matched or not any(t[2]==leader for t in matched):
                failures.append({**detail,'reason':'Word有效制表位、右对齐或引导符配置不符','tabs':tabs,'target_pt':target});continue
            if not cfg['right_align'] and len(matched)>1:
                warnings.append({**detail,'reason':'非右对齐例外存在多个制表位，须确认页码实际使用的引导符'})
            # Positional COM arguments keep backward/stop/format flags effective.
            number=entry.Duplicate
            if not number.Find.Execute(token,True,False,False,False,False,False,0,False):
                warnings.append({**detail,'reason':'Word无法定位页码字符'});continue
            try:doc.ActiveWindow.ScrollIntoView(number,True)
            except Exception:pass
            end=doc.Range(int(number.End),int(number.End));x=float(end.Information(5))
            if x<0:
                warnings.append({**detail,'reason':'Word未返回可靠页码横坐标'});continue
            right=float(ps.LeftMargin)+target
            if cfg['right_align'] and abs(x-right)>3:
                # Word 2019 can return an indent-shifted caret at a nested PAGEREF
                # boundary, even when the exported PDF is aligned. Do not hard-fail
                # that ambiguous coordinate; configuration errors above still fail.
                fields=getattr(p.Range,'Fields',None)
                field_count=int(fields.Count) if fields is not None else 0
                indent=float(getattr(p.Format,'LeftIndent',0))
                if field_count and abs(indent)>3 and abs(x+indent-right)<=3:
                    warnings.append({**detail,'reason':'原生目录嵌套域光标坐标疑似受左缩进影响，须按实际渲染核实右边界','actual_x_pt':x,'expected_x_pt':right,'left_indent_pt':indent})
                else:
                    failures.append({**detail,'reason':'页码实际右边界未对齐或超出版心','actual_x_pt':x,'expected_x_pt':right})
            before=doc.Range(int(entry.Start),int(number.Start));title=raw.rsplit('\t',1)[0].rstrip()
            if title and before.Find.Execute(title[-1],True,False,False,False,False,False,0,False):
                y1=float(before.Information(6));y2=float(number.Information(6))
                if page_of_range(before)!=page_of_range(number):
                    failures.append({**detail,'reason':'页码与标题末行分在不同页'})
                elif y1<0 or y2<0:
                    warnings.append({**detail,'reason':'标题末行/页码纵坐标不足，不能确认是否孤行'})
                elif abs(y1-y2)>max(6,min(40,float(p.Format.LineSpacing) if 8<=float(p.Format.LineSpacing)<=80 else 18)*.7):
                    failures.append({**detail,'reason':'页码单独掉行或与标题末行分离'})
            else:warnings.append({**detail,'reason':'无法可靠确认标题末行与页码是否同一行'})
            if i<len(expected) and token.isdigit():
                h=expected[i];hr=doc.Paragraphs(h['index']).Range
                if normalize_toc_title(rec['text'])==normalize_spaces(h['text']) and int(token)!=int(hr.Information(1)):
                    failures.append({**detail,'reason':'目录页码与当前Word正文实际显示页码不一致','toc_page':token,'actual_page':int(hr.Information(1))})
            elif not token.isdigit():warnings.append({**detail,'reason':'非阿拉伯页码须结合项目页码格式核实对应关系'})
            measured.append({**detail,'actual_right_x_pt':x,'expected_right_x_pt':right})
        except Exception as exc:
            warnings.append({**detail,'reason':'Word目录测量未完成：'+str(exc)})
    if not entries:failures.append({'reason':'当前要求有目录，Word前置范围内未定位到目录项'})
    return [CheckResult('目录页码对齐_Word','FAIL' if failures else 'PASS','按Word有效属性、实际坐标、标题末行及正文显示页码核验。',failures or measured),
            CheckResult('目录视觉与页码_待复核','WARN' if warnings else 'PASS','需结合Word实际渲染处理目录疑点；测量未完成不能当通过。' if warnings else ('已检查不显示页码的项目规则；仍须核对目录名称及实际渲染。' if not cfg['show_page_numbers'] else '目录实际测量完成，仍按既有流程查看渲染效果。'),warnings or None)]


# Data-table layout is checked here, using the existing static/Word audit stages.
TABLE_LAYOUT_DEFAULT = {'table_alignment':'center', 'header_alignment':'center',
                        'short_text_alignment':'center', 'long_text_alignment':'left',
                        'vertical_alignment':'center', 'zero_indents':True}


def table_layout(rules):
    raw=rules.get('table_layout',{})
    if not isinstance(raw,dict):raise AuditError('table_layout必须为对象')
    cfg={**TABLE_LAYOUT_DEFAULT,**raw}
    for key in ('table_alignment','header_alignment','short_text_alignment','long_text_alignment'):
        if cfg[key] not in ('left','center','right','both') or (key=='table_alignment' and cfg[key]=='both'):
            raise AuditError('表格对齐配置无效：'+key)
    if cfg['vertical_alignment'] not in ('top','center','bottom') or type(cfg['zero_indents']) is not bool:
        raise AuditError('表格垂直对齐/缩进配置无效')
    if any(cfg[k]!=v for k,v in TABLE_LAYOUT_DEFAULT.items()) and not str(cfg.get('source_ref','')).strip():
        raise AuditError('偏离默认表格格式须引用有效格式依据source_ref')
    entries=cfg.get('tables',[])
    if not isinstance(entries,list):raise AuditError('table_layout.tables必须为数组')
    seen=set()
    for item in entries:
        if not isinstance(item,dict) or type(item.get('index')) is not int or item['index']<1 or item['index'] in seen:
            raise AuditError('表格index须为唯一正整数，按主文档中非嵌套表格顺序计数')
        seen.add(item['index'])
        if item.get('kind','data') not in ('data','layout','figure','photo'):
            raise AuditError('表格kind无效')
        if item.get('kind','data')!='data' and not str(item.get('reason','')).strip():
            raise AuditError('排除非数据表须写明reason')
        roles=item.get('column_roles',[])
        if not isinstance(roles,list) or any(x not in ('short','long') for x in roles):
            raise AuditError('column_roles须为short/long数组')
        if type(item.get('header_rows',1)) is not int or item.get('header_rows',1)<0:
            raise AuditError('header_rows须为非负整数')
    return cfg


def _table_result(failures,warnings,checked,word=False):
    suffix='Word' if word else '静态'
    return [CheckResult('表格排版_'+suffix,'FAIL' if failures else 'PASS',
            '核对数据表整体、表头/列角色对齐、有效缩进及版心；不代替实际视觉复核。', failures or {'checked_tables':checked}),
            CheckResult('表格渲染_待复核' if word else '表格排版_待复核','WARN' if warnings else 'PASS',
            '逐项查看Word实际渲染，复核短名称、数字单位、自然换行及跨页衔接。' if warnings else '未发现可识别表格疑点，仍须按既有流程复核渲染。',warnings or None)]


def check_table_layout_static(document,rules):
    import zipfile
    import xml.etree.ElementTree as ET
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};W='{'+ns['w']+'}'
    cfg=table_layout(rules);failures=[];warnings=[];checked=0
    entries={x['index']:x for x in cfg.get('tables',[])}
    with zipfile.ZipFile(document) as z:
        body=ET.fromstring(z.read('word/document.xml')).find('w:body',ns)
        styles=ET.fromstring(z.read('word/styles.xml')) if 'word/styles.xml' in z.namelist() else ET.Element('styles')
        sm={s.get(W+'styleId'):s for s in styles.findall('w:style',ns)}
        defaults=styles.find('w:docDefaults/w:pPrDefault/w:pPr',ns)
        def chain(sid):
            out=[];seen=set()
            while sid and sid not in seen and sid in sm:
                seen.add(sid);node=sm[sid];out.insert(0,node);parent=node.find('w:basedOn',ns);sid=parent.get(W+'val') if parent is not None else None
            return out
        def val(el,key,default=None):return el.get(W+key,default) if el is not None else default
        def merge(props):
            result={}
            for pr in props:
                if pr is None:continue
                for node in pr:
                    name=node.tag.rsplit('}',1)[-1]
                    result.setdefault(name,{}).update({k.rsplit('}',1)[-1]:v for k,v in node.attrib.items()})
            return result
        default_para=next((k for k,n in sm.items() if n.get(W+'type')=='paragraph' and n.get(W+'default') in ('1','true')),None)
        def top_tables(el):
            for child in el:
                if child.tag==W+'tbl':yield child
                else:yield from top_tables(child)
        all_tables=list(top_tables(body))
        for ti,t in enumerate(all_tables,1):
            item=entries.get(ti,{});rows=t.findall('w:tr',ns);text=''.join(x.text or '' for x in t.iter(W+'t'));detail={'table':ti,'text':text[:100]}
            if item.get('kind','data')!='data':continue
            if ('实景图占位' in text or '需插入现场照片' in text) and len(rows)==1 and len(rows[0].findall('w:tc',ns))==1:continue
            if t.findall('.//w:drawing',ns) or t.findall('.//w:pict',ns):
                warnings.append({**detail,'reason':'含图件/对象的表格须按实际用途复核，不套普通数据表规则'});continue
            counts=[len(r.findall('w:tc',ns)) for r in rows]
            if not counts or max(counts)==0:
                warnings.append({**detail,'reason':'无法可靠识别表格行列，须按实际用途复核'});continue
            if max(counts)<2 and not item:
                warnings.append({**detail,'reason':'用途未确认的单列表须由AI分类；不能作为已审数据表通过'});continue
            if t.find('.//w:tbl',ns) is not None or t.find('.//w:gridSpan',ns) is not None or t.find('.//w:hMerge',ns) is not None or t.find('.//w:vMerge',ns) is not None or len(set(counts))!=1:
                warnings.append({**detail,'reason':'嵌套/合并/不规则表格需Word逐项核查，静态不猜列角色或版心'});continue
            checked+=1;cols=counts[0];roles=item.get('column_roles',[]);header=item.get('header_rows',1)
            if ti not in entries or len(roles)!=cols or header>len(rows):
                warnings.append({**detail,'reason':'须由AI按项目内容确认本表header_rows及每列short/long角色；未确认不强判内容对齐'})
                roles=[]
            default_table=next((k for k,n in sm.items() if n.get(W+'type')=='table' and n.get(W+'default') in ('1','true')),None)
            pr=t.find('w:tblPr',ns);sid=val(pr.find('w:tblStyle',ns) if pr is not None else None,'val',default_table);tc=chain(sid)
            # Conditional table paragraph/cell styles cannot be resolved from a partial rule guess.
            conditional=any(n.find('w:pPr',ns) is not None or n.find('w:tcPr',ns) is not None
                            for st in tc for n in st.findall('w:tblStylePr',ns))
            if conditional:
                warnings.append({**detail,'reason':'条件表格样式影响段落/单元格属性，须Word读取有效格式'});continue
            tp=merge([s.find('w:tblPr',ns) for s in tc]+[pr])
            floating=pr is not None and pr.find('w:tblpPr',ns) is not None
            if not floating and tp.get('jc',{}).get('val','left')!=cfg['table_alignment']:
                failures.append({**detail,'reason':'表格整体未按有效规则对齐','expected':cfg['table_alignment']})
            if floating:
                warnings.append({**detail,'reason':'浮动表须用Word实际页面核实整体对齐及版心边界，不凭行对齐属性硬判'})
            if abs(float(tp.get('tblInd',{}).get('w','0')))>.1:
                warnings.append({**detail,'reason':'表格带非零整体缩进，须核对Word实际左右边界'})
            # Fixed grid overflow is definite; auto-fit/preferred widths need native Word.
            parent=next((el for el in body if el is t or t in el.iter()),t);children=list(body);at=children.index(parent)
            sec=next((el.find('w:pPr/w:sectPr',ns) for el in children[at:] if el.find('w:pPr/w:sectPr',ns) is not None),body.find('w:sectPr',ns))
            sz=sec.find('w:pgSz',ns) if sec is not None else None;mar=sec.find('w:pgMar',ns) if sec is not None else None;columns=sec.find('w:cols',ns) if sec is not None else None
            grid=[int(x.get(W+'w','0')) for x in t.findall('w:tblGrid/w:gridCol',ns)]
            if sz is not None and mar is not None and not int(mar.get(W+'gutter','0')) and (columns is None or int(columns.get(W+'num','1'))==1):
                width=int(sz.get(W+'w','0'))-int(mar.get(W+'left','0'))-int(mar.get(W+'right','0'))
                if grid and sum(grid)>width+40:
                    target=failures if not floating and tp.get('tblLayout',{}).get('type')=='fixed' else warnings
                    target.append({**detail,'reason':'表格网格总宽超出版心；自动调整表须用Word确认实际宽度','grid_twips':sum(grid),'body_twips':width})
            else:warnings.append({**detail,'reason':'多栏/装订线或页面信息不足，须Word确认表格实际版心'})
            for ri,row in enumerate(rows):
                for ci,cell in enumerate(row.findall('w:tc',ns)):
                    cp=merge([s.find('w:tcPr',ns) for s in tc]+[cell.find('w:tcPr',ns)])
                    loc={**detail,'row':ri+1,'column':ci+1}
                    if cp.get('vAlign',{}).get('val','top')!=cfg['vertical_alignment']:
                        failures.append({**loc,'reason':'单元格未按有效规则垂直对齐','expected':cfg['vertical_alignment']})
                    for p in cell.findall('w:p',ns):
                        pt=''.join(x.text or '' for x in p.iter(W+'t'))
                        if not pt.strip():continue
                        pp=p.find('w:pPr',ns);ps=val(pp.find('w:pStyle',ns) if pp is not None else None,'val',default_para)
                        effective=merge([defaults]+[s.find('w:pPr',ns) for s in tc]+[s.find('w:pPr',ns) for s in chain(ps)]+[pp])
                        indent=effective.get('ind',{})
                        if cfg['zero_indents'] and 'numPr' not in effective and any(float(v)!=0 for k,v in indent.items() if k in ('left','right','start','end','firstLine','hanging','leftChars','rightChars','startChars','endChars','firstLineChars','hangingChars')):
                            failures.append({**loc,'text':pt[:100],'reason':'单元格继承/设置了非零正文缩进','effective_indents':indent})
                        elif 'numPr' in effective:
                            warnings.append({**loc,'reason':'单元格内列表缩进须按实际层级复核，不机械清零'})
                        expected=cfg['header_alignment'] if ri<header and roles else (cfg['short_text_alignment'] if roles and roles[ci]=='short' else cfg['long_text_alignment'] if roles else None)
                        if expected and effective.get('jc',{}).get('val','left')!=expected:
                            failures.append({**loc,'text':pt[:100],'reason':'表头/列角色水平对齐不符','expected':expected})
                        if p.find('.//w:br',ns) is not None or p.find('.//w:tab',ns) is not None or re.search(r'[ \u3000]{3,}',pt):
                            warnings.append({**loc,'text':pt[:100],'reason':'存在手动换行/制表符/连续空格，须确认未用于凑排版；不自动删除有效内容'})
        absent=sorted(set(entries)-set(range(1,len(all_tables)+1)))
        if absent:failures.append({'reason':'逐表配置指向不存在的表格','table_indices':absent})
    return _table_result(failures,warnings,checked)


def check_table_layout_word(doc,rules):
    cfg=table_layout(rules);entries={x['index']:x for x in cfg.get('tables',[])};failures=[];warnings=[];checked=0
    align={'left':0,'center':1,'right':2,'both':3};vertical={'top':0,'center':1,'bottom':3}
    for ti in range(1,int(doc.Tables.Count)+1):
        detail={'table':ti};item=entries.get(ti,{})
        if item.get('kind','data')!='data':continue
        try:
            t=doc.Tables(ti);text=clean_text(t.Range.Text);detail.update(page=page_of_range(t.Range),text=text[:100])
            if int(getattr(t,'NestingLevel',1))>1 or int(t.Range.Tables.Count)>1:
                warnings.append({**detail,'reason':'嵌套表须按Word渲染逐项复核'});continue
            if ('实景图占位' in text or '需插入现场照片' in text) and int(t.Range.Cells.Count)==1:continue
            if int(t.Range.InlineShapes.Count)>0:
                warnings.append({**detail,'reason':'图形/照片布局表按实际用途复核'});continue
            cols=int(t.Columns.Count)
            if cols<1:
                warnings.append({**detail,'reason':'Word未返回可靠列数，须按实际用途复核'});continue
            if cols<2 and not item:
                warnings.append({**detail,'reason':'用途未确认的单列表须由AI分类；不能作为已审数据表通过'});continue
            rows=int(t.Rows.Count);roles=item.get('column_roles',[]);header=item.get('header_rows',1)
            if len(roles)!=cols or header>rows:
                warnings.append({**detail,'reason':'本表列角色/表头范围未可靠配置，须逐表复核'});roles=[]
            checked+=1
            floating=bool(getattr(t.Rows,'WrapAroundText',False))
            if floating:warnings.append({**detail,'reason':'Word浮动表须核对实际页面位置，不凭行对齐属性认定边界'})
            if not floating and int(t.Rows.Alignment)!=align[cfg['table_alignment']]:
                failures.append({**detail,'reason':'Word有效表格整体对齐不符','expected':cfg['table_alignment']})
            left_indent=float(t.Rows.LeftIndent)
            if abs(left_indent)>.5:warnings.append({**detail,'reason':'Word表格带非零整体缩进，须查看实际左右边界','left_indent_pt':left_indent})
            ps=t.Range.Sections(1).PageSetup
            if floating or float(ps.Gutter)>0 or int(ps.TextColumns.Count)>1 or bool(getattr(ps,'MirrorMargins',False)):
                warnings.append({**detail,'reason':'特殊页面版心须查看Word实际渲染'})
            else:
                width=sum(float(t.Rows(1).Cells(i).Width) for i in range(1,int(t.Rows(1).Cells.Count)+1))
                usable=float(ps.PageWidth)-float(ps.LeftMargin)-float(ps.RightMargin)
                if width>usable+2:failures.append({**detail,'reason':'Word实际表格总宽超出版心','width_pt':width,'body_width_pt':usable})
            for ri in range(1,rows+1):
                for ci in range(1,cols+1):
                    cell=t.Cell(ri,ci);loc={**detail,'row':ri,'column':ci}
                    if int(cell.VerticalAlignment)!=vertical[cfg['vertical_alignment']]:failures.append({**loc,'reason':'Word有效单元格垂直对齐不符'})
                    for pi in range(1,int(cell.Range.Paragraphs.Count)+1):
                        p=cell.Range.Paragraphs(pi)
                        if not clean_text(p.Range.Text):continue
                        pf=p.Format;expected=cfg['header_alignment'] if roles and ri<=header else (cfg['short_text_alignment'] if roles and roles[ci-1]=='short' else cfg['long_text_alignment'] if roles else None)
                        if expected and int(pf.Alignment)!=align[expected]:failures.append({**loc,'reason':'Word有效表头/列角色水平对齐不符','expected':expected})
                        is_list=int(p.Range.ListFormat.ListType)!=0
                        ind={k:float(getattr(pf,k)) for k in ('LeftIndent','RightIndent','FirstLineIndent')}
                        if cfg['zero_indents'] and not is_list and any(abs(v)>.1 and abs(v)<999000 for v in ind.values()):failures.append({**loc,'reason':'Word有效单元格段落缩进不为零','indents_pt':ind})
                        if is_list:warnings.append({**loc,'reason':'单元格列表缩进须按层级复核'})
            warnings.append({**detail,'reason':'核对实际渲染中短名称/数字单位换行、列宽比例及跨页表头；程序属性通过不代表视觉合格'})
        except Exception as exc:warnings.append({**detail,'reason':'Word表格测量未完整完成（合并/复杂表可能不支持规则网格）：'+str(exc)})
    return _table_result(failures,warnings,checked,True)


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
    actual = [normalize_spaces(h["text"]) for h in headings if h["level"] == 1]
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
        for h in headings if h.get("page_break_before") and not (rules.get("allow_front_matter_break") and headings and h is headings[0] and body_start_page > 1 and h["page"] == body_start_page)
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
        heading_level, _ = classify_heading(first["text"], rules)
        detail = {
            "section": idx,
            "page": first["page"],
            "first_text": first["text"],
            "section_start": start_type,
            "layout_changed": layout_changed,
        }
        if (rules.get('allow_front_matter_break') and headings and body_start_page > 1
                and first['start'] == headings[0]['start'] and first['page'] == body_start_page):
            continue  # 仅允许必要的前置页/正文边界；正文内部仍连续排版
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
        level, _ = classify_heading(rec["text"], rules)
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
        headings, heading_checks = extract_heading_structure(records, int(rules["body_start_page"]), rules)
        results.extend(heading_checks)
        results.extend(check_front_page_and_body(doc, records, headings, rules, page_count))
        results.extend(check_toc_layout_word(doc, records, headings, rules))
        results.extend(check_table_layout_word(doc, rules))
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
            "audit_version": "1.9.0",
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
EDITORIAL_RE = re.compile(r'本[章节段](?:应|需|可)(?:写|补充|描述|引用)|(?:后续|此处|这里)(?:补写|扩写|插入正文)|(?:引用|调用|根据)知识库|(?:编制|审核|硬审|页数|门禁)脚本|待完善正文|编制时应|本段可用|本[章节段](?:需|尚需|有待)(?:进一步)?完善|建议补充相关措施|(?:审查|审核|编制)(?:意见|状态)[:：]|(?:本稿|本章|本节|本段)(?:已|尚未|未)(?:通过|完成)(?:审核|审查|复核)|(?:本稿|本章|本节|本段)(?:质量优秀|已达优秀)|待办[:：]|(?:差异|冲突)(?:尚未|未)闭合')
EDITORIAL_WARN_RE = re.compile(r'疑似符号|得分点|\bskill\b|建议(?:进一步)?补充|(?:正文|本稿|本[章节段])[^。；\n]{0,20}(?:待核实|待完善|待确认|自评|评分)|(?:待补充|待核实|待完善)|(?:资料|参数|差异|冲突)[^。；\n]{0,12}(?:未闭合|尚未关闭)',re.I)
LANGUAGE_CONTEXT_RE = re.compile(r'监理|审图|图纸会审|审查意见|审核意见|设计单位|整改通知|禁止|不得|严禁|例如|示例|[“「「]|依据[^。；]{0,16}意见')


def language_checks(records):
    failures, warnings = [], []
    for rec in records:
        text = normalize_spaces(rec.get('text',''))
        high = EDITORIAL_RE.search(text)
        low = EDITORIAL_WARN_RE.search(text)
        if not high and not low:
            continue
        hit = high or low
        detail = {**{k:v for k,v in rec.items() if k!='text'}, 'text':text[:240], 'match':hit.group(0)}
        # 报告/引用/禁用示例可能属于正常施工过程；必须语义复核，不能自动放行。
        if high and not LANGUAGE_CONTEXT_RE.search(text):
            failures.append(detail)
        else:
            warnings.append(detail)
    return [CheckResult('正文内部编制语言','FAIL' if failures else 'PASS',
             '明确内部审查/编制语言须定点修正；不得编造事实掩盖冲突。' if failures else '未命中高确定性内部语言；不替代AI语义复核。', failures or None),
            CheckResult('正文异常用语_语境复核','WARN' if warnings else 'PASS',
             '需逐条结合施工语境判断，记录处置，禁止机械删除技术表达。' if warnings else '未发现待判定用语。', warnings or None)]


def document_language_records(path):
    import zipfile
    import xml.etree.ElementTree as ET
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    records=[]
    with zipfile.ZipFile(path) as z:
        for part in z.namelist():
            if not re.fullmatch(r'word/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml',part):continue
            root=ET.fromstring(z.read(part))
            for i,para in enumerate(root.findall('.//w:p',ns),1):
                # 排除修订删除文字；w:t覆盖表格、域缓存、文本框，跨run合并。
                texts=[t.text or '' for t in para.iter() if t.tag in ('{'+ns['w']+'}t','{'+ns['a']+'}t')]
                if texts:records.append({'part':part,'paragraph':i,'text':''.join(texts)})
    return records


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
    out.extend(language_checks(document_language_records(document)))
    out.extend(check_toc_layout_static(document,rules))
    out.extend(check_table_layout_static(document,rules))
    return out


def static_audit(document, rules_path, report_path,page_dir=None):
    rules = load_rules(rules_path,page_dir)
    results = portable_checks(document, rules)
    overall = 'FAIL' if any(x.status=='FAIL' for x in results) else 'PASS'
    report = {'audit_version':'1.9.0', 'audit_type':'embedded_docx_static', 'document_sha256':sha256_file(document),
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
