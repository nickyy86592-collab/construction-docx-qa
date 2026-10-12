#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
施工组织设计编制流程门禁（Microsoft Word 阶段校准版）

职责：
- 正文生成前读取同目录中唯一的项目页数硬审脚本；
- 建立并校验《篇幅规划状态.json》；
- 约 40% / 75% 阶段接收 Microsoft Word 实测报告并校核；
- Word 暂不可用时允许 LibreOffice 临时估值，但只能记为 PROVISIONAL；
- 正式交付前执行 final-ready：40% / 75% 两个阶段都必须最终有 Microsoft Word 实测 PASS。

不负责正文写作、技术方案判断、最终 Word 页数或版式硬审。
"""
from __future__ import annotations

import argparse
import ast
import json
import math
import sys
from datetime import datetime
from pathlib import Path

STATE_DEFAULT = "篇幅规划状态.json"
STAGES = (40, 75)
PAGE_DIR = None


def now_iso():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _literal_constants(path: Path):
    tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
    vals = {}
    wanted = {"EXPECTED_BODY_PAGES", "PAGE_TOLERANCE", "FRONT_MATTER_PAGES"}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            name = node.targets[0].id
            if name in wanted:
                try:
                    vals[name] = ast.literal_eval(node.value)
                except Exception:
                    pass
    return vals


def discover_page_script(directory=None):
    here = Path(directory or PAGE_DIR or Path(__file__).resolve().parent).resolve()
    if not here.is_dir():
        raise RuntimeError("页数脚本目录不存在："+str(here))
    found = []
    for p in sorted(here.glob("*页数硬审*.py")):
        if p.name == Path(__file__).name:
            continue
        vals = _literal_constants(p)
        if "EXPECTED_BODY_PAGES" in vals:
            found.append((p, vals))
    if not found:
        return None, None
    if len(found) > 1:
        raise RuntimeError("发现多个项目页数硬审脚本，页数来源不唯一：" + "、".join(p.name for p, _ in found))
    p, vals = found[0]
    vals.setdefault("PAGE_TOLERANCE", 0)
    vals.setdefault("FRONT_MATTER_PAGES", 0)
    if any(type(vals[k]) is not int for k in ("EXPECTED_BODY_PAGES","PAGE_TOLERANCE","FRONT_MATTER_PAGES")) or vals["EXPECTED_BODY_PAGES"]<=0 or vals["PAGE_TOLERANCE"]<0 or vals["FRONT_MATTER_PAGES"]<0:
        raise RuntimeError("项目页数规则必须为有效整数目标、非负容差及非负前置页数")
    return p, vals


def load_state(path: Path):
    if not path.exists():
        raise RuntimeError(f"未找到状态文件：{path}，请先执行 init")
    return json.loads(path.read_text(encoding="utf-8-sig"))


def save_state(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp=path.with_suffix(path.suffix+".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(float(v))


def current_rule(page_script, vals):
    return {
        "source_script": page_script.name if page_script else None,
        "body_pages_target": int(vals["EXPECTED_BODY_PAGES"]),
        "page_tolerance": int(vals["PAGE_TOLERANCE"]),
        "front_matter_pages_excluded": int(vals["FRONT_MATTER_PAGES"]),
    }


def cmd_init(state_path: Path, force: bool):
    page_script, vals = discover_page_script()

    if state_path.exists() and not force:
        print("FLOW_GATE_STATUS=BLOCKED")
        print(f"STATE_EXISTS={state_path}")
        print("请继续完成规划并运行 preflight；如需重置请使用 init --force")
        return 2

    if vals and (any(type(vals[k]) is not int for k in ("EXPECTED_BODY_PAGES","PAGE_TOLERANCE","FRONT_MATTER_PAGES")) or vals["EXPECTED_BODY_PAGES"]<=0 or vals["PAGE_TOLERANCE"]<0 or vals["FRONT_MATTER_PAGES"]<0):
        raise RuntimeError("项目页数常量必须为有效正目标/非负整数")
    target = int(vals["EXPECTED_BODY_PAGES"]) if vals else 0
    tol = int(vals["PAGE_TOLERANCE"]) if vals else 0
    front = int(vals["FRONT_MATTER_PAGES"]) if vals else 0

    state = {
        "schema_version": "2.0",
        "startup": startup_template(),
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "phase": "PLANNING",
        "page_rule": {
            "source_script": page_script.name if page_script else None,
            "body_pages_target": target,
            "page_tolerance": tol,
            "body_pages_allowed_range": [target - tol, target + tol],
            "front_matter_pages_excluded": front,
            "total_pages_allowed_range": [target - tol + front, target + tol + front],
        },
        "plan": {
            "status": "NOT_READY",
            "planned_total_body_pages": None,
            "chapter_plan": [],
            "figure_table_pages_estimate": None,
            "reserve_pages": None,
            "planning_notes": "",
            "checkpoint_expected_page_ranges": {"40": None, "75": None},
        },
        "checkpoints": {},
        "gate": {
            "preflight": "BLOCKED",
            "final_ready": "BLOCKED",
            "reason": "未完成篇幅规划，禁止开始正式正文",
        },
    }
    if page_script is None:
        state["page_rule"] = None
    save_state(state_path, state)

    print("FLOW_GATE_STATUS=BLOCKED")
    print(f"PAGE_RULE_SOURCE={page_script.name if page_script else None}")
    print(f"BODY_PAGE_TARGET={target}")
    print(f"PAGE_TOLERANCE={tol}")
    print(f"FRONT_MATTER_EXCLUDED={front}")
    print(f"STATE={state_path}")
    print("PRE_DRAFT_BLOCKED")
    return 2


def _rule_mismatch(state, page_script, vals):
    rule = state.get("page_rule") or {}
    expected = current_rule(page_script, vals)
    return [f"{k} 与当前页数脚本不一致" for k, v in expected.items() if rule.get(k) != v]


def cmd_preflight(state_path: Path):
    page_script, vals = discover_page_script()
    state = load_state(state_path)
    if page_script is None:
        errors=startup_errors(state,state_path.parent)
        state["gate"]["preflight"]="FAIL" if errors else "PASS"
        save_state(state_path,state)
        return emit_gate(errors,"PRE_DRAFT_ALLOWED")
    errors = _rule_mismatch(state, page_script, vals) + startup_errors(state,state_path.parent)
    plan = state.get("plan") or {}
    target = int(vals["EXPECTED_BODY_PAGES"])

    planned = plan.get("planned_total_body_pages")
    if not is_num(planned):
        errors.append("未填写 planned_total_body_pages")
    else:
        corridor = max(2, math.ceil(target * 0.08))
        if not target - corridor <= float(planned) <= target + corridor:
            errors.append(
                f"规划总正文页数 {planned} 明显偏离目标 {target}，"
                f"应先收敛至约 {target-corridor}–{target+corridor} 页"
            )

    chapters = plan.get("chapter_plan")
    if not isinstance(chapters, list) or len(chapters) < 2:
        errors.append("chapter_plan 至少应有两个实际章节/关键内容规划")
    else:
        for i, x in enumerate(chapters, 1):
            if not isinstance(x, dict):
                errors.append(f"chapter_plan 第{i}项格式错误")
                continue
            rng = x.get("page_range")
            if not str(x.get("chapter", "")).strip():
                errors.append(f"chapter_plan 第{i}项缺少章节名")
            if not str(x.get("basis", "")).strip():
                errors.append(f"chapter_plan 第{i}项缺少评分/重难点/复杂度依据")
            if not (
                isinstance(rng, list)
                and len(rng) == 2
                and is_num(rng[0])
                and is_num(rng[1])
                and 0 <= float(rng[0]) <= float(rng[1])
            ):
                errors.append(f"chapter_plan 第{i}项 page_range 无效")

    if isinstance(chapters,list) and chapters and all(isinstance(x,dict) and isinstance(x.get("page_range"),list) and len(x["page_range"])==2 and all(is_num(v) for v in x["page_range"]) for x in chapters) and is_num(planned) and is_num(plan.get("reserve_pages")):
        lo=sum(x["page_range"][0] for x in chapters)+plan["reserve_pages"]
        hi=sum(x["page_range"][1] for x in chapters)+plan["reserve_pages"]
        if not lo<=planned<=hi:
            errors.append("规划总量与章区间（含图表）及机动量不相容")
    if is_num(plan.get("figure_table_pages_estimate")) and is_num(planned) and not 0<=plan["figure_table_pages_estimate"]<=planned:
        errors.append("图表估算不能为负或超过全篇页数")
    if is_num(plan.get("reserve_pages")) and plan["reserve_pages"]<0:
        errors.append("机动页数不能为负")
    if not is_num(plan.get("figure_table_pages_estimate")):
        errors.append("未填写 figure_table_pages_estimate")
    if not is_num(plan.get("reserve_pages")):
        errors.append("未填写 reserve_pages")
    if len(str(plan.get("planning_notes", "")).strip()) < 12:
        errors.append("planning_notes 过短，需说明收敛逻辑")

    ranges = plan.get("checkpoint_expected_page_ranges") or {}
    prev_mid = -1.0
    # 仅用于防止明显荒谬的规划，不要求机械线性。
    broad = {40: (0.22, 0.65), 75: (0.50, 0.96)}
    for stage in STAGES:
        rng = ranges.get(str(stage))
        if not (
            isinstance(rng, list)
            and len(rng) == 2
            and is_num(rng[0])
            and is_num(rng[1])
            and 0 <= float(rng[0]) <= float(rng[1])
        ):
            errors.append(f"{stage}% 阶段预计页数范围未填写或无效")
            continue
        mid = (float(rng[0]) + float(rng[1])) / 2
        if mid <= prev_mid:
            errors.append(f"{stage}% 阶段预计页数应高于前一阶段")
        prev_mid = mid
        ratio = mid / target
        lo, hi = broad[stage]
        if not lo <= ratio <= hi:
            errors.append(f"{stage}% 阶段预计页数与总目标关系明显不合理")

    state["updated_at"] = now_iso()
    if errors:
        state["phase"] = "PLANNING"
        state["gate"]["preflight"] = "FAIL"
        state["gate"]["reason"] = errors
        save_state(state_path, state)
        print("FLOW_GATE_STATUS=FAIL")
        print("PRE_DRAFT_BLOCKED")
        for e in errors:
            print("- " + e)
        return 2

    state["phase"] = "DRAFTING_ALLOWED"
    state["plan"]["status"] = "READY"
    state["gate"]["preflight"] = "PASS"
    state["gate"]["preflight_passed_at"] = now_iso()
    state["gate"]["final_ready"] = "BLOCKED"
    save_state(state_path, state)

    print("FLOW_GATE_STATUS=PASS")
    print("PRE_DRAFT_ALLOWED")
    print(f"TARGET_BODY_PAGES={target}")
    return 0


def _load_word_report(report_path: Path, expected_stage: int, page_script, vals):
    if not report_path.exists():
        raise RuntimeError(f"未找到 Microsoft Word 阶段报告：{report_path}")

    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    errors = []

    if report.get("page_source") != "microsoft_word":
        errors.append("阶段报告 page_source 不是 microsoft_word")
    if report.get("status") != "MEASURED":
        errors.append("阶段报告 status 不是 MEASURED")
    if int(report.get("stage_percent", -1)) != int(expected_stage):
        errors.append("阶段报告 stage_percent 与当前阶段不一致")

    actual = report.get("actual") or {}
    body_pages = actual.get("word_body_pages")
    if not is_num(body_pages):
        errors.append("阶段报告缺少 actual.word_body_pages")

    rr = report.get("project_rule") or {}
    expected = current_rule(page_script, vals)
    if rr.get("target_body_pages") != expected["body_pages_target"]:
        errors.append("阶段报告目标正文页数与当前项目页数脚本不一致")
    if rr.get("page_tolerance") != expected["page_tolerance"]:
        errors.append("阶段报告容差与当前项目页数脚本不一致")
    if rr.get("front_matter_pages_excluded") != expected["front_matter_pages_excluded"]:
        errors.append("阶段报告前置排除页数与当前项目页数脚本不一致")

    if errors:
        raise RuntimeError("Microsoft Word 阶段报告校验失败：" + "；".join(errors))

    return float(body_pages), report


def cmd_checkpoint(
    state_path: Path,
    stage: int,
    report_path: Path | None,
    actual_pages: float | None,
    engine: str,
    allow_estimate: bool,
    note: str,
    document: Path,
):
    page_script, vals = discover_page_script()
    if page_script is None:
        print("FLOW_GATE_STATUS=SKIP")
        return 0

    state = load_state(state_path)
    if (state.get("gate") or {}).get("preflight") != "PASS":
        print("FLOW_GATE_STATUS=FAIL")
        print("PRE_DRAFT_NOT_PASSED")
        return 2

    errors = _rule_mismatch(state, page_script, vals)+startup_errors(state,state_path.parent)
    if stage==75 and (state.get("checkpoints",{}).get("40",{}).get("status") not in ("PASS","PROVISIONAL")):
        errors.append("75%前必须完成40%校准")
    if errors:
        print("FLOW_GATE_STATUS=FAIL")
        for e in errors:
            print("- " + e)
        return 2

    rng = ((state.get("plan") or {}).get("checkpoint_expected_page_ranges") or {}).get(str(stage))
    if not (
        isinstance(rng, list)
        and len(rng) == 2
        and is_num(rng[0])
        and is_num(rng[1])
    ):
        raise RuntimeError(f"缺少 {stage}% 阶段预计页数范围")

    lo, hi = float(rng[0]), float(rng[1])

    source = None
    report_meta = None
    if report_path is not None:
        actual, report = _load_word_report(report_path, stage, page_script, vals)
        if report.get("document_sha256") != sha256(document):
            raise RuntimeError("Word阶段报告与实际快照哈希不匹配")
        source = "microsoft_word"
        report_meta = {
            "report_file": str(report_path),
            "document_sha256": report.get("document_sha256"),
            "word_version": (report.get("actual") or {}).get("word_version"),
        }
    else:
        if actual_pages is None:
            raise RuntimeError("checkpoint 必须提供 --report（优先）或 --actual-pages")
        eng = (engine or "").strip().lower()
        if not allow_estimate or "libreoffice" not in eng:
            raise RuntimeError(
                "非 Microsoft Word 阶段数据只允许作为 LibreOffice 临时估值；"
                "需同时使用 --allow-estimate --engine libreoffice_estimate"
            )
        actual = float(actual_pages)
        source = "libreoffice_estimate"

    if not is_num(actual) or actual<=0:
        raise RuntimeError("阶段页数必须为正有限数值")
    within = lo <= actual <= hi
    if source == "microsoft_word":
        status = "PASS" if within else "FAIL"
        phase = f"CHECKPOINT_{stage}_{status}"
        rc = 0 if within else 2
    else:
        status = "PROVISIONAL" if within else "FAIL"
        phase = f"CHECKPOINT_{stage}_{status}"
        rc = 0 if within else 2

    cp = {
        "snapshot_path": str(document.resolve()),
        "document_sha256": sha256(document),
        "checked_at": now_iso(),
        "actual_pages": actual,
        "expected_page_range": [lo, hi],
        "page_source": source,
        "status": status,
        "note": note,
    }
    if report_meta:
        cp["word_report"] = report_meta

    state.setdefault("checkpoints", {})[str(stage)] = cp
    state["phase"] = phase
    state["updated_at"] = now_iso()
    state["gate"]["final_ready"] = "BLOCKED"
    save_state(state_path, state)

    print(f"FLOW_GATE_STATUS={status}")
    print(f"CHECKPOINT_STAGE={stage}")
    print(f"ACTUAL_PAGES={actual}")
    print(f"EXPECTED_RANGE={lo:g}-{hi:g}")
    print(f"PAGE_SOURCE={source}")

    if status == "PASS":
        print("DRAFTING_ALLOWED")
    elif status == "PROVISIONAL":
        print("DRAFTING_ALLOWED_WITH_ESTIMATE")
        print("WORD_REMEASUREMENT_REQUIRED_BEFORE_FINAL")
    else:
        print("DRAFTING_BLOCKED_UNTIL_ADJUSTED")
    return rc


def cmd_final_ready(state_path: Path):
    state = load_state(state_path)
    errors = startup_errors(state,state_path.parent)
    page_script,vals=discover_page_script()
    if page_script:
        errors += _rule_mismatch(state,page_script,vals)

    if (state.get("gate") or {}).get("preflight") != "PASS":
        errors.append("preflight 未通过")

    cps = state.get("checkpoints") or {}
    for stage in STAGES if page_script else ():
        cp = cps.get(str(stage))
        if not isinstance(cp, dict):
            errors.append(f"缺少 {stage}% 阶段 checkpoint")
            continue
        if not fresh_file({"path":cp.get("snapshot_path",""),"sha256":cp.get("document_sha256")},state_path.parent):
            errors.append(f"{stage}%快照缺失或已改变")
        if cp.get("status") != "PASS":
            errors.append(f"{stage}% 阶段 checkpoint 不是 PASS")
        if cp.get("page_source") != "microsoft_word":
            errors.append(f"{stage}% 阶段尚未使用 Microsoft Word 实测")

    state["updated_at"] = now_iso()
    if errors:
        state["gate"]["final_ready"] = "FAIL"
        state["gate"]["final_ready_reason"] = errors
        save_state(state_path, state)
        print("FLOW_GATE_STATUS=FAIL")
        print("FINAL_AUDIT_BLOCKED")
        for e in errors:
            print("- " + e)
        return 2

    state["gate"]["final_ready"] = "PASS"
    state["gate"]["final_ready_passed_at"] = now_iso()
    state["phase"] = "FINAL_AUDIT_ALLOWED"
    save_state(state_path, state)

    print("FLOW_GATE_STATUS=PASS")
    print("FINAL_AUDIT_ALLOWED")
    return 0


def cmd_status(state_path: Path):
    if not state_path.exists():
        print("FLOW_GATE_STATUS=NOT_INITIALIZED")
        return 2
    print(state_path.read_text(encoding="utf-8-sig"))
    return 0


# Evidence-based startup and delivery gates. Human judgments remain human judgments.
REVIEW_ITEMS = ('评分标准','硬性响应与否决格式','知识库冲突','数字条款原话','点名重难点','缺失待核实','企业资料','规范控制参数','初稿原文溯源','定稿最新文件与系统要求')


def sha256(path):
    import hashlib
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def fresh_file(item, base):
    p = Path(item.get('path',''))
    if not p.is_absolute():
        p = base/p
    return bool(p.is_file() and item.get('sha256') == sha256(p))


def startup_template():
    return {'sources': [], 'latest_updates': {'status':'unanswered','asked_at':'','reply_text':'','reply_source':''},
            'critical_questions':[], 'questions_review': '',
            'batch_plan': {'count':None,'reason':'','batches':[]},
            'word_channel':'', 'word_capability_evidence':'', 'real_photo_slots':[]}


def startup_errors(state, base):
    a = state.get('startup') or {}
    errors = []
    sources = a.get('sources')
    if not isinstance(sources,list) or not sources:
        errors.append('未登记原始资料文件及哈希')
    else:
        if not any(x.get('role')=='tender' for x in sources):
            errors.append('未登记招标原文')
        for x in sources:
            if not x.get('role') or not fresh_file(x,base):
                errors.append('资料缺失或版本哈希已改变：'+str(x.get('path')))
    latest = a.get('latest_updates') or {}
    if latest.get('status') not in ('confirmed_none','provided') or not all(latest.get(x) for x in ('asked_at','reply_text','reply_source')):
        errors.append('最新澄清/答疑/补遗/补充文件/系统要求尚无用户明确答复及来源')
    if latest.get('status')=='provided' and not latest.get('provided_paths'):
        errors.append('用户提供更新文件但尚未登记路径')
    if latest.get('status')=='provided' and sources:
        known = {x.get('path') for x in sources}
        if any(p not in known for p in latest.get('provided_paths',[])):
            errors.append('最新提供文件尚未纳入资料索引')
    if not a.get('questions_review'):
        errors.append('未判断关键缺失/冲突是否影响正式编制')
    for x in a.get('critical_questions',[]):
        if x.get('status') not in ('resolved','bounded') or not x.get('evidence'):
            errors.append('关键问题未关闭：'+str(x.get('question')))
        if x.get('status')=='bounded' and not x.get('safe_scope'):
            errors.append('条件性继续缺少安全编写范围')
    b = a.get('batch_plan') or {}
    count = b.get('count')
    if type(count) is not int or not 1 <= count <= 3 or not b.get('reason') or len(b.get('batches',[])) != count:
        errors.append('需预判一次完成能力并给出1至3批交付计划')
    if a.get('word_channel') not in ('local_word','github_runner','unavailable') or not a.get('word_capability_evidence'):
        errors.append('未记录可实际执行的Word通道/阻塞证据')
    if not a.get('real_photo_slots') or any(not x.get('name') or not x.get('location') or not x.get('basis') for x in a.get('real_photo_slots',[])):
        errors.append('实景图主题/插入位置/价值依据必须至少一项')
    return errors


def emit_gate(errors, success):
    print('FLOW_GATE_STATUS='+('FAIL' if errors else 'PASS'))
    print('DELIVERY_BLOCKED' if errors else success)
    for e in errors:
        print('- '+e)
    return 2 if errors else 0


def validate_content_review(document, review_path):
    r = json.loads(review_path.read_text(encoding='utf-8'))
    errors = []
    if r.get('document_sha256') != sha256(document):
        errors.append('内容核验记录不属于当前DOCX')
    entries = r.get('source_checks') or []
    for name in REVIEW_ITEMS:
        matched = [x for x in entries if x.get('item')==name]
        if len(matched)!=1:
            errors.append('缺少唯一回查记录：'+name)
            continue
        x=matched[0]
        if x.get('status') not in ('verified','not_applicable') or not x.get('conclusion'):
            errors.append('回查尚未关闭：'+name)
        if x.get('status')=='verified':
            evidence=x.get('evidence') or []
            if not evidence or any(not v.get('locator') or not fresh_file(v,review_path.parent) for v in evidence):
                errors.append('回查证据文件/哈希/页码条款缺失：'+name)
        elif not x.get('reason'):
            errors.append('不适用缺少依据：'+name)
        if name in ('评分标准','硬性响应与否决格式','初稿原文溯源','定稿最新文件与系统要求') and x.get('status')!='verified':
            errors.append('核心回查不能标不适用：'+name)
    latest = r.get('latest_confirmation') or {}
    if not latest.get('confirmed_at') or not latest.get('reply_source') or not latest.get('conclusion'):
        errors.append('未完成定稿前最新文件/提交系统核验')
    import zipfile
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(document) as z:
        tree=ET.fromstring(z.read('word/document.xml'))
        ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        body_text=''.join(x.text or '' for x in tree.findall('.//w:t',ns))
    parameters = r.get('parameters')
    if not isinstance(parameters,list) or (not parameters and not r.get('no_parameters_reason')):
        errors.append('未登记用于正文的关键参数及跨来源核对')
    for x in parameters or []:
        candidates=x.get('occurrences') or []
        if not x.get('name') or not x.get('scope') or not x.get('cross_source_search') or not candidates:
            errors.append('关键参数跨来源搜索记录不足')
        for v in candidates:
            if v.get('value') is None or not all(v.get(k) for k in ('unit','scope','locator')) or not fresh_file(v,review_path.parent):
                errors.append('关键参数候选值/部位/原图证据不完整：'+str(x.get('name')))
        if x.get('status') not in ('consistent','resolved') or not x.get('decision') or not x.get('body_locations'):
            errors.append('关键参数未关闭或缺正文定位：'+str(x.get('name')))
        for use in x.get('body_locations') or []:
            if not isinstance(use,dict) or not use.get('location') or not use.get('quote') or use.get('quote') not in body_text:
                errors.append('关键参数正文证据不存在：'+str(x.get('name')))
                continue
            import re
            numeric=str(use.get('value',''))
            if not re.search(r'(?<![\d.])'+re.escape(numeric)+r'(?![\d.])',use.get('quote','')) or str(use.get('unit','')) not in use.get('quote',''):
                errors.append('参数引文与登记数值/单位不一致：'+str(x.get('name')))
            allowed={(str(v.get('value')),v.get('unit'),v.get('scope')) for v in candidates}
            if (str(use.get('value')),use.get('unit'),use.get('scope')) not in allowed:
                errors.append('正文参数与登记适用部位/候选值不匹配：'+str(x.get('name')))
        if x.get('status')=='resolved' and not x.get('resolution_evidence'):
            errors.append('参数冲突缺有效决定依据：'+str(x.get('name')))
        # Differing values require an explicit scope/version resolution, not a generic PASS.
        values={(str(v.get('value')),v.get('unit')) for v in candidates}
        if len(values)>1 and x.get('status')!='resolved':
            errors.append('不同候选值尚未解释适用部位/有效版本：'+str(x.get('name')))
    for x in r.get('conflicts',[]):
        if x.get('status')!='resolved' or not x.get('evidence') or not x.get('decision'):
            errors.append('未关闭冲突：'+str(x.get('name')))
    if (r.get('construction_language_review') or {}).get('all_paragraphs_reviewed') is not True or not (r.get('construction_language_review') or {}).get('conclusion'):
        errors.append('未完成逐段施工执行语言复核')
    if any((r.get('visual_review') or {}).get(k) is not True for k in ('three_color_figures_reviewed','caption_photo_reviewed')) or not (r.get('visual_review') or {}).get('conclusion'):
        errors.append('未完成彩色三图/照片框实际视觉复核')
    return errors, r


def cmd_content_review(document, review_path):
    errors,_ = validate_content_review(document,review_path)
    return emit_gate(errors,'CONTENT_REVIEW_ALLOWED')


def cmd_delivery(a):
    digest=sha256(a.document)
    errors, review=validate_content_review(a.document,a.review)
    state=None
    page_script, vals=discover_page_script()
    if a.mode=='full':
        state=load_state(a.state)
        errors+=startup_errors(state,a.state.parent)
        if (state.get('gate') or {}).get('preflight')!='PASS':
            errors.append('启动preflight未通过')
        if page_script and cmd_final_ready(a.state)!=0:
            errors.append('阶段final-ready未通过')
    reports={}
    for name,path in [('static',a.static_report),('word',a.word_report),('score',a.score_report),('page',a.page_report)]:
        if path is None:
            if name!='page' or page_script:
                errors.append('缺少审核报告：'+name)
            continue
        r=json.loads(path.read_text(encoding='utf-8'))
        reports[name]=r
        status = r.get('delivery_status') if name in ('word','static') else r.get('status')
        if status!='PASS' or r.get('document_sha256')!=digest:
            errors.append('报告未通过或文件哈希不匹配：'+name)
        if name=='score' and r.get('phase')!='final':
            errors.append('评分报告不是最终阶段')
        if name=='word' and r.get('layout_engine')!='Microsoft Word':
            errors.append('版式审核不是Microsoft Word')
        if name=='static' and r.get('audit_type')!='embedded_docx_static':
            errors.append('静态报告类型错误')
        if name=='page':
            if r.get('page_source')!='microsoft_word':
                errors.append('页数报告不是Microsoft Word')
            if vals:
                expected=current_rule(page_script,vals)
                if any((r.get('rule') or {}).get(k)!=v for k,v in expected.items() if k!='source_script'):
                    errors.append('页数报告规则与当前项目脚本不一致')
    static,word=reports.get('static',{}),reports.get('word',{})
    if not static.get('rules_sha256') or static.get('rules_sha256')!=word.get('rules_sha256'):
        errors.append('静态核验和Word审核使用不同格式规则')
    rules=word.get('rules') or {}
    import hashlib
    actual_rules_digest=hashlib.sha256(json.dumps(rules,sort_keys=True,ensure_ascii=False).encode()).hexdigest()
    if actual_rules_digest!=word.get('rules_sha256'):
        errors.append('Word报告格式规则指纹错误')
    # Re-run embedded checks so an empty or edited static report cannot bypass pictures.
    import importlib.util
    common_path=Path(__file__).with_name('Word正式交付硬审（不含页数）.py')
    spec=importlib.util.spec_from_file_location('_delivery_embedded',common_path)
    common=importlib.util.module_from_spec(spec)
    sys.modules[spec.name]=common
    spec.loader.exec_module(common)
    for check in common.portable_checks(a.document,rules):
        if check.status=='FAIL':
            errors.append(check.name+': '+check.message)
    stage=rules.get('photo_delivery_stage')
    if stage not in ('draft','final') or static.get('photo_delivery_stage')!=stage:
        errors.append('照片交付阶段未一致确认')
    slots=rules.get('required_real_photo_slots') or []
    if not slots:
        errors.append('实景图清单为空')
    if state:
        expected={x['name'] for x in state['startup'].get('real_photo_slots',[])}
        if not expected.issubset({x.get('name') for x in slots}):
            errors.append('启动约定的照片主题被省略')
    # Every WARN/SKIP in the actual reports must have a version-specific disposition.
    warning_review=json.loads(a.warning_review.read_text(encoding='utf-8')) if a.warning_review else {}
    if warning_review.get('document_sha256')!=digest:
        errors.append('未提供绑定当前文件的WARN/SKIP处理记录')
    dispositions=warning_review.get('items') or []
    for name,r in reports.items():
        for check in r.get('checks',[]):
            if check.get('status') not in ('WARN','SKIP'):
                continue
            m=[x for x in dispositions if x.get('report')==name and x.get('check')==check.get('name')]
            if len(m)!=1 or m[0].get('status') not in ('reviewed','not_applicable') or not m[0].get('evidence') or not m[0].get('conclusion'):
                errors.append('未处理审核提示：'+name+'/'+str(check.get('name')))
    if errors:
        return emit_gate(errors,'')
    return emit_gate([], 'FINAL_DELIVERY_ALLOWED' if stage=='final' else 'DRAFT_DELIVERY_ALLOWED')


def cmd_page_rule():
    script, vals = discover_page_script()
    if script is None:
        raise RuntimeError('未找到唯一项目页数硬审脚本')
    rule = current_rule(script, vals)
    rule['source_path'] = str(script.resolve())
    rule['body_pages_allowed_range'] = [rule['body_pages_target']-rule['page_tolerance'],rule['body_pages_target']+rule['page_tolerance']]
    print(json.dumps(rule,ensure_ascii=False))
    return 0


def cmd_run_page_audit(document, report_path, stage=None):
    script, vals=discover_page_script()
    if script is None:
        raise RuntimeError('未找到唯一项目页数硬审脚本')
    if stage is None:
        import subprocess
        return subprocess.run([sys.executable,str(script.resolve()),str(document.resolve()),'--report',str(report_path.resolve())]).returncode
    if sys.platform!='win32':
        raise RuntimeError('阶段实测要求Windows及可调用的Microsoft Word')
    if not document.is_file():
        raise RuntimeError('阶段DOCX快照不存在')
    import win32com.client
    word=doc=None
    before=sha256(document)
    try:
        word=win32com.client.DispatchEx('Word.Application')
        word.Visible=False
        word.DisplayAlerts=0
        doc=word.Documents.Open(str(document.resolve()),ReadOnly=True,AddToRecentFiles=False)
        doc.Repaginate()
        total=int(doc.ComputeStatistics(2))
        front=int(vals['FRONT_MATTER_PAGES'])
        if total<=front:
            raise RuntimeError('阶段稿没有有效正文页')
        if before!=sha256(document):
            raise RuntimeError('测量期间阶段稿发生变化')
        report={'audit_type':'word_stage_measurement','page_source':'microsoft_word','status':'MEASURED',
                'document':str(document.resolve()),'document_sha256':before,'stage_percent':stage,
                'measured_at':now_iso(),'project_rule_source':str(script.resolve()),
                'project_rule':{'target_body_pages':vals['EXPECTED_BODY_PAGES'],'page_tolerance':vals['PAGE_TOLERANCE'],'front_matter_pages_excluded':front},
                'actual':{'word_total_pages':total,'word_body_pages':total-front,'word_version':str(word.Version)}}
        save_state(report_path,report)
        print('WORD_STAGE_STATUS=MEASURED\nPAGE_SOURCE=microsoft_word\nREPORT='+str(report_path))
        return 0
    finally:
        if doc is not None:doc.Close(False)
        if word is not None:word.Quit()


def main():
    global PAGE_DIR
    p = argparse.ArgumentParser(description="施工组织设计预编制与 Microsoft Word 阶段篇幅流程门禁")
    p.add_argument("--page-dir",type=Path,default=None,help="项目页数脚本目录；默认自动查找同目录唯一项目脚本")
    p.add_argument("--state", type=Path, default=Path(STATE_DEFAULT))
    sub = p.add_subparsers(dest="command", required=True)

    s = sub.add_parser("init")
    s.add_argument("--force", action="store_true")

    sub.add_parser("preflight")
    sub.add_parser("page-rule")
    s=sub.add_parser("run-page-audit")
    s.add_argument("--document",type=Path,required=True)
    s.add_argument("--report",type=Path,required=True)
    s.add_argument("--stage",type=int,choices=STAGES)

    s = sub.add_parser("checkpoint")
    s.add_argument("--stage", type=int, choices=STAGES, required=True)
    s.add_argument("--document", type=Path, required=True, help="实际阶段DOCX快照，必须与报告哈希一致")
    s.add_argument("--report", type=Path, default=None, help="本机Word或远程Runner实际产出的阶段报告")
    s.add_argument("--actual-pages", type=float, default=None, help="仅 LibreOffice 临时估值时使用")
    s.add_argument("--engine", default="", help="临时估值时应为 libreoffice_estimate")
    s.add_argument("--allow-estimate", action="store_true")
    s.add_argument("--note", default="")

    sub.add_parser("final-ready")
    sub.add_parser("status")
    s=sub.add_parser("content-review")
    s.add_argument("--document", type=Path, required=True)
    s.add_argument("--review", type=Path, required=True)
    s=sub.add_parser("delivery")
    s.add_argument("--document",type=Path,required=True)
    s.add_argument("--mode",choices=("full","local-review"),default="full")
    for arg in ("static-report","word-report","score-report","review"):
        s.add_argument("--"+arg,type=Path,required=True)
    s.add_argument("--page-report",type=Path)
    s.add_argument("--warning-review",type=Path)

    a = p.parse_args()
    PAGE_DIR=a.page_dir
    try:
        if a.command=="page-rule":
            return cmd_page_rule()
        if a.command=="run-page-audit":
            return cmd_run_page_audit(a.document,a.report,a.stage)
        if a.command == "init":
            return cmd_init(a.state, a.force)
        if a.command == "preflight":
            return cmd_preflight(a.state)
        if a.command == "checkpoint":
            return cmd_checkpoint(
                a.state,
                a.stage,
                a.report,
                a.actual_pages,
                a.engine,
                a.allow_estimate,
                a.note,
                a.document,
            )
        if a.command == "final-ready":
            return cmd_final_ready(a.state)
        if a.command == "content-review":
            return cmd_content_review(a.document,a.review)
        if a.command == "delivery":
            return cmd_delivery(a)
        if a.command == "status":
            return cmd_status(a.state)
    except Exception as e:
        print(f"FLOW_GATE_STATUS=ERROR\nERROR: {e}", file=sys.stderr)
        return 3
    return 3


if __name__ == "__main__":
    raise SystemExit(main())
