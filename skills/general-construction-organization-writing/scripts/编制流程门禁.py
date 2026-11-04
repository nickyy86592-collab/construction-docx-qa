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

from visual_policy import visual_policy_errors, visual_policy_template

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
        "schema_version": "3.1",
        "module_registry": {},
        "pilot": {},
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
            "revision": 1,
            "requirements": [],
            "modules": [],
            "pilot_plan": {},
            "checkpoint_module_ids": {},
            "checkpoint_completion_basis": {},
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
        errors=startup_errors(state,state_path.parent)+detailed_plan_errors(state,state_path.parent,0)
        state["gate"]["plan_hash"]=plan_hash(state)
        state["gate"]["draft_ready"]="FAIL" if errors else "PASS"
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
        corridor = int(vals["PAGE_TOLERANCE"])
        if not target - corridor <= float(planned) <= target + corridor:
            errors.append(
                f"规划总正文页数 {planned} 不在项目脚本准入范围，"
                f"必须先收敛至 {target-corridor}–{target+corridor} 页"
            )

    chapters = plan.get("chapter_plan")
    if not isinstance(chapters, list) or len(chapters) < 1:
        errors.append("chapter_plan 至少应有一个实际章节/关键内容规划")
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

    errors += detailed_plan_errors(state,state_path.parent,target)
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

    state["phase"] = "PILOT_ONLY" if planning_policy(target)["pilot_required"] else "DRAFTING_ALLOWED"
    state["plan"]["status"] = "READY"
    state["gate"]["preflight"] = "PASS"
    state["gate"]["preflight_passed_at"] = now_iso()
    state["gate"]["final_ready"] = "BLOCKED"
    save_state(state_path, state)

    print("FLOW_GATE_STATUS=PASS")
    state["gate"]["plan_hash"] = plan_hash(state)
    state["gate"]["draft_ready"] = "BLOCKED" if planning_policy(target)["pilot_required"] else "PASS"
    save_state(state_path,state)
    print("PILOT_DRAFT_ALLOWED" if planning_policy(target)["pilot_required"] else "PRE_DRAFT_ALLOWED")
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

    errors = _rule_mismatch(state, page_script, vals)+startup_errors(state,state_path.parent)+production_errors(state,state_path.parent,vals["EXPECTED_BODY_PAGES"])+checkpoint_modules_errors(state,state_path.parent,stage,document)
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
            "report_file": str(report_path.resolve()),
            "report_sha256": sha256(report_path),
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
    selected=state["plan"]["checkpoint_module_ids"][str(stage)]
    projected=forecast_remaining(state,selected,actual)
    target=vals["EXPECTED_BODY_PAGES"]; tolerance=vals["PAGE_TOLERANCE"]
    within = lo <= actual <= hi and target-tolerance <= projected <= target+tolerance
    if source == "microsoft_word":
        status = "PASS" if within else "FAIL"
        phase = f"CHECKPOINT_{stage}_{status}"
        rc = 0 if within else 2
    else:
        status = "PROVISIONAL" if within else "FAIL"
        phase = f"CHECKPOINT_{stage}_{status}"
        rc = 0 if within else 2

    cp = {
        "module_ids": selected,
        "module_source_hashes": {mid:state["module_registry"][mid]["source_sha256"] for mid in selected},
        "projected_final_body_pages": projected,
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
    print(f"PROJECTED_FINAL_BODY_PAGES={projected}")

    if status == "PASS":
        print("DRAFTING_ALLOWED")
    elif status == "PROVISIONAL":
        print("DRAFTING_ALLOWED_WITH_ESTIMATE")
        print("WORD_REMEASUREMENT_REQUIRED_BEFORE_FINAL")
    else:
        print("DRAFTING_BLOCKED_UNTIL_ADJUSTED")
    return rc


def cmd_final_ready(state_path: Path, document=None):
    state = load_state(state_path)
    errors = startup_errors(state,state_path.parent)
    page_script,vals=discover_page_script()
    if page_script:
        errors += _rule_mismatch(state,page_script,vals)

    errors += production_errors(state,state_path.parent,vals["EXPECTED_BODY_PAGES"] if vals else 0)
    if document is None:
        errors.append("final-ready须提供实际合并后的DOCX")
    else:
        errors += registry_errors(state,state_path.parent,document=document)
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
        if any(state.get("module_registry",{}).get(mid,{}).get("source_sha256")!=h for mid,h in cp.get("module_source_hashes",{}).items()):
            errors.append(f"{stage}%阶段模块源已改变，须重新测量")
        wr=cp.get("word_report") or {}
        if not fresh_file({"path":wr.get("report_file"),"sha256":wr.get("report_sha256")},state_path.parent):
            errors.append(f"{stage}% Word测量报告丢失或已改变")
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
            'word_channel':'', 'word_capability_evidence':'', 'real_photo_slots':[],
            'visual_policy': visual_policy_template()}


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
    slots = a.get('real_photo_slots')
    errors += visual_policy_errors(a.get('visual_policy'), slots)
    if isinstance(slots, list) and any(not isinstance(x, dict) or not x.get('name') or not x.get('location') or not x.get('basis') for x in slots):
        errors.append('已选实景图须登记主题、插入位置及价值依据')
    return errors


def emit_gate(errors, success):
    print('FLOW_GATE_STATUS='+('FAIL' if errors else 'PASS'))
    print('DELIVERY_BLOCKED' if errors else success)
    for e in errors:
        print('- '+e)
    return 2 if errors else 0


TECHNICAL_REVIEW_ITEMS = (
    'applicability', 'constraints', 'number_basis',
    'quantity_capacity_schedule_resources', 'issue_recheck',
)
REFERENCE_TARGET_ITEMS = (
    'length_38000',
    'numbers_6_per_1000', 'charts_15',
)


def technical_review_errors(record, final=False):
    """Check execution records, never infer engineering quality or award points."""
    review = record.get('technical_review')
    if not isinstance(review, dict) or type(review.get('version')) is not int or review['version'] != 1:
        return ['缺五项专业复核technical_review v1（复用原记录，不另开阶段）']
    checks = review.get('checks')
    errors = []
    if not isinstance(checks, dict):
        checks = {}
    for key in TECHNICAL_REVIEW_ITEMS:
        item = checks.get(key)
        if not isinstance(item, dict) or item.get('status') not in ('reviewed', 'not_applicable') or not isinstance(item.get('conclusion'), str) or not item['conclusion'].strip():
            errors.append('专业复核漏项或缺实际结论/不适用依据：' + key)
    if final:
        refs = review.get('reference_targets')
        if not isinstance(refs, dict):
            refs = {}
        for key in REFERENCE_TARGET_ITEMS:
            item = refs.get(key)
            if not isinstance(item, dict) or item.get('status') not in ('met', 'adapted', 'not_met') or not isinstance(item.get('conclusion'), str) or not item['conclusion'].strip():
                errors.append('经验参考未留采用/调整回应（数值本身不作门槛）：' + key)
    return errors


def handled_conflict_errors(item, body_text, base):
    """A handled drafting conflict is not an official source reconciliation."""
    errors = []
    name = str(item.get('name'))
    if item.get('handling') not in ('bill_primary', 'avoid_detail'):
        errors.append('编制处置缺清单为主/收敛争议的明确方式：' + name)
    if not all(isinstance(item.get(k), str) and item[k].strip() for k in ('source_status', 'decision', 'recheck')):
        errors.append('编制处置缺原始资料状态、采用口径或关联复验：' + name)
    evidence = item.get('occurrences') if 'occurrences' in item else item.get('evidence')
    if not isinstance(evidence, list) or not evidence:
        errors.append('编制处置缺真实来源证据：' + name)
        evidence = []
    if any(not isinstance(v, dict) or not v.get('locator') or not fresh_file(v, base) for v in evidence):
        errors.append('编制处置来源文件/哈希/位置不完整：' + name)
    if item.get('handling') == 'bill_primary' and not any(isinstance(v, dict) and v.get('role') == 'bill' and v.get('locator') and fresh_file(v, base) for v in evidence):
        errors.append('清单口径缺role=bill的有效清单证据：' + name)
    uses = item.get('body_locations')
    if not isinstance(uses, list) or not uses or any(not isinstance(v, dict) or not v.get('location') or not v.get('quote') or v['quote'] not in body_text for v in uses):
        errors.append('编制处置缺当前正文措施及定位：' + name)
    return errors


def validate_content_review(document, review_path):
    r = json.loads(review_path.read_text(encoding='utf-8'))
    errors = technical_review_errors(r, final=True)
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
        if x.get('kind') not in ('project_fact', 'standard_value', 'calculated', 'proposed'):
            errors.append('关键参数未区分项目事实/规范值/计算/策划：'+str(x.get('name')))
        if x.get('kind') == 'calculated' and not x.get('calculation'):
            errors.append('计算参数缺公式、输入及单位依据：'+str(x.get('name')))
        if x.get('kind') == 'proposed' and not x.get('feasibility_basis'):
            errors.append('拟采用参数缺适用条件与可行性依据：'+str(x.get('name')))
        if not x.get('name') or not x.get('scope') or not x.get('cross_source_search') or not candidates:
            errors.append('关键参数跨来源搜索记录不足')
        for v in candidates:
            if v.get('value') is None or not all(v.get(k) for k in ('unit','scope','locator')) or not fresh_file(v,review_path.parent):
                errors.append('关键参数候选值/部位/原图证据不完整：'+str(x.get('name')))
        if x.get('status') not in ('consistent','resolved','handled') or not x.get('decision') or not x.get('body_locations'):
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
        if x.get('status') == 'handled':
            errors += handled_conflict_errors(x, body_text, review_path.parent)
            if x.get('handling') != 'bill_primary':
                errors.append('正文已使用的数值不能以收敛争议代替取值依据：'+str(x.get('name')))
            bill_values = {(str(v.get('value')),v.get('unit'),v.get('scope')) for v in candidates if v.get('role') == 'bill'}
            for use in x.get('body_locations') or []:
                if isinstance(use, dict) and (str(use.get('value')),use.get('unit'),use.get('scope')) not in bill_values:
                    errors.append('正文处理值与有效清单口径不匹配：'+str(x.get('name')))
        # Differing values require an explicit scope/version resolution, not a generic PASS.
        values={(str(v.get('value')),v.get('unit')) for v in candidates}
        if len(values)>1 and x.get('status') not in ('resolved','handled'):
            errors.append('不同候选值尚未解释适用部位/有效版本：'+str(x.get('name')))
    for x in r.get('conflicts',[]):
        if x.get('status') == 'handled':
            errors += handled_conflict_errors(x, body_text, review_path.parent)
        elif x.get('status')!='resolved' or not x.get('evidence') or not x.get('decision'):
            errors.append('未关闭冲突：'+str(x.get('name')))
    if (r.get('construction_language_review') or {}).get('all_paragraphs_reviewed') is not True or not (r.get('construction_language_review') or {}).get('conclusion'):
        errors.append('未完成逐段施工执行语言复核')
    if any((r.get('visual_review') or {}).get(k) is not True for k in ('three_color_figures_reviewed','caption_photo_reviewed')) or not (r.get('visual_review') or {}).get('conclusion'):
        errors.append('未完成按有效配色要求的三图、增补/替代示意图及适用照片框实际视觉复核')
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
        if cmd_final_ready(a.state,a.document)!=0:
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
    fresh_checks = common.portable_checks(a.document,rules)
    for check in fresh_checks:
        if check.status=='FAIL':
            errors.append(check.name+': '+check.message)
    # 新增语境WARN不能借旧静态报告逃过现有WARN处置门禁。
    fresh_warnings=[vars(c) for c in fresh_checks if c.status in ('WARN','SKIP')]
    warning_names={x['name'] for x in fresh_warnings}
    static['checks']=[x for x in static.get('checks',[]) if x.get('name') not in warning_names]+fresh_warnings
    stage=rules.get('photo_delivery_stage')
    if stage not in ('draft','final') or static.get('photo_delivery_stage')!=stage:
        errors.append('照片交付阶段未一致确认')
    slots=rules.get('required_real_photo_slots') or []
    if state:
        if state['startup'].get('visual_policy') != rules.get('visual_policy'):
            errors.append('启动与交付图文决策不一致；按当前有效要求同步更新项目判断及图件计划')
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


def cmd_run_page_audit(document, report_path, stage=None, pilot=False):
    script, vals=discover_page_script()
    if script is None:
        raise RuntimeError('未找到唯一项目页数硬审脚本')
    if stage is None and not pilot:
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
                'measurement_role':'pilot' if pilot else 'checkpoint',
                'measured_at':now_iso(),'project_rule_source':str(script.resolve()),
                'project_rule':{'target_body_pages':vals['EXPECTED_BODY_PAGES'],'page_tolerance':vals['PAGE_TOLERANCE'],'front_matter_pages_excluded':front},
                'actual':{'word_total_pages':total,'word_body_pages':total-front,'word_version':str(word.Version)}}
        if pilot:
            import re
            start=doc.GoTo(What=1,Which=1,Count=front+1).Start
            body=doc.Range(start,doc.Content.End)
            count=lambda txt:len(re.sub(r'\s+','',str(txt).replace('\x07','')))
            table_intervals=[]
            for i in range(1,int(body.Tables.Count)+1):
                r=body.Tables(i).Range
                table_intervals.append((max(start,int(r.Start)),min(int(body.End),int(r.End))))
            merged=[]
            for lo,hi in sorted(table_intervals):
                if merged and lo<=merged[-1][1]:merged[-1]=(merged[-1][0],max(merged[-1][1],hi))
                else:merged.append((lo,hi))
            tables=sum(count(doc.Range(lo,hi).Text) for lo,hi in merged if hi>lo)
            total_chars=count(body.Text)
            report['actual'].update(word_body_characters=total_chars,word_table_characters=tables,word_normal_text_characters=max(0,total_chars-tables))
            evidence=report_path.with_suffix('.word.pdf')
            evidence.parent.mkdir(parents=True,exist_ok=True)
            doc.ExportAsFixedFormat(str(evidence.resolve()),17)
            report['word_pdf_evidence']={'path':evidence.name,'sha256':sha256(evidence)}
        save_state(report_path,report)
        print('WORD_STAGE_STATUS=MEASURED\nPAGE_SOURCE=microsoft_word\nREPORT='+str(report_path))
        return 0
    finally:
        if doc is not None:doc.Close(False)
        if word is not None:word.Quit()


# Detailed planning and module persistence are generic; the page script alone supplies targets.
def planning_policy(target):
    return {'pilot_required':target>=200,
            'module_page_cap':max(12,min(30,math.ceil(target*.05))) if target else 20,
            'pilot_page_range':[max(6,min(12,math.ceil(target*.02))),max(8,min(20,math.ceil(target*.04)))] if target else [6,12]}


def plan_hash(state):
    import hashlib
    plan={k:v for k,v in (state.get('plan') or {}).items() if k!='status'}
    return hashlib.sha256(json.dumps(plan,ensure_ascii=False,sort_keys=True).encode()).hexdigest()


def _modules(state):
    return (state.get('plan') or {}).get('modules') or []


def _ids(state):
    return [x.get('id') for x in _modules(state)]


def _normal(s):
    import re
    return re.sub(r'\s+','',s)


def _manuscript(path):
    if path.suffix.lower() in ('.docx','.docm'):
        import zipfile
        import xml.etree.ElementTree as ET
        ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        with zipfile.ZipFile(path) as z:tree=ET.fromstring(z.read('word/document.xml'))
        body=tree.find('w:body',ns)
        full='\n'.join(''.join(t.text or '' for t in node.findall('.//w:t',ns)) for node in body)
        normal='\n'.join(''.join(t.text or '' for t in node.findall('.//w:t',ns)) for node in body if node.tag.endswith('}p'))
        return full,len(_normal(normal))
    if path.suffix.lower() not in ('.md','.txt'):
        raise RuntimeError('模块正文源只允许md/txt/docx，不用生成代码或状态JSON冒充正文')
    text=path.read_text(encoding='utf-8-sig')
    import re
    lines=[re.sub(r'^\s*#{1,6}\s+','',line) for line in text.splitlines() if not line.lstrip().startswith(('|','!['))]
    return text,len(_normal('\n'.join(lines)))


REVIEW_CONTRACT = '2026-10-22'


def module_requirements_digest(state,module):
    import hashlib
    reqs=[r for r in state['plan']['requirements'] if r['id'] in module['requirement_ids']]
    return hashlib.sha256(json.dumps(reqs,sort_keys=True,ensure_ascii=False).encode()).hexdigest()


def module_review_errors(state,module,text,review,base):
    errors=technical_review_errors(review)
    language=review.get('construction_language_review') or {}
    if language.get('reviewed') is not True or (type(language.get('unresolved_count')) is not int or language.get('unresolved_count') != 0) or not language.get('conclusion'):
        errors.append('模块施工语言未复核或仍有内部审查语言待处理')
    scoring=review.get('scoring_review') or {}
    reqs={r['id']:r for r in state['plan']['requirements'] if r['id'] in module['requirement_ids']}
    expected={rid for rid,r in reqs.items() if r.get('scoring') is True}
    conditions=scoring.get('conditions')
    if not isinstance(conditions,list) or not scoring.get('conclusion'):
        return errors+['模块缺优秀条件轻量复核scoring_review及结论']
    ids=[x.get('requirement_id') for x in conditions if isinstance(x,dict)]
    if len(ids)!=len(conditions) or len(ids)!=len(set(ids)) or not set(ids).issubset(reqs) or not expected.issubset(ids):
        errors.append('模块评分条件复核缺项、重复或引用非本模块责任')
    if not conditions and not scoring.get('not_applicable_reason'):
        errors.append('无评分责任须说明不适用依据，不得省略已有评分条件')
    for x in conditions:
        if not isinstance(x,dict):continue
        rid=x.get('requirement_id');req=reqs.get(rid,{})
        if req.get('scoring') is False:errors.append('普通施工需求不能冒充招标评分条件：'+str(rid))
        if x.get('criterion_quote')!=req.get('scope'):
            errors.append('优秀条件须对应需求中保留的评分原文：'+str(rid))
        for key in ('targeted','technical_depth','project_basis','substantive_response'):
            if x.get(key) is not True:errors.append('评分措施缺项或深度不足须先补强：'+str(rid)+'/'+key)
        evidence=x.get('evidence') or []
        if not evidence or any(not e.get('locator') or not fresh_file(e,base) for e in evidence):
            errors.append('优秀条件复核缺真实项目/招标依据：'+str(rid))
        quote=x.get('response_quote','')
        if not isinstance(quote,str) or len(_normal(quote))<20 or _normal(quote) not in _normal(text):
            errors.append('优秀条件复核缺当前模块施工措施原文：'+str(rid))
        if x.get('status') not in ('satisfied','deferred') or not x.get('conclusion'):
            errors.append('优秀条件复核缺状态或判断依据：'+str(rid))
        if x.get('status')=='deferred' and not all(x.get(k) for k in ('remaining_condition','responsible_module_id')):
            errors.append('跨模块未完成条件须登记剩余内容及责任模块：'+str(rid))
        if x.get('status')=='deferred' and x.get('responsible_module_id') not in _ids(state):
            errors.append('未完成条件引用未知责任模块：'+str(rid))
        if 'self_score' in x or 'self_score' in scoring:
            errors.append('模块轻量复核不强行评分，模拟评分沿用40/75/final清单')
    return errors


def registry_errors(state,base,ids=None,document=None):
    errors=[]; registry=state.get('module_registry') or {}
    selected=ids if ids is not None else _ids(state)
    word_text=_normal(_manuscript(document)[0]) if document else None
    for mid in selected:
        row=registry.get(mid)
        if not row or row.get('status')!='complete':
            errors.append('模块未完成并保存：'+str(mid));continue
        if not fresh_file({'path':row.get('source_path'),'sha256':row.get('source_sha256')},base):
            errors.append('模块正文源丢失或改变：'+mid)
        if not fresh_file({'path':row.get('review_path'),'sha256':row.get('review_sha256')},base):
            errors.append('模块核查记录丢失或改变：'+mid)
        module=next((m for m in _modules(state) if m.get('id')==mid),None)
        if row.get('review_contract') == REVIEW_CONTRACT and module and row.get('requirements_sha256') != module_requirements_digest(state,module):
            errors.append('模块评分/范围依据已改变，须定点复核：'+mid)
        if word_text is not None:
            quotes=row.get('probe_quotes') or []
            if len(quotes)<2 or any(_normal(q) not in word_text for q in quotes):
                errors.append('合并DOCX中缺少模块正文探针：'+mid)
            if fresh_file({'path':row.get('source_path'),'sha256':row.get('source_sha256')},base):
                import re
                raw=_manuscript(Path(row['source_path']))[0]
                content=re.sub(r'[^\w]','',raw)
                actual=re.sub(r'[^\w]','',_manuscript(document)[0])
                chunks=[content[i:i+100] for i in range(0,len(content),100)]
                total=sum(map(len,chunks))
                present=sum(len(x) for x in chunks if x in actual)
                if total and present/total<.90:
                    errors.append('合并DOCX中模块实质内容不足90%，不能只保留探针：'+mid)
    return errors


def _prefix_record(state,base):
    # A measured failed checkpoint remains useful evidence for replanning, never a PASS.
    all_ids=_ids(state); registry=state.get('module_registry') or {}
    for stage in reversed(STAGES):
        cp=(state.get('checkpoints') or {}).get(str(stage)) or {}
        mids=cp.get('module_ids') or []
        if cp.get('page_source')!='microsoft_word' or not mids or mids!=all_ids[:len(mids)]:continue
        if not fresh_file({'path':cp.get('snapshot_path'),'sha256':cp.get('document_sha256')},base):continue
        if any(registry.get(mid,{}).get('source_sha256')!=(cp.get('module_source_hashes') or {}).get(mid) for mid in mids):continue
        wr=cp.get('word_report') or {}
        if not fresh_file({'path':wr.get('report_file'),'sha256':wr.get('report_sha256')},base):continue
        if registry_errors(state,base,mids):continue
        return cp
    return None


def effective_budget(state,base):
    modules=_modules(state); cp=_prefix_record(state,base)
    measured_ids=set(cp.get('module_ids',[])) if cp else set()
    total=float(cp['actual_pages']) if cp else 0.0
    total+=sum(float(x['planned_pages']) for x in modules if x.get('id') not in measured_ids)
    return total+float(state['plan']['reserve_pages'])


def detailed_plan_errors(state,base,target):
    errors=[];p=state.get('plan') or {};mods=p.get('modules');reqs=p.get('requirements')
    if type(p.get('revision')) is not int or p['revision']<1:errors.append('规划须有正整数revision')
    if not isinstance(reqs,list) or not reqs:
        errors.append('缺少评分/范围/强制响应的内容需求清单requirements');reqs=[]
    req_ids=[]
    for req in reqs:
        if not isinstance(req,dict) or not req.get('id') or not req.get('scope'):
            errors.append('内容需求缺id或实质scope');continue
        req_ids.append(req['id']);evidence=req.get('evidence') or []
        if state.get('schema_version') == '3.1' and type(req.get('scoring')) is not bool:
            errors.append('内容需求须由AI标记是否承担评分条件scoring：'+req['id'])
        if not evidence or any(not x.get('locator') or not fresh_file(x,base) for x in evidence):errors.append('需求缺真实资料/位置证据：'+req['id'])
    if len(req_ids)!=len(set(req_ids)):errors.append('内容需求ID重复')
    if not isinstance(mods,list) or not mods:
        errors.append('缺少章节—节—内容模块规划modules');return errors
    capacity=p.get('capacity_review') or {}
    for key in ('all_requirements_mapped','scope_supports_volume','no_padding_or_repetition'):
        if capacity.get(key) is not True:errors.append('内容容量专业判断尚未完成：'+key)
    if not capacity.get('conclusion') or capacity.get('unresolved_gaps')!=[]:errors.append('内容容量缺结论或仍有资料缺口，先质询/关闭')
    cap=planning_policy(target)['module_page_cap'];mids=[];covered=set();primary=[];figures=0.0
    chapters={x.get('chapter') for x in p.get('chapter_plan',[]) if isinstance(x,dict)}
    for m in mods:
        if not isinstance(m,dict):errors.append('模块不是对象');continue
        mid=m.get('id');mids.append(mid)
        if not mid or not m.get('chapter') or not m.get('section') or not m.get('scope') or not m.get('basis'):errors.append('模块缺id/章/节/实质内容/分配依据')
        if m.get('chapter') not in chapters:errors.append('模块未归入真实章节：'+str(mid))
        rids=m.get('requirement_ids') or []
        if not rids or not set(rids).issubset(set(req_ids)):errors.append('模块需求映射缺失或引用未知需求：'+str(mid))
        covered.update(rids)
        owners=m.get('primary_for') or []
        if not set(owners).issubset(set(rids)):errors.append('主响应责任必须属于本模块需求：'+str(mid))
        primary.extend(owners)
        for key in ('text_pages','table_pages','figure_pages','estimated_normal_chars','planned_pages') if target else ():
            if not is_num(m.get(key)) or m[key]<0:errors.append('模块预算字段无效：'+str(mid)+'/'+key)
        if all(is_num(m.get(k)) for k in ('text_pages','table_pages','figure_pages','planned_pages')):
            if not 0<m['planned_pages']<=cap:errors.append(f'模块{mid}须拆成不超过{cap}页的实质单元')
            if abs(m['text_pages']+m['table_pages']+m['figure_pages']-m['planned_pages'])>.05:errors.append('模块正文/表格/图片重复计页或总量不符：'+str(mid))
            figures+=m['table_pages']+m['figure_pages']
        if is_num(m.get('text_pages')) and is_num(m.get('estimated_normal_chars')):
            if (m['text_pages']>0)!=(m['estimated_normal_chars']>0):errors.append('正文页数与字符量估计不对应：'+str(mid))
        evidence=m.get('evidence') or []
        if not evidence or any(not x.get('locator') or not fresh_file(x,base) for x in evidence):errors.append('模块内容缺资料支撑：'+str(mid))
    if target:
        for chapter in p.get('chapter_plan',[]):
            grouped=[m for m in mods if m.get('chapter')==chapter.get('chapter')]
            if all(is_num(m.get('planned_pages')) for m in grouped):
                total=sum(m['planned_pages'] for m in grouped)
                rng=chapter.get('page_range') or []
                if len(rng)==2 and all(is_num(v) for v in rng) and not rng[0]<=total<=rng[1]:
                    errors.append('章节区间未涵盖所属模块预算：'+str(chapter.get('chapter')))
    if len(mids)!=len(set(mids)):errors.append('模块ID重复')
    if set(req_ids)!=covered:errors.append('评分/范围/强制要求没有全部落入内容模块')
    if sorted(primary)!=sorted(req_ids):errors.append('每项内容需求须有且仅有一个主响应模块')
    if is_num(p.get('figure_table_pages_estimate')) and abs(figures-p['figure_table_pages_estimate'])>.05:errors.append('全篇图表估算与模块图表之和不符')
    if not errors and is_num(p.get('reserve_pages')) and is_num(p.get('planned_total_body_pages')):
        if abs(effective_budget(state,base)-p['planned_total_body_pages'])>.05:errors.append('全篇规划须等于已实测前缀＋剩余模块＋机动量，不能仅声明目标总页数')
    prefix=[]
    for stage in STAGES if target else ():
        selected=(p.get('checkpoint_module_ids') or {}).get(str(stage))
        if not isinstance(selected,list) or not selected or selected!=mids[:len(selected)]:
            errors.append(f'{stage}%必须规划连续的模块前缀');continue
        if stage==75 and (len(selected)<=len(prefix) or selected==mids):errors.append('75%节点须晚于40%且保留后续内容')
        prefix=selected
        if not p.get('checkpoint_completion_basis',{}).get(str(stage)):errors.append(f'{stage}%缺按章节/评分责任判断完成度的依据')
        if all(is_num(m.get('planned_pages')) for m in mods):
            cp=_prefix_record(state,base)
            center=float(cp['actual_pages']) if cp and cp.get('module_ids')==selected else sum(m['planned_pages'] for m in mods if m['id'] in selected)
            slack=max(1,math.ceil(center*.05))
            expected=[round(max(1,center-slack),2),round(center+slack,2)]
            rng=p.get('checkpoint_expected_page_ranges',{}).get(str(stage))
            if not isinstance(rng,list) or len(rng)!=2 or not all(is_num(v) for v in rng) or not expected[0]<=rng[0]<=rng[1]<=expected[1]:errors.append(f'{stage}%范围须依据模块预算/真实实测且不得任意放宽；最大范围{expected}')
    if planning_policy(target)['pilot_required']:
        pilot=p.get('pilot_plan') or {}; selected=pilot.get('module_ids') or []
        if len(set(selected))<2 or not set(selected).issubset(set(mids)):errors.append('长文样章须取至少两个真实模块的代表内容')
        if set(pilot.get('coverage',[]))!={'narrative','table','figure'} or not pilot.get('basis'):errors.append('样章须覆盖正文、表格、图片并说明代表性')
    return errors


def pilot_errors(state,base,target):
    if not planning_policy(target)['pilot_required']:return []
    row=state.get('pilot') or {};errors=[]
    if row.get('status')!='PASS' or row.get('page_source')!='microsoft_word':return ['长文尚无Microsoft Word样章试排PASS']
    for key in ('snapshot','report','analysis','format_rules'):
        item=row.get(key) or {}
        if not fresh_file(item,base):errors.append('样章证据失效：'+key)
    if row.get('page_rule')!=state.get('page_rule'):errors.append('样章页数来源与当前项目不一致')
    if row.get('module_ids')!=(state.get('plan') or {}).get('pilot_plan',{}).get('module_ids'):errors.append('样章代表模块改变，需重新确认/试排')
    for item in row.get('visual_evidence',[]):
        if not fresh_file(item,base):errors.append('样章Word导出证据改变或丢失')
    density=row.get('normal_chars_per_page')
    if not is_num(density) or density<=0:errors.append('样章缺有效文字密度')
    else:
        for m in _modules(state):
            if abs(m.get('estimated_normal_chars',0)/density-m.get('text_pages',0))>max(.15,m.get('text_pages',0)*.03):errors.append('模块字符量与Word样章密度不符，先重算：'+str(m.get('id')))
    return errors


def production_errors(state,base,target):
    errors=[]
    if (state.get('gate') or {}).get('preflight')!='PASS':errors.append('预编制规划未通过')
    if (state.get('gate') or {}).get('draft_ready')!='PASS':errors.append('draft-ready未通过，当前仅允许规划或样章')
    if (state.get('gate') or {}).get('plan_hash')!=plan_hash(state):errors.append('规划改变后未重新preflight/draft-ready')
    errors+=pilot_errors(state,base,target)
    return errors


def cmd_pilot(state_path,document,report_path,analysis_path):
    state=load_state(state_path);script,vals=discover_page_script()
    if not script:raise RuntimeError('样章试排须有项目页数来源')
    errors=startup_errors(state,state_path.parent)+_rule_mismatch(state,script,vals)
    if state.get('gate',{}).get('preflight')!='PASS' or state.get('gate',{}).get('plan_hash')!=plan_hash(state):errors.append('样章前规划未真实通过或已改变')
    report=json.loads(report_path.read_text(encoding='utf-8-sig'));analysis=json.loads(analysis_path.read_text(encoding='utf-8-sig'))
    if report.get('page_source')!='microsoft_word' or report.get('status')!='MEASURED' or report.get('measurement_role')!='pilot' or report.get('document_sha256')!=sha256(document):errors.append('样章须用绑定当前DOCX的Word实际报告，禁止估值冒充')
    expected={'target_body_pages':vals['EXPECTED_BODY_PAGES'],'page_tolerance':vals['PAGE_TOLERANCE'],'front_matter_pages_excluded':vals['FRONT_MATTER_PAGES']}
    if any((report.get('project_rule') or {}).get(k)!=v for k,v in expected.items()):errors.append('样章报告项目页数规则不符')
    pages=(report.get('actual') or {}).get('word_body_pages'); chars=(report.get('actual') or {}).get('word_normal_text_characters')
    lo,hi=planning_policy(vals['EXPECTED_BODY_PAGES'])['pilot_page_range']
    if not is_num(pages) or not lo<=pages<=hi:errors.append(f'代表性样章正文页数应在{lo}–{hi}，不足或过大须调整样章')
    if not is_num(chars) or chars<=0:errors.append('样章报告缺Word测得的非表格文字量')
    if analysis.get('document_sha256')!=sha256(document) or analysis.get('module_ids')!=(state['plan'].get('pilot_plan') or {}).get('module_ids'):errors.append('样章分析未绑定快照和代表模块')
    format_item=analysis.get('format_rules') or {}
    if not fresh_file(format_item,analysis_path.parent):errors.append('样章缺实际项目格式规则快照')
    for key in ('font_and_spacing_reviewed','table_layout_reviewed','figure_readability_reviewed','photo_frame_layout_reviewed'):
        if analysis.get(key) is not True:errors.append('未完成样章实际版式复核：'+key)
    exported=report.get('word_pdf_evidence') or {}
    if not fresh_file(exported,report_path.parent):errors.append('缺样章测量器导出的实际Word PDF')
    visuals=analysis.get('visual_evidence') or []
    if not visuals or any(not fresh_file(x,analysis_path.parent) for x in visuals):errors.append('缺Word导出PDF或页面证据文件')
    if exported.get('sha256') not in [x.get('sha256') for x in visuals]:errors.append('样章视觉核验未引用本次Word实际导出PDF')
    component_pages=[analysis.get(k) for k in ('normal_text_pages','table_pages','figure_pages')]
    if not all(is_num(x) and x>0 for x in component_pages) or not is_num(pages) or abs(sum(component_pages)-pages)>.1:errors.append('样章正文/表格/图片空间估算须与实测总页数相容')
    import zipfile,xml.etree.ElementTree as ET
    with zipfile.ZipFile(document) as z:xml=ET.fromstring(z.read('word/document.xml'))
    if not any(x.tag.endswith('}tbl') for x in xml.iter()) or not any(x.tag.endswith('}blip') for x in xml.iter()):errors.append('样章实际缺表格或嵌入图片')
    if errors:return emit_gate(errors,'')
    density=chars/analysis['normal_text_pages']
    state['pilot']={'visual_evidence':[{**item,'path':str((analysis_path.parent/item['path']).resolve())} for item in visuals],'status':'PASS','page_source':'microsoft_word','page_rule':state['page_rule'],'module_ids':analysis['module_ids'],
                    'normal_chars_per_page':density,'actual_body_pages':pages,
                    **{key:{'path':str(path.resolve()),'sha256':sha256(path)} for key,path in [('snapshot',document),('report',report_path),('analysis',analysis_path)]},
                    'format_rules':{'path':str((analysis_path.parent/format_item['path']).resolve()) if not Path(format_item['path']).is_absolute() else format_item['path'],'sha256':format_item['sha256']}}
    state['gate']['draft_ready']='BLOCKED';state['phase']='PILOT_BUDGET_CALIBRATION';save_state(state_path,state)
    print('PILOT_WORD_STATUS=PASS\nNORMAL_CHARS_PER_PAGE='+str(round(density,2)))
    print('RECALCULATE_MODULE_BUDGETS_THEN_DRAFT_READY')
    return 0


def cmd_draft_ready(state_path):
    state=load_state(state_path);script,vals=discover_page_script();target=vals['EXPECTED_BODY_PAGES'] if vals else 0
    errors=startup_errors(state,state_path.parent)
    if script:errors+=_rule_mismatch(state,script,vals)
    if state.get('gate',{}).get('preflight')!='PASS' or state.get('gate',{}).get('plan_hash')!=plan_hash(state):errors.append('须先完成当前版本preflight')
    errors+=detailed_plan_errors(state,state_path.parent,target)+pilot_errors(state,state_path.parent,target)
    state['gate']['draft_ready']='FAIL' if errors else 'PASS';state['phase']='PLANNING' if errors else 'DRAFTING_ALLOWED';save_state(state_path,state)
    return emit_gate(errors,'PRE_DRAFT_ALLOWED')


def cmd_module_checkpoint(state_path,mid,source_path,review_path):
    state=load_state(state_path);script,vals=discover_page_script();target=vals['EXPECTED_BODY_PAGES'] if vals else 0
    errors=production_errors(state,state_path.parent,target)+startup_errors(state,state_path.parent)
    if any(x.get('status') in ('FAIL','STALE') for x in state.get('checkpoints',{}).values()) and not (mid in state.get('module_registry',{}) and state['module_registry'][mid].get('source_sha256')!=sha256(source_path)):errors.append('阶段校准FAIL尚未返修复测，禁止盲写下一模块；允许更新已保存的实际返修模块')
    module=next((x for x in _modules(state) if x.get('id')==mid),None)
    if not module:errors.append('未知模块ID')
    if module and mid not in state.get('module_registry',{}):
        order=_ids(state);position=order.index(mid)
        errors+=registry_errors(state,state_path.parent,order[:position])
        for stage in STAGES if target else ():
            prefix=state['plan']['checkpoint_module_ids'][str(stage)]
            if position>=len(prefix):
                cp=state.get('checkpoints',{}).get(str(stage)) or {}
                if cp.get('status') not in ('PASS','PROVISIONAL') or cp.get('module_ids')!=prefix:
                    errors.append(f'已到{stage}%节点，先完成真实阶段校准再写后续模块')
                elif not fresh_file({'path':cp.get('snapshot_path'),'sha256':cp.get('document_sha256')},state_path.parent):
                    errors.append(f'{stage}%阶段快照失效')
                elif cp.get('page_source')=='microsoft_word':
                    wr=cp.get('word_report') or {}
                    if not fresh_file({'path':wr.get('report_file'),'sha256':wr.get('report_sha256')},state_path.parent):errors.append(f'{stage}%阶段Word报告失效')
    text,normal_chars=_manuscript(source_path);review=json.loads(review_path.read_text(encoding='utf-8-sig'))
    if len(_normal(text))<200:errors.append('模块正文源过少，不能用标题或占位冒充完成')
    if review.get('module_id')!=mid or review.get('source_sha256')!=sha256(source_path):errors.append('模块复核记录未绑定实际正文源')
    for key in ('fact_sources_reviewed','technical_depth_reviewed','interfaces_reviewed','repetition_reviewed'):
        if review.get(key) is not True:errors.append('模块实质复核缺项：'+key)
    if module:
        errors += module_review_errors(state,module,text,review,review_path.parent)
    if not review.get('conclusion'):errors.append('模块复核缺结论')
    if module and set(review.get('covered_requirement_ids') or [])!=set(module['requirement_ids']):errors.append('模块没有完成其评分/范围责任')
    quotes=review.get('probe_quotes') or []
    if len(quotes)<2 or any(len(_normal(q))<40 or _normal(q) not in _normal(text) for q in quotes):errors.append('须有两段确实来自正文的合并探针，每段至少40个非空白字符')
    if len(set(map(_normal,quotes)))!=len(quotes):errors.append('正文探针重复')
    if module and module.get('estimated_normal_chars',0)>0 and planning_policy(target)['pilot_required']:
        ratio=normal_chars/module['estimated_normal_chars']
        if not .6<=ratio<=1.5:errors.append('实际模块文字量明显偏离经样章校准的预算，须检查内容并调整规划')
    for key,row in state.get('module_registry',{}).items():
        if key!=mid and row.get('source_sha256')==sha256(source_path):errors.append('不同模块不能重复使用完全相同正文')
    if errors:return emit_gate(errors,'')
    state.setdefault('module_registry',{})[mid]={'status':'complete','source_path':str(source_path.resolve()),'source_sha256':sha256(source_path),
        'review_contract':REVIEW_CONTRACT,'requirements_sha256':module_requirements_digest(state,module),
        'review_path':str(review_path.resolve()),'review_sha256':sha256(review_path),'normal_characters':normal_chars,'probe_quotes':quotes,'saved_at':now_iso()}
    for cp in state.get('checkpoints',{}).values():
        if mid in cp.get('module_ids',[]) and cp.get('module_source_hashes',{}).get(mid)!=sha256(source_path):
            cp['status']='STALE'
            cp['note']='模块正文更新，旧阶段数据失效，须重新Word测量'
    state['gate']['final_ready']='BLOCKED'
    density=(state.get('pilot') or {}).get('normal_chars_per_page')
    if target and is_num(density) and density>0:
        cp=_prefix_record(state,state_path.parent)
        measured_ids=set(cp.get('module_ids',[])) if cp else set()
        estimated=float(cp['actual_pages']) if cp else 0.0
        for m in _modules(state):
            if m['id'] in measured_ids:continue
            row=state.get('module_registry',{}).get(m['id']) or {}
            estimated+=row['normal_characters']/density+m['table_pages']+m['figure_pages'] if row.get('status')=='complete' else m['planned_pages']
        estimated+=state['plan']['reserve_pages']
        tol=vals['PAGE_TOLERANCE']
        state['rolling_forecast']={'estimated_body_pages':round(estimated,2),'basis':'Word样章文字密度＋模块实际文字量＋图表预算；非Word最终实测',
                                   'status':'REVIEW_REQUIRED' if not target-tol<=estimated<=target+tol else 'IN_BUDGET','checked_at':now_iso()}
        print('MODEL_FORECAST_BODY_PAGES='+str(round(estimated,2)))
        if state['rolling_forecast']['status']=='REVIEW_REQUIRED':print('MODEL_FORECAST_REVIEW_REQUIRED：先复核内容/图表预算，必要时提前Word测量；估算警示不替代阶段硬门禁')
    save_state(state_path,state)
    print('MODULE_SAVED='+mid)
    return 0


def checkpoint_modules_errors(state,base,stage,document):
    selected=(state.get('plan') or {}).get('checkpoint_module_ids',{}).get(str(stage)) or []
    return registry_errors(state,base,selected,document)


def forecast_remaining(state,selected,actual):
    remaining=sum(m['planned_pages'] for m in _modules(state) if m['id'] not in selected)
    return round(actual+remaining+state['plan']['reserve_pages'],3)


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
    s.add_argument("--pilot",action="store_true",help="Word样章试排，测量文字密度并导出实际PDF")

    s = sub.add_parser("checkpoint")
    s.add_argument("--stage", type=int, choices=STAGES, required=True)
    s.add_argument("--document", type=Path, required=True, help="实际阶段DOCX快照，必须与报告哈希一致")
    s.add_argument("--report", type=Path, default=None, help="本机Word或远程Runner实际产出的阶段报告")
    s.add_argument("--actual-pages", type=float, default=None, help="仅 LibreOffice 临时估值时使用")
    s.add_argument("--engine", default="", help="临时估值时应为 libreoffice_estimate")
    s.add_argument("--allow-estimate", action="store_true")
    s.add_argument("--note", default="")

    s=sub.add_parser("final-ready")
    s.add_argument("--document",type=Path,required=True)
    sub.add_parser("draft-ready")
    s=sub.add_parser("pilot")
    for key in ("document","report","analysis"):
        s.add_argument("--"+key,type=Path,required=True)
    s=sub.add_parser("module-checkpoint")
    s.add_argument("--id",required=True)
    s.add_argument("--source",type=Path,required=True)
    s.add_argument("--review",type=Path,required=True)
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
            if a.pilot and a.stage is not None:
                raise RuntimeError("--pilot与--stage不能同时使用")
            return cmd_run_page_audit(a.document,a.report,a.stage,a.pilot)
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
        if a.command == "pilot":
            return cmd_pilot(a.state,a.document,a.report,a.analysis)
        if a.command == "draft-ready":
            return cmd_draft_ready(a.state)
        if a.command == "module-checkpoint":
            return cmd_module_checkpoint(a.state,a.id,a.source,a.review)
        if a.command == "final-ready":
            return cmd_final_ready(a.state,a.document)
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
