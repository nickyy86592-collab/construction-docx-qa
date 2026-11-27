"""Project-specific constraints verified against the active source units and submitted manuscript."""
from collections import defaultdict
from decimal import Decimal, DecimalException
from datetime import date
import re

from hard_checks import AERIAL_FIGURE, source_lookup, document_rows
from source_coverage import fresh, resolve
import json
from source_coverage import normal


FIRE_TOPICS = ("动火防火","可燃物防火","临设防火","消防通行与疏散","消防水源与器材")
ENVIRONMENT_TOPICS = ("扬尘","噪声与振动","污废水","固体废弃物")
SCREENING_TOPICS = ("国家强制要求","属地监管要求","工艺","安全","岗位资格","危大工程","临电",
                    *FIRE_TOPICS,*ENVIRONMENT_TOPICS)
SITE_TOPICS = ("场地边界与作业面","出入口与运输通行","邻近设施与周边条件",
               "地质与地下管线","专业勘察与补充勘探")
SITE_CORE = SITE_TOPICS[:3]
DIFFICULTY_FIELDS = ("name","cause","impact","resources")
CONTROL_FIELDS = ("scope","stage","exception")


def template():
    return {"jurisdiction": "", "as_of": "", "score_sources": {}, "format": {"source": {}, "fonts": {}, "pages": None,
            "color": "", "anonymous": None, "identifiers": [], "deliverable": "docx"},
            "screening": {}, "obligations": [], "decisions": [], "score_effects": {},
            "annexes": {}, "crews": {}, "equipment": {},
            "site": {"conditions": {}, "visit": {"required": None, "source": {}, "reason": ""}},
            "difficulties": {"source": {}, "reason": "", "named": [], "items": []},
            "aerial": {"layout_path": "", "layout_locator": "", "source": {}}}


def meaningful(value):
    return isinstance(value, str) and value.strip() not in (
        "", "无", "待定", "待核实", "已复核", "符合要求", "符合规范", "按需配置", "适量", "按计划", "/", "—")


def posint(n):
    return type(n) is int and n > 0


def source_ok(ref, lookup, roles=None):
    if not isinstance(ref, dict) or not meaningful(ref.get("quote")):
        return False
    if roles is not None and ref.get("source_id") not in roles:
        return False
    try:
        return normal(ref["quote"]) in normal(lookup(ref["source_id"], ref["unit_id"]))
    except (KeyError, OSError, ValueError):
        return False


def numeric_match(value, text):
    return re.search(r"(?<![\d.])"+re.escape(str(value))+r"(?![\d.])", text) is not None


def catalog_errors(plan, tasks):
    errors = []
    for kind in ("crews", "equipment"):
        catalog = plan[kind]
        used = {name for t in tasks for name in t.get(kind, [])}
        for name in used:
            item = catalog.get(name, {})
            if not meaningful(item.get("basis")) or not meaningful(item.get("entry")):
                errors.append(f"{name} 缺配置依据或投入安排")
            if kind == "crews":
                trades = item.get("trades")
                if not trades or not all(meaningful(k) and posint(v) for k,v in trades.items()) or not meaningful(item.get("shift")):
                    errors.append(f"{name} 缺分工种每班人数或班次")
            elif not posint(item.get("count")) or not meaningful(item.get("spec")) or not meaningful(item.get("capacity")):
                errors.append(f"{name} 缺设备台数、型号或规格、工况能力")
    for t in tasks:
        if not t.get("crews"):
            errors.append(t["id"]+" 关键工序缺实际作业班组")
        if not meaningful(t.get("productivity_basis")):
            errors.append(f"{t['id']} 须解释班组构成与设备能力如何支撑所用工效")
    return errors


def plan_errors(state, base):
    p = state["plan"].get("bid_requirements")
    if not p:
        return ["须按r22.5补录项目强制要求、格式、资源、附表和四项专项计划；保留原项目状态"]
    lookup = source_lookup(state, base)
    try:
        errors=validate_plan(state, p, lookup)
        checklist=json.loads(resolve(state["plan"]["scoring_basis"]["path"],base).read_text("utf-8-sig"))
        tender={s["id"] for s in state["startup"]["sources"] if s.get("role") in ("tender","amendment")}
        items=checklist.get("items",[])
        if set(p["score_sources"])!={item["id"] for item in items}:
            errors.append("实际赋分项与招标满分依据不一一对应")
        for item in items:
            ref=p["score_sources"].get(item["id"],{})
            if not source_ok(ref,lookup,tender) or not numeric_match(item["max_score"],ref.get("quote","")):
                errors.append(item["id"]+" 项目满分缺招标原文依据")
        return errors
    except (KeyError, TypeError, ValueError, IndexError, AttributeError, OSError, DecimalException) as exc:
        return ["r22.5编前登记缺项或格式错误："+str(exc)]


def validate_plan(state, p, lookup):
    errors = []
    reqs = {r["id"]:r for r in state["plan"]["requirements"]}
    modules = {m["id"] for m in state["plan"]["modules"]}
    tasks = state["plan"]["hard_checks"]["schedule"]["tasks"]
    tender = {s["id"] for s in state["startup"]["sources"] if s["role"] in ("tender","amendment")}
    f = p["format"]
    if not meaningful(p["jurisdiction"]) or not source_ok(f["source"], lookup, tender):
        errors.append("须核项目所在地、有效招标原文及格式边界")
    if type(f["anonymous"]) is not bool or f["deliverable"] not in ("docx","pdf") or f["color"] not in ("","color","monochrome","custom"):
        errors.append("匿名、实际提交格式或配色决定不明确")
    if f["anonymous"] and (not f["identifiers"] or not all(meaningful(x) for x in f["identifiers"])):
        errors.append("暗标须登记投标人名称、简称及其他已知身份标识")
    for key in ("fonts","color","pages","anonymous"):
        if f[key] and not source_ok(f.get("sources",{}).get(key),lookup,tender):
            errors.append("招标硬格式缺可回查依据："+key)
    if f["color"] and f["color"] != state["startup"]["visual_policy"]["color_figures"]["mode"]:
        errors.append("图件配色与有效招标冲突")
    if f["pages"]:
        limit=f["pages"]
        if limit["scope"] not in ("body","total") or not posint(limit["max"]):
            errors.append("页数须明确正文或含前置页及有效上限")
        else:
            rule=state.get("page_rule") or {}
            limits=rule.get(limit["scope"]+"_pages_allowed_range")
            limit_high=rule.get("body_pages_target",0)+rule.get("page_tolerance",0)
            if limit["scope"]=="total":
                limit_high+=rule.get("front_matter_pages_excluded",0)
            if not rule or limit_high>limit["max"] or (limits and limits[1]>limit["max"]):
                errors.append("项目页数脚本允许超过招标上限，须修正脚本并重测")
    for role,value in f["fonts"].items():
        if role not in ("body","heading1","heading2","heading3","table","caption"):
            errors.append("未知强制字号部位："+role)
    for loc,tokens in f.get("permitted",{}).items():
        if not tokens or not set(tokens)<=set(f["identifiers"]) or not source_ok(f.get("permit_sources",{}).get(loc),lookup,tender):
            errors.append("非匿名区域例外缺具体允许条款："+loc)
    for key in dict.fromkeys([*SCREENING_TOPICS,*p["screening"]]):
        s=p["screening"].get(key,{})
        if type(s.get("applicable")) is not bool or not meaningful(s.get("reason")) or not source_ok(s.get("source"),lookup):
            errors.append("缺地域/工况筛查依据："+key)
        elif s["applicable"] and not any(key in x["topics"] for x in p["obligations"]):
            errors.append("已触发要求缺逐工序控制："+key)
    if not p["obligations"] or not p["decisions"]:
        errors.append("须落实本项目关键合规义务与项目条件引起的施工决定")
    ids=set()
    for row in [*p["obligations"],*p["decisions"]]:
        if row["id"] in ids or row["requirement"] not in reqs or row["module"] not in modules or not source_ok(row["source"],lookup):
            errors.append("项目措施重复、无原文或责任无效："+row["id"])
        ids.add(row["id"])
        owner=next((m for m in state["plan"]["modules"] if m["id"]==row["module"]),{})
        if row["requirement"] not in owner.get("primary_for",[]):
            errors.append("项目措施与需求主响应位置不一致："+row["id"])
    for key in dict.fromkeys([*SCREENING_TOPICS,*p["screening"]]):
        if p["screening"].get(key,{}).get("applicable") is False and any(key in o["topics"] for o in p["obligations"]):
            errors.append("筛查与实际登记义务冲突："+key)
    for o in p["obligations"]:
        if not o["topics"] or not all(meaningful(o.get(k)) for k in ("authority","version","clause","territory","scope","operation","role","check","acceptance")):
            errors.append(o["id"]+" 缺依据效力/版本条款、适用地域工况或责任与验收")
        if not meaningful(o.get("level")) or o["level"] not in ("国家","行业","省","市","县","项目"):
            errors.append(o["id"]+" 缺适用层级")
        if o["level"] in ("省","市","县") and o["territory"] not in p["jurisdiction"] and not source_ok(o.get("adoption"),lookup,tender):
            errors.append(o["id"]+" 属地与项目所在地不符，缺招标适用依据")
        if not source_ok(o.get("validity"),lookup):
            errors.append(o["id"]+" 缺有效版本/地域适用性的核实依据")
        validity=o["validity"]
        checked=date.fromisoformat(p["as_of"])
        start=date.fromisoformat(validity["effective_from"])
        until=validity.get("effective_until")
        if validity.get("status")!="有效" or checked<start or (until and checked>date.fromisoformat(until)):
            errors.append(o["id"]+" 依据失效、尚未生效或适用日期错误")
        if "岗位资格" in o["topics"]:
            if not o["topics"] or not all(meaningful(o.get(k)) for k in ("qualification","verification_time")):
                errors.append(o["id"]+" 缺岗位证照、核验时点")
            if o.get("person_confirmed") is True and not source_ok(o.get("person_proof"),lookup):
                errors.append(o["id"]+" 拟派人员已持证声明缺证明")
    for d in p["decisions"]:
        if not all(meaningful(d.get(k)) for k in ("condition","choice","verification")):
            errors.append(d["id"]+" 缺限制条件、具体施工决定或核验办法")
    errors+=specialty_plan_errors(state,p,lookup,tender)
    scoring={r["id"] for r in reqs.values() if r.get("scoring")}
    if set(p["score_effects"])!=scoring:
        errors.append("评分后果须逐项覆盖原评分子条件")
    for rid,e in p["score_effects"].items():
        if e["kind"] not in ("独立项零分","综合项评审","明确扣分","否决","未明示") or not meaningful(e.get("scope")) or not source_ok(e["source"],lookup,tender):
            errors.append(rid+" 缺由原文支持的后果及适用层级")
    names={x["name"] for x in state["startup"]["visual_policy"].get("named_exhibits",[]) if x["kind"]=="表"}
    if set(p["annexes"]) != names:
        errors.append("附表实质核查范围与招标指定表不一致")
    for name,a in p["annexes"].items():
        errors.extend(annex_plan_errors(name,a,modules,lookup,tender))
    for name,a in p["annexes"].items():
        errors+=resource_scope_errors(name,a,p,tasks)
    for task in tasks:
        errors+=capacity_errors(task,p,lookup)
    return errors+catalog_errors(p,tasks)


def specialty_plan_errors(state,p,lookup,tender):
    errors=[]
    decisions={d["id"]:d for d in p["decisions"]}
    site=p.get("site",{})
    conditions=site.get("conditions",{})
    for name in SITE_TOPICS:
        item=conditions.get(name,{})
        status=item.get("status")
        if (status not in ("confirmed","verify","not_applicable")
                or (name in SITE_CORE and status=="not_applicable")
                or not meaningful(item.get("reason")) or not source_ok(item.get("source"),lookup)):
            errors.append("现场踏勘/勘察："+name+" 缺实际条件、适用判断或可回查依据")
        elif status!="not_applicable" and item.get("decision") not in decisions:
            errors.append("现场踏勘/勘察："+name+" 缺施工决定及核验责任")
    visit=site.get("visit",{})
    if (type(visit.get("required")) is not bool or not meaningful(visit.get("reason"))
            or not source_ok(visit.get("source"),lookup,tender)):
        errors.append("现场踏勘：缺对招标是否要求投标阶段实际踏勘的确认")
    elif visit["required"] and (not source_ok(visit.get("visit_record"),lookup)
            or visit["visit_record"]["source_id"] in tender or visit.get("decision") not in decisions):
        errors.append("现场踏勘：招标要求实地踏勘，缺实际踏勘记录或响应模块")

    diff=p.get("difficulties",{})
    if not source_ok(diff.get("source"),lookup,tender) or not meaningful(diff.get("reason")):
        errors.append("工程重难点：缺招标点名清单核查依据")
    named=diff.get("named",[])
    items=diff.get("items",[])
    expected={normal(x["name"]) for x in named}
    actual={normal(x["name"]) for x in items if x.get("origin")=="tender"}
    if (len(expected)!=len(named) or expected!=actual or
            any(not source_ok(x.get("source"),lookup,tender)
                or normal(x["name"]) not in normal(x["source"]["quote"]) for x in named)):
        errors.append("工程重难点：招标点名项未按原文逐项对应")
    if not items or len({normal(x["name"]) for x in items})!=len(items) or len({x["decision"] for x in items})!=len(items):
        errors.append("工程重难点：须逐项识别并落实相应措施，名称及责任决定须各自唯一")
    if not any(x.get("origin")=="project" for x in items) and (
            not meaningful(diff.get("project_reason")) or not source_ok(diff.get("project_source"),lookup)):
        errors.append("工程重难点：未从本项目图纸、清单或现场识别风险，也无核查后的排除依据")
    for item in items:
        origin=item.get("origin")
        if (origin not in ("tender","project") or not source_ok(item.get("source"),lookup,tender if origin=="tender" else None)
                or item.get("decision") not in decisions or not all(meaningful(item.get(k)) for k in DIFFICULTY_FIELDS)):
            errors.append("工程重难点："+str(item.get("name") or "未命名")+" 缺项目依据、成因、影响、资源或对应措施")

    for row in p["obligations"]:
        if set(row["topics"]).intersection((*FIRE_TOPICS,*ENVIRONMENT_TOPICS)):
            if not all(meaningful(row.get(k)) for k in CONTROL_FIELDS):
                errors.append(row["id"]+" 消防/环保控制缺作业或受影响对象、阶段、异常处置")

    aerial=p.get("aerial",{})
    figures=state["plan"]["hard_checks"].get("figures",[])
    policy=state["startup"]["visual_policy"]
    names=policy.get("additional_figures",{}).get("names",[])
    figure=next((f for f in figures if f["name"]==AERIAL_FIGURE),{})
    # A tender-named mandatory drawing can already reserve this exact title.
    named_aerial=any(x["kind"]=="图" and x["name"]==AERIAL_FIGURE for x in policy.get("named_exhibits",[]))
    if ((AERIAL_FIGURE not in names and not named_aerial) or figure.get("kind")!="site"
            or "施工边界" not in figure.get("elements",[])):
        errors.append(AERIAL_FIGURE+"：须登记为施工总平面派生空间图并纳入同稿图件硬审")
    if not all(meaningful(aerial.get(k)) for k in ("layout_path","layout_locator")) or not source_ok(aerial.get("source"),lookup):
        errors.append(AERIAL_FIGURE+"：缺项目总平面图路径、版位或相关项目依据")
    return errors


def aerial_evidence_errors(p,blocks,review,base):
    aerial=p["aerial"]
    record=review.get("hard_checks",{}).get("figures",{}).get(AERIAL_FIGURE,{})
    trace=record.get("generation",{})
    caption=re.compile(r"图\s*[1-9][0-9]*\s+"+re.escape(AERIAL_FIGURE))
    errors=[]
    if sum(bool(caption.fullmatch(b["text"].strip())) for b in blocks)!=1:
        errors.append(AERIAL_FIGURE+"：本稿须有独立、连续编号的标准题注")
    if (trace.get("method")!="ai_from_layout" or trace.get("layout_locator")!=aerial["layout_locator"]
            or not meaningful(trace.get("record")) or trace.get("layout_sha256")==record.get("image_sha256")
            or not fresh({"path":aerial["layout_path"],"sha256":trace.get("layout_sha256")},base)):
        errors.append(AERIAL_FIGURE+"：缺AI生成过程、总平面图版本或位置的有效来源记录")
    return errors


def capacity_errors(task,p,lookup):
    """Compare required hourly output with the documented crew composition and machine pools."""
    model=task.get("capacity",{})
    if not source_ok(model.get("source"),lookup) or not meaningful(model.get("calculation")):
        return [task["id"]+" 缺工效输入原件及推导"]
    errors=[]
    needed=Decimal(str(task["rate"]))*len(task["crews"])
    if set(model.get("crews",{}))!=set(task["crews"]) or set(model.get("plant",{}))!=set(task["equipment"]):
        return [task["id"]+" 工效测算未覆盖实际班组或设备"]
    labour=Decimal(0)
    for name,spec in model["crews"].items():
        rate=Decimal(str(spec["rate"]))
        if not rate.is_finite() or rate<=0 or spec["trades"]!=p["crews"][name]["trades"]:
            errors.append(task["id"]+"/"+name+" 工效测算班组构成或产能失配")
        labour+=rate
    if needed>labour:
        errors.append(task["id"]+" 计划工效超过实际班组产能")
    pools=defaultdict(Decimal)
    for name,spec in model["plant"].items():
        if "support" in spec:
            if not meaningful(spec["support"]):
                errors.append(name+" 辅助机具未说明用途")
            continue
        rate=Decimal(str(spec["unit_rate"]))
        if (not rate.is_finite() or rate<=0 or spec.get("unit")!=task["unit"]
            or not meaningful(spec.get("operation")) or not meaningful(spec.get("conversion"))):
            errors.append(name+" 设备有效小时能力/计量口径无效")
        pools[spec["operation"]]+=rate*p["equipment"][name]["count"]
    for operation,output in pools.items():
        if needed>output:
            errors.append(f"{task['id']}/{operation} 每小时需{needed}，机组只能完成{output}")
    return errors


def annex_plan_errors(name,a,modules,lookup,tender):
    headers,rows=a["headers"],a["rows"]
    if (a["module"] not in modules or not source_ok(a["source"],lookup,tender) or not meaningful(a.get("basis"))
        or not headers or not all(isinstance(r,list) and r and all(isinstance(v,str) for v in r) for r in headers)
        or not posint(a["key_col"]) or not rows or not a["required_cols"]):
        return [name+" 缺表式来源、表头、规定对象或必填列"]
    keys=set()
    linked_columns=set(a.get("resource_cols",[]))
    if any(int(col) not in a["required_cols"] for col in linked_columns):
        return [name+" 资源核对列须为必填列"]
    if any(k in name for k in ("劳动力","机械","设备")) and not linked_columns:
        return [name+" 资源表缺与人员或设备计划关联的数据列"]
    if any(not posint(col) for col in a["required_cols"]):
        return [name+" 必填列须使用正整数列号"]
    for row in rows:
        if not meaningful(row["key"]) or row["key"] in keys:
            return [name+" 对象为空或重复；重名对象使用组合行键"]
        keys.add(row["key"])
        if not set(map(str,a["required_cols"]))<=set(row["cells"]):
            return [name+" 必填单元格没有逐项定义"]
        for col,spec in row["cells"].items():
            if not str(col).isdecimal() or int(col)<1 or len(set(spec)&{"value","peak","workdays","resource","na"})!=1:
                return [name+" 单元格须有唯一的数据依据"]
            if int(col) in linked_columns and not set(spec)&{"resource","peak","workdays","na"}:
                return [name+" 资源列须由人员/设备计划取值"]
            if "value" in spec and (not meaningful(str(spec["value"])) or not meaningful(spec.get("basis"))):
                return [name+" 独立值缺有效取值或来源"]
            if "na" in spec and (not isinstance(spec["na"],str) or not spec["na"].strip() or not meaningful(spec.get("reason"))):
                return [name+" 不适用格须记录工况原因"]
    return []


def peak(tasks,p,kind,key,start,end):
    """Maximum concurrent planned headcount or machines within inclusive project days."""
    if kind not in ("trade","plant") or not posint(start) or not posint(end) or end<start:
        raise ValueError("资源峰值的类型或阶段错误")
    events=defaultdict(int)
    from schedule import occupied_windows
    for t in tasks:
        for d,lo,hi in occupied_windows(t):
            if start<=d<=min(end,t["finish"]):
                n=(sum(p["crews"][x]["trades"].get(key,0) for x in t["crews"]) if kind=="trade" else
                   sum(p["equipment"][x]["count"] for x in t["equipment"] if x==key))
                events[d,Decimal(str(lo))]+=n
                events[d,Decimal(str(hi))]-=n
    active=highest=0
    for moment in sorted(events):
        active+=events[moment]
        highest=max(highest,active)
    return highest


def resource_scope_errors(name,a,p,tasks):
    """Independent row/phase denominator from the approved work breakdown and resource inventory."""
    errors=[]
    kind=a.get("resource_kind","")
    if "劳动力" in name and kind!="trade" or any(x in name for x in ("机械","设备")) and kind!="plant":
        return [name+" 需指明人员/机具的取数范围"]
    if not kind:
        return errors
    if kind not in ("trade","plant") or not a.get("task_ids"):
        return [name+" 缺覆盖活动"]
    ids=set(a["task_ids"])
    chosen=[t for t in tasks if t["id"] in ids]
    if ids!={t["id"] for t in chosen} or not chosen:
        return [name+" 活动范围不存在"]
    if ids!={t["id"] for t in tasks} and not meaningful(a.get("scope_basis")):
        errors.append(name+" 局部附表缺活动范围依据")
    listed=set()
    phases=a.get("phases",{})
    equipment_fields=a.get("equipment_fields",{})
    if kind=="plant":
        for i,header in enumerate(a["headers"][-1],1):
            if any(w in header for w in ("数量","台数","套数")):
                if equipment_fields.get(str(i))!="count":
                    errors.append(name+" 数量列未绑定设备台数")
            if any(w in header for w in ("型号","规格")) and equipment_fields.get(str(i))!="spec":
                errors.append(name+" 型号列绑定错误")
        if not phases and "count" not in equipment_fields.values():
            errors.append(name+" 设备表须锁定数量列或分阶段数量")
        if not set(map(int,equipment_fields)) <=set(a.get("resource_cols",[])):
            errors.append(name+" 设备属性列必须由资源计划取值")
    required=({k for t in chosen for team in t["crews"] for k in p["crews"][team]["trades"]} if kind=="trade"
              else {key for t in chosen for key in t["equipment"]})
    if kind=="trade" and not phases:
        errors.append(name+" 缺施工阶段")
    phase_keys=set()
    for label,days in phases.items():
        if (not meaningful(label) or len(days)!=2 or not all(posint(d) for d in days)
                or not 1<=days[0]<=days[1]<=a.get("project_days",max(t["finish"] for t in tasks))):
            errors.append(name+" 阶段定义无效："+label)
    for row in a["rows"]:
        key=row.get("resource_key",row["key"])
        for col,cell in row["cells"].items():
            if kind=="plant" and col in equipment_fields and cell.get("resource")!=["equipment",key,equipment_fields[col]]:
                errors.append(f"{name} {key}/{col} 设备字段未绑定规定的目录属性")
            if "peak" in cell or "workdays" in cell:
                k,resource,lo,hi=(cell["peak"] if "peak" in cell else cell["workdays"][:4])
                phase=cell.get("phase")
                labels=[r[int(col)-1] for r in a["headers"] if len(r)>=int(col)]+[row["key"]]
                axis=a.get("phase_labels",{}).get(phase,phase)
                if k!=kind or phase not in phases or phases[phase]!=[lo,hi] or not any(axis and axis in h for h in labels):
                    errors.append(f"{name} {row['key']}/{col} 峰值窗口与表内阶段不符")
                if resource!=key and not any(resource in h for h in labels):
                    errors.append(f"{name} {row['key']}/{col} 工种或机具与实际行列不符")
                listed.add(resource)
                phase_keys.add((resource,phase))
            elif "resource" in cell:
                family,resource,field,*rest=cell["resource"]
                if family=="equipment" and kind=="plant":
                    listed.add(resource)
                    if key!=resource:
                        errors.append(name+" 设备实体与表行不符："+key)
            elif "na" in cell and int(col) in a.get("resource_cols",[]):
                errors.append(name+" 人员或机具数量格应填实际数量（含零）")
    if not required<=listed:
        errors.append(name+" 缺覆盖工种或设备："+",".join(sorted(required-listed)))
    if kind=="trade" and {(k,phase) for k in required for phase in phases}-phase_keys:
        errors.append(name+" 工种与阶段交叉项未覆盖")
    return errors


def amount_at(key,number,lines,unit):
    token=re.escape(str(number))
    index=r"(?<![0-9.])"+token+r"(?![0-9.])"
    # Exact paired cells or a direct name/amount phrase. Do not borrow a neighbour's number.
    pattern=re.escape(normal(key))+r"(?:[，,：:为共计约]|投入|配置|每班|数量|各|[（(])*"+index+r"(?:[）)]?(?:"+unit+r")|(?=$|[，,。；;、|）)]))"
    for line in lines:
        if re.search(pattern,normal(line)):
            return True
        cells=[normal(x) for x in line.split(" | ")]
        if any(v==normal(key) and re.fullmatch(token+r"(?:"+unit+r")?",cells[i+1]) for i,v in enumerate(cells[:-1])):
            return True
    return False


def expected_cell(spec,p,tasks):
    if "workdays" in spec:
        from schedule import occupied_windows,positive
        kind,key,start,end,hours=spec["workdays"]
        total=Decimal(0)
        for t in tasks:
            count=(sum(p["crews"][name]["trades"].get(key,0) for name in t["crews"]) if kind=="trade" else
                   sum(p["equipment"][name]["count"] for name in t["equipment"] if name==key))
            total+=count*sum(Decimal(str(hi))-Decimal(str(lo)) for day,lo,hi in occupied_windows(t) if start<=day<=end)
        return format(total/positive(hours),"f")
    if "peak" in spec:
        kind,key,start,end=spec["peak"]
        if kind=="trade" and not any(key in crew["trades"] for crew in p["crews"].values()):
            raise ValueError("未知工种："+key)
        if kind=="plant" and key not in p["equipment"]:
            raise ValueError("未知机具："+key)
        return str(peak(tasks,p,kind,key,start,end))
    if "resource" in spec:
        kind,name,field,*trade=spec["resource"]
        value=p[kind][name][field]
        return str(value[trade[0]] if trade else value)
    return str(spec["na"] if "na" in spec else spec["value"])


def check_table(name,table,a,p,tasks):
    errors=[]
    headers=a["headers"]
    if [list(map(normal,r)) for r in table[:len(headers)]] != [list(map(normal,r)) for r in headers]:
        return [name+" 表头/列顺序与编前锁定表式不一致"]
    seen=set()
    specs={normal(x["key"]):x for x in a["rows"]}
    for n,row in enumerate(table[len(headers):],len(headers)+1):
        try:
            cols=a.get("key_cols",[a["key_col"]])
            key=normal("|".join(row[c-1] for c in cols))
            if key not in specs or key in seen:
                errors.append(f"{name} 第{n}行对象未规划、重复或错配")
                continue
            seen.add(key)
            for col,spec in specs[key]["cells"].items():
                scoped=[t for t in tasks if t["id"] in a["task_ids"]] if a.get("task_ids") else tasks
                value=expected_cell(spec,p,scoped)
                actual=normal(row[int(col)-1])
                agreed=normal(value)
                numeric=re.fullmatch(r"[+-]?\d+(?:\.\d+)?",agreed) and re.fullmatch(r"[+-]?\d+(?:\.\d+)?",actual)
                if not (actual==agreed or numeric and Decimal(actual)==Decimal(agreed)):
                    errors.append(f"{name} 第{n}行第{col}列 应为「{value}」，实为「{row[int(col)-1]}」")
        except (IndexError,KeyError,ValueError,TypeError) as exc:
            errors.append(f"{name} 第{n}行无有效数据列或资源依据：{exc}")
    if set(specs)!=seen:
        errors.append(name+" 缺规定对象："+",".join(sorted(set(specs)-seen)))
    return errors


def module_evidence(p,mid,locations,record):
    errors=[]
    difficulties={x["decision"]:x for x in p.get("difficulties",{}).get("items",[])}
    for row in [*p["decisions"],*p["obligations"]]:
        if row["module"]!=mid:
            continue
        fields=(["condition","choice","verification"] if "choice" in row else
                ["authority","clause","operation","role","check","acceptance"]+(["qualification","verification_time"] if "岗位资格" in row["topics"] else []))
        if row.get("topics") and (not meaningful(row.get("scope")) or not meaningful(row.get("territory"))):
            errors.append(row["id"]+" 未界定实际工况或适用地域")
        if "qualification" in fields and row.get("person_confirmed") is True and not row.get("person_proof"):
            errors.append(row["id"]+" 无拟派人资格依据")
        values=row
        if row["id"] in difficulties:
            values={**row,**difficulties[row["id"]]}
            fields+=DIFFICULTY_FIELDS
        if set(row.get("topics",[])).intersection((*FIRE_TOPICS,*ENVIRONMENT_TOPICS)):
            fields+=CONTROL_FIELDS
        for field in fields:
            loc=record.get(row["id"],{}).get(field,"")
            if not meaningful(values.get(field)) or normal(values[field]) not in normal(locations.get(loc,"")):
                errors.append(f"{row['id']}/{field} 缺本稿实际位置或施工引文")
    return errors


def evidence_errors(state,document,review,common,selected=None,*,base=None):
    errors=[]
    p=state["plan"]["bid_requirements"]
    blocks=common._docx_blocks(document)
    locations=document_rows(blocks)
    evidence=review.get("bid_evidence",{})
    modules=set(selected if selected is not None else (m["id"] for m in state["plan"]["modules"]))
    spec=state["plan"]["hard_checks"]
    if any(f["name"]==AERIAL_FIGURE and f["module"] in modules for f in spec["figures"]):
        errors+=aerial_evidence_errors(p,blocks,review,base if base is not None else document.parent)
    tasks=spec["schedule"]["tasks"]
    for mid in sorted(modules):
        errors+=module_evidence(p,mid,locations,evidence)
    for t in tasks:
        if t["module"] not in modules:
            continue
        for kind in ("crews","equipment"):
            for name in t[kind]:
                item=p[kind][name]
                source=evidence.get("resources",{}).get(t["id"],{}).get(name,{})
                cells=[locations.get(loc,"") for loc in source.get("locations",[])]
                parts=([item["shift"],*item["trades"]] if kind=="crews" else [item["spec"],item["capacity"]])+[item["entry"]]
                if not cells or not all(any(normal(v) in normal(line) for line in cells) for v in [name,*parts]):
                    errors.append(f"{t['id']}/{name} 缺可定位的人员、设备配置")
                pairs=(item["trades"] if kind=="crews" else {item["spec"]:item["count"]})
                for key,num in pairs.items():
                    if not amount_at(key,num,cells,"人" if kind=="crews" else "台|套|部|辆"):
                        errors.append(f"{name}/{key} 未核到投入数量{num}")
        loc=evidence.get("resources",{}).get(t["id"],{}).get("productivity","")
        if normal(t["productivity_basis"]) not in normal(locations.get(loc,"")):
            errors.append(t["id"]+" 正文缺与实际配置对应的工效依据")
    import project_response
    for name,a in p["annexes"].items():
        if a["module"] not in modules:
            continue
        item={"kind":"表","name":name}
        candidates=project_response.named_tables(blocks,item,common._caption_match)
        if len(candidates)!=1:
            errors.append(name+" 缺同名章节内与题注绑定的唯一实表")
        else:
            errors+=check_table(name,candidates[0]["table_content"],a,p,tasks)
    return errors
