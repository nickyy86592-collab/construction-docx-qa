"""r21 四项硬审。自动检查可复算事实，专业判断绑定实际成稿位置和图像。"""
from hashlib import sha256
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET
import posixpath
import boq
import schedule
from source_coverage import inventory
from visual_policy import base_names

ASPECTS=("工序","触发","做法","标准","异常")
ELEMENTS={"site":{"作业区","出入口","路线","空间关系","临设"},
          "network":{"关键活动","依赖关系"}, "gantt":{"关键活动","作业时段"}, "other":set()}
BAD={"已复核","已核实","符合要求","按规范执行","已检查","通过","无","待核实","不适用"}
W="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R="{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CHECK_ERRORS=(OSError, ValueError, TypeError, KeyError, IndexError, AttributeError, ImportError, zipfile.BadZipFile, ET.ParseError)


def detail(s):
    return (isinstance(s,str) and len(s.strip())>=2 and s.strip() not in BAD
            and not re.fullmatch(r"(?:本项|已经|已|均已)?(?:复核|核对|检查)(?:完成|通过|无误)?[。！!]?",s.strip()))


def template():
    return {"bills":[], "figures":[], "measures":[], "schedule": {"tasks":[]}, "gaps":[]}


def pending_review():
    return {"bills":[], "non_bill":[], "bill_scan":"", "figures":{}, "measures":{}, "schedule":{}}


def required_figures(state):
    policy=state["startup"]["visual_policy"]
    spatial,gantt,network=base_names(policy)
    kinds={spatial:"other",gantt:"gantt",network:"network"}
    for item in policy.get("named_exhibits",[]):
        if item["kind"]=="图":
            kinds.setdefault(item["name"],"other")
    for name in kinds:
        if "总平面" in name or "场布" in name:
            kinds[name]="site"
        elif "网络" in name:
            kinds[name]="network"
    return kinds


def plan_errors(state, base=None):
    try:
        spec=state["plan"]["hard_checks"]
        req={r["id"]:r for r in state["plan"]["requirements"]}
        modules={m["id"]:m for m in state["plan"]["modules"]}
        errors=[]
        bills={s["id"] for s in state["startup"]["sources"] if s["role"] in ("bill","boq")}
        if bills!={b["source"] for b in spec["bills"]} or any(not b["tables"] for b in spec["bills"]):
            errors.append("清单：有效原件/表尚未对应；失效版从资料盘点中注明替代关系")
        if base is not None:
            try:
                boq.index(state,base)
            except CHECK_ERRORS as exc:
                errors.append("清单：编前原件定位失败 "+str(exc))
        required=required_figures(state)
        figures=spec["figures"]
        if len(figures)!=len({f["name"] for f in figures}) or set(required)-{f["name"] for f in figures}:
            errors.append("图件：须锁定各指定图与基础图的必验要素")
        for f in figures:
            kind=f["kind"]
            if (kind not in ELEMENTS or (required.get(f["name"],"other")!="other" and required[f["name"]]!=kind)
                    or not f["elements"] or not ELEMENTS[kind]<=set(f["elements"]) or not detail(f["basis"])
                    or f["module"] not in modules):
                errors.append("图件：图种、依据、模块或要素不足 "+f["name"])
        measures=spec["measures"]
        if not measures or len(measures)!=len({m["id"] for m in measures}):
            errors.append("措施：缺适用措施清单或编号重复")
        for m in measures:
            if (m["requirement"] not in req or m["requirement"] not in modules[m["module"]]["primary_for"]
                    or not detail(m["completion"]) or not detail(m["basis"])):
                errors.append("措施：缺本模块所属需求、适用依据或完成条件 "+m["id"])
        if not spec["schedule"]["tasks"] or any(t["module"] not in modules for t in spec["schedule"]["tasks"]):
            errors.append("计划：缺关键工序或责任模块")
        for g in spec["gaps"]:
            if not all(detail(g[k]) for k in ("element","source","next_action","affected")):
                errors.append("资料缺口须单独登记对象、来源、影响及核验办法")
        return errors
    except CHECK_ERRORS as exc:
        return ["四项硬审规划未齐：在原状态plan.hard_checks补录，保留项目记录："+str(exc)]


def document_rows(blocks):
    rows={}
    for i,b in enumerate(blocks,1):
        if b["table_index"] is None:
            rows[f"P{i}"]=b["text"]
        else:
            for j,values in enumerate(b["table_content"],1):
                rows[f"T{b['table_index']}R{j}"]=" | ".join(values)
    return rows


def table_claim_errors(blocks,claims,tables):
    """按实际表头独立读取Word原生表内明示字段，即使复核人漏填fields也核对。"""
    by_loc={}
    for claim in claims:
        by_loc.setdefault(claim["loc"],[]).append(claim)
    errors=[]
    for block in blocks:
        tab=block["table_index"]
        if tab is None:
            continue
        columns={}
        for n,cells in enumerate(block["table_content"],1):
            fields={i:k for i,v in enumerate(cells) for k,aliases in boq.HEADERS.items() if boq.clean(v) in aliases}
            if "code" in fields.values() or ("name" in fields.values() and "quantity" in fields.values()):
                columns=fields
                continue
            loc=f"T{tab}R{n}"
            if not columns or len(by_loc.get(loc,[]))!=1:
                continue
            c=by_loc[loc][0]
            row=tables.get((c["source"],c["sheet"],c["trade"],c["title"]),{}).get(c["row"])
            if row is None:
                continue
            for i,k in columns.items():
                if i>=len(cells) or not cells[i].strip():
                    continue
                if not boq.same(k,cells[i],row[k]):
                    errors.append(f"Word {loc} 第{i+1}列 / {c['source']} {c['trade']} {c['sheet']} 第{c['row']}行 {boq.LABEL[k]} Word={cells[i]} 原行={row[k]}")
    return errors


def measure_errors(specs,reviews,locations,source_lookup):
    errors=[]
    for spec in specs:
        mid=spec["id"]
        record=reviews.get(mid,{})
        if not isinstance(record,dict):
            errors.append(f"措施 {mid}：复核须逐项定位实际正文证据")
            continue
        for aspect in ASPECTS:
            e=record.get(aspect,{})
            quote=e.get("quote")
            if (not detail(quote) or not detail(e.get("finding")) or
                    quote not in locations.get(e.get("loc"),"")):
                errors.append(f"措施 {mid}/{aspect}：缺实际正文位置、引文或充分性判断")
                continue
            if re.search(r"\d",quote):
                refs=e.get("values",[])
                numbers=set(re.findall(r"\d+(?:\.\d+)?",quote))
                supported=set()
                for ref in refs:
                    if "operation" in ref:
                        value=calculated_value(ref,source_lookup)
                        if value is not None:
                            from decimal import Decimal
                            supported.update(n for n in numbers if Decimal(n)==value)
                        continue
                    src=source_lookup(ref["source"],ref["unit_id"])
                    if detail(ref["quote"]) and ref["quote"] in src:
                        supported.update(re.findall(r"\d+(?:\.\d+)?",ref["quote"]))
                if numbers-supported:
                    errors.append(f"措施 {mid}/{aspect}：数值缺原件依据 {sorted(numbers-supported)}")
        if spec.get("conditional"):
            for key in ("实施前核验","对应处置"):
                e=record.get(key,{})
                if not detail(e.get("quote")) or e["quote"] not in locations.get(e.get("loc"),""):
                    errors.append(f"措施 {mid}：缺{key}的实际正文证据")
    return errors


def calculated_value(ref,lookup):
    """从原件引文重新计算，禁止执行任意表达式。复合公式用嵌套输入。"""
    from decimal import Decimal, DecimalException
    import operator
    operations={"add":operator.add,"subtract":operator.sub,"multiply":operator.mul,"divide":operator.truediv}
    try:
        if "operation" not in ref:
            raw=str(ref["value"])
            quote=ref["quote"]
            if (not detail(quote) or quote not in lookup(ref["source"],ref["unit_id"])
                    or raw not in re.findall(r"-?\d+(?:\.\d+)?",quote)):
                return None
            return Decimal(raw)
        if not detail(ref.get("basis")) or not ref.get("unit"):
            return None
        left,right=[calculated_value(x,lookup) for x in ref["inputs"]]
        if left is None or right is None:
            return None
        return operations[ref["operation"]](left,right)
    except (KeyError,ValueError,DecimalException):
        return None


def source_lookup(state,base):
    cache={}
    sources={s["id"]:s for s in state["startup"]["sources"]}
    def lookup(key,unit):
        if key not in cache:
            cache[key]=inventory(sources[key],base)[1]
        return cache[key].get(unit,"")
    return lookup


def module_errors(state,module,text,review,base):
    specs=state.get("plan",{}).get("hard_checks",{}).get("measures",[])
    specs=[s for s in specs if s["module"]==module["id"]]
    return measure_errors(specs,review.get("measures",{}),{"L"+str(i):s for i,s in enumerate(text.splitlines(),1)},source_lookup(state,base))


def figure_errors(specs,reviews,blocks,common,word,model,document):
    errors=[]
    pages=word.get("figure_pages",{})
    for f in specs:
        name=f["name"]
        review=reviews.get(name,{})
        caps=[(i,b["text"]) for i,b in enumerate(blocks) if (m:=common._caption_match(b["text"],"图")) and m.group(2)==name]
        if len(caps)!=1:
            errors.append("图件 "+name+"：实际Word题注不唯一")
            continue
        i,caption=caps[0]
        images=common._caption_images(blocks,i)
        page=review.get("page")
        if (len(images)!=1 or review.get("image_sha256")!=sha256(images[0]["bytes"]).hexdigest()
                or type(page) is not int or page<1 or pages.get(caption)!=[page]
                or review.get("observed_kind")!=f["kind"]):
            errors.append(f"图件 {caption} 第{page}页：图种、同页图像、指纹或Word实排页不符")
        for element in f["elements"]:
            e=review.get("elements",{}).get(element,{})
            if element=="临设" and e.get("status")=="not_applicable" and detail(f.get("no_facilities_basis")) and e.get("finding")==f["no_facilities_basis"]:
                continue
            if not detail(e.get("position")) or not detail(e.get("finding")) or e.get("status")!="verified":
                errors.append(f"图件 {caption} 第{page}页/{element}：缺图上位置、实见内容或必验要素")
        if f["kind"] in ("network","gantt") and not schedule.calculate(model):
            expected=schedule.chart(model,f["kind"])
            if expected not in svg_at_caption(document,images):
                errors.append("计划／图件 "+name+"：Word嵌入的进度图与唯一计划数据不一致")
    return errors


def svg_at_caption(document,images):
    """校验实际显示的SVG关系，拒绝仅将正确SVG作为无关附件塞入ZIP。"""
    if len(images)!=1:
        return []
    with zipfile.ZipFile(document) as z:
        rel={r.get("Id"):posixpath.normpath(posixpath.join("word",r.get("Target"))).lstrip("/")
             for r in ET.fromstring(z.read("word/_rels/document.xml.rels")) if r.get("TargetMode")!="External"}
        root=ET.fromstring(z.read("word/document.xml"))
        result=[images[0]["bytes"]]
        for blip in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            if rel.get(blip.get(R+"embed"))==images[0]["member"]:
                for el in blip.iter("{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"):
                    result.append(z.read(rel[el.get(R+"embed")]))
        return result


def audit(state,base,document,review,common,word,selected=None):
    """阶段只核已纳入快照的责任模块。全稿从实际正文枚举，反向寻找原件。"""
    errors=plan_errors(state)
    if errors:
        return errors
    try:
        spec=state["plan"]["hard_checks"]
        rec=review["hard_checks"]
        if review["document_sha256"]!=sha256(document.read_bytes()).hexdigest():
            return ["四项硬审记录已失效：Word稿已改变"]
        blocks=common._docx_blocks(document)
        rows=document_rows(blocks)
        versions={s["id"]:s["sha256"] for s in state["startup"]["sources"]}
        if not detail(rec["bill_scan"]):
            errors.append("清单：缺逐段、逐表行与图注反查结论")
        tables=boq.index(state,base)
        errors+=boq.check(rows,rec["bills"],tables,versions,rec.get("non_bill",[]))
        errors+=table_claim_errors(blocks,rec["bills"],tables)
        def included(items):
            return [x for x in items if selected is None or x["module"] in selected]
        figures=included(spec["figures"])
        if figures and (word.get("layout_engine")!="Microsoft Word" or word.get("document_sha256")!=review["document_sha256"]):
            errors.append("图件：Word页码记录未绑定同稿")
        errors+=figure_errors(figures,rec["figures"],blocks,common,word,spec["schedule"],document)
        errors+=measure_errors(included(spec["measures"]),rec["measures"],rows,source_lookup(state,base))
        model=spec["schedule"]
        errors+=schedule.calculate(model)
        for t in included(model["tasks"]):
            errors+=schedule_errors_for_task(t,rec["schedule"],rows)
            for loc,line in rows.items():
                for phrase in re.split(r"[。；;\n]",line):
                    if t["name"] not in phrase:
                        continue
                    for pair in re.finditer(r"D(\d+)\s*[—–~～至-]\s*D?(\d+)",phrase):
                        if sum(x["name"] in phrase for x in model["tasks"])==1 and (int(pair[1]),int(pair[2]))!=(t["start"],t["finish"]):
                            errors.append(f"计划 {t['id']}/{loc} 明示时段{pair[0]}与计划源不一致")
        if selected is None and spec["gaps"]:
            errors+=["未关闭的必验资料："+g["element"] for g in spec["gaps"] if g.get("required",True)]
        return list(dict.fromkeys(errors))
    except CHECK_ERRORS as exc:
        return errors+["四项硬审证据缺失或不可读："+str(exc)]


def schedule_errors_for_task(task,views,locations):
    ref=views.get(task["id"],{})
    p,r=ref.get("body",""),ref.get("table","")
    expected=schedule.table_row(task)
    cells=locations.get(r,"").split(" | ")
    cols=ref.get("columns",list(range(1,len(expected)+1)))
    if (not isinstance(cols,list) or len(cols)!=len(expected) or any(type(i) is not int or not 1<=i<=len(cells) for i in cols)
            or len(set(cols))!=len(cols)):
        return ["计划 "+task["id"]+"：计划表核心列定位失败"]
    paragraph=locations.get(p,"")
    if (not p.startswith("P") or task["name"] not in paragraph or schedule.interval(task) not in paragraph
            or not r.startswith("T") or expected!=[cells[i-1] for i in cols]):
        return ["计划 "+task["id"]+"：正文或计划表与唯一计划源不一致"]
    return []


def inspect_word(document):
    """输出同一冻结稿的实际物理页码。阶段稿也可独立运行，无须定稿版式通过。"""
    import importlib.util
    import sys
    import win32com.client
    source=Path(__file__).with_name("Word正式交付硬审（不含页数）.py")
    spec=importlib.util.spec_from_file_location("_figure_word",source)
    common=importlib.util.module_from_spec(spec)
    sys.modules[spec.name]=common
    spec.loader.exec_module(common)
    word=win32com.client.DispatchEx("Word.Application")
    try:
        word.Visible=False
        word.DisplayAlerts=0
        doc=word.Documents.Open(str(document.resolve()),ReadOnly=True,AddToRecentFiles=False)
        try:
            doc.Repaginate()
            pages=common.figure_pages(doc,common.get_paragraph_records(doc))
            return {"layout_engine":"Microsoft Word","document_sha256":sha256(document.read_bytes()).hexdigest(),
                    "word_version":str(word.Version),"figure_pages":pages}
        finally:
            doc.Close(False)
    finally:
        word.Quit()


if __name__=="__main__":
    import argparse
    import json
    parser=argparse.ArgumentParser(description="采集阶段稿实际Word图页，或从计划唯一数据源生成SVG")
    sub=parser.add_subparsers(dest="cmd",required=True)
    pages=sub.add_parser("pages")
    pages.add_argument("document",type=Path)
    pages.add_argument("output",type=Path)
    draw=sub.add_parser("chart")
    draw.add_argument("state",type=Path)
    draw.add_argument("kind",choices=("gantt","network"))
    draw.add_argument("output",type=Path)
    args=parser.parse_args()
    if args.cmd=="pages":
        args.output.write_text(json.dumps(inspect_word(args.document),ensure_ascii=False,indent=2),"utf-8")
    else:
        state=json.loads(args.state.read_text("utf-8-sig"))
        args.output.write_bytes(schedule.chart(state["plan"]["hard_checks"]["schedule"],args.kind))
