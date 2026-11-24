"""关键计划的唯一数据源。D1 起算，日内按小时登记；日历间歇独立于作业窗口。"""
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from xml.etree import ElementTree as ET


def positive(value):
    val = Decimal(str(value))
    if not val.is_finite() or val <= 0:
        raise ValueError("工程量、工效和作业时长须为有限正数")
    return val


def day(value):
    if type(value) is not int or value < 1:
        raise ValueError("日号须为正整数")
    return value


def unique_names(values):
    return isinstance(values, list) and all(isinstance(v, str) and v.strip() for v in values) and len(set(values)) == len(values)


def calculate(model):
    """按实际可作业日及班组能力逐日消耗工程量，预留清场前的整个已分配窗口。"""
    errors, spans = [], defaultdict(list)
    boundaries = {}
    try:
        tasks = model["tasks"]
        by_id = {t["id"]: t for t in tasks}
        if not tasks or len(by_id) != len(tasks) or not unique_names(list(by_id)):
            raise ValueError("关键活动清单为空或ID重复")
        for t in tasks:
            label = t["id"]
            teams, plant = t["crews"], t["equipment"]
            if (not unique_names(teams) or not teams or not unique_names(plant)
                    or not str(t["name"]).strip() or not str(t["unit"]).strip() or not t["basis"]):
                raise ValueError(label + " 缺工程量单位、工效依据或唯一班组/机具")
            remaining = positive(t["quantity"])
            rate = positive(t["rate"]) * len(teams)
            windows = t["windows"]
            if not windows:
                raise ValueError(label + " 缺有效作业窗口")
            first, last = day(t["start"]), day(t["finish"])
            previous = None
            end_day = None
            used_end = None
            for d, begin, end in windows:
                day(d)
                begin, end = Decimal(str(begin)), Decimal(str(end))
                if not 0 <= begin < end <= 24 or (previous and (d, begin) < previous):
                    raise ValueError(label + " 作业窗口乱序、重叠或超出当日")
                previous = (d, end)
                if remaining <= 0:
                    continue
                for resource in teams + plant:
                    spans[resource, d].append((begin, end, label))
                remaining -= rate * (end - begin)
                end_day = d
                used_end = (d,begin,end)
            boundaries[label]=(windows[0],used_end)
            if remaining > 0 or windows[0][0] != first or end_day != last:
                errors.append(f"{label} 持续时间不符：登记D{first}—D{last}，按窗口复算至D{end_day}，剩余量{max(0, remaining)}")
            if not isinstance(t["ready"],dict) or not set(plant)<=t["ready"].keys():
                raise ValueError(label+" 缺机具到位时点")
            if not unique_names(t["materials"]) or not set(t["materials"]) <= t["ready"].keys():
                raise ValueError(label + " 缺材料到位时点")
            if not plant and not t.get("equipment_basis") or not t["materials"] and not t.get("materials_basis"):
                errors.append(label + " 无材料或机具须登记该工序的实际依据")
            for name, arrived in t["ready"].items():
                if day(arrived) > first:
                    errors.append(f"{label} {name} 到位晚于开工")
        for t in tasks:
            first=t["start"]
            label=t["id"]
            for pred, rest in t["after"]:
                if pred not in by_id or type(rest) is not int or rest < 0: 
                    errors.append(f"{label} 紧前活动或技术间歇无效：{pred} + {rest}日")
                    continue
                finish=by_id[pred]["finish"]
                same_day=(rest==0 and first==finish and Decimal(str(boundaries[label][0][1]))>=boundaries[pred][1][2])
                if pred==label or (first<=finish+rest and not same_day):
                    errors.append(f"{label} 紧前活动或技术间歇冲突：{pred} + {rest}日")
        for (resource, d), entries in spans.items():
            entries.sort()
            for left, right in zip(entries, entries[1:]):
                if right[0] < left[1]:
                    errors.append(f"D{d} {resource} 被 {left[2]}、{right[2]} 重复占用")
    except (KeyError, TypeError, ValueError, InvalidOperation, ArithmeticError) as exc:
        errors.append("关键计划字段无效：" + str(exc))
    return errors


def occupied_windows(task):
    """Use the same entire reserved windows as calculate, up to completion of the work quantity."""
    remaining=positive(task["quantity"])
    rate=positive(task["rate"])*len(task["crews"])
    for d,lo,hi in task["windows"]:
        if remaining<=0:
            break
        yield d,lo,hi
        remaining-=rate*(Decimal(str(hi))-Decimal(str(lo)))


def interval(task):
    return f"D{task['start']}—D{task['finish']}"


def table_row(t):
    return [t["name"], interval(t), str(t["finish"]-t["start"]+1),
            "、".join(t["crews"]), "、".join(p for p, _ in t["after"]) or "无"]


def chart(model, kind, color_policy=None):
    """确定性 SVG；保留活动编号、名称、时段、紧前关系与技术间歇。"""
    errors = calculate(model)
    if errors:
        raise ValueError("；".join(errors))
    if kind not in {"network", "gantt"}:
        raise ValueError("进度图种类无效")
    tasks = sorted(model["tasks"], key=lambda t: (t["start"],t["windows"][0][1],t["id"]))
    indexes = {t["id"]: i for i, t in enumerate(tasks)}
    width, top, pitch = 1020, 65, 68
    height = top + len(tasks)*pitch
    svg = ET.Element("svg", xmlns="http://www.w3.org/2000/svg", width=str(width), height=str(height), viewBox=f"0 0 {width} {height}")
    def add(tag, **attrs):
        return ET.SubElement(svg, tag, {k.replace("_","-"): str(v) for k,v in attrs.items()})
    def label(x,y,s, color="#183746", size=17):
        add("text", x=x,y=y,fill=color,font_size=size,font_family="Microsoft YaHei, sans-serif").text=s
    add("rect",x=0,y=0,width=width,height=height,fill="#ffffff")
    label(16,30,"活动 / 计划窗口" if kind=="gantt" else "关键活动及紧前依赖（箭头指向后续工序）",size=22)
    last = max(t["finish"] for t in tasks)
    # Determine longest logic path(s), including technological lags; tied branches all qualify.
    lengths, paths = {}, {}
    for t in tasks:
        prev={p:lengths[p]+lag for p,lag in t["after"]}
        longest=max(prev.values(),default=0)
        lengths[t["id"]]=longest+t["finish"]-t["start"]+1
        paths[t["id"]]={t["id"]}.union(*(paths[p] for p,value in prev.items() if value==longest))
    critical=set().union(*(paths[k] for k,v in lengths.items() if v==max(lengths.values())))
    explicit=model.get("critical_path")
    if explicit is not None:
        if not explicit or not set(explicit)<=set(indexes) or not model.get("critical_basis"):
            raise ValueError("实际资源、日历决定的关键线路须载明活动及复算依据")
        critical=set(explicit)
    mono=(color_policy or {}).get("mode")=="monochrome"
    label(705,30,"★ 关键活动" if mono else "橙色：关键活动",color="#ab4924",size=15)
    depths = {}
    for i,t in enumerate(tasks):
        y=top+i*pitch
        if kind=="gantt":
            label(12,y,t["id"]+" "+t["name"])
            a=325+545*(t["start"]-1)/last
            b=325+545*t["finish"]/last
            add("rect",x=a,y=y-13,width=max(2,b-a),height=17,fill="#ce652f" if t["id"] in critical else "#227992")
            label(878,y,interval(t))
        else:
            depths[t["id"]]=1+max((depths[p] for p,_ in t["after"]),default=-1)
    if kind=="network":
        gap=735/max(1,max(depths.values()))
        for i,t in enumerate(tasks):
            x,y=15+depths[t["id"]]*gap,top+i*pitch
            for p,lag in t["after"]:
                px,py=15+depths[p]*gap,top+indexes[p]*pitch
                add("path",d=f"M{px+8} {py+18} L{x+8} {y-17}",stroke="#23785b",fill="none",stroke_width=2)
                add("path",d=f"M{x+2} {y-21} L{x+8} {y-13} L{x+11} {y-24} Z",fill="#23785b")
            add("rect",x=x,y=y-12,width=260,height=42,rx=5,fill="#f2c396" if t["id"] in critical else "#d5eaf3",stroke="#236780")
            label(x+5,y+4,t["id"]+" "+t["name"],size=15)
            label(x+5,y+22,interval(t)+"  "+", ".join(f"{p}+{lag}d" for p,lag in t["after"]),size=14)
    if mono:
        palette={"#ce652f":"#252525","#227992":"#969696","#f2c396":"#aaa","#d5eaf3":"#eee",
                 "#ffffff":"#fff","#183746":"#222","#ab4924":"#222","#23785b":"#444","#236780":"#333"}
        for el in svg.iter():
            for attr in ("fill","stroke"):
                if el.get(attr) in palette:
                    el.set(attr,palette[el.get(attr)])
            if el.tag=="text" and any(el.text==t["id"]+" "+t["name"] for t in tasks if t["id"] in critical):
                el.text="★ "+el.text
    return ET.tostring(svg,encoding="utf-8",xml_declaration=True)
