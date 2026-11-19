"""项目标题、篇幅及招标图表核对。技术判断由编制者回查原件完成。"""
from collections import Counter
import math
import re

LEVEL_MIN = {"2": 15, "3": 20}
SHARES = (0.20, 0.13, 0.10)


def text(value):
    return isinstance(value, str) and bool(value.strip())


def number(value):
    return type(value) in (int, float) and math.isfinite(value)


def compact(value):
    return re.sub(r"\s+", "", value)


def template():
    return {"keywords": [], "headings": {"2": [], "3": []}, "expanded": {},
            "priorities": [], "font_budget": None}


def title_coverage(spec, headings):
    """同层按术语去重；同一别名不能给两个概念加分。"""
    errors, report, vocabulary = [], {}, {}
    seen = {}
    for row in spec.get("keywords", []):
        term = row["term"]
        forms = row.get("forms", [])
        if not text(term) or not isinstance(forms, list) or not all(text(v) for v in forms):
            raise ValueError("关键词须有原文和有效变体数组")
        if term in vocabulary:
            errors.append("原文词重复登记：" + term)
        vocabulary[term] = {compact(v) for v in [term, *forms]}
        for form in vocabulary[term]:
            if form in seen and seen[form] != term:
                errors.append("同义词须归入同一原文词：" + form)
            seen[form] = term
    expanded = spec.get("expanded", {})
    if not isinstance(expanded, dict) or expanded.keys() - LEVEL_MIN.keys():
        raise ValueError("expanded只能按实际增设的2、3级登记")
    for level, floor in LEVEL_MIN.items():
        titles = headings[level]
        if not isinstance(titles, list) or not all(text(t) for t in titles):
            raise ValueError(level + "级标题须为实际标题数组")
        limits = spec.get("heading_limits", {}).get(level)
        if limits is not None:
            if not text(limits.get("prompt")):
                raise ValueError("标题数量限制须引用编制前提示词")
            lower, upper = limits.get("min", 0), limits.get("max")
            if type(lower) is not int or lower < 0 or (upper is not None and (type(upper) is not int or upper < lower)):
                raise ValueError("标题数量上下限无效")
            if len(titles) < lower or (upper is not None and len(titles) > upper):
                errors.append(f"{level}级标题数量{len(titles)}与提示词要求不符")
        goal = floor
        if level in expanded:
            row = expanded[level]
            goal = row.get("keyword_target")
            counts = [row.get(k) for k in ("usual_headings", "heading_target", "usual_keywords", "keyword_target")]
            if (not all(type(v) is int and v > 0 for v in counts)
                    or counts[1] <= counts[0] or counts[3] <= max(counts[2], floor)
                    or not text(row.get("prompt")) or not text(row.get("basis"))):
                raise ValueError(level + "级增设标题须有提示词、常规和增设后的两组数量及增量相适应的依据")
            if len(titles) < counts[1]:
                errors.append(f"{level}级标题{len(titles)}个，低于规划{counts[1]}个")
        hits = {term: [t for t in titles if any(f in compact(t) for f in forms)]
                for term, forms in vocabulary.items()}
        hits = {term: found for term, found in hits.items() if found}
        report[level] = {"headings": len(titles), "target": goal, "distinct": len(hits), "matches": hits}
        if len(hits) < goal:
            errors.append(f"{level}级标题仅覆盖{len(hits)}个不同关键词，须达到{goal}个")
    return errors, report


def priority_errors(priorities, pages, total):
    errors, details, seen = [], [], set()
    if not isinstance(priorities, list) or len(priorities) != 3 or not number(total) or total <= 0:
        return ["须按本项目重要性登记前三项及有效正文总篇幅"], []
    for rank, (item, minimum) in enumerate(zip(priorities, SHARES), 1):
        titles = item.get("chapters")
        if not isinstance(titles, list) or not titles or not all(text(v) for v in titles) or not text(item.get("basis")):
            return [f"第{rank}项缺实际章名或排序依据"], details
        if set(titles) - pages.keys() or seen.intersection(titles) or len(titles) != len(set(titles)):
            errors.append(f"第{rank}项章名缺失或与其他篇幅重复计入")
        seen.update(titles)
        count = sum(pages.get(title, 0) for title in titles)
        details.append({"chapters": titles, "pages": count, "body_pages": total, "share": count/total, "minimum": minimum})
        if count + 1e-8 < total * minimum:
            errors.append(f"第{rank}项占{count/total:.2%}，低于{minimum:.0%}")
    return errors, details


def plan_errors(state, base, fresh):
    plan = state["plan"]
    spec = plan.get("project_response")
    if not isinstance(spec, dict):
        return ["按07补齐plan.project_response，保留已有模块与测量记录"]
    try:
        errors, _ = title_coverage(spec, spec["headings"])
        refs = {r["id"]: r.get("source_refs", []) for r in plan["requirements"]}
        roles = {s["id"]: s["role"] for s in state["startup"]["sources"] if "id" in s}
        used = set()
        for row in spec["keywords"]:
            matches = {roles.get(ref.get("source_id")) for ref in refs.get(row["requirement"], [])
                       if compact(row["term"]) in compact(ref.get("quote", ""))}
            valid = matches & {"tender", "amendment", "boq", "bill"}
            if not valid:
                errors.append("原文词缺对应招文/清单需求引文：" + row["term"])
            used.update(valid)
        if not used & {"tender", "amendment"} or (set(roles.values()) & {"bill", "boq"} and not used & {"bill", "boq"}):
            errors.append("有招文与清单时须从两类资料提取原文词")
        pages = Counter()
        for mod in plan["modules"]:
            weight = mod.get("planned_pages", mod.get("relative_weight"))
            if not number(weight) or weight < 0:
                raise ValueError("模块缺实际篇幅预算；无项目页数目标时可填relative_weight")
            pages[mod["chapter"]] += weight
        errors.extend(priority_errors(spec["priorities"], pages, plan.get("planned_total_body_pages") or sum(pages.values()))[0])
        for exhibit in (state["startup"].get("visual_policy") or {}).get("named_exhibits", []):
            if not any(roles.get(ref.get("source_id")) in {"tender", "amendment"}
                       and compact(exhibit["name"]) in compact(ref.get("quote", ""))
                       for evidence in refs.values() for ref in evidence):
                errors.append("指定图表缺招文需求原文定位：" + exhibit["name"])
        font = spec.get("font_budget")
        if font is not None:
            errors.extend(font_errors(font, plan["modules"], base, fresh))
        return errors
    except (ValueError, KeyError, TypeError, AttributeError) as exc:
        return ["项目响应规划字段无效，见07：" + str(exc)]


def font_errors(font, modules, base, fresh):
    if not text(font.get("prompt")) or not number(font.get("size_pt")) or font["size_pt"] <= 0:
        return ["指定正文字号须登记提示原文和磅值"]
    samples = [font[k] for k in ("sample_four", "sample_selected")]
    if any(not fresh(s, base) or not number(s.get("word_pages")) or s["word_pages"] <= 0 for s in samples):
        return ["请绑定同内容的四号和指定字号Word混排小样，并记录Word实测等效页数"]
    if not text(font.get("layout_basis")) or not text(font.get("word_version")):
        return ["须说明Word版本、取样代表性、标题密度、图表和实际版式"]
    original, adjusted = font["four_chars"], font["estimated_chars"]
    if not all(type(x) is int and x > 0 for x in (original, adjusted)):
        return ["须分别估算四号及指定字号的全篇容量"]
    front = font.get("front_chars", 0)
    if type(front) is not int or not 0 <= front < min(original, adjusted):
        return ["封面目录固定字数须在全篇估算范围内"]
    if font["size_pt"] != 14 and samples[0]["sha256"] == samples[1]["sha256"]:
        return ["不同字号的小样须分别实排并绑定各自文件"]
    expected = front + round((original - front) * samples[0]["word_pages"] / samples[1]["word_pages"])
    if abs(expected - adjusted) > 1:
        return [f"同内容小样占页比例对应约{expected}字，须重算容量"]
    allocations = font["allocation"]
    valid_ids = {m["id"] for m in modules}
    if (not isinstance(allocations, list) or any(a["id"] not in valid_ids or type(a["delta"]) is not int
            or not text(a.get("content")) for a in allocations)
            or sum(a["delta"] for a in allocations) != adjusted - original):
        return ["字号增减量须分配到实际模块和具体施工内容，合计与容量差一致"]
    return []


def named_errors(blocks, items, caption_match):
    errors = []
    for item in items:
        name, label = item["name"], item["kind"]
        # 序号不属于章节目名。招文锁定的名称（含标点与空格）逐字核对。
        numbered = re.compile(r"^(?:第[一二三四五六七八九十百千万零〇两\d]+[章节篇部分]+[、：:.．\s]*|[（(]?[一二三四五六七八九十百千万零〇两\d]+(?:[.．]\d+)*[）)、.．\s]+)?"+re.escape(name)+r"$")
        heads = [i for i,b in enumerate(blocks) if b["outline"] in (1,2,3) and numbered.fullmatch(b["text"].strip())]
        captions = [i for i,b in enumerate(blocks) if b["table_index"] is None and b["outline"] not in (1,2,3)
                    and (m := caption_match(b["text"], label)) and m.group(2) == name]
        if len(heads) != 1 or len(captions) != 1:
            errors.append("须有逐字同名的正文标题及唯一编号题注："+name)
            continue
        first, cap = heads[0], captions[0]
        end = next((i for i in range(first+1,len(blocks)) if blocks[i]["outline"] in range(1,blocks[first]["outline"]+1)),len(blocks))
        if not first < cap < end:
            errors.append("题注不在同名章节目内："+name)
            continue
        candidates = [blocks[cap]]
        for step in (-1, 1):
            i = cap + step
            while first < i < end:
                b = blocks[i]
                if b["text"].strip() or b["table_index"] is not None or b["drawing_count"]:
                    candidates.append(b)
                    break
                i += step
        valid = (any(b["drawing_count"] for b in candidates) if label == "图" else
                 any(b["table_index"] is not None and any(any(c for c in row) for row in b["table_content"]) for b in candidates))
        if not valid:
            errors.append("同名题注旁缺少实际图或数据表："+name)
    return errors
