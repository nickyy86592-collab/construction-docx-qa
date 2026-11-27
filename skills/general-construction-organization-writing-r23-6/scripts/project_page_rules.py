"""Validation of the optional per-review-point project page rule and its Word evidence."""
from collections import defaultdict
from itertools import groupby
import math


POINT_CONSTANTS = {"COVER_PAGES", "REVIEW_POINT_COUNT", "REVIEW_POINT_PAGE_LIMIT", "REVIEW_POINT_HEADINGS"}


def normalized(text):
    return "".join(text.split())


def point_rule(values):
    if not POINT_CONSTANTS.intersection(values):
        return {}
    if not POINT_CONSTANTS.issubset(values):
        raise ValueError("逐项限页脚本缺少封面及评审点字面量配置")
    cover = values["COVER_PAGES"]
    front = values["FRONT_MATTER_PAGES"]
    maximum = values["EXPECTED_BODY_PAGES"]
    count = values["REVIEW_POINT_COUNT"]
    limit = values["REVIEW_POINT_PAGE_LIMIT"]
    titles = values["REVIEW_POINT_HEADINGS"]
    if (not all(type(n) is int for n in (cover, front, maximum, count, limit))
            or values["PAGE_TOLERANCE"] != 0
            or not 0 <= cover <= front or maximum <= front - cover or count < 1 or limit < 1):
        raise ValueError("逐项限页的非封面上限、封面页数或评审点规则无效")
    if (not isinstance(titles, list) or (titles and len(titles) != count)
            or not all(isinstance(t, str) and normalized(t) for t in titles)
            or len({normalized(t) for t in titles}) != len(titles)):
        raise ValueError("评审点须使用足量、唯一的一级章名")
    return {
        "non_cover_page_limit": maximum,
        "cover_pages_excluded": cover,
        "review_point_count": count,
        "review_point_page_limit": limit,
        "review_point_headings": titles,
    }


def body_bounds(rule):
    if "review_point_count" in rule:
        directory = rule["front_matter_pages_excluded"] - rule["cover_pages_excluded"]
        return 1, rule["non_cover_page_limit"] - directory
    target, tolerance = rule["body_pages_target"], rule["page_tolerance"]
    return target - tolerance, target + tolerance


def plan_errors(state, score_items):
    """Validate the mapping from the actual scoring checklist to chapters and their budgets."""
    plan = state["plan"]
    rule = state.get("page_rule") or {}
    if "review_point_count" not in rule:
        return []
    names = rule["review_point_headings"]
    if len(names) != rule["review_point_count"]:
        return ["从有效招标评分表登记全部评审点的实际一级章名后才能开始正文"]
    if len(score_items) != len(names):
        return ["评分清单评审点数与项目限页脚本不一致"]
    errors = []
    score_chapters = {}
    for item, heading in zip(score_items, names):
        original = normalized(item["name"])
        if not original or original not in normalized(heading):
            errors.append("一级章名须包含对应评审点原文名称：" + item["name"])
        score_chapters[item["id"]] = heading
    if len(score_chapters) != len(names):
        errors.append("评审点评分编号存在重复")
    chapters = plan.get("chapter_plan") or []
    if [c.get("chapter") for c in chapters] != names:
        errors.append("章计划须按评审点顺序逐项对应，所有附表须归入相应评审点")
    modules = plan.get("modules") or []
    if [chapter for chapter, _ in groupby(m.get("chapter") for m in modules)] != names:
        errors.append("正文模块须按评审点顺序连续归属")
    requirements = {r["id"]: r for r in plan["requirements"]}
    planned = defaultdict(float)
    cap = rule["review_point_page_limit"]
    for module in modules:
        chapter = module["chapter"]
        budget = module.get("planned_pages")
        if type(budget) is not int and (type(budget) is not float or not math.isfinite(budget)):
            errors.append("评审点模块无有效页数预算：" + module["id"])
        else:
            planned[chapter] += budget
        for rid in module.get("primary_for", []):
            req = requirements.get(rid) or {}
            if req.get("scoring"):
                refs = req.get("score_refs") or []
                if not refs or any(score_chapters.get(ref["item_id"]) != chapter for ref in refs):
                    errors.append("评分子条件主响应未放入所属评审点：" + rid)
    for chapter in chapters:
        name = chapter.get("chapter")
        page_range = chapter.get("page_range")
        if (not isinstance(page_range, list) or len(page_range) != 2
                or not all(type(n) in (int, float) and math.isfinite(n) for n in page_range)
                or not 0 < page_range[0] <= page_range[1] <= cap
                or not page_range[0] <= planned[name] <= page_range[1]):
            errors.append(f"评审点 {name} 的实际模块预算和章区间须在 {cap} 页以内")
    return errors


def evidence_errors(report, rule, script_hash, expected_headings=None, reference_allowed=False):
    """Recalculate counts and compare rule, source, status and complete ordered point evidence."""
    if "review_point_count" not in rule:
        return []
    reference = report.get("reference_only") is True
    if reference and not reference_allowed:
        return ["参考合册只供内容核对，须有同一批经签名的十份实际稿实排回执"]
    reported_rule = report.get("rule") or {}
    compared = ("non_cover_page_limit", "cover_pages_excluded", "review_point_count",
                "review_point_page_limit", "review_point_headings", "body_pages_target",
                "page_tolerance", "front_matter_pages_excluded")
    if report.get("page_script_sha256") != script_hash or any(
            reported_rule.get(key) != rule[key] for key in compared):
        return ["逐项页数报告与当前项目脚本不一致，须用同一版本重新实排"]
    headings = rule["review_point_headings"] if expected_headings is None else expected_headings
    if (not headings or (expected_headings is None and len(headings) != rule["review_point_count"])
            or headings != rule["review_point_headings"][:len(headings)]):
        return ["阶段或定稿的评审点范围与项目顺序不一致"]
    actual = report.get("actual") or {}
    total = actual.get("word_total_pages")
    front = rule["front_matter_pages_excluded"]
    cover = rule["cover_pages_excluded"]
    if (type(total) is not int or total <= front or (not reference and total - cover > rule["non_cover_page_limit"])
            or actual.get("body_pages") != total - front
            or actual.get("non_cover_pages") != total - cover
            or actual.get("excluded_cover_pages") != cover
            or actual.get("excluded_front_matter_pages") != front
            or (expected_headings is not None and actual.get("word_body_pages") != total - front)):
        return ["逐项页数报告整册实排、封面扣除或总上限不符"]
    rows = actual.get("review_points")
    if not isinstance(rows, list) or len(rows) != len(headings):
        return ["缺按评审点齐全登记的 Word 实排页数"]
    errors = []
    previous = front + 1
    for index, (heading, row) in enumerate(zip(headings, rows), 1):
        if not isinstance(row, dict) or row.get("index") != index or row.get("heading") != heading:
            errors.append(f"第{index}项标题缺失、重复或顺序不符")
            continue
        start, end, pages = row.get("start_page"), row.get("end_page"), row.get("pages")
        if (not all(type(n) is int for n in (start, end, pages))
                or (index == 1 and start != front + 1)
                or not previous <= start <= end <= total or pages != end - start + 1):
            errors.append(f"{heading} 缺有效的 Word 起止物理页或计算有误")
            continue
        previous = end
        if ((not reference and pages > rule["review_point_page_limit"]) or row.get("status") != "PASS"
                or row.get("has_content") is not True):
            errors.append(f"{heading} 未通过逐项限页，实排 {pages} 页")
    if report.get("point_limit_status") != "PASS" or report.get("failures") != []:
        errors.append("逐项限页实排结论未通过")
    return errors
