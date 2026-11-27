"""对每一份实际提交的技术标逐点实排，并核对内容母稿与实际文件。"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict
from hashlib import sha256
from pathlib import Path
import re
import posixpath
from xml.etree import ElementTree as ET
import zipfile

from release_task import GateError, digest, file_sha, inside, read, require, write
from word_native import open_word

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
DOC_PROPERTIES = {"pPr", "rPr", "tblPr", "tblGrid", "trPr", "tcPr", "sectPr",
                  "bookmarkStart", "bookmarkEnd", "lastRenderedPageBreak", "proofErr",
                  "commentRangeStart", "commentRangeEnd", "permStart", "permEnd"}


def texts(value):
    return " ".join(value.split())


def title_body(value, common, rules):
    custom = (rules.get("heading_patterns") or {}).get("1")
    match = (re.compile(custom) if custom else common.RE_L1).match(value.strip())
    require(match is not None, "一级标题缺可解析的项目编号")
    tail = (value.strip()[match.end():] if custom else (match.group(2) or ""))
    # A custom pattern may consume the whole heading. Its first group still anchors numbering.
    if custom and not tail:
        tail = re.sub(r"^[\s.、:：)）章]+", "", value.strip()[match.end(1):])
    return texts(tail)


def semantic_xml(element, related, resolve_note=None, ancestors=()):
    local = element.tag.rsplit("}", 1)[-1]
    if resolve_note and element.tag in (W + "footnoteReference", W + "endnoteReference", W + "commentReference"):
        return [(element.tag, resolve_note(local, element.get(W + "id")))]
    if element.tag in {WP + n for n in ("extent", "effectExtent", "positionH", "positionV", "simplePos")} or (
            "pic" in ancestors and not set(ancestors).intersection(("grpSp", "wgp", "group"))
            and element.tag in (A + "off", A + "ext")
            and set(element.attrib) in ({"x", "y"}, {"cx", "cy"})):
        return []
    if local == "rPr" and element.tag.startswith(W):
        visibility = tuple((p.tag, p.get(W + "val", "1")) for p in element
                           if p.tag in (W + "vanish", W + "webHidden", W + "specVanish"))
        return [("visibility", visibility)] if visibility else []
    if element.tag.startswith(W) and local in DOC_PROPERTIES:
        # These fields carry layout or editing metadata. The actual files undergo their own Word audit.
        if local == "tcPr":
            return [tuple((x.tag, x.get(W + "val")) for x in element
                          if x.tag in (W + "gridSpan", W + "vMerge", W + "hMerge"))]
        return []
    if element.tag in (W + "t", W + "instrText", W + "delText"):
        return [element.text or ""]
    if element.tag == W + "r":
        return [v for child in element for v in semantic_xml(child, related, resolve_note, (*ancestors, local))]
    if element.tag == W + "br" and element.get(W + "type") == "page":
        return []
    # Only translucent containers are unwrapped; tracked revisions and fields remain evidence.
    if element.tag in (W + "sdt", W + "sdtContent", W + "customXml"):
        return [v for child in element if child.tag not in (W + "sdtPr", W + "customXmlPr")
                for v in semantic_xml(child, related, resolve_note, (*ancestors, local))]
    attrs = tuple(sorted((key, related(val) if key.startswith(R) else val)
                         for key, val in element.attrib.items()
                         if not key.startswith(W + "rsid") and not key.endswith(("}paraId", "}textId"))
                         and not (local in ("docPr", "cNvPr") and key in ("id", "name"))))
    content = []
    for child in element:
        for part in semantic_xml(child, related, resolve_note, (*ancestors, local)):
            if isinstance(part, str) and content and isinstance(content[-1], str):
                content[-1] += part
            else:
                content.append(part)
    return [(element.tag, attrs, element.text if not len(element) else None, content)]


def document_sections(document, common):
    """Ordered content blocks, including embedded-content digests and table structure."""
    fmt_blocks = common._docx_blocks(document, text_only=True)
    with zipfile.ZipFile(document) as z:
        root = ET.fromstring(z.read("word/document.xml")).find(W + "body")

        def rels(member):
            folder, name = posixpath.split(member)
            location = posixpath.join(folder, "_rels", name + ".rels")
            if location not in z.namelist():
                return {}
            return {x.get("Id"): x for x in ET.fromstring(z.read(location))}

        def closure(member, ancestors):
            def related(rid):
                ref = links.get(rid)
                require(ref is not None, "图表或文档关系在 DOCX 中无法闭合")
                target = ref.get("Target")
                kind = ref.get("Type")
                if ref.get("TargetMode") == "External":
                    return (kind, "external", target)
                name = (target.lstrip("/") if target.startswith("/") else
                        posixpath.normpath(posixpath.join(posixpath.dirname(member), target)))
                require(name not in ancestors, "DOCX 嵌入部件存在循环关系")
                data = z.read(name)
                if name.endswith((".xml", ".vml")):
                    return (kind, digest(semantic_xml(ET.fromstring(data), closure(name, ancestors | {name}))))
                return (kind, sha256(data).hexdigest())

            links = rels(member)
            return related

        def body_parts(node):
            for el in node:
                if el.tag in (W + "p", W + "tbl"):
                    yield el
                elif el.tag in (W + "sdt", W + "sdtContent", W + "customXml"):
                    yield from body_parts(el)

        main_member = "word/document.xml"
        related = closure(main_member, {main_member})
        document_rels = rels(main_member)
        note_categories = {"footnoteReference": "footnotes", "endnoteReference": "endnotes",
                           "commentReference": "comments"}
        notes = {}

        def resolve_note(kind, ident):
            if kind not in notes:
                role = note_categories[kind]
                relation = next((rel for rel in document_rels.values() if
                                 rel.get("Type", "").rsplit("/", 1)[-1] == role), None)
                require(relation is not None, "注释引用缺原部件")
                target = relation.get("Target")
                member = target.lstrip("/") if target.startswith("/") else posixpath.normpath("word/" + target)
                notes[kind] = (ET.fromstring(z.read(member)), closure(member, {member}))
            definitions, relations = notes[kind]
            record = next((note for note in definitions if note.get(W + "id") == ident), None)
            require(record is not None, "注释与正文引用不一致")
            record = deepcopy(record)
            record.attrib.pop(W + "id", None)
            return digest(semantic_xml(record, relations))
        sections = []
        for block, xml in zip(fmt_blocks, body_parts(root), strict=True):
            is_heading = block["outline"] == 1 and block["text"].strip() and block["table_index"] is None
            if is_heading:
                sections.append({"heading": texts(block["text"]), "fingerprints": [], "table_indices": []})
            elif sections and (block["text"].strip() or block["table_index"] is not None or
                               xml.find(".//" + W + "drawing") is not None or
                               xml.find(".//" + W + "pict") is not None or
                               xml.find(".//" + W + "object") is not None or
                               any(xml.find(".//" + W + kind) is not None for kind in note_categories)):
                sections[-1]["fingerprints"].append((block["outline"], digest(semantic_xml(xml, related, resolve_note))))
                if block["table_index"] is not None:
                    sections[-1]["table_indices"].append(block["table_index"])
        return sections


def cover_allowed(quote):
    # Actual separate files may deduct a cover only if the clause permits it for that scope.
    if not isinstance(quote, str):
        return False
    negative = r"(?:不得|禁止|不可|不许).{0,10}(?:封面|封皮)|不(?:得|能|予|应|可)?扣除.{0,10}(?:封面|封皮)|(?:封面|封皮)[^。；！？]{0,32}(?:不(?:得|能|予|应|可)?扣除|(?<!不)计入|应计|须计)"
    if re.search(negative, quote) or re.search(r"统一封面|总封面|(?i:common cover|cover.+must be counted|not excluding)", quote):
        return False
    china = r"(?:封面|封皮)[^。；！？]{0,16}(?:不(?:计|占|纳入|列入|计入|包含)|除外)|(?:不(?:含|计|包括|包含|纳入)|除去|扣除|除外)[^。；！？]{0,12}(?:封面|封皮)"
    if re.search(r"评审点|评分点|分册|分件|各份|各点|逐项|每份|每点|每个|每项|单独|分别", quote) and re.search(china, quote):
        return True
    english = r"\bcover(?:\s+pages?)?\b[^.;]{0,20}\b(?:excluded|not\s+counted)\b|\b(?:excluding|exclude)\s+(?:the\s+)?cover\b"
    return bool(re.search(r"per criterion|each criterion|per part|each file|per submission", quote, re.I)
                and re.search(english, quote, re.I))


def point_specs(project, rules, page_rule, common):
    items = read(project / "parts.json")["parts"]
    headings = page_rule["review_point_headings"]
    require(len(items) == len(headings) == page_rule["review_point_count"] == 10 and
            [x.get("index") for x in items] == list(range(1, 11)), "需按招标评审点顺序提供完整的十份实际稿")
    require(len({str(inside(project, x["docx"]).resolve()).casefold() for x in items}) == len(items),
            "十个评审点不能重复使用同一份文件")
    require(len({file_sha(inside(project, x["docx"])) for x in items}) == len(items),
            "实际提交存在完全重复的评审点文件")
    for item, original in zip(items, headings, strict=True):
        docx = inside(project, item["docx"])
        require(docx.suffix.casefold() == ".docx" and docx != project / "design.docx", "评审点须提供独立 DOCX")
        front, cover = item["front_pages"], item["cover_pages"]
        require(type(front) is int and type(cover) is int and 0 <= cover <= front and
                isinstance(item["source_ref"], str) and item["source_ref"].strip() and isinstance(item.get("source"), dict),
                "逐点封面、目录及正文起始页须登记招标依据")
        require(cover == 0 or cover_allowed(item["source"].get("quote")), "扣除封面须有原文明确依据")
        require(title_body(item["heading"], common, rules) == title_body(original, common, rules),
                "实际评审点编号与题名无法对应原评分章")
    return items


def compare_content(project, specs, page_rule, common, rules):
    reference = document_sections(project / "design.docx", common)
    require(len(reference) == len(page_rule["review_point_headings"]) and all(
        sec["heading"] in (texts(expected), title_body(expected, common, rules))
        for sec, expected in zip(reference, page_rule["review_point_headings"], strict=True)),
            "内容母稿的一级标题与评审点计划不同")
    indices = []
    for ref, spec in zip(reference, specs, strict=True):
        current = document_sections(inside(project, spec["docx"]), common)
        require(len(current) == 1 and current[0]["heading"] in (
                texts(spec["heading"]), title_body(spec["heading"], common, rules)) and
                current[0]["fingerprints"] and current[0]["fingerprints"] == ref["fingerprints"],
                f"第 {spec['index']} 项的实际稿与已审内容母稿正文、图表或顺序不同")
        indices.append(dict(zip(ref["table_indices"], current[0]["table_indices"], strict=True)))
    return indices


def point_format(rules, part, table_indices):
    scoped = deepcopy(rules)
    front, cover = part["front_pages"], part["cover_pages"]
    scoped.update(body_start_page=front + 1, required_level1_titles=[part["heading"]])
    fm = deepcopy(scoped["front_matter"])
    fm.update(cover_required=bool(cover), cover_page_range=[1, cover] if cover else [],
              toc_levels=part.get("toc_levels", fm["toc_levels"] if front else []),
              toc_page_range=part.get("toc_page_range", [1, front] if front else []),
              source_ref=part["source_ref"])
    levels, span = fm["toc_levels"], fm["toc_page_range"]
    require(isinstance(levels, list) and all(type(x) is int and x in (1, 2, 3) for x in levels) and
            levels == sorted(set(levels)) and
            (not levels or (isinstance(span, list) and len(span) == 2 and
                            all(type(x) is int for x in span) and 1 <= span[0] <= span[1] <= front)),
            "实际逐点目录页范围或目录层级无效")
    scoped["front_matter"] = fm
    if scoped.get("table_layout"):
        global_tables = {table["index"]: table for table in scoped["table_layout"].get("tables", [])}
        scoped["table_layout"]["tables"] = [
            global_tables[master_index] | {"index": part_index} for master_index, part_index in table_indices.items()
            if master_index in global_tables]
    locked = scoped.get("locked_level2_titles")
    if locked:
        scoped["locked_level2_titles"] = [
            rec | {"chapter": part["heading"]} for rec in locked
            if rec["chapter"] == part.get("master_heading", part["heading"])]
    return scoped


def native_measure(project, out, specs, page_rule, common, rules, point_word):
    indices = compare_content(project, specs, page_rule, common, rules)
    parts = []
    for part, tables, master_heading in zip(specs, indices, page_rule["review_point_headings"], strict=True):
        path = inside(project, part["docx"])
        scoped = point_format(rules, part | {"master_heading": master_heading}, tables)
        before = file_sha(path)
        with open_word(path) as (word, doc):
            total = int(doc.ComputeStatistics(2))
            front, cover = part["front_pages"], part["cover_pages"]
            actual, issues = point_word.measure_points(doc, page_rule | {
                "review_point_headings": [part["heading"]],
                "front_matter_pages_excluded": front})
            if len(actual) != 1 or actual[0]["heading"] != part["heading"] or actual[0]["start_page"] != front + 1:
                issues.append("未实排到本点唯一正文起始标题")
            records = common.get_paragraph_records(doc)
            headings, checks = common.extract_heading_structure(records, front + 1, scoped)
            checks.extend(common.check_front_page_and_body(doc, records, headings, scoped, total))
            checks.extend(common.check_toc_layout_word(doc, records, headings, scoped))
            checks.extend(common.check_table_layout_word(doc, scoped))
            checks.extend(common.check_pagination_controls(doc, records, headings, scoped))
            checks.extend(common.check_blank_and_low_content_pages(doc, total, scoped))
            checks.extend([common.check_heading_level_presence(headings),
                           common.check_heading_alignment(headings, scoped, doc),
                           common.check_heading_sequence(headings, common.classify_heading(part["heading"], scoped)[1]),
                           common.check_page_orientation(doc, scoped),
                           common.check_last_page_whitespace(doc, total, scoped),
                           common.check_fonts(doc, records, scoped)])
            issues.extend(x.name for x in checks if x.status == "FAIL")
            if not 1 <= total - cover <= page_rule["review_point_page_limit"]:
                issues.append("扣除允许的封面后，实际文件超过单点页数上限")
            record = {"index": part["index"], "docx": part["docx"], "document_sha256": before,
                      "word_version": str(word.Version), "physical_pages": total, "cover_pages": cover,
                      "non_cover_pages": total - cover, "failures": issues, "checks": [asdict(x) for x in checks]}
            if rules.get("submission_format") == "pdf":
                pdf = out / f"part-{part['index']:02d}.pdf"
                doc.ExportAsFixedFormat(str(pdf.resolve()), 17)
                from fitz import open as open_pdf
                with open_pdf(pdf) as exported:
                    if len(exported) != total:
                        record["failures"].append("实际 PDF 页数与同一稿件的 Word 实排不符")
                record.update(pdf=pdf.name, pdf_sha256=file_sha(pdf))
        parts.append(record)
    report = {"page_source": "microsoft_word", "parts": parts,
              "non_cover_pages": sum(row["non_cover_pages"] for row in parts)}
    total_rule = (read(project / "篇幅规划状态.json")["plan"]["bid_requirements"]["format"].get("pages") or {})
    extra = total_rule.get("additional_total")
    report["status"] = "PASS" if not page_failures(report, page_rule, extra, specs) else "FAIL"
    write(out / "part-pages.json", report)
    return report


def page_failures(report, rule, extra=None, specs=None):
    rows = report["parts"]
    failures = []
    if report["page_source"] != "microsoft_word" or [x.get("index") for x in rows] != list(range(1, 11)):
        return ["缺少按顺序实排的十份 Word 证据"]
    for row, spec in zip(rows, specs or rows, strict=True):
        physical, cover = row["physical_pages"], row["cover_pages"]
        if (row["failures"] or not row.get("word_version") or type(physical) is not int or type(cover) is not int or
                cover < 0 or row["non_cover_pages"] != physical - cover or
                not 1 <= physical - cover <= rule["review_point_page_limit"] or row["docx"] != spec["docx"] or
                cover != spec["cover_pages"]):
            failures.append(f"第 {row['index']} 项未通过实际提交件页数或版式审核")
    total = sum(x["non_cover_pages"] for x in rows)
    if report["non_cover_pages"] != total or total > rule["non_cover_page_limit"]:
        failures.append("实际十份稿件的非封面合计超限或计算不一致")
    if extra:
        scope = extra["scope"]
        if scope == "total":
            actual = sum(x["physical_pages"] for x in rows)
        elif scope == "body":
            require(specs is not None, "正文附加总限须登记逐份目录页范围")
            actual = sum(x["physical_pages"] - s["front_pages"] for x, s in zip(rows, specs, strict=True))
        else:
            actual = total
        if actual > extra["max"]:
            failures.append("实际全部提交件超过招标另定的整批上限")
    return failures


def verify_tender_scope(project, state, specs, layout):
    from bid_requirements import source_ok
    from hard_checks import source_lookup

    roles = {s["id"] for s in state["startup"]["sources"] if s["role"] in ("tender", "amendment")}
    lookup = source_lookup(state, project)
    require(source_ok(layout["source"], lookup, roles), "提交方式未绑定有效招标原文")
    require(all(source_ok(part["source"], lookup, roles) and
                (part["cover_pages"] == 0 or cover_allowed(part["source"]["quote"])) for part in specs),
            "逐点封面及前置页登记缺有效招标原文依据")


def release_candidates(project, report_dir, specs, page_rule, fmt, review, identity):
    report = read(report_dir / "part-pages.json")
    failures = page_failures(report, page_rule, (fmt.get("pages") or {}).get("additional_total"), specs)
    require(not failures and report.get("status") == "PASS", "；".join(failures) or "逐点报告未通过")
    artifacts = {}
    for row in report["parts"]:
        src = inside(project, row["docx"])
        require(file_sha(src) == row["document_sha256"], "逐点文件在 Word 审核后变化")
        if fmt["deliverable"] == "pdf":
            path = inside(report_dir, row["pdf"])
            checksum = row["pdf_sha256"]
            name = src.with_suffix(".pdf").name
        else:
            path, checksum, name = src, row["document_sha256"], src.name
        require(file_sha(path) == checksum, "实际待发布件与实排证据字节不一致")
        record = review.get("part_identity", {}).get(str(row["index"]), {})
        problems = identity.errors(path, fmt, record, logical_name=name)
        require(not problems, f"第 {row['index']} 项暗标核验未通过")
        require(name.casefold() not in {key.casefold() for key in artifacts}, "实际待发布文件同名")
        logical = (path.relative_to(project).as_posix() if path.is_relative_to(project) else
                   f"release-evidence/measure/{row['pdf']}")
        artifacts[name] = {"path": logical, "source": path, "sha256": checksum}
    return artifacts

