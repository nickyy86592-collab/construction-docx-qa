#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Original-source coverage and substantive-response evidence, v1.

Enumerates physical reading units independently of the AI's requirements list.
Checks records and quotes, never decides which engineering obligations exist.
No macros, formula evaluation, or third-party reviewer rules are executed.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


VERSION = 1
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def normal(text):
    return re.sub(r'\s+', '', str(text))


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def file_hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def resolve(path, base):
    p = Path(path)
    return p.resolve() if p.is_absolute() else (Path(base) / p).resolve()


def fresh(item, base):
    return (isinstance(item, dict) and isinstance(item.get('path'), str)
            and resolve(item['path'], base).is_file()
            and file_hash(resolve(item['path'], base)) == item.get('sha256'))


def scope_errors(state, base):
    """Inventory the designated source folders, including unregistered files."""
    startup = state.get('startup') or {}
    scope = startup.get('source_scope') or {}
    roots = scope.get('roots')
    if not isinstance(roots, list) or not roots:
        return ['缺原资料实际目录source_scope.roots，不能只盘点已登记文件']
    files, errors = set(), []
    for value in roots:
        path = resolve(value, base)
        if not path.exists():
            errors.append('资料目录/文件不存在：' + str(path)); continue
        files.update(p.resolve() for p in path.rglob('*') if p.is_file()) if path.is_dir() else files.add(path)
    known = {resolve(s['path'], base) for s in startup.get('sources') or []}
    excluded = scope.get('excluded') or []
    excluded_paths = []
    for item in excluded:
        if not fresh(item, base) or not item.get('reason'):
            errors.append('资料盘点排除项缺版本或具体原因')
        else:
            excluded_paths.append(resolve(item['path'], base))
    if len(excluded_paths) != len(set(excluded_paths)) or known.intersection(excluded_paths):
        errors.append('资料排除重复或与正式来源冲突')
    if files != known.union(excluded_paths):
        errors.append('原资料目录与登记/排除文件不一致：' + ','.join(
            p.name for p in sorted(files.symmetric_difference(known.union(excluded_paths)))[:8]))
    if scope.get('reviewed') is not True or not scope.get('conclusion'):
        errors.append('未完成原资料目录与附件总索引的范围核对')
    return errors


def state_basis(state):
    startup = state.get('startup') or {}
    return digest({'sources': startup.get('sources'), 'source_scope': startup.get('source_scope'),
                   'source_coverage': startup.get('source_coverage'),
                   'plan': {k: v for k, v in (state.get('plan') or {}).items() if k != 'status'},
                   'module_registry': state.get('module_registry') or {}})


def pdf_units(path):
    """Commit one whole extraction; fallback never duplicates a partial prefix."""
    import tempfile
    failures = []
    for engine in ('pypdf', 'pymupdf'):
        # CAD-export streams can exceed pypdf's decompression guard. Do not
        # disable that guard; use an independent reader and retain visual review.
        with tempfile.SpooledTemporaryFile(max_size=2 * 1024 * 1024, mode='w+t', encoding='utf-8') as spool:
            try:
                if engine == 'pypdf':
                    from pypdf import PdfReader
                    reader = PdfReader(str(path))
                    pages = reader.pages
                    count = len(pages)
                    for i, page in enumerate(pages, 1):
                        spool.write(json.dumps([f'PDF第{i}页', page.extract_text() or '', ['visual']], ensure_ascii=False) + '\n')
                else:
                    import pymupdf
                    with pymupdf.open(path) as reader:
                        if reader.needs_pass:
                            raise ValueError('PDF加密且未解锁')
                        count = len(reader)
                        for i, page in enumerate(reader, 1):
                            spool.write(json.dumps([f'PDF第{i}页', page.get_text() or '', ['visual']], ensure_ascii=False) + '\n')
                if not count:
                    raise ValueError('PDF无可枚举页面')
            except Exception as exc:
                failures.append(f'{engine}: {type(exc).__name__}: {exc}')
                continue
            spool.seek(0)
            if engine == 'pymupdf':
                print(f'PDF_READER_FALLBACK=pymupdf\nSOURCE={path}\nREASON={failures[0]}', file=sys.stderr)
            for line in spool:
                locator, text, flags = json.loads(line)
                yield locator, text, flags
            return
    raise ValueError(f'PDF读取失败：{path}；' + '；'.join(failures))


def native_units(path):
    """Yield (locator, text, manual_flags); no keyword-based filtering."""
    suffix = path.suffix.lower()
    if suffix == '.pdf':
        yield from pdf_units(path)
    elif suffix in ('.docx', '.docm'):
        with zipfile.ZipFile(path) as z:
            parts = ['word/document.xml'] + sorted(n for n in z.namelist()
                if re.fullmatch(r'word/(?:header\d+|footer\d+|footnotes|endnotes)\.xml', n))
            for part in parts:
                tree = ET.fromstring(z.read(part))
                container = tree.find(W + 'body') if part == 'word/document.xml' else tree
                for i, node in enumerate(list(container), 1):
                    flags = ['visual'] if node.tag == W + 'altChunk' or any(x.tag.endswith(('}drawing', '}pict', '}object'))
                                             for x in node.iter()) else []
                    if node.tag == W + 'tbl':
                        for j, row in enumerate(node.findall(W + 'tr'), 1):
                            text = ' | '.join(''.join(x.text or '' for x in cell.iter(W + 't'))
                                              for cell in row.findall(W + 'tc'))
                            yield f'{part}表块{i}行{j}', text, flags
                    else:
                        text = ''.join(x.text or '' for x in node.iter(W + 't'))
                        if text.strip() or flags:
                            yield f'{part}块{i}', text, flags
    elif suffix in ('.xlsx', '.xlsm'):
        import openpyxl
        formulas = openpyxl.load_workbook(path, data_only=False)
        values = openpyxl.load_workbook(path, data_only=True)
        try:
            for ws in formulas.worksheets:
                merges = [str(x) for x in ws.merged_cells.ranges]
                meta = f'工作表={ws.title};可见性={ws.sheet_state};合并区域={merges}'
                has_visual = bool(ws._images or ws._charts)
                yield f'工作表[{ws.title}]说明', meta, ['visual'] if has_visual else []
                for row in ws.iter_rows():
                    parts, flags = [], set()
                    for c in row:
                        if c.value is None and not c.comment:
                            continue
                        text = f'{c.coordinate}={c.value!s}'
                        if c.data_type == 'f':
                            text += f';缓存值={values[ws.title][c.coordinate].value!s}'
                            flags.add('formula')
                        if c.comment:
                            text += ';批注=' + c.comment.text
                        parts.append(text)
                    if parts:
                        hidden = bool(ws.row_dimensions[row[0].row].hidden)
                        yield f'工作表[{ws.title}]第{row[0].row}行(隐藏={hidden})', ' | '.join(parts), sorted(flags)
        finally:
            formulas.close()
            values.close()
    elif suffix == '.xls':
        import xlrd
        book = xlrd.open_workbook(str(path), formatting_info=True)
        try:
            for ws in book.sheets():
                # xlrd exposes cached results, not the original BIFF formula text.
                yield f'工作表[{ws.name}]说明', (
                    f'可见性={ws.visibility};合并区域={ws.merged_cells};'
                    '原表公式及图形必须回看，禁止把缓存值当成已核验公式'), ['formula', 'visual']
                notes = getattr(ws, 'cell_note_map', {})
                for r in range(ws.nrows):
                    parts = []
                    for c in range(ws.ncols):
                        value = ws.cell_value(r, c)
                        note = notes.get((r, c))
                        if value != '' or note:
                            text = f'列{c+1}={value!s}'
                            if note:
                                text += ';批注=' + note.text
                            parts.append(text)
                    if parts:
                        hidden = bool(getattr(ws.rowinfo_map.get(r), 'hidden', False))
                        yield f'工作表[{ws.name}]第{r+1}行(隐藏={hidden})', ' | '.join(parts), []
        finally:
            book.release_resources()
    elif suffix == '.csv':
        with path.open(encoding='utf-8-sig', newline='') as f:
            for i, row in enumerate(csv.reader(f), 1):
                yield f'CSV第{i}行', ' | '.join(row), []
    elif suffix in ('.txt', '.md'):
        for i, line in enumerate(path.read_text(encoding='utf-8-sig').splitlines(), 1):
            if line.strip():
                yield f'文本第{i}行', line, ['visual'] if '![' in line else []
    else:
        raise ValueError(f'不能直接枚举{suffix}，须登记经对应核验的readable_copy，禁止跳过原件')


def inventory(source, base):
    p = resolve(source['path'], base)
    if not fresh(source, base):
        raise ValueError('原件不存在或哈希已改变：' + str(p))
    copy = source.get('readable_copy')
    copy_binding = None
    if copy is not None:
        if (not fresh(copy, base) or copy.get('original_sha256') != source['sha256']
                or copy.get('correspondence_reviewed') is not True or not copy.get('conclusion')):
            raise ValueError('可读转换件缺原件绑定或完整对应核验：' + str(p))
        p = resolve(copy['path'], base)
        copy_binding = {k: copy[k] for k in ('path', 'sha256', 'original_sha256')}
    units, texts = [], {}
    for i, (locator, text, flags) in enumerate(native_units(p), 1):
        uid = f"{source['id']}:U{i:06d}"
        units.append({'id': uid, 'locator': locator, 'text_sha256': digest(text), 'flags': flags})
        texts[uid] = text
    if not units:
        raise ValueError('原件未产生可核查读取单元，须核对空文件/转换件：' + str(p))
    row = {'source_id': source['id'], 'path': source['path'], 'sha256': source['sha256'],
           'readable_copy': copy_binding, 'units': units}
    return row, texts


def scan(state, base):
    """Refresh only the reading denominator. Never set reading/adequacy to PASS."""
    sources = state.get('startup', {}).get('sources')
    if not isinstance(sources, list) or not sources:
        raise ValueError('先在startup.sources登记现有全部原件及版本哈希')
    problems = scope_errors(state, base)
    if problems:
        raise ValueError('；'.join(problems))
    rows, texts = [], {}
    used = set()
    for source in sources:
        source.setdefault('id', 'S' + digest(str(resolve(source['path'], base)))[:12])
        if source['id'] in used:
            raise ValueError('资料ID重复：' + source['id'])
        used.add(source['id'])
        row, content = inventory(source, base)
        rows.append(row)
        texts.update(content)
    previous = state['startup'].get('source_coverage') or {}
    old = {u['id']: u for r in previous.get('inventory', []) for u in r.get('units', [])}
    current = {u['id']: u for r in rows for u in r['units']}
    old_bindings = {u['id']: (r.get('sha256'), r.get('readable_copy'))
                    for r in previous.get('inventory', []) for u in r.get('units', [])}
    new_bindings = {u['id']: (r['sha256'], r['readable_copy']) for r in rows for u in r['units']}
    # Retain completed reading only for unchanged physical units.
    retained = [x for x in previous.get('dispositions', [])
                if x.get('unit_ids') and all(uid in old and old[uid] == current.get(uid)
                                            and old_bindings[uid] == new_bindings.get(uid)
                                            for uid in x['unit_ids'])]
    state['startup']['source_coverage'] = {
        'version': VERSION, 'inventory': rows, 'dispositions': retained,
        'extraction_review': previous.get('extraction_review', {}) if previous.get('inventory') == rows else {},
    }
    state['schema_version'] = '3.2'
    state.setdefault('gate', {}).update(preflight='BLOCKED', draft_ready='BLOCKED', final_ready='BLOCKED')
    return rows, texts


def scoring_basis(data):
    keys = ('id', 'name', 'source_ref', 'criteria', 'max_score', 'excellent_range', 'excellent_note')
    return {'tender_source': data.get('tender_source'),
            'items': [{k: item[k] for k in keys if k in item} for item in data.get('items', [])]}


def coverage_errors(state, base):
    """Require complete physical dispositions AND every scoring subcondition."""
    errors = scope_errors(state, base)
    coverage = state.get('startup', {}).get('source_coverage')
    if not isinstance(coverage, dict) or coverage.get('version') != VERSION:
        return errors + ['缺原资料覆盖source_coverage v1；先scan，再按原文与评分两个方向完成提取']
    sources = state.get('startup', {}).get('sources', [])
    rows, texts = [], {}
    try:
        for source in sources:
            row, content = inventory(source, base)
            rows.append(row)
            texts.update(content)
    except (KeyError, OSError, ValueError, ImportError, zipfile.BadZipFile) as exc:
        return ['原资料读取范围无法复核：' + str(exc)]
    ids = [r['source_id'] for r in rows]
    if not ids or len(ids) != len(set(ids)):
        errors.append('原资料ID为空或重复')
    if coverage.get('inventory') != rows:
        errors.append('原资料读取范围已改变或被删减，须重新scan并复查受影响内容')
    units = {u['id']: u for r in rows for u in r['units']}
    owners = {u['id']: r['source_id'] for r in rows for u in r['units']}
    reqs = (state.get('plan') or {}).get('requirements') or []
    reqmap = {r.get('id'): r for r in reqs if isinstance(r, dict)}
    seen, responded = set(), {}
    dispositions = coverage.get('dispositions')
    if not isinstance(dispositions, list):
        dispositions = []
    for d in dispositions:
        if not isinstance(d, dict):
            errors.append('原文处置记录必须是对象'); continue
        selected = d.get('unit_ids')
        if not isinstance(selected, list) or not selected or any(not isinstance(x, str) for x in selected):
            errors.append('原文处置缺实际unit_ids'); continue
        if len(selected) != len(set(selected)) or seen.intersection(selected) or not set(selected).issubset(units):
            errors.append('原文处置单元重复或引用未知范围')
        seen.update(selected)
        if len({owners.get(uid) for uid in selected}) != 1:
            errors.append('同一处置组不得跨原始文件')
        if d.get('status') != 'reviewed' or not d.get('conclusion'):
            errors.append('原文单元尚未实际阅读或缺结论：' + str(selected[:1]))
        kind = d.get('disposition')
        if kind == 'requirements':
            rids = d.get('requirement_ids')
            if not isinstance(rids, list) or not rids or not set(rids).issubset(reqmap):
                errors.append('原文重点尚未落入有效requirements：' + str(selected[:1]))
            else:
                for uid in selected:
                    responded[uid] = set(rids)
        elif kind == 'no_technical_response':
            if not d.get('reason') or d.get('requirement_ids'):
                errors.append('无需技术正文响应须留实际依据，不能混填需求')
        else:
            errors.append('原文重点未处置，不能以待核实/已读取放行')
        flags = {flag for uid in selected if uid in units for flag in units[uid]['flags']}
        for flag in flags:
            if d.get(flag + '_reviewed') is not True:
                errors.append('图像/版面或公式须回看原件：' + flag + '/' + str(selected[:1]))
        if flags and (not d.get('manual_evidence') or any(
                not fresh(x, base) or not x.get('locator') for x in d['manual_evidence'])):
            errors.append('原图/公式复核缺有效文件及定位证据')
    if seen != set(units):
        errors.append('原资料有未逐项处置的读取单元：' + ','.join(sorted(set(units) - seen)[:8]))
    linked = {}
    for rid, req in reqmap.items():
        refs = req.get('source_refs')
        if not isinstance(refs, list) or not refs:
            errors.append('内容需求缺原文单元定位source_refs：' + str(rid)); refs = []
        for ref in refs:
            if not isinstance(ref, dict):
                errors.append('内容需求原文引用须为对象：' + str(rid)); continue
            uid = ref.get('unit_id')
            quote = ref.get('quote')
            if (uid not in units or ref.get('source_id') != owners.get(uid)
                    or not isinstance(quote, str) or len(normal(quote)) < 4):
                errors.append('内容需求缺有效原文定位/引文：' + str(rid)); continue
            # Scanned graphics and original BIFF formulas require documented visual verification.
            if normal(quote) not in normal(texts[uid]):
                if not units[uid]['flags'] or ref.get('manual_transcription_verified') is not True:
                    errors.append('需求原文引文不在登记单元，图文转录须另行实际核验：' + str(rid))
            if rid not in responded.get(uid, set()):
                errors.append('原文单元处置与内容需求未双向对应：' + str(rid))
            linked.setdefault(uid, set()).add(rid)
        expectations = req.get('response_requirements')
        if (not isinstance(expectations, list) or not expectations
                or any(not isinstance(x, str) or not x.strip() for x in expectations)
                or len(expectations) != len(set(expectations))):
            errors.append('需求缺具体论述完成条件response_requirements：' + str(rid))
        if req.get('scoring') is True:
            roles = {s.get('id'): s.get('role') for s in sources}
            original = ''.join(x.get('quote', '') for x in refs if isinstance(x, dict)
                               and roles.get(x.get('source_id')) in ('tender', 'amendment'))
            if normal(req.get('scope', '')) not in normal(original):
                errors.append('评分子条件未保留招标/有效补遗原文引文：' + str(rid))
    for uid, rids in responded.items():
        if rids != linked.get(uid, set()):
            errors.append('原文重点处置引用了无本单元依据的需求：' + uid)
    extraction = coverage.get('extraction_review') or {}
    for key in ('source_pass', 'scoring_pass', 'scope_reconciliation'):
        value = extraction.get(key) or {}
        if value.get('reviewed') is not True or not value.get('conclusion'):
            errors.append('缺原文遍历/评分反查/范围核对实际结论：' + key)
    basis = (state.get('plan') or {}).get('scoring_basis') or {}
    try:
        data = json.loads(resolve(basis['path'], base).read_text(encoding='utf-8-sig'))
        if basis.get('basis_sha256') != digest(scoring_basis(data)):
            errors.append('评分原文条件已改变或缺scoring_basis指纹')
        item_list = data.get('items') or []
        item_ids = [x['id'] for x in item_list]
        if not item_list or len(item_ids) != len(set(item_ids)):
            errors.append('评分清单为空或ID重复')
        expected = {(item['id'], i): text for item in item_list
                    for i, text in enumerate(item.get('criteria') or [])}
        if any(not item.get('criteria') or not item.get('source_ref') for item in item_list):
            errors.append('评分项缺优秀/评判条件或招标原文位置')
        covered = []
        for rid, req in reqmap.items():
            refs = req.get('score_refs') or []
            if req.get('scoring') is True and len(refs) != 1:
                errors.append('每个评分需求须对应一个明确评分子条件：' + str(rid))
            for ref in refs:
                key = (ref.get('item_id'), ref.get('criterion_index'))
                if type(key[1]) is not int or key not in expected:
                    errors.append('需求引用未知评分子条件：' + str(rid)); continue
                if req.get('scoring') is True:
                    covered.append(key)
                    if req.get('scope') != expected[key]:
                        errors.append('评分需求scope须保留对应条件原文：' + str(rid))
        if set(covered) != set(expected) or len(covered) != len(set(covered)):
            errors.append('评分清单的子条件未完整、唯一分配到评分需求')
    except (KeyError, TypeError, OSError, ValueError) as exc:
        errors.append('无法核对既有项目评分清单scoring_basis：' + str(exc))
    return errors


def response_errors(requirements, checks, body_text):
    errors = []
    reqmap = {r['id']: r for r in requirements}
    if not isinstance(checks, list):
        return ['缺逐项正文充分性复核response_review.checks']
    seen = []
    for row in checks:
        if not isinstance(row, dict):
            errors.append('正文充分性记录必须为对象'); continue
        rid = row.get('requirement_id')
        seen.append(rid)
        if rid not in reqmap:
            errors.append('正文充分性引用未知需求：' + str(rid)); continue
        if row.get('status') != 'adequate' or not row.get('conclusion'):
            errors.append('正文响应尚不充分：' + rid)
        expectations = reqmap[rid].get('response_requirements') or []
        aspects = row.get('aspects')
        if not isinstance(aspects, list):
            aspects = []
        actual = [x.get('expectation') for x in aspects if isinstance(x, dict)]
        if not expectations or sorted(actual) != sorted(expectations):
            errors.append('正文充分性未逐项核对完成条件：' + rid)
        for aspect in aspects:
            if not isinstance(aspect, dict):
                errors.append('论述条件复核须为对象：' + rid); continue
            quote = aspect.get('quote')
            if (not isinstance(quote, str) or len(normal(quote)) < 20
                    or normal(quote) not in normal(body_text) or not aspect.get('location')):
                errors.append('论述条件缺当前正文实际措施引文及位置：' + rid)
            if aspect.get('adequacy') != 'adequate' or not aspect.get('conclusion'):
                errors.append('论述条件尚未充分落实：' + rid)
    if len(seen) != len(set(seen)) or set(seen) != set(reqmap):
        errors.append('正文充分性复核遗漏、重复或超出责任需求')
    return errors


def final_response_errors(record, base, body_text, expected_state=None):
    review = record.get('source_response_review')
    if not isinstance(review, dict) or review.get('version') != VERSION:
        return ['缺最终原件独立反查及充分性复核source_response_review v1']
    binding = review.get('state') or {}
    if not binding.get('path') or not resolve(binding['path'], base).is_file():
        return ['最终原件反查未绑定当前项目状态文件']
    state_path = resolve(binding['path'], base)
    if expected_state is not None and state_path != Path(expected_state).resolve():
        return ['最终原件反查与交付使用了不同项目状态']
    state = json.loads(state_path.read_text(encoding='utf-8-sig'))
    errors = coverage_errors(state, state_path.parent)
    if binding.get('basis_sha256') != state_basis(state):
        errors.append('最终原件反查的来源、规划或模块版本已经失效')
    coverage = state.get('startup', {}).get('source_coverage') or {}
    inventories = {r['source_id']: r for r in coverage.get('inventory', [])}
    rechecks = review.get('original_recheck') or []
    seen = []
    for check in rechecks:
        sid = check.get('source_id')
        seen.append(sid)
        inv = inventories.get(sid)
        if inv is None or check.get('inventory_sha256') != digest(inv):
            errors.append('原件独立反查未绑定实际读取范围：' + str(sid)); continue
        ranges = check.get('reviewed_ranges') or []
        indices = []
        for pair in ranges:
            if (not isinstance(pair, list) or len(pair) != 2 or any(type(x) is not int for x in pair)
                    or not 1 <= pair[0] <= pair[1] <= len(inv['units'])):
                errors.append('原件独立反查范围无效：' + sid); continue
            indices.extend(range(pair[0], pair[1] + 1))
        if sorted(indices) != list(range(1, len(inv['units']) + 1)):
            errors.append('原件独立反查漏页/漏表或重复范围：' + sid)
        if check.get('status') != 'reviewed' or not check.get('conclusion'):
            errors.append('原件独立反查尚未完成：' + sid)
    if len(seen) != len(set(seen)) or set(seen) != set(inventories):
        errors.append('最终反查遗漏原始文件或重复登记')
    for key in ('scoring_recheck', 'body_to_source_recheck'):
        item = review.get(key) or {}
        if item.get('reviewed') is not True or not item.get('conclusion'):
            errors.append('最终缺评分优秀条件/正文依据反查：' + key)
    if review.get('unresolved') != [] or not review.get('conclusion'):
        errors.append('最终重点响应仍有未完成项或缺反查结论')
    checks = list(review.get('responses') or [])
    for module in review.get('module_reviews') or []:
        if not fresh(module, base):
            errors.append('复用的模块充分性记录缺失或已变更'); continue
        path = resolve(module['path'], base)
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        registered = (state.get('module_registry') or {}).get(data.get('module_id')) or {}
        if registered.get('review_sha256') != file_hash(path):
            errors.append('复用充分性记录不是当前登记的模块复核'); continue
        source = {'path': registered.get('source_path'), 'sha256': data.get('source_sha256')}
        if not fresh(source, state_path.parent):
            errors.append('复用充分性记录对应的模块正文已失效'); continue
        checks.extend((data.get('response_review') or {}).get('checks') or [])
    errors += response_errors((state.get('plan') or {}).get('requirements') or [], checks, body_text)
    return errors


def main():
    parser = argparse.ArgumentParser(description='枚举原资料真实范围，不自动声称已读或已充分响应')
    parser.add_argument('--state', type=Path, required=True)
    parser.add_argument('--text-output', type=Path, help='可选原文读取缓存，不能代替原图核对')
    parser.add_argument('--scoring-checklist', type=Path, help='绑定既有项目评分清单的原文依据指纹')
    parser.add_argument('--check', action='store_true', help='只核验已填记录并输出最终反查绑定值，不修改状态')
    args = parser.parse_args()
    try:
        state = json.loads(args.state.read_text(encoding='utf-8-sig'))
        if args.check:
            if args.scoring_checklist or args.text_output:
                raise ValueError('--check不能同时修改来源或输出原文缓存')
            problems = coverage_errors(state, args.state.parent)
            coverage = state.get('startup', {}).get('source_coverage') or {}
            print(json.dumps({'status': 'FAIL' if problems else 'PASS', 'failures': problems,
                'state_basis_sha256': state_basis(state),
                'inventories': [{'source_id': r['source_id'], 'inventory_sha256': digest(r),
                                  'unit_count': len(r['units'])} for r in coverage.get('inventory', [])]},
                ensure_ascii=False, indent=2))
            return 2 if problems else 0
        rows, texts = scan(state, args.state.parent)
        if args.scoring_checklist:
            score_path = args.scoring_checklist.resolve()
            data = json.loads(score_path.read_text(encoding='utf-8-sig'))
            state.setdefault('plan', {})['scoring_basis'] = {
                'path': str(score_path), 'basis_sha256': digest(scoring_basis(data))}
        # Save atomically only after all source formats were successfully enumerated.
        temp = args.state.with_suffix(args.state.suffix + '.tmp')
        temp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding='utf-8')
        temp.replace(args.state)
        if args.text_output:
            args.text_output.parent.mkdir(parents=True, exist_ok=True)
            args.text_output.write_text('\n\n'.join(
                f"[{u['id']}] {row['path']} {u['locator']}\n{texts[u['id']]}"
                for row in rows for u in row['units']), encoding='utf-8')
        print(f"SOURCE_INVENTORY_CREATED={sum(len(r['units']) for r in rows)}")
        print('READING_AND_EXTRACTION_REVIEW_REQUIRED')
        return 0
    except (KeyError, TypeError, OSError, ValueError, ImportError, zipfile.BadZipFile) as exc:
        print('SOURCE_SCAN_ERROR=' + str(exc))
        return 3


if __name__ == '__main__':
    raise SystemExit(main())
