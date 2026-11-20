"""Portable synthetic source/response regressions; no Windows Word measurement."""
from pathlib import Path
import copy
import importlib.util
import json
import subprocess
import struct
import sys
import tempfile
import unittest


ROOT = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 and Path(sys.argv[1]).is_dir() else Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / 'scripts' if (ROOT / 'scripts').is_dir() else ROOT
sys.path.insert(0, str(SCRIPTS))
import source_coverage as sc


def write(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    return path


def evidence(path, locator='synthetic original fixture'):
    return {'path': str(path), 'sha256': sc.file_hash(path), 'locator': locator}


def response(rid, expectation, quote):
    return {'requirement_id': rid, 'status': 'adequate', 'conclusion': 'synthetic evidence contract',
            'aspects': [{'expectation': expectation, 'quote': quote, 'location': '第一章测试措施',
                         'adequacy': 'adequate', 'conclusion': 'synthetic adequacy record'}]}


def make_trace(base, body_text):
    """A minimal trace for the separate legacy format-regression fixtures."""
    folder = base / 'trace-originals'
    folder.mkdir(exist_ok=True)
    quote = sc.normal(body_text)[-120:]
    assert len(sc.normal(quote)) >= 20
    source = folder / 'synthetic-tender.txt'
    source.write_text(quote, encoding='utf-8')
    scores = write(base / 'trace-scores.json', {'tender_source': 'synthetic fixture line 1', 'items': [
        {'id': 'T', 'name': 'synthetic criterion', 'source_ref': 'line 1', 'criteria': [quote], 'max_score': 1}]})
    state = {'schema_version': '3.2', 'startup': {
        'sources': [evidence(source) | {'role': 'tender', 'id': 'S'}],
        'source_scope': {'roots': [str(folder)], 'excluded': [], 'reviewed': True, 'conclusion': 'fixture scope'}},
        'plan': {}, 'module_registry': {}}
    rows, texts = sc.scan(state, base)
    uid = rows[0]['units'][0]['id']
    state['plan'].update(requirements=[{
        'id': 'R', 'scope': quote, 'scoring': True, 'evidence': [evidence(source)],
        'source_refs': [{'source_id': 'S', 'unit_id': uid, 'quote': quote}],
        'score_refs': [{'item_id': 'T', 'criterion_index': 0}], 'response_requirements': ['fixture measure']}],
        scoring_basis={'path': str(scores), 'basis_sha256': sc.digest(sc.scoring_basis(json.loads(scores.read_text())))})
    state['startup']['source_coverage'].update(dispositions=[{
        'unit_ids': [uid], 'status': 'reviewed', 'disposition': 'requirements', 'requirement_ids': ['R'],
        'conclusion': 'fixture physical unit reviewed'}], extraction_review={
            key: {'reviewed': True, 'conclusion': 'fixture independent reading'}
            for key in ('source_pass', 'scoring_pass', 'scope_reconciliation')})
    path = write(base / 'trace-state.json', state)
    trace = {'version': 1, 'state': {'path': str(path), 'basis_sha256': sc.state_basis(state)},
        'original_recheck': [{'source_id': 'S', 'inventory_sha256': sc.digest(rows[0]),
            'reviewed_ranges': [[1, 1]], 'status': 'reviewed', 'conclusion': 'fixture original recheck'}],
        'scoring_recheck': {'reviewed': True, 'conclusion': 'fixture scoring recheck'},
        'body_to_source_recheck': {'reviewed': True, 'conclusion': 'fixture body recheck'},
        'unresolved': [], 'conclusion': 'synthetic test only', 'module_reviews': [],
        'responses': [response('R', 'fixture measure', quote)]}
    return trace, state['plan']['scoring_basis']['basis_sha256']


class SourceCoverageTests(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook
        from openpyxl.comments import Comment
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.base = Path(self.tmp.name)
        self.raw = self.base / 'originals'
        self.raw.mkdir()
        self.tender = self.raw / '招标原文.txt'
        self.criterion = '施工措施应针对现场限制并具有可操作性'
        self.tender.write_text(self.criterion + '\n覆土前完成第三方测绘并提交记录', encoding='utf-8')
        self.boq = self.raw / '工程量清单.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = '工程量'
        ws.append(['项目名称', '数量', '要求'])
        ws.append(['标线施工', 879, '基层清理及施工后保护；热熔涂料、玻璃珠、放样、试划、底漆、涂层厚度、逆反射检测'])
        fees = wb.create_sheet('费用及技术要求')
        fees.sheet_state = 'hidden'
        fees.append(['安全文明施工费', '=100*0.03', '投入使用审批及专款检查'])
        fees['C1'].comment = Comment('智慧工地费用对应实施责任', 'fixture')
        fees.row_dimensions[1].hidden = True
        wb.save(self.boq)
        wb.close()
        self.scores = write(self.base / '评分清单.json', {'tender_source': '招标原文第1行', 'items': [{
            'id': 'Q1', 'name': '针对性', 'source_ref': '第1行', 'criteria': [self.criterion], 'max_score': 10,
            'excellent_note': 'synthetic test criterion, no fixed excellent band'}]})
        self.state = {'schema_version': '3.2', 'startup': {
            'source_scope': {'roots': [str(self.raw)], 'excluded': [], 'reviewed': True, 'conclusion': '实际目录盘点'},
            'sources': [evidence(self.tender) | {'id': 'T', 'role': 'tender'},
                        evidence(self.boq) | {'id': 'B', 'role': 'boq'}]}, 'plan': {}, 'module_registry': {}}
        rows, texts = sc.scan(self.state, self.base)
        self.texts = texts
        self.rows = rows
        self.req = []
        self.dispositions = []
        for row in rows:
            for unit in row['units']:
                text = texts[unit['id']]
                technical = unit['locator'].startswith('文本') or 'A2=' in text or '安全文明施工费' in text
                d = {'unit_ids': [unit['id']], 'status': 'reviewed', 'conclusion': 'synthetic original-unit judgment'}
                if technical:
                    rid = f'R{len(self.req)+1}'
                    scoring = text == self.criterion
                    self.req.append({'id': rid, 'scope': text, 'scoring': scoring, 'evidence': [evidence(self.tender if row['source_id']=='T' else self.boq)],
                        'source_refs': [{'source_id': row['source_id'], 'unit_id': unit['id'], 'quote': text}],
                        'response_requirements': ['落实该项对象、做法及检查责任'],
                        'score_refs': [{'item_id': 'Q1', 'criterion_index': 0}]})
                    d.update(disposition='requirements', requirement_ids=[rid])
                else:
                    d.update(disposition='no_technical_response', reason='synthetic table metadata/header only')
                for flag in unit['flags']:
                    d[flag + '_reviewed'] = True
                if unit['flags']:
                    d['manual_evidence'] = [evidence(self.boq, unit['locator'])]
                self.dispositions.append(d)
        self.state['startup']['source_coverage'].update(dispositions=self.dispositions, extraction_review={
            k: {'reviewed': True, 'conclusion': 'synthetic full-source and scoring pass'}
            for k in ('source_pass', 'scoring_pass', 'scope_reconciliation')})
        self.state['plan'].update(requirements=self.req, scoring_basis={
            'path': str(self.scores), 'basis_sha256': sc.digest(sc.scoring_basis(json.loads(self.scores.read_text())))})
        self.quote = '项目经理按施工阶段落实费用投入和责任分工，专款使用履行审批并建立台账，检查结果形成记录。'

    def test_full_coverage_positive(self):
        self.assertEqual(sc.coverage_errors(self.state, self.base), [])

    def test_hidden_sheet_comments_formula_and_line_items_enumerated(self):
        text = '\n'.join(self.texts.values()) + '\n' + '\n'.join(u['locator'] for r in self.rows for u in r['units'])
        for content in ('安全文明施工费', '智慧工地费用', '=100*0.03', 'hidden', '隐藏=True', '标线施工'):
            self.assertIn(content, text)
        self.assertTrue(any('formula' in u['flags'] for r in self.rows for u in r['units']))

    def test_unregistered_attachment_blocks(self):
        (self.raw / '后续补遗.txt').write_text('现场不具备堆场条件')
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_omitted_registered_source_blocks(self):
        self.state['startup']['sources'].pop()
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_deleted_unit_cannot_reduce_denominator(self):
        self.state['startup']['source_coverage']['inventory'][1]['units'].pop()
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_reading_mark_alone_is_not_disposition(self):
        self.state['startup']['source_coverage']['dispositions'] = []
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_fee_requirement_unmapped_blocks(self):
        fee = next(r for r in self.req if '安全文明施工费' in r['scope'])
        self.state['plan']['requirements'].remove(fee)
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_no_response_requires_reason(self):
        next(d for d in self.dispositions if d['disposition'] == 'no_technical_response').pop('reason')
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_wrong_source_quote_blocks(self):
        self.req[1]['source_refs'][0]['quote'] = '原件不存在的项目要求'
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_missing_scoring_subcondition_blocks(self):
        data = json.loads(self.scores.read_text())
        data['items'][0]['criteria'].append('新增的检验与验收要求')
        write(self.scores, data)
        self.state['plan']['scoring_basis']['basis_sha256'] = sc.digest(sc.scoring_basis(data))
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_score_only_updates_preserve_basis(self):
        data = json.loads(self.scores.read_text())
        data['items'][0].update(completion='complete', self_score=9, evidence=[{'response':'current body'}])
        write(self.scores, data)
        self.assertEqual(sc.coverage_errors(self.state, self.base), [])

    def test_formula_and_original_visual_review_required(self):
        next(d for d in self.dispositions if d.get('formula_reviewed')).pop('formula_reviewed')
        self.assertTrue(sc.coverage_errors(self.state, self.base))

    def test_source_change_invalidates_and_scan_retains_only_unchanged_reading(self):
        before = len(self.dispositions)
        self.tender.write_text(self.tender.read_text() + '\n新增交通约束')
        self.assertTrue(sc.coverage_errors(self.state, self.base))
        self.state['startup']['sources'][0]['sha256'] = sc.file_hash(self.tender)
        sc.scan(self.state, self.base)
        retained = self.state['startup']['source_coverage']['dispositions']
        self.assertLess(len(retained), before)
        self.assertTrue(all(not d['unit_ids'][0].startswith('T:') for d in retained))

    def test_complete_response_positive(self):
        checks = [response(r['id'], r['response_requirements'][0], self.quote) for r in self.req]
        self.assertEqual(sc.response_errors(self.req, checks, self.quote), [])

    def test_response_omission_blocks(self):
        self.assertTrue(sc.response_errors(self.req, [], self.quote))

    def test_keyword_or_true_alone_blocks(self):
        row = response(self.req[0]['id'], self.req[0]['response_requirements'][0], '按规范执行')
        self.assertTrue(sc.response_errors(self.req[:1], [row], '按规范执行'))

    def test_partial_response_blocks(self):
        row = response(self.req[0]['id'], self.req[0]['response_requirements'][0], self.quote)
        row['aspects'][0]['adequacy'] = 'partial'
        self.assertTrue(sc.response_errors(self.req[:1], [row], self.quote))

    def test_fabricated_body_quote_blocks(self):
        row = response(self.req[0]['id'], self.req[0]['response_requirements'][0], self.quote)
        self.assertTrue(sc.response_errors(self.req[:1], [row], '其他正文'))

    def test_final_trace_and_gate_bookkeeping(self):
        trace, basis = make_trace(self.base, self.quote)
        record = {'source_response_review': trace}
        self.assertEqual(sc.final_response_errors(record, self.base, self.quote), [])
        path = Path(trace['state']['path'])
        state = json.loads(path.read_text())
        state.update(gate={'final_ready':'PASS'}, updated_at='new audit bookkeeping')
        write(path, state)
        self.assertEqual(sc.final_response_errors(record, self.base, self.quote), [])

    def test_final_requires_original_recheck_and_no_open_findings(self):
        trace, _ = make_trace(self.base, self.quote)
        trace['original_recheck'][0]['reviewed_ranges'] = []
        self.assertTrue(sc.final_response_errors({'source_response_review':trace}, self.base, self.quote))

    def test_final_cannot_bind_another_project_state(self):
        trace, _ = make_trace(self.base, self.quote)
        self.assertTrue(sc.final_response_errors({'source_response_review':trace}, self.base, self.quote, self.base/'other.json'))

    def test_unsupported_original_is_an_error_not_a_successful_read(self):
        p = self.raw / '设计.dwg'
        p.write_bytes(b'synthetic unsupported CAD')
        source = evidence(p) | {'id':'CAD', 'role':'drawing'}
        with self.assertRaises(ValueError):
            sc.inventory(source, self.base)

    def test_docx_tables_and_pdf_pages_are_units(self):
        from docx import Document
        from pypdf import PdfWriter
        p = self.base / 'source.docx'
        d = Document()
        d.add_paragraph('现场条件与工艺限制')
        d.add_table(rows=1, cols=1).cell(0,0).text = '竣工前完成CCTV检测'
        d.save(p)
        units = list(sc.native_units(p))
        self.assertEqual(len(units), 2)
        self.assertTrue(any('CCTV' in text for _, text, _ in units))
        pdf = self.base / 'scan.pdf'
        writer = PdfWriter()
        writer.add_blank_page(100, 100)
        writer.add_blank_page(100, 100)
        with pdf.open('wb') as f:
            writer.write(f)
        units = list(sc.native_units(pdf))
        self.assertEqual(len(units), 2)
        self.assertTrue(all('visual' in flags for _, _, flags in units))

    def test_scan_cli_does_not_claim_reading_complete(self):
        path = write(self.base/'state.json', self.state)
        result = subprocess.run([sys.executable, str(SCRIPTS/'source_coverage.py'), '--state', str(path),
            '--text-output', str(self.base/'cache.txt')], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('READING_AND_EXTRACTION_REVIEW_REQUIRED', result.stdout)
        self.assertEqual(json.loads(path.read_text())['gate']['preflight'], 'BLOCKED')

    def test_actual_legacy_biff_xls_is_read_with_manual_formula_boundary(self):
        # Small actual BIFF8 stream, parsed by xlrd rather than a mocked reader.
        rec = lambda n, p: struct.pack('<HH', n, len(p)) + p
        bof = lambda kind: rec(0x0809, struct.pack('<HHHHII', 0x0600, kind, 0x0DBB, 1997, 0x41, 6))
        name = '费用'.encode('utf-16-le')
        bounds = lambda offset: rec(0x0085, struct.pack('<IBBBB', offset, 1, 0, 2, 1) + name)
        prefix = bof(5) + rec(0x0042, struct.pack('<H', 1200)) + rec(0x00E0, b'\x00' * 20)
        old_globals = prefix + bounds(0) + rec(0x000A, b'')
        label = '安全文明施工费'.encode('utf-16-le')
        sheet = (bof(0x10) + rec(0x0200, struct.pack('<IIHHH', 0, 1, 0, 2, 0))
            + rec(0x0204, struct.pack('<HHHHB', 0, 0, 0, len(label)//2, 1) + label)
            + rec(0x0203, struct.pack('<HHHd', 0, 1, 0, 0.03)) + rec(0x000A, b''))
        path = self.base / 'fees.xls'
        path.write_bytes(prefix + bounds(len(old_globals)) + rec(0x000A, b'') + sheet)
        units = list(sc.native_units(path))
        self.assertTrue(any('安全文明施工费' in text for _, text, _ in units))
        self.assertIn('formula', units[0][2])
        self.assertIn('可见性=1', units[0][1])

    def test_flow_initializes_current_contract_and_blocks_unreviewed_sources(self):
        spec = importlib.util.spec_from_file_location('source_test_flow', SCRIPTS/'编制流程门禁.py')
        flow = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(flow)
        flow.PAGE_DIR = self.base / 'empty-page-rules'
        flow.PAGE_DIR.mkdir()
        path = self.base / 'init.json'
        self.assertEqual(flow.cmd_init(path, False), 2)
        state = json.loads(path.read_text())
        self.assertEqual(state['schema_version'], '3.2')
        self.assertEqual(flow.cmd_preflight(path), 2)

    def test_module_gate_requires_non_scoring_responses_too(self):
        spec = importlib.util.spec_from_file_location('source_module_flow', SCRIPTS/'编制流程门禁.py')
        flow = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(flow)
        module = {'id':'M', 'requirement_ids':[r['id'] for r in self.req], 'primary_for':[r['id'] for r in self.req]}
        self.state['plan']['modules'] = [module]
        review = {'technical_review':{'version':1,'checks':{
            k:{'status':'reviewed','conclusion':'synthetic technical judgment'} for k in flow.TECHNICAL_REVIEW_ITEMS}},
            'construction_language_review':{'reviewed':True,'unresolved_count':0,'conclusion':'synthetic'},
            'scoring_review':{'conclusion':'synthetic','conditions':[{
                'requirement_id':'R1','criterion_quote':self.criterion,'response_quote':self.quote,
                'evidence':[evidence(self.tender)],'targeted':True,'technical_depth':True,
                'project_basis':True,'substantive_response':True,'status':'satisfied','conclusion':'synthetic'}]},
            'response_review':{'checks':[response(r['id'],r['response_requirements'][0],self.quote) for r in self.req]}}
        self.assertEqual(flow.module_review_errors(self.state,module,self.quote,review,self.base), [])
        fee = next(r for r in self.req if '安全文明施工费' in r['scope'])
        review['response_review']['checks'] = [r for r in review['response_review']['checks'] if r['requirement_id'] != fee['id']]
        self.assertTrue(flow.module_review_errors(self.state,module,self.quote,review,self.base))

    def test_no_page_preflight_updates_phase_and_reason(self):
        spec = importlib.util.spec_from_file_location('source_nopage_flow', SCRIPTS/'编制流程门禁.py')
        flow = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(flow)
        flow.PAGE_DIR = self.base / 'no-pages'
        flow.PAGE_DIR.mkdir()
        startup = self.state['startup']
        startup.update(latest_updates={'status':'confirmed_none','asked_at':'fixture','reply_text':'no updates','reply_source':'synthetic reply'},
            critical_questions=[],questions_review='synthetic scope review',
            batch_plan={'count':1,'reason':'synthetic small scope','batches':['whole']},
            word_channel='unavailable',word_capability_evidence='synthetic portable test; no Word',real_photo_slots=[],
            visual_policy={'version':1,'real_photos':{'decision':'schematic','reason':'synthetic narrow road','replacement_figures':['施工接口示意']},
                           'color_figures':{'mode':'color'},'additional_figures':{'names':['施工接口示意'],'reason':'synthetic interface'}})
        self.state['plan'].update(revision=1,chapter_plan=[{'chapter':'第一章 施工安排'}],
            modules=[{'id':'M','chapter':'第一章 施工安排','section':'第一节 施工组织',
                'scope':'synthetic technical scope','basis':'synthetic score and source mapping',
                'requirement_ids':[r['id'] for r in self.req],'primary_for':[r['id'] for r in self.req],
                'evidence':[evidence(self.tender)]}],
            capacity_review={'all_requirements_mapped':True,'scope_supports_volume':True,
                'no_padding_or_repetition':True,'conclusion':'synthetic capacity review','unresolved_gaps':[]})
        plan = self.state['plan']
        chapter = plan['modules'][0]
        chapter['planned_pages'] = 5
        for n, name in enumerate(('第二章 施工方法', '第三章 质量控制'), 2):
            plan['chapter_plan'].append({'chapter':name})
            plan['modules'].append(dict(chapter, id='M'+str(n), chapter=name, planned_pages=3, primary_for=[]))
        words = '施工措施 现场限制 可操作性 覆土 第三方测绘 提交记录 标线施工 基层清理 施工后保护 安全文明施工费 投入使用审批 专款检查 智慧工地费用 实施责任 热熔涂料 玻璃珠 放样 试划 底漆 涂层厚度 逆反射检测'.split()
        plan['project_response'] = {
            'keywords':[{'term':t, 'requirement':next(r['id'] for r in self.req if t in r['scope'])} for t in words],
            'headings':{'2':words[:15], '3':words},
            'priorities':[{'chapters':[c['chapter']], 'basis':'synthetic scope'} for c in plan['chapter_plan']]}
        plan['hard_checks'] = {'bills':[{'source':'B','sha256':sc.file_hash(self.boq),
            'tables':[{'sheet':'工程量','title':'工程量','trade':'标线施工','trade_cell':[2,1]}]}],
            'figures':[{'name':n,'kind':k,'module':'M','basis':'测试项目图件规划',
                'elements':list(flow.hard_checks.ELEMENTS[k]) or ['作业区']}
                for n,k in flow.hard_checks.required_figures(self.state).items()],
            'measures':[{'id':'标线','module':'M','requirement':self.req[0]['id'],'basis':'试件招标条件','completion':'实测标线厚度'}],
            'schedule':{'tasks':[{'id':'A','module':'M'}]},'gaps':[]}
        self.state['phase']='PLANNING' 
        self.state['gate']={'reason':'未完成篇幅规划，禁止开始正式正文'}
        path=write(self.base/'no-page-state.json',self.state)
        self.assertEqual(flow.cmd_preflight(path),0)
        state=json.loads(path.read_text())
        self.assertEqual(state['phase'],'DRAFTING_ALLOWED')
        self.assertEqual(state['plan']['status'],'READY')
        self.assertNotIn('未完成',state['gate']['reason'])


if __name__ == '__main__':
    # Optional root path is consumed above, not passed to unittest.
    unittest.main(argv=[sys.argv[0]], verbosity=2)
