"""Portable process regressions from failed-run patterns; no Windows Word claims."""
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import importlib.util
import io
import json
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / 'scripts'
sys.path.insert(0, str(SCRIPTS))
import source_coverage as sc
spec = importlib.util.spec_from_file_location('workflow_under_test', SCRIPTS / '编制流程门禁.py')
flow = importlib.util.module_from_spec(spec)
spec.loader.exec_module(flow)


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.base = Path(self.tmp.name)
        self.state_path = self.base / 'state.json'
        self.pages = self.base / 'page-rules'
        self.pages.mkdir()
        (self.pages / '测试页数硬审.py').write_text('EXPECTED_BODY_PAGES=50\nPAGE_TOLERANCE=0\nFRONT_MATTER_PAGES=1\n')
        self.old_pages = flow.PAGE_DIR
        flow.PAGE_DIR = self.pages
        self.addCleanup(setattr, flow, 'PAGE_DIR', self.old_pages)
        from docx import Document
        self.document = self.base / 'current.docx'
        doc = Document()
        doc.add_paragraph('项目经理落实分区施工，隐蔽验收合格后方可转序，异常立即停作并留存复验记录。')
        doc.save(self.document)

    def write(self, path, value):
        path.write_text(json.dumps(value, ensure_ascii=False), encoding='utf-8')
        return path

    def state(self, completed=0):
        modules = [{'id': f'M{i}', 'requirement_ids': [], 'primary_for': []} for i in range(1, 5)]
        script, vals = flow.discover_page_script()
        rule = flow.current_rule(script, vals) if script else None
        state = {'schema_version': '3.2', 'page_rule': rule, 'startup': {'sources': []}, 'plan': {
            'modules': modules, 'requirements': [], 'scoring_basis': {'basis_sha256': 'test-basis'},
            'checkpoint_module_ids': {'40': ['M1'], '75': ['M1', 'M2', 'M3']}},
            'module_registry': {}, 'checkpoints': {}, 'gate': {'preflight': 'PASS', 'draft_ready': 'PASS'}}
        for module in modules[:completed]:
            mid = module['id']
            source = self.base / (mid + '.md')
            source.write_text('真实保存的测试模块正文与对应复核文件。' * 20)
            review = self.write(self.base / (mid + '.json'), {'module_id': mid})
            state['module_registry'][mid] = {'status': 'complete', 'source_path': str(source),
                'source_sha256': flow.sha256(source), 'review_path': str(review), 'review_sha256': flow.sha256(review)}
        state['gate']['plan_hash'] = flow.plan_hash(state)
        state['gate']['startup_sha256'] = sc.digest(state['startup'])
        self.write(self.state_path, state)
        return state

    def next(self):
        out = io.StringIO()
        with redirect_stdout(out):
            rc = flow.cmd_next(self.state_path)
        return rc, json.loads(out.getvalue())

    def score(self, stage=40):
        return self.write(self.base / f'score-{stage}.json', {'phase': str(stage), 'status': 'PASS',
            'document_sha256': flow.sha256(self.document), 'scoring_basis_sha256': 'test-basis'})

    def delivery_args(self, mode='full'):
        return SimpleNamespace(state=self.state_path, document=self.document, mode=mode, result=None,
            static_report=None, word_report=None, page_report=None, score_report=None,
            review=None, warning_review=None)

    def test_preflight_cannot_be_replaced_by_existing_body_files(self):
        state = self.state(completed=4)
        state['gate']['preflight'] = 'ERROR'
        self.write(self.state_path, state)
        rc, status = self.next()
        self.assertEqual(rc, 2)
        self.assertEqual(status['next_action'], 'preflight')
        self.assertFalse(status['task_complete'])

    def test_due_checkpoint_is_returned_before_next_new_module(self):
        self.state(completed=1)
        rc, status = self.next()
        self.assertEqual(rc, 2)
        self.assertEqual(status['next_action'], 'checkpoint')
        self.assertEqual(status['details'][0], '40')

    def test_old_measurement_without_score_cannot_advance(self):
        state = self.state(completed=1)
        state['checkpoints']['40'] = {'status': 'PASS', 'module_ids': ['M1'],
            'snapshot_path': str(self.document), 'document_sha256': flow.sha256(self.document)}
        self.write(self.state_path, state)
        self.assertEqual(self.next()[1]['next_action'], 'checkpoint')

    def test_removing_project_page_script_does_not_remove_gate(self):
        self.state()
        self.pages.joinpath('测试页数硬审.py').unlink()
        self.assertEqual(self.next()[1]['next_action'], 'preflight')

    def test_force_init_cannot_delete_existing_history(self):
        self.state()
        before = self.state_path.read_bytes()
        with redirect_stdout(io.StringIO()):
            self.assertEqual(flow.cmd_init(self.state_path, True), 2)
        self.assertEqual(self.state_path.read_bytes(), before)

    def test_first_checkpoint_cannot_be_backfilled_after_later_registered_modules(self):
        state = self.state(completed=4)
        errors = flow.checkpoint_modules_errors(state, self.base, 40, self.document)
        self.assertTrue(any('不能事后补造' in x for x in errors))

    def test_next_module_after_fresh_word_and_score_evidence(self):
        state = self.state(completed=1)
        word = self.write(self.base / 'word.json', {'status': 'MEASURED'})
        score = self.score()
        state['checkpoints']['40'] = {'status': 'PASS', 'module_ids': ['M1'],
            'snapshot_path': str(self.document), 'document_sha256': flow.sha256(self.document),
            'module_source_hashes': {'M1': state['module_registry']['M1']['source_sha256']},
            'word_report': {'report_file': str(word), 'report_sha256': flow.sha256(word)},
            'score_report': {'path': str(score), 'sha256': flow.sha256(score)}}
        self.write(self.state_path, state)
        status = self.next()[1]
        self.assertEqual(status['next_action'], 'module-checkpoint')
        self.assertEqual(status['details'][0], 'M2')

    def test_temporary_estimate_remains_provisional_not_final_success(self):
        state = self.state(completed=1)
        score = self.score()
        state['checkpoints']['40'] = {'status': 'PROVISIONAL', 'page_source': 'libreoffice_estimate', 'module_ids': ['M1'],
            'snapshot_path': str(self.document), 'document_sha256': flow.sha256(self.document),
            'score_report': {'path': str(score), 'sha256': flow.sha256(score)}}
        self.write(self.state_path, state)
        status = self.next()[1]
        self.assertEqual(status['details'][0], 'M2')
        self.assertFalse(status['task_complete'])
        with redirect_stdout(io.StringIO()):
            self.assertEqual(flow.cmd_final_ready(self.state_path, self.document), 2)
        reasons = json.loads(self.state_path.read_text())['gate']['final_ready_reason']
        self.assertIn('40% 阶段尚未使用 Microsoft Word 实测', reasons)

    def test_stage_score_must_match_document_phase_and_scoring_basis(self):
        state = self.state()
        self.assertTrue(flow.stage_score_errors(state, self.document, 40, None))
        score = self.score()
        self.assertEqual(flow.stage_score_errors(state, self.document, 40, score), [])
        self.assertTrue(flow.stage_score_errors(state, self.document, 75, score))
        state['plan']['scoring_basis']['basis_sha256'] = 'changed'
        self.assertTrue(flow.stage_score_errors(state, self.document, 40, score))

    def test_changed_authorization_invalidates_previous_readiness(self):
        state = self.state()
        state['startup']['latest_updates'] = {'status': 'authorized_current_sources'}
        self.write(self.state_path, state)
        self.assertEqual(self.next()[1]['next_action'], 'preflight')
        self.assertTrue(any('授权记录改变' in x for x in flow.production_errors(state, self.base, 50)))

    def test_changed_readable_copy_invalidates_previous_readiness(self):
        state = self.state()
        source = self.base / 'original.bin'
        source.write_bytes(b'original')
        copy = self.base / 'readable.txt'
        copy.write_text('first extraction')
        state['startup']['sources'] = [{'path': str(source), 'sha256': flow.sha256(source),
            'readable_copy': {'path': str(copy), 'sha256': flow.sha256(copy)}}]
        state['gate']['startup_sha256'] = sc.digest(state['startup'])
        self.write(self.state_path, state)
        copy.write_text('changed after extraction')
        self.assertEqual(self.next()[1]['next_action'], 'preflight')

    def test_next_distinguishes_pilot_budget_calibration_from_expired_evidence(self):
        # Bound fixtures exercise next's decisions without claiming a Word run.
        (self.pages / '测试页数硬审.py').write_text(
            'EXPECTED_BODY_PAGES=300\nPAGE_TOLERANCE=0\nFRONT_MATTER_PAGES=1\n')
        state = self.state()
        state['plan']['pilot_plan'] = {'module_ids': ['M1', 'M2']}
        module = state['plan']['modules'][0]
        module.update(estimated_normal_chars=5000, text_pages=2)
        evidence = {'snapshot': self.document}
        for name in ('report', 'analysis', 'format_rules'):
            evidence[name] = self.write(self.base / f'pilot-{name}.json', {'fixture': name})
        state['pilot'] = {
            'status': 'PASS', 'page_source': 'microsoft_word', 'normal_chars_per_page': 1000,
            'page_rule': state['page_rule'], 'module_ids': ['M1', 'M2'],
            **{name: {'path': str(path), 'sha256': flow.sha256(path)} for name, path in evidence.items()},
        }
        state['gate']['draft_ready'] = 'BLOCKED'
        state['gate']['plan_hash'] = flow.plan_hash(state)
        self.write(self.state_path, state)
        rc, status = self.next()
        self.assertEqual(rc, 2)
        self.assertEqual(status['next_action'], 'preflight')
        self.assertTrue(any('重算' in detail and 'M1' in detail for detail in status['details']))

        module['text_pages'] = 5
        state['gate']['plan_hash'] = flow.plan_hash(state)
        self.write(self.state_path, state)
        self.assertEqual(self.next()[1]['next_action'], 'draft-ready')
        evidence['report'].write_text('changed after pilot')
        _, status = self.next()
        self.assertEqual(status['next_action'], 'pilot')
        self.assertIn('样章证据失效：report', status['details'])

    def test_authorization_is_distinct_from_no_updates(self):
        state = self.state()
        latest = {'status': 'authorized_current_sources', 'asked_at': 'actual-time',
            'reply_text': '暂按现有资料', 'reply_source': 'user reply',
            'safe_scope': '不擅定现场事实，按有依据的条件措施继续', 'pending_confirmation': '正式提交前核最新文件'}
        state['startup']['latest_updates'] = latest
        messages = flow.startup_errors(state, self.base)
        self.assertFalse(any('尚无用户明确答复' in x or '按现有资料继续须' in x for x in messages))
        latest.pop('safe_scope')
        self.assertTrue(any('按现有资料继续须' in x for x in flow.startup_errors(state, self.base)))

    def test_pending_is_not_misreported_as_not_applicable(self):
        review = {'document_sha256': flow.sha256(self.document), 'source_checks': [
            {'item': name, 'status': 'pending', 'conclusion': '尚未核验'} for name in flow.REVIEW_ITEMS]}
        path = self.write(self.base / 'pending.json', review)
        errors, _ = flow.validate_content_review(self.document, path)
        self.assertTrue(any('回查尚未关闭' in x for x in errors))
        self.assertFalse(any('不适用' in x for x in errors))

    def test_templates_preserve_unreviewed_status_and_refuse_overwrite(self):
        self.state()
        args = SimpleNamespace(state=self.state_path, kind='module', id='M1',
            source=self.base / 'M1.md', document=None, output=self.base / 'template.json')
        with redirect_stdout(io.StringIO()):
            flow.cmd_template(args)
        record = json.loads(args.output.read_text())
        self.assertFalse(record['technical_depth_reviewed'])
        self.assertEqual(record['technical_review']['checks']['number_basis']['status'], 'pending')
        with self.assertRaises(RuntimeError):
            flow.cmd_template(args)

    def test_failed_delivery_overwrites_old_release_and_returns_nonzero(self):
        self.state()
        path = self.write(self.base / '流程交付结果.json', {'task_complete': True})
        args = self.delivery_args()
        args.result = self.base / 'custom-result.json'
        with redirect_stdout(io.StringIO()):
            rc = flow.delivery_result(args, ['内容未核验', '内容未核验'])
        result = json.loads(path.read_text())
        self.assertEqual(rc, 2)
        self.assertFalse(result['task_complete'])
        self.assertFalse(result['formal_release'])
        self.assertEqual(result['failures'], ['内容未核验'])
        self.assertEqual(args.result.read_bytes(), path.read_bytes())

    def test_draft_and_local_result_cannot_become_full_formal_release(self):
        self.state()
        for mode, stage in [('full', 'draft'), ('local-review', 'final')]:
            with redirect_stdout(io.StringIO()):
                self.assertEqual(flow.delivery_result(self.delivery_args(mode), [], stage), 0)
            result = json.loads((self.base / '流程交付结果.json').read_text())
            self.assertFalse(result['task_complete'])
            self.assertFalse(result['formal_release'])
            self.assertEqual(result['status'], 'LOCAL_REVIEW_ALLOWED' if mode == 'local-review' else 'DRAFT_READY')

    def test_full_final_completion_and_report_change_invalidate_next(self):
        # No-page projects still require all content modules and bound release evidence.
        self.pages.joinpath('测试页数硬审.py').unlink()
        state = self.state(completed=4)
        state['gate']['final_ready'] = 'PASS'
        self.write(self.state_path, state)
        args = self.delivery_args()
        default_result = self.base / '流程交付结果.json'
        for result_path in (None, self.base / 'archive' / 'custom-result.json'):
            with self.subTest(result_path=result_path):
                default_result.unlink(missing_ok=True)
                args.result = result_path
                for key in ('static_report', 'word_report', 'score_report', 'review', 'warning_review'):
                    setattr(args, key, self.write(self.base / (key + '.json'), {'fixture': key}))
                with redirect_stdout(io.StringIO()):
                    flow.delivery_result(args, [], 'final')
                rc, status = self.next()
                self.assertEqual(rc, 0)
                self.assertTrue(status['task_complete'])
                if result_path:
                    self.assertEqual(result_path.read_bytes(), default_result.read_bytes())
                previous = json.loads(default_result.read_text())
                self.write(default_result, {**previous, 'version': 1})
                self.assertEqual(self.next()[1]['next_action'], 'delivery')
                self.write(default_result, previous)
                args.review.write_text('changed after audit')
                self.assertFalse(self.next()[1]['task_complete'])

    def test_cli_export_error_saves_failure_in_default_result(self):
        state = self.state()
        state['gate']['final_ready'] = 'PASS'
        self.write(self.state_path, state)
        default_result = self.write(self.base / '流程交付结果.json', {'task_complete': True})
        custom_result = self.base / 'result-is-a-directory.json'
        custom_result.mkdir()
        argv = ['gate', '--page-dir', str(self.pages), '--state', str(self.state_path),
                'delivery', '--document', str(self.document), '--result', str(custom_result)]
        for key in ('static-report', 'word-report', 'score-report', 'review'):
            argv.extend(['--' + key, str(self.base / (key + '.json'))])
        with patch.object(sys, 'argv', argv), \
                patch.object(flow, 'cmd_delivery', side_effect=lambda args: flow.delivery_result(args, [], 'final')), \
                redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            self.assertEqual(flow.main(), 3)
        result = json.loads(default_result.read_text())
        self.assertFalse(result['task_complete'])
        self.assertFalse(result['formal_release'])
        self.assertEqual(json.loads(self.state_path.read_text())['gate']['final_ready'], 'ERROR')

    def test_cli_failure_is_nonzero_and_saves_false_result(self):
        self.state()
        cmd = [sys.executable, '-B', str(SCRIPTS / '编制流程门禁.py'), '--page-dir', str(self.pages),
            '--state', str(self.state_path), 'delivery', '--document', str(self.document)]
        for key in ('static-report', 'word-report', 'score-report', 'review'):
            cmd.extend(['--' + key, str(self.base / ('missing-' + key + '.json'))])
        run = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(run.returncode, 3)
        result = json.loads((self.base / '流程交付结果.json').read_text())
        self.assertFalse(result['task_complete'])

    def test_execution_exception_invalidates_previous_preflight_pass(self):
        self.state()
        with patch.object(sys, 'argv', ['gate', '--page-dir', str(self.pages), '--state', str(self.state_path), 'preflight']), \
                patch.object(flow, 'cmd_preflight', side_effect=RuntimeError('PDF decompression failure')), \
                redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            self.assertEqual(flow.main(), 3)
        self.assertEqual(json.loads(self.state_path.read_text())['gate']['preflight'], 'ERROR')


class FormatAuditTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        spec = importlib.util.spec_from_file_location(
            'format_audit_under_test', SCRIPTS / 'Word正式交付硬审（不含页数）.py')
        cls.audit = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = cls.audit
        spec.loader.exec_module(cls.audit)

    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.document = Path(directory.name) / 'format.docx'

    def check(self, doc, name, rules=None):
        doc.save(self.document)
        return next(result for result in self.audit.portable_checks(self.document, rules or {})
                    if result.name == name)

    def test_font_command_and_audit_resolve_prompt_size_with_per_scope_fallback(self):
        rules_path = self.document.with_suffix('.json')
        baseline = {'font_rules': {'body': {'name_east_asia': '宋体', 'size_pt': 14},
                                   'heading1': {'name_east_asia': '宋体', 'size_pt': '三号'}}}
        rules_path.write_text(json.dumps(baseline | {'prompt_font_sizes': {'body': '小四号'}}), encoding='utf-8')
        command = [sys.executable, str(SCRIPTS / 'Word正式交付硬审（不含页数）.py'),
                   '--rules', str(rules_path), '--print-font-rules']
        output = json.loads(subprocess.check_output(command, text=True, encoding='utf-8'))
        rules = self.audit.load_rules(rules_path)
        self.assertEqual(output['font_rules'], rules['font_rules'])
        self.assertEqual(output['font_rules']['body'], {'name_east_asia': '宋体', 'size_pt': 12})
        self.assertEqual(output['font_rules']['heading1']['size_pt'], 16)
        self.assertEqual(output['font_body_scope'], 'text')
        self.assertEqual(self.audit.effective_font_settings(baseline)['font_rules']['body']['size_pt'], 14)
        self.assertEqual(self.audit.effective_font_settings({'font_rules': {'body': {'size_pt': 14}},
                         'prompt_font_sizes': {'heading3': 15}})['font_rules']['heading3']['size_pt'], 15)
        with self.assertRaises(self.audit.AuditError):
            self.audit.effective_font_settings(baseline | {'prompt_font_sizes': {'body': '略小'}})

    def test_body_size_scope_excludes_tables_and_captions_and_reports_mismatch(self):
        def paragraph(size, table=False, style='Normal'):
            return SimpleNamespace(Style=style, Range=SimpleNamespace(
                size=size, Information=lambda flag: table))

        paragraphs = [paragraph(12), paragraph(9, table=True), paragraph(10, style='Caption')]
        records = [{'index': i, 'page': 2, 'text': text} for i, text in enumerate(
            ['普通正文', '表格内容', '图1 施工示意'], 1)]
        document = SimpleNamespace(Paragraphs=lambda i: paragraphs[i - 1])
        rules = {'body_start_page': 2, 'font_body_scope': 'text', 'font_rules': {'body': {'size_pt': 12}}}
        with patch.object(self.audit, 'inspect_range_font', side_effect=lambda rng: {'size_pt': rng.size}):
            result = self.audit.check_fonts(document, records, rules)
            self.assertEqual((result.status, result.details['checked_paragraphs']), ('PASS', 1))
            paragraphs[0].Range.size = 14
            result = self.audit.check_fonts(document, records, rules)
            self.assertEqual(result.status, 'FAIL')
            self.assertEqual(result.details[0]['mismatch']['size_pt'], {'expected': 12, 'actual': 14})
            paragraphs[0].Range.size = 12
            rules['font_rules']['table'] = {'size_pt': 9}
            self.assertEqual(self.audit.check_fonts(document, records, rules).details['checked_paragraphs'], 2)

    def test_body_citations_verify_inline_sources_and_deduplicate(self):
        from docx import Document

        doc = Document()
        doc.add_paragraph('封面引用《施工总说明》图A01。')
        doc.add_heading('第一章 施工部署', 1)
        doc.add_heading('第一节 编制依据', 2)
        doc.add_paragraph('列入《施工总说明》图A01是项目设计依据。')
        doc.add_heading('第二节 施工方法', 2)
        doc.add_paragraph('依据《施工总说明》图A01，东侧保留原有通道。')
        doc.add_paragraph('依据《施工总说明》图A02，西侧设置临时围挡。')
        doc.save(self.document)
        item = {'subject': '东侧保通', 'source_name': '施工总说明', 'locator': '图A01',
                'source_ref': '项目施工总说明.pdf 图A01', 'body_quote': '依据《施工总说明》图A01，东侧保留原有通道。'}
        second = item | {'subject': '西侧围挡', 'locator': '图A02',
                         'body_quote': '依据《施工总说明》图A02，西侧设置临时围挡。'}
        review = {'reviewed': True, 'conclusion': '复核了这两项部署', 'items': [item, item, second]}
        self.assertEqual(self.audit.body_citation_review(self.document, review), (2, []))
        for quote in ('封面引用《施工总说明》图A01。', '列入《施工总说明》图A01是项目设计依据。',
                      '依据《施工总说明》图A99，东侧保留原有通道。', '《施工总说明》图A01'):
            count, errors = self.audit.body_citation_review(
                self.document, review | {'items': [item | {'body_quote': quote}]})
            self.assertEqual(count, 0)
            self.assertTrue(errors)
        self.assertTrue(self.audit.body_citation_review(self.document, None)[1])
        self.assertEqual([flow.body_citation_target(p) for p in (49, 99, 399)], [15, 25, 100])

    def test_tender_numbered_source_lists_and_split_table_citations_do_not_count(self):
        from docx import Document

        doc = Document()
        doc.add_heading('2 施工部署', 1)
        doc.add_heading('2.2 编制依据', 2)
        source_list = '依据《设计说明》图A01，列入本项目的采用文件。'
        doc.add_paragraph(source_list)
        doc.add_heading('2.3 组织安排', 2)
        cell = doc.add_table(1, 1).cell(0, 0)
        cell.paragraphs[0].text = '东侧保留原有通道。'
        cell.add_paragraph('依据《设计说明》图A01。')
        active = '依据《设计说明》图A01，西侧设置临时围挡。'
        cell.add_paragraph(active)
        doc.save(self.document)
        entry = {'subject': '围挡或保通', 'source_name': '设计说明', 'locator': '图A01', 'source_ref': 'A01原图'}
        record = {'reviewed': True, 'conclusion': '复核', 'items': [
            entry | {'body_quote': text} for text in
            (source_list, '东侧保留原有通道。依据《设计说明》图A01。', active)]}
        tender = {'heading_patterns': {'1': r'^(\d+)\s+', '2': r'^\d+\.(\d+)\s+', '3': r'^\((\d+)\)\s+'}}
        for rules in (None, tender):
            count, errors = self.audit.body_citation_review(self.document, record, rules)
            self.assertEqual((count, len(errors)), (1, 2))

    def test_body_title_alignment_uses_style_inheritance_and_direct_formatting(self):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH as Align

        doc = Document()
        doc.add_paragraph('目录', style='Heading 1').alignment = Align.LEFT
        doc.styles['Heading 1'].paragraph_format.alignment = Align.CENTER
        doc.styles['Heading 2'].base_style = doc.styles['Heading 1']
        doc.add_heading('第一章 施工部署', 1)
        section = doc.add_heading('第一节 作业组织', 2)
        subsection = doc.add_heading('一、分区实施', 3)
        name = '正文标题对齐_静态'
        self.assertEqual(self.check(doc, name).status, 'PASS')

        section.alignment = Align.LEFT
        subsection.alignment = Align.CENTER
        result = self.check(doc, name)
        self.assertEqual(result.status, 'FAIL')
        self.assertEqual([(x['level'], x['expected'], x['actual']) for x in result.details],
                         [(2, 'center', 'left'), (3, 'left', 'center')])

    def test_chart_minimum_counts_bound_body_objects_once(self):
        from docx import Document
        from docx.shared import Inches
        from PIL import Image

        def picture(color):
            stream = io.BytesIO()
            Image.new('RGB', (10, 12), color).save(stream, format='PNG')
            stream.seek(0)
            return stream

        def data_table(number, contents):
            doc.add_paragraph(f'表{number} 检查记录')
            doc.add_table(1, 2).cell(0, 0).text = contents

        doc = Document()
        doc.add_picture(picture('yellow'), width=Inches(.5))
        doc.add_paragraph('图1 封面图片')
        doc.add_heading('第一章 施工部署', 1)
        for number, color in enumerate(('red', 'blue', 'green', 'red'), 1):
            doc.add_picture(picture(color), width=Inches(.5))
            doc.add_paragraph(f'图{number} 工序示意')
        doc.add_table(1, 1).cell(0, 0).text = '【实景图占位：进场道路】'
        doc.add_paragraph('图5 进场道路占位')
        doc.add_table(1, 1).cell(0, 0).text = '纯排版框'
        for i in range(1, 12):
            data_table(i, f'分区{i}的检查安排')
        data_table(12, '分区1的检查安排')
        data_table(13, '')
        doc.add_paragraph('表99 无对应对象的题注')
        rules = {'table_layout': {'tables': [{'index': 2, 'kind': 'layout', 'reason': '排版框'}]}}
        name = self.audit.BODY_CHART_COUNT_CHECK
        result = self.check(doc, name, rules)
        self.assertEqual(result.status, 'WARN')
        self.assertEqual(result.details, {'figures': 3, 'tables': 11, 'total': 14, 'minimum': 15, 'shortfall': 1})

        data_table(14, '分区12的检查安排')
        result = self.check(doc, name, rules)
        self.assertEqual((result.status, result.details['total'], result.details['shortfall']), ('PASS', 15, 0))


class PdfFallbackTests(unittest.TestCase):
    def test_partial_primary_extraction_does_not_duplicate_pages(self):
        import pymupdf
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'cad-export.pdf'
            doc = pymupdf.open()
            doc.new_page().insert_text((72, 72), 'page one')
            doc.new_page().insert_text((72, 72), 'page two')
            doc.save(path)
            doc.close()
            digest = sc.file_hash(path)
            reader = SimpleNamespace(pages=[SimpleNamespace(extract_text=lambda: 'primary prefix'),
                SimpleNamespace(extract_text=lambda: (_ for _ in ()).throw(ValueError('Limit reached while decompressing')))])
            with patch('pypdf.PdfReader', return_value=reader), redirect_stderr(io.StringIO()):
                units = list(sc.native_units(path))
            self.assertEqual(len(units), 2)
            self.assertIn('page one', units[0][1])
            self.assertNotIn('primary prefix', ''.join(x[1] for x in units))
            self.assertTrue(all(x[2] == ['visual'] for x in units))
            self.assertEqual(sc.file_hash(path), digest)

    def test_both_readers_fail_loudly_instead_of_returning_no_units(self):
        with patch('pypdf.PdfReader', side_effect=ValueError('primary failed')), \
                patch('pymupdf.open', side_effect=ValueError('fallback failed')):
            with self.assertRaisesRegex(ValueError, 'PDF读取失败'):
                list(sc.native_units(Path('unreadable.pdf')))


if __name__ == '__main__':
    unittest.main(verbosity=2)
