"""Regressions for the r23 planning, Word page evidence, and final release gate."""
from argparse import Namespace
from contextlib import ExitStack, redirect_stdout
from copy import deepcopy
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from test_hard_checks import common, flow, write

from docx import Document
import project_page_rules as policy


SCRIPT = next((Path(__file__).resolve().parents[1] / "scripts").glob("*页数硬审*.py"))
spec = importlib.util.spec_from_file_location("review_point_page_script", SCRIPT)
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)
POINT_NAMES = [f"评审项目{n}" for n in range(1, 11)]
HEADINGS = [f"第{n}章 {name}" for n, name in enumerate(POINT_NAMES, 1)]


class WordLayout:
    """Test double for Word's main-story paragraph positions and physical pages."""

    def __init__(self, total):
        self.total = total
        self.Paragraphs = []
        self.Shapes = []
        self.positions = {}
        self.full_text = ""
        self.repaginated = False

    def paragraph(self, text, page, level=10, end_page=None, number="", table=False):
        start = len(self.full_text)
        self.full_text += text
        end = len(self.full_text)
        self.positions.update({i: page for i in range(start, end)})
        if end_page is not None:
            # Paragraphs terminate with a CR and may end in a page/section break before it.
            last_visible = len(text.rstrip(audit.LAYOUT)) - 1
            self.positions[start + last_visible] = end_page
        rng = self.Range(start, end)
        rng.ListFormat = SimpleNamespace(ListString=number)
        rng.Information = lambda _: table
        self.Paragraphs.append(SimpleNamespace(Range=rng, OutlineLevel=level))

    @property
    def Content(self):
        return SimpleNamespace(End=len(self.full_text))

    def Range(self, start, end):
        return SimpleNamespace(Start=start, End=end, Text=self.full_text[start:end],
                               Information=lambda _: self.positions[start])

    def GoTo(self, *, What, Which, Count):
        return SimpleNamespace(Start=next(pos for pos, page in self.positions.items() if page == Count))

    def Repaginate(self):
        self.repaginated = True

    def ComputeStatistics(self, statistic):
        if statistic != 2 or not self.repaginated:
            raise AssertionError("expected Word actual repagination first")
        return self.total


class ReviewPointPages(unittest.TestCase):

    def setUp(self):
        work = tempfile.TemporaryDirectory()
        self.addCleanup(work.cleanup)
        self.tmp = Path(work.name)
        self.project_script = self.tmp / SCRIPT.name
        source = SCRIPT.read_text("utf-8").replace("REVIEW_POINT_HEADINGS = []",
                                                   "REVIEW_POINT_HEADINGS = " + repr(HEADINGS))
        self.project_script.write_text(source, "utf-8")
        self.rule = flow.page_range_snapshot(self.project_script, flow._literal_constants(self.project_script))
        doc = Document()
        doc.add_paragraph("测试稿封面")
        for heading in HEADINGS:
            doc.add_heading(heading, 1)
            doc.add_paragraph("用于绑定文件哈希的测试正文及图表说明。")
        self.document = self.tmp / "design.docx"
        doc.save(self.document)

    def make_layout(self, counts, *, front=1, total=None, shared_after=(), headings=HEADINGS):
        point_total = sum(counts) + front - len(shared_after)
        word = WordLayout(point_total if total is None else total)
        word.paragraph("封面\r", 1)
        for n in range(2, front + 1):
            word.paragraph("目录\r", n)
        page = front + 1
        for i, (heading, count) in enumerate(zip(headings, counts), 1):
            word.paragraph(heading + "\r", page, level=1)
            word.paragraph("实际施工技术和图表承诺\f\r", page, end_page=page + count - 1)
            page += count - int(i in shared_after)
        return word

    def measurement(self, layout, *, stage=None, front=1):
        with patch.multiple(audit, REVIEW_POINT_HEADINGS=HEADINGS, FRONT_MATTER_PAGES=front), \
                patch.object(audit, "__file__", str(self.project_script)):
            return audit.build_report(layout, "16.0", self.document, flow.sha256(self.document),
                                      audit.rule_snapshot(), stage=stage)

    def evidence(self, report, headings=None):
        return policy.evidence_errors(report, self.rule, flow.sha256(self.project_script), headings)

    def test_100_and_shorter_are_valid_excluding_the_cover(self):
        for counts in ([10] * 10, [8] * 10):
            with self.subTest(counts=counts):
                report = self.measurement(self.make_layout(counts))
                self.assertEqual(report["status"], "PASS")
                self.assertEqual(report["actual"]["non_cover_pages"], sum(counts))
                self.assertFalse(self.evidence(report))

    def test_11_pages_fails_even_when_whole_file_is_100(self):
        report = self.measurement(self.make_layout([11, 8] + [10] * 8))
        self.assertEqual(report["actual"]["word_total_pages"], 100)
        self.assertEqual(report["actual"]["review_points"][0]["pages"], 11)
        self.assertEqual(report["status"], "FAIL")
        self.assertTrue(self.evidence(report))

    def test_independent_toc_consumes_the_whole_document_allowance(self):
        report = self.measurement(self.make_layout([10] * 10, front=2), front=2)
        self.assertEqual(report["actual"]["body_pages"], 100)
        self.assertEqual(report["actual"]["non_cover_pages"], 101)
        self.assertEqual(report["status"], "FAIL")

    def test_same_page_counts_for_both_points_and_once_in_total(self):
        report = self.measurement(self.make_layout([10] * 10, shared_after=(1,)))
        first, second = report["actual"]["review_points"][:2]
        self.assertEqual(first["end_page"], second["start_page"])
        self.assertEqual([first["pages"], second["pages"]], [10, 10])
        self.assertEqual(report["actual"]["non_cover_pages"], 99)
        self.assertFalse(self.evidence(report))

    def test_split_table_cell_extending_beyond_others_is_counted(self):
        word = WordLayout(12)
        word.paragraph("封面\r", 1)
        word.paragraph(HEADINGS[0] + "\r", 2, level=1)
        word.paragraph("跨页表长单元格末行\x07", 2, end_page=12, table=True)
        word.paragraph("邻列短单元格\x07", 2, end_page=2, table=True)
        report = self.measurement(word, stage=40)
        self.assertEqual(report["actual"]["review_points"][0]["pages"], 11)
        self.assertEqual(report["point_limit_status"], "FAIL")

    def test_missing_duplicate_and_wrong_order_fail_in_frozen_manuscript(self):
        for titles in (HEADINGS[:-1], [HEADINGS[0], *HEADINGS[:-1]], [*HEADINGS[1:], HEADINGS[0]]):
            with self.subTest(titles=titles):
                report = self.measurement(self.make_layout([8] * len(titles), headings=titles))
                self.assertEqual(report["status"], "FAIL")
                self.assertTrue(self.evidence(report))

    def test_stage_can_measure_a_prefix_and_checks_the_expected_prefix(self):
        report = self.measurement(self.make_layout([9, 9, 9, 9]), stage=40)
        self.assertEqual((report["status"], report["point_limit_status"]), ("MEASURED", "PASS"))
        self.assertFalse(self.evidence(report, HEADINGS[:4]))
        self.assertTrue(self.evidence(report, HEADINGS[:5]))
        report["actual"]["review_points"][1]["pages"] = 2
        self.assertTrue(self.evidence(report, HEADINGS[:4]))

    def test_page_rule_is_discovered_and_used_in_planning_and_format(self):
        script, values = flow.discover_page_script(self.tmp)
        self.assertEqual(script, self.project_script)
        self.assertEqual(policy.body_bounds(flow.current_rule(script, values)), (1, 100))
        rules = common.load_rules(None, page_dir=self.tmp)
        self.assertIs(rules["review_point_page_audit"], True)
        self.assertEqual(rules["front_matter"]["cover_page_range"], [1, 1])
        items = [{"id": str(i), "name": name} for i, name in enumerate(POINT_NAMES, 1)]
        plan = {"requirements": [], "modules": [{"id": title, "chapter": title, "planned_pages": 9}
                                                      for title in HEADINGS],
                "chapter_plan": [{"chapter": title, "page_range": [7, 10]} for title in HEADINGS]}
        state = {"page_rule": self.rule, "plan": plan}
        self.assertFalse(policy.plan_errors(state, items))
        plan["modules"][0]["planned_pages"] = 11
        self.assertTrue(policy.plan_errors(state, items))

    def test_preflight_accepts_a_realistic_90_page_budget(self):
        state_path = self.tmp / "state.json"
        plan = {"planned_total_body_pages": 90, "figure_table_pages_estimate": 8, "reserve_pages": 0,
                "planning_notes": "按各评审点的工程约束、图表、篇幅和具体施工内容分配预算",
                "chapter_plan": [{"chapter": title, "basis": "按实际评分义务",
                                  "page_range": [8, 10]} for title in HEADINGS],
                "checkpoint_expected_page_ranges": {"40": [34, 38], "75": [68, 76]}}
        state = {"page_rule": self.rule, "plan": plan, "startup": {}, "gate": {}}
        write(state_path, state)
        with patch.object(flow, "PAGE_DIR", self.tmp), patch.object(flow, "startup_errors", return_value=[]), \
                patch.object(flow, "detailed_plan_errors", return_value=[]), redirect_stdout(io.StringIO()):
            self.assertEqual(flow.cmd_preflight(state_path), 0)
        self.assertEqual(flow.load_state(state_path)["gate"]["preflight"], "PASS")

    def test_stage_gate_rejects_an_over_limit_point_with_an_in_range_forecast(self):
        state_path = self.tmp / "state.json"
        score_path = self.tmp / "score.json"
        hard_path = self.tmp / "hard.json"
        report_path = self.tmp / "stage.json"
        write(score_path, {})
        write(hard_path, {})
        selected = [str(n) for n in range(4)]
        plan = {"reserve_pages": 0, "checkpoint_module_ids": {"40": selected},
                "checkpoint_expected_page_ranges": {"40": [35, 37]},
                "modules": [{"id": str(n), "chapter": title, "planned_pages": 9}
                            for n, title in enumerate(HEADINGS)]}
        seed = {"page_rule": self.rule, "plan": plan, "gate": {"preflight": "PASS"},
                "module_registry": {mid: {"source_sha256": mid} for mid in selected}}
        for counts, expected in (([9, 9, 9, 9], 0), ([11, 7, 9, 9], 2)):
            write(state_path, seed)
            write(report_path, self.measurement(self.make_layout(counts), stage=40))
            with ExitStack() as isolated:
                isolated.enter_context(patch.object(flow, "PAGE_DIR", self.tmp))
                isolated.enter_context(patch.object(flow, "load_format_audit", return_value=common))
                for name in ("startup_errors", "production_errors", "checkpoint_modules_errors", "stage_score_errors"):
                    isolated.enter_context(patch.object(flow, name, return_value=[]))
                isolated.enter_context(patch.object(flow.hard_checks, "audit", return_value=[]))
                isolated.enter_context(patch.object(flow.bid_requirements, "evidence_errors", return_value=[]))
                isolated.enter_context(redirect_stdout(io.StringIO()))
                rc = flow.cmd_checkpoint(state_path, 40, report_path, None, "", False, "", self.document,
                                         score_path, hard_path)
            self.assertEqual(rc, expected)
            state = flow.load_state(state_path)
            self.assertEqual(state["checkpoints"]["40"]["projected_final_body_pages"], 90)

    def release(self, page_report):
        # Exercise the normal delivery decision with the unrelated r22 content validators isolated.
        paths = {name: self.tmp / (name + ".json") for name in
                 ("state", "review", "score_report", "warning_review", "static_report",
                  "word_report", "page_report", "result")}
        args = Namespace(mode="full", document=self.document, **paths)
        digest = flow.sha256(self.document)
        response = {"keywords": []}
        rules = {"review_point_page_audit": True, "project_response": response,
                 "photo_delivery_stage": "final", "body_start_page": 2, "visual_policy": {}}
        shared = {"document_sha256": digest, "audit_version": common.AUDIT_VERSION,
                  "rules_sha256": common.rules_digest(rules), "photo_delivery_stage": "final",
                  "delivery_status": "PASS", "checks": []}
        word = dict(shared, layout_engine="Microsoft Word", rules=rules,
                    checks=[{"name": n, "status": "PASS"} for n in [common.WORD_HEADING_ALIGNMENT_CHECK,
                                                                  common.FONT_CHECK]])
        write(paths["state"], {"gate": {"preflight": "PASS"}, "startup": {"visual_policy": {}},
                               "plan": {"project_response": response,
                                        "bid_requirements": {"format": {"deliverable": "docx", "fonts": {}}}}})
        write(paths["review"], {})
        write(paths["static_report"], dict(shared, audit_type="embedded_docx_static"))
        write(paths["word_report"], word)
        write(paths["warning_review"], {"document_sha256": digest, "dispositions": []})
        write(paths["score_report"], {"document_sha256": digest, "status": "PASS", "phase": "final"})
        write(paths["page_report"], page_report)
        with ExitStack() as isolated:
            for target, names in [
                (flow.hard_checks, ("audit",)), (flow.bid_requirements, ("plan_errors", "evidence_errors")),
                (flow.submission_identity, ("errors",)),
                (flow, ("startup_errors", "final_response_errors", "word_pagination_evidence_errors")),
                (common, ("section_heading_review_errors", "portable_checks")),
            ]:
                for name in names:
                    isolated.enter_context(patch.object(target, name, return_value=[]))
            for name, value in {"PAGE_DIR": self.tmp, "cmd_final_ready": 0, "load_format_audit": common,
                                "validate_content_review": ([], {}), "_manuscript": ("正文", []),
                                "body_citation_target": 0}.items():
                if name == "PAGE_DIR":
                    isolated.enter_context(patch.object(flow, name, value))
                else:
                    isolated.enter_context(patch.object(flow, name, return_value=value))
            isolated.enter_context(patch.object(common, "body_citation_review", return_value=(1, [])))
            isolated.enter_context(redirect_stdout(io.StringIO()))
            rc = flow.cmd_delivery(args)
        return rc, json.loads(paths["result"].read_text("utf-8"))

    def test_formal_gate_rejects_failed_and_incomplete_word_evidence(self):
        no_excess = self.measurement(self.make_layout([9] * 10))
        rc, decision = self.release(no_excess)
        self.assertEqual(rc, 0, decision)
        self.assertIs(decision["formal_release"], True)
        over = self.measurement(self.make_layout([11, 8] + [10] * 8))
        # A fabricated summary status still cannot override the individual physical page span.
        over.update(status="PASS", point_limit_status="PASS", failures=[])
        over["actual"]["review_points"][0]["status"] = "PASS"
        missing = deepcopy(no_excess)
        missing["actual"]["review_points"].pop()
        for report in (over, missing):
            rc, decision = self.release(report)
            self.assertEqual(rc, 2, decision)
            self.assertIs(decision["formal_release"], False)

    @unittest.skipUnless(sys.platform == "win32", "requires real Windows Microsoft Word")
    def test_windows_word_cli_on_a_natively_paginated_fixture(self):
        from subprocess import run

        for counts in ([10] * 10, [11, 8] + [10] * 8):
            doc = Document()
            doc.add_paragraph("封面")
            for heading, count in zip(HEADINGS, counts):
                doc.add_page_break()
                doc.add_heading(heading, 1)
                doc.add_paragraph("评审点正文")
                for _ in range(count - 1):
                    doc.add_page_break()
                    doc.add_paragraph("评审点续页技术正文")
            doc.save(self.document)
            report_path = self.tmp / "native-pages.json"
            result = run([sys.executable, str(self.project_script), str(self.document),
                          "--report", str(report_path)], capture_output=True, text=True, encoding="utf-8",
                         env={**os.environ, "PYTHONIOENCODING": "utf-8"}, timeout=300)
            self.assertEqual(result.returncode, 0 if counts[0] == 10 else 2, result.stderr)
            report = json.loads(report_path.read_text("utf-8"))
            self.assertEqual([row["pages"] for row in report["actual"]["review_points"]], counts)


if __name__ == "__main__":
    unittest.main()
