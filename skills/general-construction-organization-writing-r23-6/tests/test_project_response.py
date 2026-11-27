"""r20 regressions using Word document objects and simulated chapter positions."""
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from docx import Document
from docx.shared import Cm

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
import project_response as response
import visual_policy

spec = importlib.util.spec_from_file_location("r20_word_format", next((ROOT/"scripts").glob("Word正式*.py")))
audit = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = audit
spec.loader.exec_module(audit)


class ProjectResponseTests(unittest.TestCase):
    def setUp(self):
        self.terms = [f"工序{i:02}" for i in range(1, 36)]
        self.plan = {"keywords": [{"term": t} for t in self.terms],
                     "headings": {"2": self.terms[:15], "3": self.terms[15:]}}

    def test_levels_are_independent_and_aliases_count_once(self):
        self.plan["keywords"][0]["forms"] = ["独立变体"]
        heads = {"2": ["独立变体", "独立变体", *self.terms[1:14]], "3": self.terms}
        errors, rows = response.title_coverage(self.plan, heads)
        self.assertEqual((len(errors), rows["2"]["distinct"], rows["3"]["distinct"]), (1, 14, 35))

    def test_increased_title_count_alone_cannot_pass(self):
        head = self.plan["headings"]
        self.plan["expanded"] = {"3":{"prompt":"尽可能增加三级标题", "usual_headings":20,
            "heading_target":26, "usual_keywords":20, "keyword_target":25, "basis":"六个新控制点"}}
        head["3"] += head["3"][:6]
        self.assertTrue(response.title_coverage(self.plan, head)[0])
        head["3"][-6:] = self.terms[:6]
        self.assertFalse(response.title_coverage(self.plan, head)[0])
        self.plan["expanded"]["3"]["keyword_target"] = 20
        with self.assertRaises(ValueError):
            response.title_coverage(self.plan, head)

    def test_document_coverage_excludes_body_and_front_toc(self):
        doc = Document()
        doc.add_paragraph("目录")
        doc.add_paragraph(" ".join(self.terms), "TOC Heading")
        doc.add_heading("第一章 技术方案", 1)
        doc.add_heading("第一节 "+"与".join(self.terms[:15]), 2)
        doc.add_heading("一、"+"与".join(self.terms[15:34]), 3)
        doc.add_paragraph(self.terms[-1])
        policy = visual_policy.visual_policy_template()
        policy.update(real_photos={"decision":"schematic","reason":"test","replacement_figures":["现场保通图"]},
                      additional_figures={"names":["现场保通图"],"reason":"test"})
        figures = [{"name":n,"hard":True,"color_required":True, "location":"本节",
                    "purpose":"test","context_anchor":"test"} for n in visual_policy.planned_figure_names(policy)]
        rules = dict(audit.DEFAULT_RULES, project_response=self.plan, visual_policy=policy,
                     required_real_photo_slots=[], required_figures=figures)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp)/"sample.docx"
            doc.save(path)
            result = next(c for c in audit.portable_checks(path,rules) if c.name=="标题关键词分层硬审")
        self.assertEqual(result.status, "FAIL")
        self.assertEqual([result.details[k]["distinct"] for k in ("2","3")], [15,19])

    def test_priority_shares_each_use_the_same_denominator(self):
        priorities = [{"chapters":[n],"basis":"原件"} for n in "甲乙丙"]
        pages = dict(甲=10,乙=6.5,丙=5)
        self.assertFalse(response.priority_errors(priorities, pages, 50)[0])
        pages.update(甲=16,乙=6.49)
        self.assertEqual(len(response.priority_errors(priorities, pages, 50)[0]), 1)
        priorities[2]["chapters"] = ["甲"]
        self.assertTrue(response.priority_errors(priorities,pages,50)[0])

    def test_word_chapter_boundary_inside_same_page(self):
        priorities = [{"chapters":[n],"basis":"评分与工况"} for n in "甲乙丙"]
        # Two front pages excluded; chapters start at body positions 0, 1.25 and 1.75.
        doc = MagicMock()
        doc.Range.side_effect = lambda start,end: start
        headings = [dict(level=1, start=i, text=n) for i,n in enumerate("甲乙丙")]
        rules = {"body_start_page":3,"project_response":{"priorities":priorities}}
        with patch.object(audit,"_page_body_box",return_value={"body_top_pt":50,"usable_height_pt":100}), \
             patch.object(audit,"page_of_range",side_effect=[3,4,4]), \
             patch.object(audit,"_range_vertical_y",side_effect=[50,75,125]):
            result = audit.measured_priority_shares(doc,headings,rules,5)
        self.assertEqual(result.status,"PASS")
        self.assertEqual([d["pages"] for d in result.details],[1.25,.5,1.25])

    def test_source_named_figure_and_table(self):
        image = next((ROOT/"assets/examples").glob("*.png"))
        name = "施工总平面布置图"
        items = [{"kind":"图","name":name}, {"kind":"表","name":"劳动力计划表"}]
        doc = Document()
        doc.add_heading("第一章 "+name,1)
        doc.add_picture(str(image), width=Cm(2))
        fig = doc.add_paragraph("图1 "+name)
        doc.add_heading("第二章 劳动力计划表",1)
        doc.add_paragraph("表1 劳动力计划表")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp)/"sample.docx"
            def errors():
                doc.save(path)
                return response.named_errors(audit._docx_blocks(path),items,audit._caption_match)
            self.assertEqual(len(errors()),1)
            doc.add_table(1,1).cell(0,0).text = "班组与阶段投入"
            self.assertEqual(errors(),[])
            fig.text = "图1 施工功能分区示意图"
            self.assertEqual(len(errors()),1)
            fig.text = "图1 "+name
            doc.paragraphs[0].text = "第一章 "+name+"控制措施"
            self.assertEqual(len(errors()),1)

    def test_required_name_replaces_only_corresponding_default_figure(self):
        name = "施工总平面布置图"
        policy = {"named_exhibits":[{"kind":"图","name":name,"source_ref":"第42页",
                  "quote":"应提交施工总平面布置图","replaces":visual_policy.BASE_FIGURES[0]}]}
        self.assertEqual(visual_policy.planned_figure_names(policy), [name,*visual_policy.BASE_FIGURES[1:]])

    def test_font_delta_is_allocated_to_real_modules(self):
        font = {"prompt":"正文小四","size_pt":12,"word_version":"fixture","layout_basis":"同内容混排",
                "sample_four":{"sha256":"a","word_pages":4},"sample_selected":{"sha256":"b","word_pages":3.5},
                "four_chars":35100,"front_chars":100,"estimated_chars":40100,
                "allocation":[{"id":"M","delta":5000,"content":"深化施工窗口与分段验收"}]}
        fresh = lambda *_: True
        self.assertFalse(response.font_errors(font,[{"id":"M"}],None,fresh))
        font["allocation"][0]["delta"] = 4900
        self.assertTrue(response.font_errors(font,[{"id":"M"}],None,fresh))
