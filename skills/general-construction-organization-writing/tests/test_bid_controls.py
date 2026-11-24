"""Portable project-specific release gates. Values in these tests are synthetic."""
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
import tempfile
import sys
import unittest
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/"scripts"))
from docx import Document
import bid_requirements as bid
import schedule
import submission_identity as identity
from test_hard_checks import case, common, flow, schedule_model, hc


from bid_fixture import project_profile

class GateTests(unittest.TestCase):
    def setUp(self):
        tmp=tempfile.TemporaryDirectory(); self.addCleanup(tmp.cleanup)
        self.base=Path(tmp.name)
        self.tasks=schedule_model()["tasks"]
        self.p=project_profile({},"M","R",[],[],["甲班","乙班"])
        self.table={"headers":[["工种","第一阶段","第二阶段"]],"key_col":1,"required_cols":[2,3],
                    "resource_cols":[2,3],"resource_kind":"trade","task_ids":["A","B","C"],
                    "phases":{"第一阶段":[1,1],"第二阶段":[2,3]},
                    "rows":[{"key":trade,"cells":{str(i+2):{"phase":phase,"peak":["trade",trade,*days]}
                         for i,(phase,days) in enumerate({"第一阶段":[1,1],"第二阶段":[2,3]}.items())}}
                         for trade in ("安装工","普工")]}
        self.rows=[["工种","第一阶段","第二阶段"],["安装工","4","2"],["普工","6","3"]]

    def test_table_pass_and_cell_failures(self):
        self.assertFalse(bid.resource_scope_errors("劳动力计划表",self.table,self.p,self.tasks))
        self.assertEqual(bid.check_table("劳动力计划表",self.rows,self.table,self.p,self.tasks),[])
        for rows in ([["班组与阶段投入"]],self.rows[:2],[self.rows[0],["安装工","",""],self.rows[2]],
                     [self.rows[0],["安装工","2","4"],self.rows[2]]):
            self.assertTrue(bid.check_table("劳动力计划表",rows,self.table,self.p,self.tasks))

    def test_deleted_trade_and_bad_phase(self):
        a=deepcopy(self.table);a["rows"].pop()
        self.assertTrue(bid.resource_scope_errors("劳动力计划表",a,self.p,self.tasks))
        a=deepcopy(self.table);a["rows"][0]["cells"]["2"]["peak"][-1]=3
        self.assertTrue(bid.resource_scope_errors("劳动力计划表",a,self.p,self.tasks))

    def test_resource_quantity_pairing(self):
        self.assertTrue(bid.amount_at("D-8",2,["电钻型号D-8，共2台"],"台"))
        self.assertFalse(bid.amount_at("安装工",2,["安装工2台"],"人"))
        self.assertFalse(bid.amount_at("安装工",2,["安装工3人、普工2人"],"人"))
        self.assertTrue(bid.amount_at("安装工",2,["安装工 | 2人"],"人"))

    def test_no_double_allocation_of_unused_spare_window(self):
        t=deepcopy(self.tasks[0]);t.update(quantity=4,windows=[[1,8,12],[1,13,17]])
        u=dict(t,id="new",windows=[[1,13,17]])
        self.assertEqual(schedule.calculate({"tasks":[t,u]}),[])
        self.assertEqual(bid.peak([t,u],self.p,"trade","安装工",1,1),2)

    def test_resource_quantities_cannot_be_replaced_by_names(self):
        p=deepcopy(self.p);p["crews"]["甲班"]["trades"]={}
        self.assertTrue(bid.catalog_errors(p,self.tasks))

    def test_wrong_region_outdated_basis_and_source(self):
        rule={"source_id":"S","unit_id":"S:U000001","quote":"试件条款"}
        p=project_profile(rule,"M","R",[],[],[])
        state={"startup":{"sources":[{"id":"S","role":"tender"}],"visual_policy":{}},
               "plan":{"modules":[{"id":"M","primary_for":["R"]}],"requirements":[{"id":"R"}],"hard_checks":{"schedule":{"tasks":[]}}}}
        get=lambda source,unit:"试件条款"
        self.assertEqual(bid.validate_plan(state,p,get),[])
        for field,value in (("source",dict(rule,quote="其他条款")),("validity",dict(rule,effective_from="2025-01-01",effective_until="2026-01-01",status="有效"))):
            bad=deepcopy(p);bad["obligations"][0][field]=value
            self.assertTrue(bid.validate_plan(state,bad,get))
        p["obligations"][0].update(level="市",territory="合肥市")
        self.assertTrue(bid.validate_plan(state,p,get))

    def test_qualification_and_condition_must_reach_manuscript(self):
        self.p["obligations"][0].update(topics=["岗位资格"],qualification="相应岗位证",verification_time="入场前")
        loc={"L1":"试件操作标准 试件第1条 检查接口 班长 核对记录 复查接口后移交 场地受限 分时交接 逐段验收"}
        rec={r["id"]:{k:"L1" for k in r} for r in [*self.p["decisions"],*self.p["obligations"]]}
        self.assertTrue(any("qualification" in e for e in bid.module_evidence(self.p,"M",loc,rec)))
        loc["L1"]+=" 入场前核相应岗位证"
        self.assertFalse(bid.module_evidence(self.p,"M",loc,rec))

    def test_monochrome_figures_share_plan_and_check(self):
        model=schedule_model()
        policy={"mode":"monochrome"}
        old_chart=schedule.chart
        with patch.object(schedule,"chart",side_effect=lambda m,k: old_chart(m,k,policy)):
            s,r,w,d=case(self.base)
        s["startup"]["visual_policy"]["color_figures"]=policy
        self.assertEqual(hc.audit(s,d.parent,d,r,common,w),[])
        s["plan"]["hard_checks"]["schedule"]["tasks"][0]["name"]="另项活动"
        self.assertTrue(any("进度图" in e for e in hc.audit(s,d.parent,d,r,common,w)))
        for kind in ("gantt","network"):
            self.assertNotIn("橙色".encode(),schedule.chart(model,kind,policy))

    def test_plant_template_binds_quantity_and_specification(self):
        task=deepcopy(self.tasks[0])
        name=task["equipment"][0]
        self.p["equipment"][name]={"count":2,"spec":"D-8"}
        a={"headers":[["机械","数量","型号"]],"resource_kind":"plant","resource_cols":[2,3],
           "equipment_fields":{"2":"count","3":"spec"},"task_ids":[task["id"]],
           "rows":[{"key":name,"cells":{"2":{"resource":["equipment",name,"count"]},"3":{"resource":["equipment",name,"spec"]}}}]}
        self.assertEqual(bid.resource_scope_errors("机械设备表",a,self.p,[task]),[])
        a["rows"][0]["cells"]["2"]["resource"][-1]="spec"
        self.assertTrue(bid.resource_scope_errors("机械设备表",a,self.p,[task]))

    def test_workdays_and_production_bottleneck(self):
        from bid_fixture import fixture_capacity
        t=deepcopy(self.tasks[0])
        self.assertEqual(bid.expected_cell({"workdays":["trade","安装工",1,1,8]},self.p,[t]),"2")
        ref={"source_id":"S","unit_id":"U","quote":"台班记录"}
        t["capacity"]=fixture_capacity(ref,t["crews"])
        name=t["equipment"][0]
        self.p["equipment"][name]={"count":2}
        model=dict(operation="切割",unit=t["unit"],unit_rate=0.5,conversion="按现场净作业时段折减")
        t["capacity"]["plant"][name]=model
        lookup=lambda sid,uid:"台班记录"
        self.assertEqual(bid.capacity_errors(t,self.p,lookup),[])
        model["unit_rate"]=0.25
        self.assertTrue(bid.capacity_errors(t,self.p,lookup))

    def test_actual_annex_and_formal_publishing(self):
        import json
        import release_docx
        from annex_fixture import annex_case
        for wrong in (False,True):
            state,review,word,doc=annex_case(self.base,incorrect=wrong)
            self.assertEqual(bid.plan_errors(state,doc.parent),[])
            self.assertEqual(hc.audit(state,doc.parent,doc,review,common,word),[])
            failures=bid.evidence_errors(state,doc,review,common)
            if not wrong:
                self.assertEqual(failures,[])
                continue
            expected=["劳动力计划表 第2行第2列 应为「4」，实为「2」"]
            self.assertEqual(failures,expected)
            out=self.base/"formal"/"design.docx"
            self.assertNotEqual(release_docx.release(doc.parent,self.base/"reports",out),0)
            result=json.loads((doc.parent/"流程交付结果.json").read_text("utf-8"))
            self.assertEqual(result["failures"],expected)
            self.assertIs(result["formal_release"],False)
            self.assertFalse(out.exists())

    def test_tender_font_precedence(self):
        rules={"font_rules":{"body":{"size_pt":14}},"prompt_font_sizes":{"body":"五号"},"tender_font_sizes":{"body":"小四"}}
        self.assertEqual(common.effective_font_settings(rules)["font_rules"]["body"]["size_pt"],12)

    def test_identity_in_split_runs_footer_and_metadata(self):
        d=self.base/"正式.docx";doc=Document()
        p=doc.sections[0].footer.paragraphs[0];p.add_run("测试");p.add_run("企业")
        doc.core_properties.author="测试企业";doc.save(d)
        fmt={"anonymous":True,"identifiers":["测试企业"]}
        r={"sha256":sha256(d.read_bytes()).hexdigest()}
        errs=identity.errors(d,fmt,r)
        self.assertTrue(any("footer" in e for e in errs))
        self.assertTrue(any("core.xml" in e for e in errs))

    def test_submitted_pdf_requires_actual_pages_and_identity(self):
        import fitz
        p=self.base/"actual.pdf"
        with fitz.open() as pdf:
            pdf.new_page().insert_text((40,40),"BIDDER_INC");pdf.save(p)
        fmt={"anonymous":True,"identifiers":["BIDDER_INC"]}
        rec={"sha256":sha256(p.read_bytes()).hexdigest()}
        self.assertTrue(any("第1页" in e for e in identity.errors(p,fmt,rec)))
        fmt["identifiers"]=["OTHER_INC"]
        self.assertTrue(identity.errors(p,fmt,rec))
        rec["visual"]={loc:{"sha256":image,"finding":"核过本页图文","clean":True} for loc,_,image in identity.surfaces(p) if image}
        self.assertEqual(identity.errors(p,fmt,rec),[])

if __name__=="__main__":
    unittest.main()
