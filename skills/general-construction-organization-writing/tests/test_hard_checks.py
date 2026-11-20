"""泾县医院两条清单源自原表。其余数据是回归用的小样，不作为工程参数。"""
import copy
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from hashlib import sha256

from docx import Document
from docx.shared import Inches
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.oxml import parse_xml
from openpyxl import Workbook
from PIL import Image, ImageDraw

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/"scripts"))
import boq
import hard_checks as hc
import schedule
import importlib.util
spec=importlib.util.spec_from_file_location("test_r21_flow",ROOT/"scripts"/"编制流程门禁.py")
flow=importlib.util.module_from_spec(spec)
spec.loader.exec_module(flow)
common=flow.load_format_audit()
TABLE="分部分项工程量清单计价表"
TRADE="装饰工程"
FEATURE="所有病房卫生间内需考虑瓷砖与洗脸盆以及门框边缝隙打胶处理"
MEASURE=["病房卫生间边缝嵌胶","旧胶开裂或脱粘时","清除旧胶并清洁干燥后连续施胶","胶缝连续且与两侧粘结牢固","发现渗漏先排查管根再修复复验"]
SITE={"作业区":("院区西侧","西侧A区为施工区"),"出入口":("南侧边界","南门为施工入口"),
      "路线":("南门向西","进门沿箭头抵达西区"),"临设":("东侧矩形","临时库房位于东侧"),
      "空间关系":("南门和西区间","通道与库房分开")}


def write(path,data):
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2),"utf-8")


def fixture_book(path):
    b=Workbook()
    ws=b.active
    ws.title="装饰分册"
    ws.append([TABLE])
    ws.append(["工程名称：泾县医院总院装饰改造工程-"+TRADE])
    ws.append(["序号","项目编码","","项目名称","项目特征描述","","计量单位","","工程量"])
    for n,r in [(50,["1","011606003001","","拆除工程","部分科室施工期间会正常运转","","项","",1]),
                (417,["45","010518017001","","嵌缝、打胶","1．"+FEATURE,"","项","",1])]:
        for i,value in enumerate(r,1):
            ws.cell(n,i,value)
    b.save(path)


def schedule_model():
    first=dict(id="A",name="围挡",unit="米",quantity=8,rate=1,crews=["甲班"],equipment=["切割机A"],
               materials=["隔板"],ready={"隔板":1,"切割机A":1},windows=[[1,8,16]],start=1,finish=1,
               after=[],basis="试件工效记录，按班组每小时完成一米",module="M")
    second=dict(first,id="B",name="放线",crews=["乙班"],equipment=["仪器B"],materials=[],
                materials_basis="仪器放样不耗用主要材料",ready={"仪器B":1})
    third=dict(first,id="C",name="安装",start=3,finish=3,after=[["A",1],["B",0]],windows=[[3,8,16]])
    return {"tasks":[first,second,third]}


def png(svg):
    # Render the same vector for Word versions which display its PNG fallback.
    import fitz
    with fitz.open(stream=svg,filetype="svg") as pdf:
        return pdf[0].get_pixmap().tobytes("png")


def put_image(doc,raw,svg=None,idx=0):
    picture=doc.add_picture(io.BytesIO(raw),width=Inches(3.4))
    if svg:
        part=Part(PackURI(f"/word/media/schedule-{idx}.svg"),"image/svg+xml",svg,doc.part.package)
        rid=doc.part.relate_to(part,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        blip=picture._inline.xpath(".//a:blip")[0]
        blip.append(parse_xml(f'<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/></a:ext></a:extLst>'))


def site_image(flowchart=False):
    im=Image.new("RGB",(480,240),"white")
    d=ImageDraw.Draw(im)
    if flowchart:
        for i,s in enumerate(("START","REMOVE","CHECK")):
            x=10+i*150
            d.rectangle((x,100,x+96,145),fill="#d0e9e5",outline="navy")
            d.text((x+8,115),s,fill="black")
            if i<2:d.line((x+96,125,x+150,125),fill="navy",width=2)
    else:
        d.rectangle((15,10,465,225),outline="navy",width=3)
        d.rectangle((40,40,180,120),fill="#d0e9e5")
        d.text((55,70),"WORK ZONE A",fill="black")
        d.rectangle((330,50,430,130),fill="#efe0c1")
        d.text((338,85),"STORE",fill="black")
        d.line([(245,220),(245,165),(115,165),(115,125)],fill="green",width=5)
        d.polygon([(110,132),(120,132),(115,122)],fill="green")
        d.text((255,200),"SOUTH GATE",fill="black")
    buf=io.BytesIO()
    im.save(buf,format="PNG")
    return buf.getvalue()


def case(base,wrong=None):
    project=base/"project"
    reports=base/"reports"
    project.mkdir(parents=True,exist_ok=True)
    reports.mkdir(exist_ok=True)
    book=project/"清单摘录.xlsx"
    fixture_book(book)
    sha=flow.sha256(book)
    src={"id":"S","role":"boq","path":book.name,"sha256":sha}
    bills=[{"source":"S","sha256":sha,"tables":[{"sheet":"装饰分册","title":TABLE,"trade":TRADE,"title_cell":[1,1],"trade_cell":[2,1]}]}]
    model=schedule_model()
    names=["施工总平面布置图","施工进度横道图","施工进度网络图"]
    figs=[dict(name=n,kind=k,elements=sorted(hc.ELEMENTS[k]),basis="测试项目指定图种及现场情况",module="M") for n,k in zip(names,("site","gantt","network"))]
    m=dict(id="封胶",module="M",requirement="R",basis="病房卫生间饰面清单",completion="胶缝连续粘结、通水复验")
    spec=dict(bills=bills,figures=figs,measures=[m],schedule=model,gaps=[])
    state={"plan":{"hard_checks":spec,"modules":[{"id":"M","primary_for":["R"]}],"requirements":[{"id":"R"}]},
           "startup":{"sources":[src],"visual_policy":{"named_exhibits":[{"name":names[0],"kind":"图","replaces":"施工功能分区示意图"}]}}}
    doc=Document()
    rec=hc.pending_review()
    rec["bill_scan"]="按样稿各正文段、表行及题注反查所见引用"
    claims=[(417,dict(item="45",code="010518017001",name="嵌缝、打胶",quantity="1",unit="项",feature=FEATURE)),
            (50,dict(item="1",code=("011601001001" if wrong=="bill" else "011606003001"),name="拆除工程"))]
    for i,(row,fields) in enumerate(claims,1):
        quote=f"装饰清单第{fields['item']}项 {fields['name']}，编码{fields['code']}"
        if row==417:quote+="，1项，"+FEATURE
        doc.add_paragraph(quote)
        rec["bills"].append(dict(loc=f"P{i}",quote=quote,source="S",sha256=sha,title=TABLE,
                                 trade=TRADE,sheet="装饰分册",row=row,fields=fields))
    doc.add_paragraph("；".join(MEASURE))
    rec["measures"]["封胶"]={k:dict(loc="P3",quote=v,finding="核对本工序相应操作及完成条件") for k,v in zip(hc.ASPECTS,MEASURE)}
    if wrong=="measure":rec["measures"]["封胶"]={"结论":"已复核"}
    snapshots={}
    for i,f in enumerate(figs,1):
        k=f["kind"]
        svg=None if k=="site" else schedule.chart(model,k)
        raw=png(svg) if svg else site_image(wrong=="figure")
        put_image(doc,raw,svg,i)
        doc.paragraphs[-1].paragraph_format.keep_with_next=True
        caption=f"图{i} {f['name']}"
        doc.add_paragraph(caption)
        snapshots[caption]=[i]
        details=SITE if k=="site" else {"关键活动":("活动行","围挡、放线及安装"),"作业时段":("各行右端","安装位于第三日"),"依赖关系":("安装节点左侧","围挡和放线箭头同时汇入安装")}
        observed="flowchart" if wrong=="figure" and k=="site" else k
        rec["figures"][f["name"]]=dict(page=i,image_sha256=sha256(raw).hexdigest(),observed_kind=observed,
             elements={key:dict(status="verified",position=details[key][0],finding=details[key][1]) for key in f["elements"] if observed==k})
    if wrong=="plan":model["tasks"][1]["crews"]=["甲班"]
    for i,t in enumerate(model["tasks"]):
        doc.add_paragraph(t["name"]+"安排于"+schedule.interval(t))
        rec["schedule"][t["id"]]={"body":f"P{10+i}","table":f"T1R{i+1}"}
    tab=doc.add_table(rows=0,cols=5)
    for task in model["tasks"]:
        for cell,val in zip(tab.add_row().cells,schedule.table_row(task)):cell.text=val
    manuscript=project/"design.docx"
    doc.save(manuscript)
    docsha=flow.sha256(manuscript)
    review={"document_sha256":docsha,"hard_checks":rec}
    word={"layout_engine":"Microsoft Word","document_sha256":docsha,"figure_pages":snapshots,"fixture":"portable synthetic page map"}
    write(project/"篇幅规划状态.json",state)
    write(project/"review.json",review)
    write(reports/"delivery.json",word)
    return state,review,word,manuscript


class HardTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.base=Path(self.tmp.name)

    def audit(self,wrong=None):
        s,r,w,d=case(self.base,wrong)
        return hc.audit(s,d.parent,d,r,common,w)

    def test_four_positive_controls_in_final_word(self):
        self.assertEqual(self.audit(),[])

    def test_item_45_original_row(self):
        s,r,w,d=case(self.base)
        rows=hc.document_rows(common._docx_blocks(d))
        c=r["hard_checks"]["bills"][0]
        self.assertEqual((c["row"],c["fields"]["code"],c["fields"]["quantity"],c["fields"]["unit"]),(417,"010518017001","1","项"))
        self.assertEqual(boq.check({"P1":rows["P1"]},[c],boq.index(s,d.parent),{"S":c["sha256"]}),[])

    def test_site_plan_observed_spatial_elements(self):
        s,r,w,d=case(self.base)
        spec=s["plan"]["hard_checks"]
        self.assertEqual(hc.figure_errors(spec["figures"][:1],r["hard_checks"]["figures"],common._docx_blocks(d),common,w,spec["schedule"],d),[])

    def test_measure_module_evidence(self):
        s,r,w,d=case(self.base)
        rows=r["hard_checks"]["measures"]
        for e in rows["封胶"].values():
            e["loc"]="L1"
        self.assertEqual(hc.module_errors(s,{"id":"M"},"；".join(MEASURE),{"measures":rows},d.parent),[])
        self.assertTrue(hc.module_errors(s,{"id":"M"},"；".join(MEASURE),{"measures":{"封胶":{"结论":"已复核"}}},d.parent))

    def test_parallel_separate_crews(self):
        self.assertEqual(schedule.calculate(schedule_model()),[])

    def test_final_picture_replacement_invalidates_visual_review(self):
        s,r,w,d=case(self.base)
        good=copy.deepcopy(r["hard_checks"]["figures"])
        s,r,w,d=case(self.base,"figure")
        r["hard_checks"]["figures"]=good
        self.assertTrue(any("指纹" in e for e in hc.audit(s,d.parent,d,r,common,w)))

    def test_derived_measure_value_uses_original_inputs(self):
        source={"q":"工程量48米","r":"每班组每小时完成6米"}
        data={"operation":"divide","unit":"小时","basis":"工程量除以每小时产能","inputs":[
            {"source":"S","unit_id":i,"quote":source[i],"value":n} for i,n in (("q","48"),("r","6"))]}
        lookup=lambda sid,uid:source[uid]
        self.assertEqual(hc.calculated_value(data,lookup),8)
        data["inputs"][0]["value"]="49"
        self.assertIsNone(hc.calculated_value(data,lookup))

    def test_wrong_demolition_code(self):
        errors=self.audit("bill")
        self.assertTrue(any("P2" in e and "第50行" in e and "011606003001" in e and "011601001001" in e for e in errors),errors)

    def test_same_caption_flowchart(self):
        self.assertTrue(any("图件 图1" in e for e in self.audit("figure")))

    def test_review_without_body_quote(self):
        self.assertTrue(any("封胶/做法" in e for e in self.audit("measure")))

    def test_shared_crew(self):
        self.assertEqual(self.audit("plan"),["D1 甲班 被 A、B 重复占用"])

    def test_unlisted_table_and_caption(self):
        rows={"T1R5":"第45项","P9":"图1 第1项拆除"}
        self.assertEqual(len(boq.check(rows,[],{},{})),2)

    def test_wrong_version_trade_and_row(self):
        s,r,w,d=case(self.base)
        claim=r["hard_checks"]["bills"][0]
        for key,value in [("sha256","earlier"),("trade","安装工程"),("row",50),("row",418)]:
            with self.subTest(key=key,value=value):
                x=copy.deepcopy(r)
                x["hard_checks"]["bills"][0][key]=value
                self.assertTrue(hc.audit(s,d.parent,d,x,common,w))

    def test_material_time_rest_output_and_calculation(self):
        for key,value in [("ready",{"隔板":4,"切割机A":1}),("after",[["A",2]]),("quantity",9),("finish",4)]:
            model=schedule_model()
            model["tasks"][2][key]=value
            self.assertTrue(schedule.calculate(model),key)
        s,r,w,d=case(self.base)
        s["plan"]["hard_checks"]["schedule"]["tasks"][2]["name"]="更改活动"
        self.assertTrue(any("进度图与唯一计划数据不一致" in e for e in hc.audit(s,d.parent,d,r,common,w)))

    def test_stale_figure_and_wrong_page(self):
        s,r,w,d=case(self.base)
        w["figure_pages"]["图1 施工总平面布置图"]=[99]
        self.assertTrue(any("页不符" in e for e in hc.audit(s,d.parent,d,r,common,w)))


if __name__=="__main__":
    if len(sys.argv)==3 and sys.argv[1]=="--runner-fixture":
        case(Path(sys.argv[2]),"plan")
    else:
        unittest.main()
