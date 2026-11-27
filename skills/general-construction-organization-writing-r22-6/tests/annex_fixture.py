"""Isolated tender, resource plan and actual Word annex for the publishing regression."""
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/"scripts"))
from docx import Document
from bid_fixture import aerial_spec, fixture_capacity, profile_evidence, profile_text, project_profile
from test_hard_checks import SITE, case, write, flow, common, hc, put_image, site_image


def annex_case(base, incorrect=False, incorrect_aerial=False):
    state, review, word, manuscript = case(base)
    project = manuscript.parent
    source = project / "测试要求.md"
    quote = "试件要求按阶段列明安装工与普工人数，按试件操作标准核验。"
    source.write_text(quote, encoding="utf-8")
    ref = {"source_id":"X", "unit_id":"X:U000001", "quote":quote}
    state["startup"]["sources"].append(dict(id="X", role="tender", path=source.name, sha256=flow.sha256(source)))
    write(project/"scoring.json", {"items":[]})
    state["plan"]["scoring_basis"] = {"path":"scoring.json"}
    profile = project_profile(ref, "M", "R", [], [], ["甲班","乙班"])
    state["plan"]["bid_requirements"] = profile
    tasks = state["plan"]["hard_checks"]["schedule"]["tasks"]
    doc = Document(manuscript)
    site = doc.add_paragraph(profile_text(profile)).text
    resources = {}
    for task in tasks:
        task["capacity"] = fixture_capacity(ref,task["crews"])
        task["productivity_basis"] = task["name"]+"每小时完成1米"
        for name in task["equipment"]:
            profile["equipment"][name] = dict(count=2,spec="D-8",capacity="单台每小时1米",entry="开工日到位",basis="试件参数")
            task["capacity"]["plant"][name] = dict(operation="切割或测量",unit="米",unit_rate=1,conversion="按试件有效作业时间核算")
        resources[task["id"]] = doc.add_paragraph("；".join([task["productivity_basis"],*task["crews"],
            "日班，开工日到位，安装工2人，普工3人",*task["equipment"],"D-8共2台，单台每小时1米"])).text
    name = "劳动力计划表"
    policy = state["startup"]["visual_policy"]
    policy["named_exhibits"].append({"kind":"表","name":name})
    figure=aerial_spec("M")
    state["plan"]["hard_checks"]["figures"].append(figure)
    policy["additional_figures"]={"names":[figure["name"]],"reason":"用试件示意图核对边界与出入口"}
    # The distinct raster bytes are synthetic test data, never presented as a real AI rendering.
    layout=project/profile["aerial"]["layout_path"]
    raw=site_image()
    layout.write_bytes(raw+b"\n")
    opts={**SITE,"施工边界":("图示外围","核对场地外轮廓与出入口相接")}
    review["hard_checks"]["figures"][figure["name"]]=dict(page=1,image_sha256=hc.sha256(raw).hexdigest(),observed_kind="site",
        elements={k:{"status":"verified","position":v[0],"finding":v[1]} for k,v in opts.items()},
        generation=dict(method="missing" if incorrect_aerial else "ai_from_layout",
                        layout_sha256=flow.sha256(layout),layout_locator=profile["aerial"]["layout_locator"],record="合成回归夹具图源记录"))
    put_image(doc,raw)
    doc.paragraphs[-1].paragraph_format.keep_with_next=True
    caption="图4 "+figure["name"]
    doc.add_paragraph(caption)
    word["figure_pages"][caption]=[1]
    rows = [["工种","开工阶段","安装阶段"],["安装工","4","2"],["普工","6","3"]]
    phases = {"开工阶段":[1,1],"安装阶段":[2,3]}
    profile["annexes"][name] = {"source":ref, "basis":"按试件活动和两班实际合计", "module":"M",
        "headers":[rows[0]], "key_col":1, "required_cols":[2,3], "resource_cols":[2,3],
        "resource_kind":"trade", "task_ids":[t["id"] for t in tasks], "phases":phases,
        "rows":[{"key":trade,"cells":{str(i):{"phase":phase,"peak":["trade",trade,*days]}
            for i,(phase,days) in enumerate(phases.items(),2)}} for trade in ("安装工","普工")]}
    doc.add_heading(name,2)
    doc.add_paragraph("表1 "+name)
    tab = doc.add_table(rows=0,cols=3)
    if incorrect:
        rows[1][1]="2"
    for row in rows:
        for cell,value in zip(tab.add_row().cells,row):
            cell.text=value
    doc.save(manuscript)
    positions = {text:loc for loc,text in hc.document_rows(common._docx_blocks(manuscript)).items()}
    review["document_sha256"] = word["document_sha256"] = flow.sha256(manuscript)
    review["bid_evidence"] = profile_evidence(profile,positions[site])
    review["bid_evidence"]["resources"] = {t["id"]:{"productivity":positions[resources[t["id"]]],
        **{n:{"locations":[positions[resources[t["id"]]]]} for n in [*t["crews"],*t["equipment"]]}} for t in tasks}
    word["rules"] = {"visual_policy":policy}
    write(project/"篇幅规划状态.json",state)
    write(project/"review.json",review)
    write(base/"reports"/"delivery.json",word)
    return state,review,word,manuscript


if __name__ == "__main__":
    annex_case(Path(sys.argv[1]), incorrect=True, incorrect_aerial=True)
