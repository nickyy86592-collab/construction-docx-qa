"""Isolated tender, resource plan and actual Word annex for the publishing regression."""
from pathlib import Path
from docx import Document
from bid_fixture import project_profile, fixture_capacity
from test_hard_checks import case, write, flow, common, hc


def annex_case(base, incorrect=False):
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
    site = doc.add_paragraph("场地受限，分时交接并逐段验收。试件操作标准：试件第1条由班长检查接口，核对记录，复查接口后移交。").text
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
    review["bid_evidence"] = {r["id"]:{field:positions[site] for field in r}
                               for r in [*profile["decisions"],*profile["obligations"]]}
    review["bid_evidence"]["resources"] = {t["id"]:{"productivity":positions[resources[t["id"]]],
        **{n:{"locations":[positions[resources[t["id"]]]]} for n in [*t["crews"],*t["equipment"]]}} for t in tasks}
    word["rules"] = {"visual_policy":policy}
    write(project/"篇幅规划状态.json",state)
    write(project/"review.json",review)
    write(base/"reports"/"delivery.json",word)
    return state,review,word,manuscript


if __name__ == "__main__":
    import sys
    annex_case(Path(sys.argv[1]), incorrect=True)
