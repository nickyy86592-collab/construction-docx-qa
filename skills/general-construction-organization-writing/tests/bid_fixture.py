"""Synthetic fixtures, used only by the regression suite."""
import bid_requirements as bid

def project_profile(ref, module, rid, scoring, score_items, teams):
    p=bid.template()
    p.update(jurisdiction="试验省试验市",as_of="2026-11-23")
    p["format"].update(source=ref,anonymous=False)
    p["score_effects"]={r:{"kind":"未明示","scope":"项目综合项","source":ref} for r in scoring}
    p["score_sources"]={i:ref for i in score_items}
    topics=["国家强制要求","属地监管要求","工艺","安全","岗位资格","危大工程","临电","扬尘"]
    active=["工艺","安全","国家强制要求"]
    p["screening"]={k:{"applicable":k in active,"reason":"试件范围已核实","source":ref} for k in topics}
    common_row=dict(module=module,requirement=rid,source=ref)
    p["obligations"]=[dict(common_row,id="O1",authority="试件操作标准",clause="试件第1条",topics=active,
        version="测试版",territory="全国",level="国家",scope="工序交接",operation="检查接口",role="班长",
        check="核对记录",acceptance="复查接口后移交",validity=dict(ref,effective_from="2026-01-01",status="有效"))]
    p["decisions"]=[dict(common_row,id="D1",condition="场地受限",choice="分时交接",verification="逐段验收")]
    p["crews"]={n:{"basis":"试件班组工作记录","entry":"开工日到位","shift":"日班","trades":{"安装工":2,"普工":3}} for n in teams}
    return p



def fixture_capacity(ref,teams):
    return {"source":ref,"calculation":"每小时一个作业单元，使用一个完整班组",
            "crews":{n:{"trades":{"安装工":2,"普工":3},"rate":1} for n in teams},"plant":{}}
