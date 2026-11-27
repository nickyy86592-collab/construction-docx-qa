"""Synthetic fixtures, used only by the regression suite."""
import bid_requirements as bid

def project_profile(ref, module, rid, scoring, score_items, teams):
    p=bid.template()
    p.update(jurisdiction="试验省试验市",as_of="2026-11-23")
    p["format"].update(source=ref,anonymous=False)
    p["score_effects"]={r:{"kind":"未明示","scope":"项目综合项","source":ref} for r in scoring}
    p["score_sources"]={i:ref for i in score_items}
    active=["工艺","安全","国家强制要求","动火防火","扬尘"]
    p["screening"]={k:{"applicable":k in active,"reason":"试件工况与范围独立核查","source":ref}
                    for k in bid.SCREENING_TOPICS}
    common_row=dict(module=module,requirement=rid,source=ref)
    p["obligations"]=[dict(common_row,id="O1",authority="试件操作标准",clause="试件第1条",topics=active,
        version="测试版",territory="全国",level="国家",scope="工序交接",stage="拆装工序",exception="停工并重新核查",operation="检查接口",role="班长",
        check="核对记录",acceptance="复查接口后移交",validity=dict(ref,effective_from="2026-01-01",status="有效"))]
    p["decisions"]=[dict(common_row,id="D1",condition="场地受限",choice="分时交接",verification="逐段验收")]
    p["site"]["conditions"]={k:{"status":"confirmed" if k in bid.SITE_CORE else "not_applicable",
        "source":ref,"reason":"试件西侧作业区、南门运输及东侧邻接库房已核对","decision":"D1"} for k in bid.SITE_TOPICS}
    p["site"]["visit"]={"required":False,"source":ref,"reason":"试件招标未明确要求投标阶段实地踏勘"}
    p["difficulties"].update(source=ref,reason="试件招标未点名工程重难点",items=[
        dict(origin="project",name="狭窄通道交接",cause="入口单车位",impact="交接容易交叉",
             resources="配置一名交通疏导员",source=ref,decision="D1")])
    p["aerial"].update(source=ref,layout_path="fixture-layout.png",layout_locator="试件总平面图第1页")
    p["crews"]={n:{"basis":"试件班组工作记录","entry":"开工日到位","shift":"日班","trades":{"安装工":2,"普工":3}} for n in teams}
    return p



def fixture_capacity(ref,teams):
    return {"source":ref,"calculation":"每小时一个作业单元，使用一个完整班组",
            "crews":{n:{"trades":{"安装工":2,"普工":3},"rate":1} for n in teams},"plant":{}}


def aerial_spec(module):
    from hard_checks import AERIAL_FIGURE, ELEMENTS
    return dict(name=AERIAL_FIGURE,kind="site",basis="试件施工总平面上的边界和通行关系承接方案",
                module=module,elements=sorted(ELEMENTS["site"]|{"施工边界"}))


def profile_evidence(p,location):
    result={r["id"]:{k:location for k in r} for r in [*p["decisions"],*p["obligations"]]}
    for item in p["difficulties"]["items"]:
        result[item["decision"]].update({k:location for k in bid.DIFFICULTY_FIELDS})
    return result


def profile_text(p):
    return "，".join([
        *(o[k] for o in p["obligations"] for k in ("authority","clause","operation","role","check","acceptance",*bid.CONTROL_FIELDS)),
        *(d[k] for d in p["decisions"] for k in ("condition","choice","verification")),
        *(d[k] for d in p["difficulties"]["items"] for k in bid.DIFFICULTY_FIELDS),
    ])+"。"
