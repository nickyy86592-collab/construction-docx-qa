"""在隔离的 Runner 回归目录记录真实 Word 图页，检查生产交付端的结果。"""
from pathlib import Path
import json
import sys


def run(root, action):
    sys.path.insert(0,str((root/"scripts").resolve()))
    project=root/"input"/"project"
    if action=="pages":
        from hard_checks import inspect_word
        actual=inspect_word(project/"design.docx")
        print("WORD_NATIVE",actual)
        review=json.loads((project/"review.json").read_text("utf-8"))
        assert len(actual["figure_pages"])==len(review["hard_checks"]["figures"])
        for caption,pages in actual["figure_pages"].items():
            assert len(pages)==1,(caption,pages)
            review["hard_checks"]["figures"][caption.split(" ",1)[1]]["page"]=pages[0]
        (project/"review.json").write_text(json.dumps(review,ensure_ascii=False),"utf-8")
        (root/"input/reports/delivery.json").write_text(json.dumps(actual,ensure_ascii=False),"utf-8")
        return
    if action!="verify":
        raise ValueError(action)
    result=json.loads((project/"流程交付结果.json").read_text("utf-8"))
    print(json.dumps(result,ensure_ascii=False))
    assert result["failures"]==["D1 甲班 被 A、B 重复占用"],result["failures"]
    assert result["status"]=="FAIL" and result["formal_release"] is False and result["task_complete"] is False
    assert not list((root/"formal").glob("*"))
    print("EXPECTED_CREW_CONFLICT_ONLY; formal_release=false; PUBLISHED_FILES=0")


if __name__=="__main__":
    run(Path.cwd(),sys.argv[1])
