"""正式交付端。每次调用原 delivery 命令，复核本次结果后才发布对应字节。"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile


def release(project, reports, destination):
    project, reports = project.resolve(), reports.resolve()
    source = project / "design.docx"
    state = project / "篇幅规划状态.json"
    script = Path(__file__).with_name("编制流程门禁.py")
    reports.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(dir=reports) as run:
        result = Path(run) / "decision.json"
        bootstrap = "import runpy,sys;from pathlib import Path;sys.argv.pop(0);sys.path.insert(0,str(Path(sys.argv[0]).parent));runpy.run_path(sys.argv[0],run_name='__main__')"
        args=[sys.executable,"-c",bootstrap,str(script),"--state",str(state),"--page-dir",str(project),
              "delivery","--mode","full","--document",str(source),"--result",str(result)]
        for flag,path in {"review":project/"review.json","score-report":project/"score.json",
                          "warning-review":project/"warnings.json","static-report":reports/"static.json",
                          "word-report":reports/"delivery.json"}.items():
            args.extend(["--"+flag,str(path)])
        if (reports/"pages.json").exists():
            args.extend(["--page-report",str(reports/"pages.json")])
        rc=subprocess.run(args).returncode
        data=json.loads(result.read_text("utf-8")) if result.is_file() else {}
        if rc or data.get("formal_release") is not True or data.get("task_complete") is not True or data.get("status")!="PASS":
            print("PUBLISH_BLOCKED formal_release=false delivery_exit="+str(rc))
            return rc or 2
        from submission_identity import folded
        state_data=json.loads(state.read_text("utf-8-sig"))
        fmt=state_data["plan"]["bid_requirements"]["format"]
        if destination.suffix.lower()!="."+fmt["deliverable"] or (fmt["anonymous"] and any(folded(t) in folded(destination.name) for t in fmt["identifiers"])):
            print("PUBLISH_BLOCKED 交付文件名或扩展名与项目要求冲突")
            return 2
        artifact=data.get("submission_pdf") or {"path":source,"sha256":data["document_sha256"]}
        payload=Path(artifact["path"]).read_bytes()
        if hashlib.sha256(payload).hexdigest()!=artifact["sha256"]:
            print("PUBLISH_BLOCKED document changed after audit")
            return 2
        destination.parent.mkdir(parents=True,exist_ok=True)
        # Same-volume atomic no-clobber publication; simultaneous runs cannot mix bytes.
        with tempfile.NamedTemporaryFile(dir=destination.parent,delete=False) as tmp:
            tmp.write(payload)
        try:
            os.link(tmp.name,destination)
        finally:
            Path(tmp.name).unlink(missing_ok=True)
        print("PUBLISHED "+str(destination))
        return 0


if __name__=="__main__":
    cli=argparse.ArgumentParser(description=__doc__)
    for name in ("project","reports","destination"):
        cli.add_argument("--"+name,type=Path,required=True)
    ns=cli.parse_args()
    raise SystemExit(release(ns.project,ns.reports,ns.destination))
