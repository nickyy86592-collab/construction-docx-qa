#!/usr/bin/env python3
"""Generic Word runner: project facts/rules are runtime inputs, never defaults here."""
from __future__ import annotations
import argparse
import importlib.util
import json
import subprocess
import sys
from pathlib import Path

COMMON = Path(__file__).resolve().parent
# The Windows runner uses an embeddable Python with an isolated ._pth file.
sys.path.insert(0, str(COMMON))

def execute(argv):
    bootstrap = (
        "import runpy, sys; from pathlib import Path; "
        "sys.argv.pop(0); "
        "sys.path.insert(0, str(Path(sys.argv[0]).resolve().parent)); "
        "runpy.run_path(sys.argv[0], run_name='__main__')"
    )
    return subprocess.run([sys.executable, "-c", bootstrap, *map(str, argv)]).returncode

def load_flow(page_dir):
    spec=importlib.util.spec_from_file_location('_project_flow',COMMON/'编制流程门禁.py')
    m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
    m.PAGE_DIR=page_dir
    return m

def readiness(project_dir, mode):
    if mode=='formal':
        return (project_dir/'design.docx').is_file() and (project_dir/'format_rules.json').is_file()
    if mode=='pilot':return (project_dir/'pilot.docx').is_file()
    return any((project_dir/f'stage-{stage}.docx').is_file() for stage in (40,75))

def run(project_dir, outdir, mode, stage, photo_stage):
    flow=load_flow(project_dir)
    page_script,values=flow.discover_page_script()
    if page_script is None:
        raise RuntimeError('Runtime project directory needs exactly one project page-audit script')
    outdir.mkdir(parents=True,exist_ok=True)
    failures=[]
    if mode=='pilot':
        document=project_dir/'pilot.docx'
        return execute([COMMON/'编制流程门禁.py','--page-dir',project_dir,'run-page-audit','--pilot','--document',document,'--report',outdir/'pilot-word-pages.json'])
    if mode=='stage':
        stages=[stage] if stage is not None else [n for n in (40,75) if (project_dir/f'stage-{n}.docx').is_file()]
        if not stages:
            raise RuntimeError('No stage DOCX snapshot provided')
        for n in stages:
            document=project_dir/f'stage-{n}.docx'
            rc=execute([COMMON/'编制流程门禁.py','--page-dir',project_dir,'run-page-audit','--document',document,'--stage',n,'--report',outdir/f'stage-{n}-word-pages.json'])
            if rc: failures.append(f'stage-{n}')
        return 2 if failures else 0
    document=project_dir/'design.docx'
    rules_path=project_dir/'format_rules.json'
    rules=json.loads(rules_path.read_text(encoding='utf-8-sig'))
    if photo_stage!='auto':rules['photo_delivery_stage']=photo_stage
    if rules.get('photo_delivery_stage') not in ('draft','final'):
        raise RuntimeError('Runtime rules must declare draft/final photo stage')
    effective=outdir/'effective-format-rules.json'
    effective.write_text(json.dumps(rules,ensure_ascii=False,indent=2),encoding='utf-8')
    for label,args in [
        ('static',[COMMON/'Word正式交付硬审（不含页数）.py',document,'--page-dir',project_dir,'--rules',effective,'--static-only','--report',outdir/'static.json']),
        ('pages',[COMMON/'编制流程门禁.py','--page-dir',project_dir,'run-page-audit','--document',document,'--report',outdir/'pages.json']),
        ('word',[COMMON/'Word正式交付硬审（不含页数）.py',document,'--page-dir',project_dir,'--rules',effective,'--report',outdir/'delivery.json']),
        ('word_pdf',[COMMON/'export_pdf.py',document,outdir/'design-word.pdf'])]:
        code=execute(args)
        if code:failures.append(label)
    # This is a component-audit result, not the skill's complete delivery gate.
    summary={'audit_type':'word_runner_components','status':'FAIL' if failures else 'PASS',
             'document_sha256':flow.sha256(document),'project_rule_source':page_script.name,
             'failed_components':failures,'photo_delivery_stage':rules['photo_delivery_stage'],
             'notes':['完整交付仍需Skill核对启动/阶段状态、评分、内容溯源及WARN处理；本报告不能替代完整delivery门禁。']}
    (outdir/'runner-summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    return 2 if failures else 0

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--mode',choices=('formal','stage','pilot'),required=True)
    p.add_argument('--project-dir',type=Path,required=True)
    p.add_argument('--outdir',type=Path,default=Path('output/project'))
    p.add_argument('--stage',type=int,choices=(40,75))
    p.add_argument('--photo-stage',choices=('auto','draft','final'),default='auto')
    p.add_argument('--check-ready',action='store_true')
    a=p.parse_args()
    try:
        if a.check_ready:
            print('true' if readiness(a.project_dir,a.mode) else 'false')
            return 0
        return run(a.project_dir.resolve(),a.outdir.resolve(),a.mode,a.stage,a.photo_stage)
    except Exception as exc:
        print('RUNNER_AUDIT_STATUS=ERROR\n'+str(exc),file=sys.stderr)
        return 3
if __name__=='__main__':raise SystemExit(main())
