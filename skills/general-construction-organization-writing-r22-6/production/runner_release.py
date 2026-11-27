"""受信任的 GH Word 作业及正式发布入口。"""
from __future__ import annotations

import argparse
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile

from drive_release import Drive, check_access, publish, receive, save_evidence, test_author_upload
from release_task import (EVIDENCE, PHASES, PROBE, STATE, TASK, TRUST, GateError, archive,
                          ast_with_constants, check_probe, digest, documents_for_phase, file_sha, files_in,
                          inside, load_module, now, page_script, profile, read, require,
                          signed, unpack, verify_receipt, verify_task, write)
from word_native import probe

BOOTSTRAP = ("import runpy,sys;from pathlib import Path;sys.argv.pop(0);"
             "sys.path.insert(0,str(Path(sys.argv[0]).parent));runpy.run_path(sys.argv[0],run_name='__main__')")
WORD_FORMAT = "Word正式交付硬审（不含页数）.py"
FLOW = "编制流程门禁.py"
PRIVATE_RUN = ("repository", "branch", "id", "attempt", "commit")


def context(trust):
    require(os.environ.get("CONSTRUCTION_PRIVATE_REPOSITORY") == "true",
            "正式送审要求已登记的私有生产 GH 仓库")
    result = {a: os.environ[k] for a, k in zip(
        PRIVATE_RUN, ("GITHUB_REPOSITORY", "GITHUB_REF_NAME", "GITHUB_RUN_ID",
                      "GITHUB_RUN_ATTEMPT", "GITHUB_SHA"), strict=True)}
    require(result["repository"] == trust["repository"] and result["branch"] == trust["branch"] and
            os.environ.get("GITHUB_EVENT_NAME") == "workflow_dispatch" and
            result["id"].isdigit() and result["attempt"].isdigit() and
            re.fullmatch(r"[0-9a-f]{40}", result["commit"]), "既定仓库、分支或真实 GH 运行信息不匹配")
    return result


def log_call(script, args, project, log):
    environment = {k: v for k, v in os.environ.items() if not k.startswith(("CONSTRUCTION_", "INPUT_"))
                   and k not in ("GITHUB_TOKEN", "GH_TOKEN")}
    with Path(log).open("ab") as output:
        result = subprocess.run([sys.executable, "-I", "-X", "utf8", "-c", BOOTSTRAP,
                                 str(script.resolve()), *map(str, args)], cwd=project, env=environment,
                                stdout=output, stderr=subprocess.STDOUT, check=False)
    return result.returncode


def select(root, request, project, settings):
    task = request["task"]
    tag = task["skill_version"]
    require(tag in settings["profiles"] and re.fullmatch(r"r[0-9]+(?:\.[0-9]+)?", tag),
            "编制版本尚未登记审核配对")
    version = root / "versions" / tag
    registered = profile(version)
    verify_task(task, settings, registered)
    require(read(project / TASK) == task and request["phase"] in PHASES and
            request["schema"] == 1 and re.fullmatch(r"[0-9a-f]{32}", request["request_id"]), "任务单或阶段无效")
    expected = set(documents_for_phase(project, task, request["phase"])) | {TASK}
    require(expected <= request["files"].keys(), "这次送审缺本阶段稿件或规则")
    if request["phase"] == "probe":
        return version, registered, None
    require(read(project / TRUST).get("profiles", {}).get(tag) == registered["audit_sha256"],
            "本批编制信任配置与生产登记不一致")
    state = read(project / STATE)
    require(state["startup"].get("versioned_task_sha256") == digest(task), "项目状态和送审批次未绑定")
    check_probe(project, task, settings)
    script = page_script(project, version, registered)
    if not script:
        fmt = (state["plan"]["bid_requirements"].get("format") or {})
        require(state.get("page_rule") is None and not fmt.get("pages"), "已要求限页的项目缺页数硬审脚本")
    return version, registered, script


def stage(project, out, scripts, target, phase, script, reference=False):
    require(script is not None, "本阶段页数校准缺项目页数来源")
    receipt = out / "pages.json"
    args = ["--page-dir", project, "run-page-audit", "--document", target, "--report", receipt]
    args += ["--pilot"] if phase == "pilot" else ["--stage", phase.split("-")[-1]]
    if reference and phase != "pilot":
        args.append("--reference")
    require(log_call(scripts / FLOW, args, project, out / "发动记录.txt") == 0 and receipt.is_file(),
            "本阶段 Word 实排失败")
    result = read(receipt)
    require(result.get("status") == "MEASURED" and result.get("page_source") == "microsoft_word" and
            result.get("document_sha256") == file_sha(target)
            and result.get("point_limit_status", "PASS") == "PASS", "本阶段缺真实 Word 实测或逐点超限")
    if phase != "pilot":
        hc = load_module("_native_figures", scripts / "hard_checks.py")
        write(out / "figure-pages.json", hc.inspect_word(target))


def plan_and_layout(version, registered, project, script, task):
    scripts = version / "scripts"
    flow = load_module("_bound_flow", scripts / FLOW)
    flow.PAGE_DIR = project
    fmt = read(project / STATE)["plan"]["bid_requirements"]["format"]
    common = flow.load_format_audit()
    rules = common.load_rules(project / "format_rules.json", project, allow_unlimited=script is None)
    require(fmt.get("deliverable") in ("docx", "pdf") and
            rules.get("submission_format", "docx") == fmt["deliverable"] and
            rules.get("photo_delivery_stage") == "final", "格式文件须对应本次实际提交格式与正式阶段")
    if registered["skill_version"] == "r23.6":
        layout = fmt.get("submission_layout") or {}
        require(layout.get("mode") == task["submission"] and isinstance(layout.get("source"), dict),
                "逐点或合册提交方式须在本项目计划登记有效招标依据，并与任务单相同")
    if task["submission"] == "review-points":
        path, vals = flow.discover_page_script(project)
        rule = flow.current_rule(path, vals)
        mod = load_module("_parts_for_release", scripts / "split_submission.py")
        specs = mod.point_specs(project, rules, rule, common)
        return flow, common, fmt, rules, mod, specs, rule
    return flow, common, fmt, rules, None, None, None


def measure(project, out, version, registered, script, task):
    scripts = version / "scripts"
    flow, common, fmt, rules, split, specs, rule = plan_and_layout(version, registered, project, script, task)
    design = project / "design.docx"
    common_args = [design, "--rules", project / "format_rules.json", "--page-dir", project]
    if script is None:
        common_args.append("--allow-unlimited")
    log = out / "发动记录.txt"
    failures = []
    if log_call(scripts / WORD_FORMAT, [*common_args, "--static-only", "--report", out / "static.json"], project, log):
        failures.append("合册静态硬审未通过")
    if script:
        flags = ["--reference"] if split else []
        if log_call(scripts / FLOW, ["--page-dir", project, "run-page-audit", "--document", design,
                                     "--report", out / "pages.json", *flags], project, log):
            failures.append("本项目 Word 页数或参考合册边界未通过")
    if log_call(scripts / WORD_FORMAT, [*common_args, "--report", out / "delivery.json"], project, log):
        failures.append("合册 Word 版式硬审未通过")
    if split:
        point_word = load_module("_bound_point_page", version / registered["page_template"])
        _, values = ast_with_constants(script)
        for key, value in values.items():
            setattr(point_word, key, value)
        parts = split.native_measure(project, out, specs, rule, common, rules, point_word)
        if parts["status"] != "PASS":
            failures.append("实际提交的评审点 Word 页数或版式未通过")
    require(not failures, "；".join(failures))


def evidence(project, task, trust, phase, stable):
    folder = project / EVIDENCE / phase
    result = verify_receipt(read(folder / "receipt.json"), task, trust, phase)
    require(result["status"] == "PASS", "缺已通过的本批 " + phase + " 回执")
    for name, expected in result["files"].items():
        require(file_sha(inside(folder, name)) == expected, "Word 回执附件与签名清单不一致")
    for name in stable:
        require(result["inputs"].get(name) == file_sha(inside(project, name)), "稿件或规则变化，须按本批版本重新实排")
    return folder, result


def stage_evidence(project, task, trust, flow, script, state):
    if not script:
        return
    for number in (40, 75):
        phase = f"stage-{number}"
        cp = state["checkpoints"][str(number)]
        snapshot = cp["snapshot_path"]
        folder, proof = evidence(project, task, trust, phase, [snapshot, script.name])
        require(cp["page_source"] == "microsoft_word" and cp["status"] == "PASS" and
                file_sha(inside(project, cp["word_report"]["report_file"])) == file_sha(folder / "pages.json"),
                "阶段校准未绑定本批受信 Word 实测报告")
    if flow.planning_policy(state["page_rule"]["body_pages_target"])["pilot_required"]:
        pilot = state["pilot"]
        folder, proof = evidence(project, task, trust, "pilot", [pilot["snapshot"]["path"], script.name])
        require(file_sha(inside(project, pilot["report"]["path"])) == file_sha(folder / "pages.json"),
                "样章校准未绑定本批受信 Word 实测报告")


def release_files(project, out, version, registered, script, task, trust):
    scripts = version / "scripts"
    state = read(project / STATE)
    flow, common, fmt, rules, split, specs, rule = plan_and_layout(version, registered, project, script, task)
    stable = documents_for_phase(project, task, "measure")
    if script:
        stable.append(script.name)
    stage_evidence(project, task, trust, flow, script, state)
    if registered["skill_version"] == "r23.6":
        from bid_requirements import source_ok
        from hard_checks import source_lookup

        roles = {s["id"] for s in state["startup"]["sources"] if s["role"] in ("tender", "amendment")}
        layout = fmt["submission_layout"]
        require(source_ok(layout["source"], source_lookup(state, project), roles), "提交方式依据无法回查有效招标原文")
        quote = layout["source"]["quote"]
        separate = re.search(r"(?:各|每|逐|按)[^。；]{0,12}评审点[^。；]{0,24}(?:分别|单独|逐项)[^。；]{0,12}(?:提交|上传|递交)|(?:按|逐项|分别|单独)[^。；]{0,16}评审点[^。；]{0,16}(?:提交|上传|递交)|(?:分成|分别)[^。；]{0,22}(?:十份|10份)[^。；]{0,14}(?:提交|上传|递交)", quote)
        require(not separate or task["submission"] == "review-points", "有效招标要求逐点提交，不能以合册放行")
    reports, proof = evidence(project, task, trust, "measure", stable)
    approved = out.parent / ("核验件." + fmt["deliverable"])
    args = ["--project", project, "--reports", reports, "--destination", approved]
    require(log_call(scripts / "release_docx.py", args, project, out / "发动记录.txt") == 0,
            "原编制版本的完整交付门禁未通过")
    decision = read(project / "流程交付结果.json")
    require(decision["status"] == "PASS" and decision["task_complete"] is True and
            decision["formal_release"] is True and decision["document_sha256"] == file_sha(project / "design.docx"),
            "完整门禁未批准本稿")
    if split:
        split.verify_tender_scope(project, state, specs, layout)
        split.compare_content(project, specs, rule, common, rules)
        identity = load_module("_bound_identity", scripts / "submission_identity.py")
        return split.release_candidates(project, reports, specs, rule, fmt, read(project / "review.json"), identity)
    if fmt["deliverable"] == "pdf":
        target = flow._resolve_record(decision["submission_pdf"]["path"], project)
        relative = target.relative_to(project).as_posix()
        filename = read(project / "review.json").get("submission_name", "施工组织设计.pdf")
    else:
        relative = "design.docx"
        filename = read(project / "review.json").get("submission_name", "施工组织设计.docx")
    source = inside(project, relative)
    require(Path(filename).name == filename and "/" not in filename and "\\" not in filename and
            Path(filename).suffix.casefold() == "." + fmt["deliverable"], "正式文件名无效")
    identity = load_module("_artifact_identity", scripts / "submission_identity.py")
    require(not fmt["anonymous"] or not any(identity.folded(t) in identity.folded(filename)
                                           for t in fmt["identifiers"]), "正式文件名违反暗标要求")
    checksum = file_sha(approved)
    require(file_sha(source) == checksum, "实际待发字节与原版正式门禁产物不一致")
    return {filename: {"path": relative, "source": source, "sha256": checksum}}


def execute(root, project, out, request, settings, publisher, folders):
    registration, registered, script = select(root, request, project, settings)
    task, phase = request["task"], request["phase"]
    if phase == "probe":
        return {"measurement": probe(project / PROBE, task)}
    if phase in ("pilot", "stage-40", "stage-75"):
        stage(project, out, registration / "scripts", project / (phase + ".docx"), phase, script,
              reference=task["submission"] == "review-points")
        return {}
    if phase == "measure":
        measure(project, out, registration, registered, script, task)
        return {}
    candidates = release_files(project, out, registration, registered, script, task, settings)
    metadata = [{"name": name, "path": value["path"], "sha256": value["sha256"]} for name, value in candidates.items()]
    # The publish function itself reads these exact bytes, stages privately and returns only after the complete move.
    url = publish(publisher, folders, digest(task), task["batch_id"],
                  {name: (item["source"], item["sha256"]) for name, item in candidates.items()})
    return {"formal_release": True, "task_complete": True, "submission": metadata, "formal_url": url}


def run(root, drive_file_id, packet_sha):
    settings = read(root / TRUST)
    run_record = context(settings)
    key = os.environ["CONSTRUCTION_SIGNING_KEY"]
    signed({}, key, settings)  # Verify the signing identity before requesting private files.
    require(re.fullmatch(r"[0-9a-f]{64}", packet_sha), "送审压缩包校验值无效")
    author = Drive.from_env("CONSTRUCTION_DRIVE_AUTHOR_OAUTH")
    publisher = Drive.from_env("CONSTRUCTION_DRIVE_PUBLISHER_OAUTH")
    folders = check_access(author, publisher, settings)
    with tempfile.TemporaryDirectory(prefix="construction-release-") as work:
        work = Path(work)
        project, out = work / "project", work / "out"
        receive(publisher, folders, drive_file_id, packet_sha, work / "request.zip")
        request = unpack(work / "request.zip", "request.json", project)
        phase, task = request["phase"], request["task"]
        require(re.fullmatch(r"[0-9a-f]{32}", request["request_id"]), "送审请求编号无效")
        existing = [file for file in publisher.children(folders["evidence"])
                    if file["name"] == "回执-" + request["request_id"] + ".zip"]
        require(len(existing) <= 1, "发现多份同号私有回执")
        if existing:
            previous = work / "previous.zip"
            previous.write_bytes(publisher.download(existing[0]["id"]))
            prior = unpack(previous, "receipt.json", work / "previous")
            result = verify_receipt(prior, task, settings, phase)
            require(result["request_id"] == request["request_id"] and result["inputs"] == request["files"],
                    "同一送审号不能更换冻结文件")
            print(f"REQUEST={request['request_id']} STATUS={result['status']} REPLAY=VERIFIED")
            return 0 if result["status"] == "PASS" else 2
        out.mkdir()
        error = None
        try:
            extra = execute(root, project, out, request, settings, publisher, folders)
        except Exception as exc:
            extra = {}
            error = str(exc) or type(exc).__name__
        result = {"schema": 1, "phase": phase, "status": "FAIL" if error else "PASS", "checked_at": now(),
                  "task_sha256": digest(task), "audit_sha256": task["audit_sha256"],
                  "request_id": request["request_id"], "inputs": request["files"], "runner": run_record,
                  "formal_release": False, "task_complete": False, "files": files_in(out), **extra}
        if error:
            result["failure"] = error
        envelope = signed(result, key, settings)
        packet = work / "receipt.zip"
        archive(packet, "receipt.json", envelope, {name: out / name for name in result["files"]})
        save_evidence(publisher, folders["evidence"], request["request_id"], packet)
    # GH stores a minimal status, without project files, GD links or diagnostic lines.
    print(f"REQUEST={request['request_id']} PHASE={phase} STATUS={result['status']} FORMAL_RELEASE={result['formal_release']}")
    return 2 if error else 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    sub = parser.add_subparsers(dest="command", required=True)
    execute_job = sub.add_parser("run")
    execute_job.add_argument("--drive-file", required=True)
    execute_job.add_argument("--packet-sha256", required=True)
    sub.add_parser("check-acl")
    args = parser.parse_args()
    if args.command == "run":
        return run(args.root.resolve(), args.drive_file, args.packet_sha256)
    settings = read(args.root / TRUST)
    author = Drive.from_env("CONSTRUCTION_DRIVE_AUTHOR_OAUTH")
    publisher = Drive.from_env("CONSTRUCTION_DRIVE_PUBLISHER_OAUTH")
    folders = check_access(author, publisher, settings)
    test_author_upload(author, folders["formal"])
    print("AUTHOR_WRITE=DENIED")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (GateError, OSError, ValueError, KeyError, TypeError, ImportError):
        print("STATUS=BLOCKED，核对私有工作流输入、授权目录和批次配置", file=sys.stderr)
        raise SystemExit(2)

