"""编制任务、冻结输入与经签名的 Windows Word 证据。"""
from __future__ import annotations

import argparse
import ast
import base64
from datetime import datetime, timezone
from hashlib import file_digest, sha256
import importlib.util
import json
import os
from pathlib import Path, PurePosixPath
import re
import sys
import tempfile
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from uuid import uuid4
import zipfile

TASK = "交付任务单.json"
STATE = "篇幅规划状态.json"
TRUST = "release-trust.json"
PROBE = "word-probe.docx"
EVIDENCE = "release-evidence"
PHASES = ("probe", "pilot", "stage-40", "stage-75", "measure", "publish")
PAGE_NAMES = {
    "EXPECTED_BODY_PAGES", "PAGE_TOLERANCE", "FRONT_MATTER_PAGES", "PROJECT_NAME",
    "COVER_PAGES", "REVIEW_POINT_COUNT", "REVIEW_POINT_PAGE_LIMIT", "REVIEW_POINT_HEADINGS",
}
PROFILE_NAME = "release-profile.json"
PRODUCTION_WORKFLOW = "construction-versioned-release.yml"
GATEWAY_WORKFLOW = "construction-word-dispatch.yml"


class GateError(Exception):
    """A required fact or receipt failed validation."""


def require(condition, message):
    if not condition:
        raise GateError(message)


def read(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=path.parent, delete=False) as tmp:
        tmp.write(canonical(value) + b"\n")
    try:
        Path(tmp.name).replace(path)
    finally:
        Path(tmp.name).unlink(missing_ok=True)


def canonical(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"),
                      allow_nan=False).encode("utf-8")


def digest(value):
    return sha256(canonical(value)).hexdigest()


def file_sha(path):
    with Path(path).open("rb") as handle:
        return file_digest(handle, "sha256").hexdigest()


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def load_module(name, path):
    path = Path(path).resolve()
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, "无法加载已登记的审核脚本")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    if str(path.parent) not in sys.path:
        sys.path.insert(0, str(path.parent))
    spec.loader.exec_module(module)
    return module


def relname(value):
    require(isinstance(value, str) and value and not any(ch in value for ch in '\\:*?"<>|'),
            "存在无效的相对文件名")
    parts = value.split("/")
    require(all(p and p not in (".", "..") and not p.endswith((" ", ".")) for p in parts),
            "存在无效的相对路径")
    for part in parts:
        require(all(ord(ch) >= 32 for ch in part) and
                not re.fullmatch(r"(?i)(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])",
                                 part.split(".")[0]), "路径含 Windows 保留名称")
    return value


def inside(root, name):
    root = Path(root).resolve()
    ref = root.joinpath(*PurePosixPath(relname(name)).parts)
    require(ref.resolve().is_relative_to(root) and not ref.is_symlink(),
            "输入文件须位于本批冻结目录")
    return ref


def files_in(root):
    root = Path(root)
    entries = {}
    for path in sorted(root.rglob("*")):
        require(not path.is_symlink(), "批次目录不得含符号链接")
        if path.is_file():
            name = relname(path.relative_to(root).as_posix())
            entries[name] = file_sha(path)
    require(len({x.casefold() for x in entries}) == len(entries), "存在 Windows 下同名文件")
    return entries


def archive(path, record_name, record, sources):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    require(not path.exists(), "输出已经存在，使用新文件名")
    snapshots = record.get("files", record.get("body", {}).get("files"))
    with tempfile.NamedTemporaryFile(dir=path.parent, delete=False) as tmp:
        staging = Path(tmp.name)
    try:
        with zipfile.ZipFile(staging, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr(record_name, canonical(record))
            for name, src in sorted(sources.items()):
                data = Path(src).read_bytes()
                require(sha256(data).hexdigest() == snapshots[name], "打包期间文件发生变化")
                z.writestr("payload/" + relname(name), data)
        os.link(staging, path)
    finally:
        staging.unlink(missing_ok=True)


def unpack(path, record_name, dest):
    dest = Path(dest)
    require(not dest.exists() or not any(dest.iterdir()), "冻结文件需解包到空目录")
    with zipfile.ZipFile(path) as z:
        names = [relname(n) for n in z.namelist()]
        require(len(names) == len({n.casefold() for n in names}), "压缩包条目重名")
        record = json.loads(z.read(record_name))
        entries = record.get("files", record.get("body", {}).get("files"))
        require(isinstance(entries, dict) and set(names) == {
            record_name, *("payload/" + relname(n) for n in entries)}, "压缩包清单和内容不一致")
        for name, expected in entries.items():
            target = inside(dest, name)
            data = z.read("payload/" + name)
            require(sha256(data).hexdigest() == expected, "冻结文件校验值不一致")
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
    return record


def profile(root):
    root = Path(root)
    data = read(root / PROFILE_NAME)
    files = data["audit_files"]
    require(isinstance(files, dict) and files and all(
        file_sha(inside(root, name)) == expected for name, expected in files.items()),
        "审核程序文件与登记版本不一致")
    require(data["audit_sha256"] == digest({k: v for k, v in data.items() if k != "audit_sha256"}),
            "审核版本指纹不一致")
    return data


def channel(trust):
    return {key: trust[key] for key in ("repository", "branch", "public_key")}


def verify_task(task, trust, registered):
    require(task["schema"] == 1 and re.fullmatch(r"[0-9a-f]{32}", task["batch_id"]) and
            isinstance(task["project_id"], str) and bool(task["project_id"].strip()),
            "任务单的批次或项目无效")
    require(task["channel"] == channel(trust), "任务单来自不同的 Word 审核通道")
    version = task["skill_version"]
    require(version == registered["skill_version"] and
            task["audit_sha256"] == registered["audit_sha256"] == trust["profiles"].get(version) and
            task["submission"] in registered["submission_modes"], "编制版本、审稿版本或交付方式未配对")


def signed(body, encoded_key, trust):
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

    key = Ed25519PrivateKey.from_private_bytes(base64.b64decode(encoded_key, validate=True))
    require(key.public_key().public_bytes_raw() == base64.b64decode(trust["public_key"], validate=True),
            "运行器签名身份与生产配置不一致")
    return {"body": body, "signature": base64.b64encode(key.sign(canonical(body))).decode("ascii")}


def verify_receipt(envelope, task, trust, phase):
    from cryptography.exceptions import InvalidSignature
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

    verify = Ed25519PublicKey.from_public_bytes(base64.b64decode(trust["public_key"], validate=True))
    body = envelope["body"]
    try:
        verify.verify(base64.b64decode(envelope["signature"], validate=True), canonical(body))
    except (InvalidSignature, ValueError) as exc:
        raise GateError("Word 或交付回执签名无效") from exc
    run = body["runner"]
    require(body["task_sha256"] == digest(task) and body["phase"] == phase and
            body["audit_sha256"] == task["audit_sha256"] and
            run["repository"] == trust["repository"] and run["branch"] == trust["branch"] and
            str(run["id"]).isdigit() and str(run["attempt"]).isdigit() and
            re.fullmatch(r"[0-9a-f]{40}", run["commit"]), "回执与本项目批次、审核版本或运行通道不匹配")
    return body


def create_probe(path, marker):
    with zipfile.ZipFile(path, "x", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '</Relationships>')
        z.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                   f'<w:body><w:p><w:r><w:t>{marker}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')


def start(project, project_id, submission, trust_path, root):
    project = Path(project)
    data = profile(root)
    trust = read(trust_path)
    project.mkdir(parents=True, exist_ok=True)
    trust_record = channel(trust) | {"profiles": {data["skill_version"]: trust["profiles"].get(data["skill_version"])}}
    if "gateway" in trust:
        dispatch_target(trust)
        trust_record["gateway"] = trust["gateway"]
    task_path = project / TASK
    if task_path.exists():
        task = read(task_path)
        verify_task(task, trust_record, data)
        require((task["project_id"], task["submission"]) == (project_id, submission), "已有任务单，按原批次续办")
    else:
        task = {"schema": 1, "project_id": project_id, "batch_id": uuid4().hex, "created_at": now(),
                "skill_version": data["skill_version"], "audit_sha256": data["audit_sha256"],
                "channel": channel(trust), "submission": submission}
        verify_task(task, trust_record, data)
        require(not (project / PROBE).exists(), "目录已有试排文件，确认原批次后接续")
        create_probe(project / PROBE, "WORD-CHANNEL-" + task["batch_id"])
        task["probe_sha256"] = file_sha(project / PROBE)
        write(task_path, task)
    write(project / TRUST, trust_record)
    state_file = project / STATE
    if not state_file.exists():
        flow = load_module("_initial_flow", Path(root) / "scripts" / "编制流程门禁.py")
        flow.PAGE_DIR = project
        flow.cmd_init(state_file, False)
    state = read(state_file)
    previous = state["startup"].get("versioned_task_sha256")
    require(previous is None or previous == digest(task), "旧状态已绑定另一批，先保留原记录")
    state["startup"]["versioned_task_sha256"] = digest(task)
    write(state_file, state)
    return task


def check_probe(project, task, trust):
    proof = verify_receipt(read(project / EVIDENCE / "probe" / "receipt.json"), task, trust, "probe")
    item = proof.get("measurement", {})
    require(proof["status"] == "PASS" and item.get("engine") == "Microsoft Word" and
            isinstance(item.get("word_version"), str) and bool(item["word_version"]) and
            type(item.get("pages")) is int and item["pages"] >= 1 and
            item.get("document_sha256") == task["probe_sha256"] == file_sha(project / PROBE),
            "缺本批实际私密收发路径的 Windows Word 样稿实排证据")
    return proof


def startup_errors(state, project, root):
    if "versioned_task_sha256" not in (state.get("startup") or {}) and not (project / TASK).is_file():
        return []
    try:
        task, trust = read(project / TASK), read(project / TRUST)
        verify_task(task, trust, profile(root))
        require(state["startup"]["versioned_task_sha256"] == digest(task), "编前状态与任务单不一致")
        check_probe(project, task, trust)
    except (GateError, OSError, ValueError, KeyError, TypeError, ImportError) as exc:
        return ["本批 Word 通道未通过实际试跑。 " + str(exc)]
    return []


def reference_allowed(project, document_sha256, report, stage=None):
    """Reference-only limits are allowed only with a signed PASS for the ten real files."""
    if report.get("reference_only") is not True:
        return False
    try:
        task, trust = read(project / TASK), read(project / TRUST)
        require(task["submission"] == "review-points", "普通合册不得用参考模式放行")
        phase = "measure" if stage is None else f"stage-{stage}"
        input_name = "design.docx" if stage is None else phase + ".docx"
        folder = project / EVIDENCE / phase
        proof = verify_receipt(read(folder / "receipt.json"), task, trust, phase)
        require(proof["status"] == "PASS" and proof["inputs"].get(input_name) == document_sha256
                and proof["files"].get("pages.json") == file_sha(folder / "pages.json")
                and read(folder / "pages.json") == report, "已签名实排证据与参考合册不一致")
        if stage is None:
            require(proof["files"].get("part-pages.json") == file_sha(folder / "part-pages.json")
                    and read(folder / "part-pages.json")["status"] == "PASS", "缺十份实际稿的实排证据")
        return True
    except (GateError, OSError, ValueError, KeyError, TypeError, ImportError):
        return False


def published(project, state):
    try:
        task, trust = read(project / TASK), read(project / TRUST)
        require(state["startup"]["versioned_task_sha256"] == digest(task), "状态与任务单不一致")
        result = verify_receipt(read(project / EVIDENCE / "publish" / "receipt.json"), task, trust, "publish")
        require(result["formal_release"] is True and result["task_complete"] is True and
                result["status"] == "PASS" and result["formal_url"], "缺正式发布回执")
        require(all(file_sha(inside(project, item["path"])) == item["sha256"] for item in result["submission"]),
                "待发稿改动后须重新验收")
        return True
    except (GateError, OSError, ValueError, KeyError, TypeError, ImportError):
        return False


def ast_with_constants(path):
    tree = ast.parse(Path(path).read_text(encoding="utf-8-sig"))
    values = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            name = node.targets[0].id
            if name in PAGE_NAMES:
                require(name not in values, "页数规则常量存在多次赋值")
                values[name] = ast.literal_eval(node.value)
                node.value = ast.Constant(None)
    return ast.dump(tree, include_attributes=False), values


def page_script(project, root, registered):
    choices = sorted(Path(project).glob("*页数硬审*.py"))
    require(len(choices) <= 1, "本批须有唯一项目页数脚本")
    if not choices:
        require(registered["skill_version"] == "r22.6", "逐点评审任务缺项目页数脚本")
        return None
    supplied, values = ast_with_constants(choices[0])
    template, defaults = ast_with_constants(Path(root) / registered["page_template"])
    require(supplied == template and set(values) == set(defaults),
            "项目页数脚本只能按招标调整已登记版本的页数常量")
    require(all(type(values[key]) is int for key in values if key not in ("REVIEW_POINT_HEADINGS", "PROJECT_NAME")),
            "页数常量须为整数")
    if "PROJECT_NAME" in values:
        require(isinstance(values["PROJECT_NAME"], str) and values["PROJECT_NAME"].strip(), "项目显示名无效")
    if registered["skill_version"] == "r23.6":
        require(values["REVIEW_POINT_COUNT"] == 10 and 0 < values["REVIEW_POINT_PAGE_LIMIT"] <= 10 and
                0 < values["EXPECTED_BODY_PAGES"] <= 100, "r23.6 的十点评审上限超出本版本能力")
    return choices[0]


def documents_for_phase(project, task, phase):
    if phase == "probe":
        return [PROBE]
    if phase == "pilot":
        return ["pilot.docx"]
    if phase.startswith("stage-"):
        return [phase + ".docx"]
    names = ["design.docx", "format_rules.json"]
    if task["submission"] == "review-points":
        names.append("parts.json")
        specs = read(project / "parts.json")["parts"]
        names.extend(relname(row["docx"]) for row in specs)
    return names


def prepare(project, phase, path, root):
    project = Path(project).resolve()
    task, trust = read(project / TASK), read(project / TRUST)
    registered = profile(root)
    verify_task(task, trust, registered)
    paths = [TASK, *documents_for_phase(project, task, phase)]
    if phase != "probe":
        check_probe(project, task, trust)
        paths.extend([STATE, TRUST, PROBE, EVIDENCE + "/probe/receipt.json"])
        script = page_script(project, root, registered)
        if script:
            paths.append(script.name)
    if phase == "publish":
        require(not Path(path).resolve().is_relative_to(project), "整批冻结压缩包须置于本批目录之外")
        paths = [n for n in files_in(project) if not n.endswith((".pyc", ".pyo"))
                 and "__pycache__" not in Path(n).parts and n != "流程交付结果.json"]
    require(len(paths) == len(set(paths)), "批次文件清单含重复条目")
    sources = {name: inside(project, name) for name in paths}
    record = {"schema": 1, "phase": phase, "task": task, "request_id": uuid4().hex,
              "files": {name: file_sha(file) for name, file in sources.items()}}
    archive(path, "request.json", record, sources)
    return record


def accept(project, packet, expected_request):
    project = Path(project)
    task, trust = read(project / TASK), read(project / TRUST)
    with tempfile.TemporaryDirectory() as directory:
        dest = Path(directory)
        envelope = unpack(packet, "receipt.json", dest)
        phase = envelope["body"]["phase"]
        require(phase in PHASES, "回执阶段无效")
        proof = verify_receipt(envelope, task, trust, phase)
        require(proof["request_id"] == expected_request, "收回的回执与本次 GH 调用不一致")
        for name, expected in proof["inputs"].items():
            require(file_sha(inside(project, name)) == expected, "回执对应的稿件或规则在本地已经变化")
        out = project / EVIDENCE / phase
        # No stray file from a previous measurement can count as current evidence.
        if out.exists():
            moved = out.with_name(phase + "-" + uuid4().hex)
            out.rename(moved)
        dest_out = out.parent / (phase + ".incoming-" + uuid4().hex)
        import shutil
        shutil.copytree(dest, dest_out)
        write(dest_out / "receipt.json", envelope)
        dest_out.rename(out)
        if phase == "probe" and proof["status"] == "PASS":
            check_probe(project, task, trust)
            state = read(project / STATE)
            require(state["startup"].get("versioned_task_sha256") == digest(task), "编前状态与任务单不一致")
            state["startup"].update(word_channel="github_runner",
                                    word_capability_evidence="release-evidence/probe/receipt.json")
            write(project / STATE, state)
    return proof


def dispatch_target(trust):
    gateway = trust.get("gateway")
    route = gateway if gateway is not None else trust
    repository, branch = route["repository"], route["branch"]
    require(re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository) and bool(branch.strip()),
            "GitHub 送审入口仓库或分支无效")
    require(gateway is None or repository.casefold() != trust["repository"].casefold(),
            "公开送审入口与私有审核须使用不同的仓库")
    return repository, branch, PRODUCTION_WORKFLOW if gateway is None else GATEWAY_WORKFLOW


def dispatch(project, packet, file_id):
    repository, branch, workflow = dispatch_target(read(Path(project) / TRUST))
    require(re.fullmatch(r"[A-Za-z0-9_-]+", file_id), "Drive 文件 ID 无效")
    data = {"ref": branch, "inputs": {"drive_file": file_id, "packet_sha256": file_sha(packet)}}
    req = Request(f"https://api.github.com/repos/{repository}/actions/workflows/{workflow}/dispatches",
                  data=canonical(data), method="POST",
                  headers={"Authorization": "Bearer " + os.environ["CONSTRUCTION_GH_DISPATCH_TOKEN"],
                           "Accept": "application/vnd.github+json", "Content-Type": "application/json"})
    try:
        with urlopen(req, timeout=45) as response:
            require(response.status in (200, 204), "GitHub 未接受本批送审任务")
    except HTTPError as exc:
        raise GateError(f"GitHub 送审请求被拒绝，HTTP {exc.code}") from exc


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", required=True, type=Path)
    sub = parser.add_subparsers(dest="action", required=True)
    begin = sub.add_parser("start")
    begin.add_argument("--project-id", required=True)
    begin.add_argument("--submission", required=True, choices=("whole", "review-points"))
    begin.add_argument("--trust", required=True, type=Path)
    build = sub.add_parser("prepare")
    build.add_argument("phase", choices=PHASES)
    build.add_argument("packet", type=Path)
    send = sub.add_parser("dispatch")
    send.add_argument("packet", type=Path)
    send.add_argument("drive_file_id")
    get = sub.add_parser("accept")
    get.add_argument("packet", type=Path)
    get.add_argument("--request-id", required=True)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    if args.action == "start":
        task = start(args.project, args.project_id, args.submission, args.trust, root)
        print("BATCH=" + task["batch_id"] + " SKILL=" + task["skill_version"])
    elif args.action == "prepare":
        result = prepare(args.project, args.phase, args.packet, root)
        print("REQUEST=" + result["request_id"])
    elif args.action == "dispatch":
        dispatch(args.project, args.packet, args.drive_file_id)
        print("DISPATCH=ACCEPTED")
    else:
        result = accept(args.project, args.packet, args.request_id)
        print(json.dumps({k: result.get(k) for k in ("phase", "status", "formal_release", "formal_url")}))
        return 0 if result["status"] == "PASS" else 2
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (GateError, OSError, ValueError, KeyError, TypeError, ImportError) as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(2)

