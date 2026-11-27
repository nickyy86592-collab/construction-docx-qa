"""Regression tests for version dispatch, receipts, frozen Word artifacts and publication."""
from copy import deepcopy
from pathlib import Path
import base64
import json
import os
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import MagicMock, patch
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "production"))
sys.path.insert(1, str(ROOT / "scripts"))

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from drive_release import DriveError, check_access, publish
from forward_request import forward
from register import register
from release_task import (EVIDENCE, GATEWAY_WORKFLOW, PROBE, PRODUCTION_WORKFLOW, STATE, TASK, TRUST,
                          GateError, accept, archive, check_probe, digest, dispatch, file_sha, page_script, prepare, profile, read,
                          signed, start, unpack, verify_receipt, verify_task, write)
from runner_release import context


class MockDrive:
    def __init__(self):
        self.objects = {name: {"id": name, "mimeType": "application/vnd.google-apps.folder",
                               "name": name, "parents": []} for name in ("formal", "staging")}
        self.data = {}
        self.moves = 0
        self.reject = None

    def children(self, parent):
        return [x for x in self.objects.values() if parent in x["parents"]]

    def create(self, name, parent, media=None, props=None):
        if name == self.reject:
            raise DriveError(503)
        key = "id" + str(len(self.objects))
        result = self.objects[key] = {"id": key, "name": name, "parents": [parent],
                   "mimeType": "application/vnd.google-apps.folder" if media is None else "application/octet-stream",
                   "appProperties": props or {}}
        if media is not None:
            self.data[key] = media
        return result

    def download(self, key):
        return self.data[key]

    def move(self, key, source, target):
        assert self.objects[key]["parents"] == [source]
        self.objects[key]["parents"] = [target]
        self.moves += 1

    def meta(self, key):
        return self.objects[key]


class WorkflowTest(unittest.TestCase):
    def setUp(self):
        self.work = TemporaryDirectory()
        self.addCleanup(self.work.cleanup)
        self.root = Path(self.work.name)
        self.project = self.root / "project"
        self.program = profile(ROOT)
        key = Ed25519PrivateKey.generate()
        self.secret = base64.b64encode(key.private_bytes_raw()).decode()
        self.trust = {"repository": "owner/private", "branch": "main",
                      "public_key": base64.b64encode(key.public_key().public_bytes_raw()).decode(),
                      "profiles": {self.program["skill_version"]: self.program["audit_sha256"]}}
        write(self.root / TRUST, self.trust)
        self.task = start(self.project, "示例项目", "whole", self.root / TRUST, ROOT)

    def prepared_probe(self):
        request = prepare(self.project, "probe", self.root / "probe.zip", ROOT)
        body = {"status": "PASS", "phase": "probe", "runner": {
            "repository": "owner/private", "branch": "main", "id": "900", "attempt": "1", "commit": "c" * 40,
        }, "request_id": request["request_id"], "inputs": request["files"], "files": {},
            "task_sha256": digest(self.task), "audit_sha256": self.program["audit_sha256"],
            "measurement": {"engine": "Microsoft Word", "word_version": "16.0", "pages": 1,
                            "document_sha256": file_sha(self.project / PROBE)}}
        return request, body

    def test_same_task_automatically_attaches_exact_version(self):
        self.assertEqual(self.task["skill_version"], self.program["skill_version"])
        self.assertEqual(read(self.project / STATE)["startup"]["versioned_task_sha256"], digest(self.task))
        self.assertEqual(start(self.project, "示例项目", "whole", self.root / TRUST, ROOT), self.task)
        with self.assertRaises(GateError):
            start(self.project, "另一个项目", "whole", self.root / TRUST, ROOT)
        with self.assertRaises(GateError):
            verify_task(self.task | {"skill_version": "r99"}, self.trust, self.program)
        self.assertNotEqual(start(self.root / "other", "同名可另建项目", "whole", self.root / TRUST, ROOT)["batch_id"],
                            self.task["batch_id"])

    def test_report_without_native_proof_fails_even_with_ready_text(self):
        from release_task import startup_errors

        state = read(self.project / STATE)
        state["startup"].update(word_channel="github_runner", word_capability_evidence="运行器开着")
        self.assertTrue(startup_errors(state, self.project, ROOT))
        request, body = self.prepared_probe()
        body["measurement"]["engine"] = "LibreOffice"
        packet = self.root / "not-word.zip"
        archive(packet, "receipt.json", signed(body, self.secret, self.trust), {})
        with self.assertRaises(GateError):
            accept(self.project, packet, request["request_id"])

    def test_accepts_a_signed_native_receipt_and_rejects_reused_or_changed_proof(self):
        request, body = self.prepared_probe()
        packet = self.root / "receipt.zip"
        proof = signed(body, self.secret, self.trust)
        archive(packet, "receipt.json", proof, {})
        with self.assertRaises(GateError):
            accept(self.project, packet, "mistaken-request")
        with self.assertRaises(GateError):
            verify_receipt(proof, self.task | {"batch_id": "0" * 32}, self.trust, "probe")
        altered = deepcopy(proof)
        altered["body"]["measurement"]["pages"] = 2
        with self.assertRaises(GateError):
            verify_receipt(altered, self.task, self.trust, "probe")
        accept(self.project, packet, request["request_id"])
        self.assertEqual(check_probe(self.project, self.task, self.trust)["status"], "PASS")
        self.assertEqual(read(self.project / STATE)["startup"]["word_channel"], "github_runner")
        (self.project / PROBE).write_bytes(b"changed")
        with self.assertRaises(GateError):
            check_probe(self.project, self.task, self.trust)

    def test_requires_private_approved_production_entry(self):
        fields = {"CONSTRUCTION_PRIVATE_REPOSITORY": "true", "GITHUB_REPOSITORY": "owner/private",
                  "GITHUB_REF_NAME": "main", "GITHUB_RUN_ID": "2", "GITHUB_RUN_ATTEMPT": "1",
                  "GITHUB_SHA": "f" * 40, "GITHUB_EVENT_NAME": "workflow_dispatch"}
        with patch.dict(os.environ, fields, clear=True):
            self.assertEqual(context(self.trust)["id"], "2")
        for changed in ({"CONSTRUCTION_PRIVATE_REPOSITORY": "false"}, {"GITHUB_REF_NAME": "other"},
                        {"GITHUB_EVENT_NAME": "push"}):
            with patch.dict(os.environ, fields | changed, clear=True), self.assertRaises(GateError):
                context(self.trust)

    def test_public_gateway_routes_the_same_frozen_file_to_private_production(self):
        packet = self.root / "probe.zip"
        request = prepare(self.project, "probe", packet, ROOT)
        gateway = {"repository": "owner/public", "branch": "main"}
        write(self.root / TRUST, self.trust | {"gateway": gateway})
        self.assertEqual(start(self.project, "示例项目", "whole", self.root / TRUST, ROOT), self.task)
        self.assertEqual(read(self.project / TRUST)["gateway"], gateway)

        with patch.dict(os.environ, {"CONSTRUCTION_GH_DISPATCH_TOKEN": "mock-gateway-token"}), \
                patch("release_task.urlopen") as send:
            send.return_value.__enter__.return_value.status = 204
            dispatch(self.project, packet, "opaque-file-id")
        submission = send.call_args.args[0]
        self.assertIn("/repos/owner/public/", submission.full_url)
        self.assertIn(GATEWAY_WORKFLOW, submission.full_url)
        payload = json.loads(submission.data)
        self.assertEqual(payload, {"ref": "main", "inputs": {"drive_file": "opaque-file-id", "packet_sha256": file_sha(packet)}})
        self.assertEqual(request["task"]["channel"]["repository"], "owner/private")

        env = {"CONSTRUCTION_REPOSITORY_PRIVATE": "false", "GITHUB_EVENT_NAME": "workflow_dispatch",
               "GITHUB_REPOSITORY": "owner/public", "GITHUB_REF": "refs/heads/main",
               "CONSTRUCTION_DEFAULT_REF": "refs/heads/main", "CONSTRUCTION_PRODUCTION_REPO": "owner/private",
               "CONSTRUCTION_PRODUCTION_BRANCH": "secure", "CONSTRUCTION_FORWARD_TOKEN": "mock-private-token",
               "CONSTRUCTION_REQUEST_FILE": "opaque-file-id", "CONSTRUCTION_REQUEST_HASH": file_sha(packet)}
        trust_send = MagicMock()
        trust_send.return_value.__enter__.return_value.status = 204
        forward(env, trust_send)
        forwarded = trust_send.call_args.args[0]
        self.assertIn("/repos/owner/private/", forwarded.full_url)
        self.assertIn(PRODUCTION_WORKFLOW, forwarded.full_url)
        self.assertEqual(json.loads(forwarded.data), payload | {"ref": "secure"})
        for change in ({"GITHUB_REF": "refs/heads/other"}, {"GITHUB_EVENT_NAME": "pull_request"},
                       {"CONSTRUCTION_REQUEST_HASH": "wrong"}, {"CONSTRUCTION_FORWARD_TOKEN": ""}):
            with self.assertRaises(ValueError):
                forward(env | change, trust_send)
        trust_send.assert_called_once()

    def test_register_public_gateway_keeps_production_configuration_private(self):
        private_checkout = self.root / "private"
        public_checkout = self.root / "public"
        for path in (private_checkout, public_checkout):
            path.mkdir()
            (path / ".git").mkdir()
        settings = self.trust | {"gateway": {"repository": "owner/public", "branch": "main"},
                                "drive": {name: name for name in ("inbox", "evidence", "staging", "formal")}}
        write(self.root / TRUST, settings)
        self.assertEqual(register([ROOT], private_checkout, self.root / TRUST, public_checkout), self.trust["profiles"])
        self.assertEqual({path.relative_to(public_checkout).as_posix() for path in public_checkout.rglob("*") if path.is_file()},
                         {f".github/workflows/{GATEWAY_WORKFLOW}", ".github/construction/forward_request.py"})
        self.assertEqual(read(private_checkout / "review-runtime" / TRUST), settings)

    def test_project_page_template_rejects_executable_change(self):
        source = ROOT / self.program["page_template"]
        copy = self.project / source.name
        copy.write_bytes(source.read_bytes())
        self.assertEqual(page_script(self.project, ROOT, self.program), copy)
        copy.write_text(copy.read_text() + "\nprint('unregistered code')\n")
        with self.assertRaises(GateError):
            page_script(self.project, ROOT, self.program)

    def test_cli_uses_the_project_page_rule_after_task_start(self):
        template = ROOT / self.program["page_template"]
        path = self.project / template.name
        changed = template.read_text().replace("EXPECTED_BODY_PAGES = 50", "EXPECTED_BODY_PAGES = 46")
        changed = changed.replace("EXPECTED_BODY_PAGES = 100", "EXPECTED_BODY_PAGES = 89")
        path.write_text(changed)
        exe = ROOT / "scripts" / "编制流程门禁.py"
        output = subprocess.run([sys.executable, str(exe), "--state", str(self.project / STATE), "page-rule"],
                                capture_output=True, text=True, check=True).stdout
        self.assertIn(str(path.resolve()), output)
        self.assertNotIn(str((ROOT / self.program["page_template"]).resolve()), output)

    def test_author_upload_role_and_read_only_formal_target(self):
        from types import SimpleNamespace
        acl = [{"id": "pub", "type": "user", "role": "owner"}]
        formal = {"canAddChildren": False, "canEdit": False, "canShare": False}
        author = SimpleNamespace(uid=lambda: "author", meta=lambda name: {
            "mimeType": "application/vnd.google-apps.folder",
            "capabilities": formal if name == "formal" else {"canAddChildren": True}})
        owner = SimpleNamespace(uid=lambda: "pub", meta=lambda name: {
            "mimeType": "application/vnd.google-apps.folder", "capabilities": {"canAddChildren": True}},
            private=lambda name: acl)
        folders = {n: n for n in ("inbox", "evidence", "staging", "formal")}
        self.assertEqual(check_access(author, owner, {"drive": folders}), folders)
        formal["canAddChildren"] = True
        with self.assertRaises(GateError):
            check_access(author, owner, {"drive": folders})

    def test_unpacked_payload_must_match_its_immutable_hash(self):
        request = prepare(self.project, "probe", self.root / "probe.zip", ROOT)
        bad = self.root / "altered.zip"
        with ZipFile(self.root / "probe.zip") as original, ZipFile(bad, "w") as new:
            for name in original.namelist():
                new.writestr(name, b"altered" if name == "payload/word-probe.docx" else original.read(name))
        with self.assertRaises(GateError):
            unpack(bad, "request.json", self.root / "other")

    def test_upload_is_private_until_all_files_match(self):
        drive = MockDrive()
        paths = {}
        for n in range(1, 11):
            path = self.root / f"{n}.docx"
            path.write_bytes(str(n).encode())
            paths[path.name] = (path, file_sha(path))
        drive.reject = "10.docx"
        folders = {"formal": "formal", "staging": "staging"}
        with self.assertRaises(DriveError):
            publish(drive, folders, "release", "batch", paths)
        self.assertEqual(drive.children("formal"), [])
        drive.reject = None
        url = publish(drive, folders, "release", "batch", paths)
        self.assertTrue(url.startswith("https://drive.google.com/"))
        self.assertEqual(drive.moves, 1)
        endpoint = drive.children("formal")
        self.assertEqual(len(endpoint), 1)
        self.assertEqual(len(drive.children(endpoint[0]["id"])), 10)
        self.assertEqual(url, publish(drive, folders, "release", "batch", paths))
        self.assertEqual(drive.moves, 1)
        paths["10.docx"][0].write_bytes(b"edited")
        with self.assertRaises(GateError):
            publish(drive, folders, "release", "batch", paths)
        self.assertEqual(drive.moves, 1)


@unittest.skipUnless((ROOT / "scripts" / "split_submission.py").is_file(), "r23.6 逐点专用")
class PointTest(unittest.TestCase):
    def test_100_total_10_per_point_scope_is_verified_separately(self):
        from split_submission import page_failures
        rows = [{"index": n, "docx": f"{n}.docx", "word_version": "16.0", "cover_pages": 1,
                 "physical_pages": 11, "non_cover_pages": 10, "failures": []} for n in range(1, 11)]
        report = {"page_source": "microsoft_word", "parts": rows, "non_cover_pages": 100}
        rule = {"non_cover_page_limit": 100, "review_point_page_limit": 10}
        self.assertEqual(page_failures(report, rule), [])
        rows[0].update(physical_pages=12, non_cover_pages=11)
        rows[1].update(physical_pages=10, non_cover_pages=9)
        self.assertEqual(len(page_failures(report, rule)), 1)
        rows.pop()
        self.assertTrue(page_failures(report, rule))

    def test_scope_and_table_rules_follow_individually_submitted_parts(self):
        from split_submission import cover_allowed, point_format

        self.assertTrue(cover_allowed("每个评审点不超过10页（不含封面）"))
        self.assertTrue(cover_allowed("每点评审文件封面不计入页数"))
        for quote in ("只允许按时上传", "封面不得扣除", "不扣除封面", "每个评审点封面计入总页数，不得扣除",
                      "统一封面不计入页数"):
            self.assertFalse(cover_allowed(quote), quote)
        mapping = {2: 1}
        rules = {"front_matter": {"toc_levels": [1]}, "table_layout": {
            "tables": [{"index": 1, "role": "layout"}, {"index": 2, "role": "data"}]}}
        part = {"front_pages": 0, "cover_pages": 0, "heading": "第二章 有效内容", "source_ref": "P18"}
        scoped = point_format(rules, part, mapping)
        self.assertEqual(scoped["table_layout"]["tables"], [{"index": 1, "role": "data"}])
        self.assertEqual(rules["table_layout"]["tables"][1]["index"], 2)

    def test_real_docx_sections_preserve_content_after_independent_renumbering(self):
        from docx import Document
        from release_task import load_module
        from split_submission import compare_content

        common = load_module("_new_split_fmt", ROOT / "scripts" / "Word正式交付硬审（不含页数）.py")
        with TemporaryDirectory() as work:
            project = Path(work)
            headings, specs = [], []
            master = Document()
            for index in range(1, 11):
                headline = f"第{index}章 评审点 {index}"
                own = f"第一章 评审点 {index}"
                headings.append(headline)
                doc = Document()
                for content, name in [(master, headline), (doc, own)]:
                    content.add_heading(name, 1)
                    content.add_heading("第一节 具体工作", 2)
                    content.add_heading("一、控制方法", 3)
                    content.add_paragraph(f"第 {index} 项已有确认的施工组织内容。")
                    t = content.add_table(1, 2)
                    t.cell(0, 0).text = "数量"
                    t.cell(0, 1).text = str(10 + index)
                doc.save(project / f"{index}.docx")
                specs.append({"index": index, "docx": f"{index}.docx", "heading": own})
            master.save(project / "design.docx")
            rule = {"review_point_headings": headings}
            compare_content(project, specs, rule, common, {})
            actual = Document(project / "5.docx")
            actual.tables[0].cell(0, 1).text = "151"
            actual.save(project / "5.docx")
            with self.assertRaisesRegex(GateError, "第 5 项"):
                compare_content(project, specs, rule, common, {})


if __name__ == "__main__":
    unittest.main()

