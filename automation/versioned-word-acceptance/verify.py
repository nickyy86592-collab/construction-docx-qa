"""Verify the registered releases using synthetic documents and actual Windows Word."""
from contextlib import redirect_stdout
import io
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from xml.sax.saxutils import escape
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[2]
PACKAGES = {
    version: ROOT / "skills" / f"general-construction-organization-writing-r{version.replace('.', '-')}"
    for version in ("22.6", "23.6")
}
sys.path.insert(0, str(PACKAGES["22.6"] / "production"))

from register import register
from release_task import GateError, file_sha, load_module, prepare, read, start, unpack, write
from runner_release import execute, select
from word_native import open_word


def fixture(path, pages, heading=None, cover=False):
    paragraphs = []
    for index in range(pages + int(cover)):
        is_heading = heading is not None and index == int(cover)
        props = '<w:pageBreakBefore/>' if index else ''
        if is_heading:
            props += '<w:outlineLvl w:val="0"/>'
        text = heading if is_heading else f"Synthetic Word page {index + 1}"
        paragraphs.append(f"<w:p><w:pPr>{props}</w:pPr><w:r><w:t>{escape(text)}</w:t></w:r></w:p>")
        if is_heading:
            paragraphs.append("<w:p><w:r><w:t>Synthetic review content</w:t></w:r></w:p>")
    with ZipFile(path, "w") as docx:
        docx.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                      '</Types>')
        docx.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                      '</Relationships>')
        docx.writestr("word/document.xml", '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                      '<w:body>' + ''.join(paragraphs) + '<w:sectPr/></w:body></w:document>')


def probe_version(work, runtime, settings_file, settings, version, mode):
    root = PACKAGES[version]
    project = work / ("project-" + version)
    task = start(project, "synthetic-word-test", mode, settings_file, root)
    packet = work / (version + ".zip")
    prepare(project, "probe", packet, root)
    frozen = work / ("frozen-" + version)
    request = unpack(packet, "request.json", frozen)
    selected, profile, _ = select(runtime, request, frozen, settings)
    assert task["skill_version"] == "r" + version and selected.name == task["skill_version"]
    try:
        select(runtime, request | {"task": task | {"audit_sha256": "0" * 64}}, frozen, settings)
    except GateError:
        pass
    else:
        raise AssertionError("Wrong audit hash was accepted")
    measurement = execute(runtime, frozen, work, request, settings, None, None)["measurement"]
    assert measurement["engine"] == "Microsoft Word" and measurement["pages"] == 1, measurement
    assert measurement["document_sha256"] == task["probe_sha256"]
    print(f"{task['skill_version']} selected {profile['audit_sha256']} and opened the frozen probe in Word {measurement['word_version']}", flush=True)
    return selected / profile["page_template"]


def verify_whole(work, gate):
    for actual, expected in ((51, 0), (50, 2)):
        docx = work / f"whole-{actual}.docx"
        report = work / "whole-pages.json"
        fixture(docx, actual)
        with redirect_stdout(io.StringIO()):
            status = gate.audit(docx, report)
        proof = read(report)
        assert status == expected and proof["page_source"] == "microsoft_word", proof
        assert proof["actual"]["word_total_pages"] == actual and proof["document_sha256"] == file_sha(docx), proof
        print(f"r22.6 actual Word total={actual}, page gate={'PASS' if status == 0 else 'BLOCK'}", flush=True)


def verify_points(work, registered, gate):
    split = load_module("_smoke_split", registered.parent / "split_submission.py")
    gate.REVIEW_POINT_HEADINGS = [f"Synthetic review point {index}" for index in range(1, 11)]
    rule = gate.rule_snapshot()

    def measure(index, pages):
        heading = rule["review_point_headings"][index - 1]
        docx = work / f"point-{index:02d}.docx"
        fixture(docx, pages, heading, cover=True)
        with open_word(docx) as (word, doc):
            physical = int(doc.ComputeStatistics(2))
            actual, failures = gate.measure_points(doc, rule | {"review_point_headings": [heading]})
            assert physical == pages + 1 and len(actual) == 1 and actual[0]["pages"] == pages, (actual, physical)
            assert actual[0]["start_page"] == 2 and actual[0]["heading"] == heading, actual
            return {"index": index, "docx": docx.name, "physical_pages": physical, "cover_pages": 1,
                    "non_cover_pages": physical - 1, "word_version": str(word.Version), "failures": failures}

    def report(parts):
        return {"page_source": "microsoft_word", "parts": parts,
                "non_cover_pages": sum(row["non_cover_pages"] for row in parts)}

    point_rows = [measure(index, 10) for index in range(1, 11)]
    assert not split.page_failures(report(point_rows), rule), point_rows
    print("r23.6 actual ten Word files, non-cover pages=100, page gate=PASS", flush=True)
    apart = [measure(1, 11), measure(2, 8), *point_rows[2:]]
    assert report(apart)["non_cover_pages"] == 99 and split.page_failures(report(apart), rule), apart
    assert split.page_failures(report(point_rows[:-1]), rule)
    print("r23.6 actual Word point 1=11, non-cover total=99, page gate=BLOCK. Missing part also blocked.", flush=True)


def main():
    assert sys.platform == "win32", "This acceptance test requires native Windows Word"
    with TemporaryDirectory(prefix="versioned-word-") as directory:
        work = Path(directory)
        checkout = work / "production"
        (checkout / ".git").mkdir(parents=True)
        settings_file = work / "settings.json"
        write(settings_file, {"repository": "example/private-word-test", "branch": "main",
                              "public_key": "synthetic-not-used-to-sign", "drive": {
                                  key: "synthetic-" + key for key in ("inbox", "evidence", "staging", "formal")}})
        register(list(PACKAGES.values()), checkout, settings_file)
        runtime = checkout / "review-runtime"
        settings_file = runtime / "release-trust.json"
        settings = read(settings_file)
        whole = probe_version(work, runtime, settings_file, settings, "22.6", "whole")
        verify_whole(work, load_module("_smoke_whole_gate", whole))
        points = probe_version(work, runtime, settings_file, settings, "23.6", "review-points")
        verify_points(work, points, load_module("_smoke_point_gate", points))


if __name__ == "__main__":
    main()
