#!/usr/bin/env python3
"""DOCX -> LibreOffice PDF preflight. This never reports a Microsoft Word PASS."""
import argparse
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def run(*args):
    return subprocess.run(args, check=True, text=True, capture_output=True).stdout


def structure(docx):
    with zipfile.ZipFile(docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        styles = ET.fromstring(z.read("word/styles.xml")) if "word/styles.xml" in z.namelist() else None
    heading_styles = {}
    if styles is not None:
        for style in styles.findall(f".//{W}style"):
            key = style.get(W + "styleId")
            level = style.find(f"./{W}pPr/{W}outlineLvl")
            if key and level is not None:
                heading_styles[key] = level.get(W + "val")
    headings = {"1": 0, "2": 0, "3": 0}
    forced_page_breaks = 0
    for para in root.findall(f".//{W}body/{W}p"):
        pstyle = para.find(f"./{W}pPr/{W}pStyle")
        if pstyle is not None:
            key = pstyle.get(W + "val", "")
            level = heading_styles.get(key)
            if level is None:
                m = re.fullmatch(r"Heading([123])", key, re.I)
                level = str(int(m.group(1)) - 1) if m else None
            if level in ("0", "1", "2"):
                headings[str(int(level) + 1)] += 1
        forced_page_breaks += len(para.findall(f"./{W}pPr/{W}pageBreakBefore"))
        forced_page_breaks += len([br for br in para.findall(f".//{W}br") if br.get(W + "type") == "page"])
    sections = root.findall(f".//{W}sectPr")
    landscape = sum(s.find(f"./{W}pgSz") is not None and s.find(f"./{W}pgSz").get(W + "orient") == "landscape" for s in sections)
    return {"heading_levels": headings, "forced_page_breaks": forced_page_breaks,
            "landscape_sections": landscape}


def audit(path, outdir, body_target, tolerance, front_pages):
    outdir.mkdir(parents=True, exist_ok=True)
    pdf = outdir / (path.stem + ".pdf")
    info = structure(path)
    with tempfile.TemporaryDirectory() as profile:
        run("soffice", "-env:UserInstallation=file://" + profile, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(path))
    if not pdf.is_file():
        raise RuntimeError("LibreOffice did not produce a PDF")
    metadata = run("pdfinfo", "-f", "1", "-l", "9999", str(pdf))
    total = int(re.search(r"^Pages:\s*(\d+)", metadata, re.M).group(1))
    page_sizes = re.findall(r"^Page\s+\d+ size:\s*[\d.]+ x [\d.]+ pts", metadata, re.M)
    pdf_landscape = []
    for line in page_sizes:
        m = re.search(r"Page\s+(\d+) size:\s*([\d.]+) x ([\d.]+)", line)
        if m and float(m.group(2)) > float(m.group(3)):
            pdf_landscape.append(int(m.group(1)))
    body = total - front_pages
    within = abs(body - body_target) <= tolerance
    warnings = []
    if info["forced_page_breaks"]:
        warnings.append("DOCX contains explicit page breaks; inspect whether the cover exception is the only one.")
    if not all(info["heading_levels"].values()):
        warnings.append("Not all three navigation heading levels were detected; inspect Word styles.")
    if info["landscape_sections"] or pdf_landscape:
        warnings.append("Landscape sections or pages detected.")
    result = {"method": "LibreOffice PDF rendering (NOT Microsoft Word)",
              "source_file": path.name, "pdf_file": pdf.name, "total_pdf_pages": total,
              "front_pages_excluded": front_pages, "body_pdf_pages": body,
              "target_body_pages": body_target, "tolerance": tolerance,
              "pdf_page_gate": "PASS" if within else "FAIL", "structure": info,
              "landscape_pdf_pages": pdf_landscape, "warnings": warnings,
              "word_page_gate": "NOT_RUN", "word_hard_audit": "NOT_RUN"}
    (outdir / "report.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    (outdir / "report.txt").write_text(
        f"PDF 渲染页数：总计 {total} 页；扣除前置 {front_pages} 页；正文 {body} 页。\n"
        f"目标：{body_target} ± {tolerance} 页；PDF 页数门槛：{result['pdf_page_gate']}。\n"
        "方法：LibreOffice 转 PDF；未运行桌面 Microsoft Word 页数脚本及原 Word 硬审脚本。\n"
        + "\n".join(warnings) + "\n", encoding="utf-8")
    print((outdir / "report.txt").read_text(encoding="utf-8"))
    return 0 if within else 2


def main():
    p = argparse.ArgumentParser()
    p.add_argument("docx", type=Path)
    p.add_argument("--outdir", type=Path, default=Path("qa-output"))
    p.add_argument("--page-dir",type=Path,default=Path("input/project"),help="Read the unique current project page-audit script")
    a = p.parse_args()
    try:
        import importlib.util
        module_path=Path(__file__).resolve().parent/'audit/common/编制流程门禁.py'
        spec=importlib.util.spec_from_file_location('_pdf_project_rules',module_path)
        module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
        script,values=module.discover_page_script(a.page_dir)
        if script is None:raise RuntimeError('Missing unique project page script')
        if not a.docx.is_file():raise RuntimeError('Missing DOCX')
        return audit(a.docx.resolve(),a.outdir.resolve(),values['EXPECTED_BODY_PAGES'],values['PAGE_TOLERANCE'],values['FRONT_MATTER_PAGES'])
    except Exception as exc:
        print(f"检测无法完成：{exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    sys.exit(main())

