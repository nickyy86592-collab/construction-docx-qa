#!/usr/bin/env python3
"""A disposable generated sample to verify the GitHub renderer and gate."""
import json
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document


with tempfile.TemporaryDirectory() as temp:
    root = Path(temp)
    docx = root / "sample.docx"
    doc = Document()
    doc.add_heading("第一章 示例", level=1)
    doc.add_heading("第一节 示例", level=2)
    doc.add_heading("一、示例", level=3)
    doc.add_paragraph("中文字体和页数检测示例。" * 20)
    doc.save(docx)
    for target, expected_code, expected_status in [(1, 0, "PASS"), (50, 2, "FAIL")]:
        output = root / f"out-{target}"
        page_dir=root/f"rules-{target}";page_dir.mkdir()
        (page_dir/"测试项目页数硬审.py").write_text(f"EXPECTED_BODY_PAGES = {target}\nPAGE_TOLERANCE = 0\nFRONT_MATTER_PAGES = 0\n",encoding="utf-8")
        result = subprocess.run(
            [sys.executable, "pdf_page_gate.py", str(docx), "--outdir", str(output),
             "--page-dir", str(page_dir)],
            text=True, capture_output=True)
        report = json.loads((output / "report.json").read_text(encoding="utf-8"))
        assert result.returncode == expected_code, (result.returncode, result.stdout, result.stderr)
        assert report["pdf_page_gate"] == expected_status, report
        assert report["word_page_gate"] == "NOT_RUN", report
        assert (output / "sample.pdf").is_file()
        print(f"target {target}: {expected_status}; {report['total_pdf_pages']} PDF page(s)")

