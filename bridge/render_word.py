import hashlib
import json
import os
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from win32com.client import DispatchEx

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def measure(word, source, target):
    saved = target / source.name
    pdf = saved.with_suffix(".pdf")
    print("OPEN_BEGIN", flush=True)
    doc = word.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
    try:
        print("OPEN_END", flush=True)
        doc.Repaginate()
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Repaginate()
        for toc in doc.TablesOfContents:
            toc.UpdatePageNumbers()
        doc.Repaginate()
        print("TOC_END", flush=True)
        doc.RemovePersonalInformation = True
        doc.SaveAs2(str(saved), FileFormat=12, AddToRecentFiles=False)
        print("SAVE_END", flush=True)
    finally:
        doc.Close(False)
    doc = word.Documents.Open(str(saved), ReadOnly=True, AddToRecentFiles=False)
    try:
        doc.Repaginate()
        print("REOPEN_END", flush=True)
        physical_pages = int(doc.ComputeStatistics(2))
        last = max(0,int(doc.Content.End)-2)
        first_body = int(doc.Bookmarks("review_start").Range.Information(3)) if doc.Bookmarks.Exists("review_start") else 0
        toc_pages = []
        for toc in doc.TablesOfContents:
            start,end=int(toc.Range.Start),max(int(toc.Range.Start),int(toc.Range.End)-2)
            toc_pages.append([int(doc.Range(start,start).Information(3)),int(doc.Range(end,end).Information(3))])
        headings=[]
        for pg in doc.Paragraphs:
            level=int(pg.OutlineLevel)
            if level in (1,2,3):
                rng=pg.Range
                headings.append({"level":level,"text":str(rng.Text).strip(),"page":int(rng.Information(3)),"numbered_page":int(rng.Information(1)),"font_asia":str(rng.Font.NameFarEast),"alignment":int(pg.Alignment),"font_size":rng.Font.Size,"font_bold":rng.Font.Bold,"font_italic":rng.Font.Italic})
        doc.ExportAsFixedFormat(str(pdf),17)
        print("EXPORT_END", flush=True)
        return dict(file=source.name,source_sha256=sha(source),document_sha256=sha(saved),pdf_sha256=sha(pdf),
                physical_pages=physical_pages,last_content_page=int(doc.Range(last,last).Information(3)),
                first_body_page=first_body,toc_pages=toc_pages,tables=int(doc.Tables.Count),
                inline_shapes=int(doc.InlineShapes.Count),headings=headings)
    finally:
        doc.Close(False)

def main():
    source, target = (Path(n).resolve() for n in sys.argv[1:])
    word = DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3
    word.Options.UpdateLinksAtOpen = False
    print("WORD_READY", flush=True)
    try:
        report = {"run_id":os.environ["GITHUB_RUN_ID"],"engine":f"Microsoft Word {word.Version}","documents":[]}
        for path in sorted(source.glob("*.docx")):
            report["documents"].append(measure(word,path,target))
    finally:
        word.Quit()
    (target / "layout.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
    print("MEASURED_DOCUMENTS",len(report["documents"]),flush=True)

if __name__=="__main__":
    main()
