"""Word 原生分页，只读打开本批 DOCX。"""
from contextlib import contextmanager
from pathlib import Path
import sys
from release_task import file_sha, require


@contextmanager
def open_word(document):
    require(sys.platform == "win32", "须在 Windows 且可调用 Microsoft Word 的运行器实测")
    require(Path(document).suffix.casefold() == ".docx", "实排入口接受冻结的 DOCX")
    import win32com.client

    before = file_sha(document)
    word = win32com.client.DispatchEx("Word.Application")
    doc = None
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        word.Options.UpdateLinksAtOpen = False
        doc = word.Documents.Open(str(Path(document).resolve()), ReadOnly=True, AddToRecentFiles=False)
        doc.Repaginate()
        yield word, doc
        require(file_sha(document) == before, "Word 实测期间输入发生变化")
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        finally:
            word.Quit()


def probe(document, task):
    marker = "WORD-CHANNEL-" + task["batch_id"]
    require(file_sha(document) == task["probe_sha256"], "试排样稿与任务单不匹配")
    with open_word(document) as (word, doc):
        pages = int(doc.ComputeStatistics(2))
        require(pages >= 1 and marker in str(doc.Content.Text), "未实读本批试排样稿")
        return {"engine": "Microsoft Word", "word_version": str(word.Version), "pages": pages,
                "document_sha256": file_sha(document)}

