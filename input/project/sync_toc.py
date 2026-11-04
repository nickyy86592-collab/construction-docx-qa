"""Refresh this document's ten cached TOC page references using Windows Word."""
from pathlib import Path
from zipfile import ZipFile
import hashlib
import json
import re
import sys

REFERENCE = re.compile(
    r'(<w:instrText\b[^>]*>\s*PAGEREF Chapter(\d+)\s*</w:instrText>'
    r'(?:(?!<w:instrText\b).)*?<w:fldChar\b[^>]*w:fldCharType="separate"[^>]*?/>'
    r'.*?<w:t(?: [^>]*)?>)(\d+)(</w:t>)',
    re.S,
)


def measured_layout(path):
    import win32com.client

    word = win32com.client.DispatchEx('Word.Application')
    doc = None
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        doc = word.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
        doc.Repaginate()
        return {
            'word_version': str(word.Version),
            'total_pages': int(doc.ComputeStatistics(2)),
            'body_pages': int(doc.ComputeStatistics(2)) - 1,
            'chapter_pages': {
                int(i): int(doc.Bookmarks(f'Chapter{i}').Range.Information(1))
                for i in range(1, 11)
            },
        }
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def refresh(path, pages):
    tmp = path.with_suffix('.toc-sync.tmp')
    with ZipFile(path) as source:
        old = source.read('word/document.xml').decode('utf-8')
        found = {int(m[2]): int(m[3]) for m in REFERENCE.finditer(old)}
        assert len(list(REFERENCE.finditer(old))) == len(found) == 10
        assert found.keys() == pages.keys()
        if found == pages:
            return found
        updated = REFERENCE.sub(lambda m: m[1] + str(pages[int(m[2])]) + m[4], old)
        with ZipFile(tmp, 'w') as output:
            for entry in source.infolist():
                content = updated.encode('utf-8') if entry.filename == 'word/document.xml' else source.read(entry)
                output.writestr(entry, content)
    tmp.replace(path)
    return found


if __name__ == '__main__':
    path = Path(sys.argv[1]).resolve()
    report = Path(sys.argv[2]).resolve()
    observed = measured_layout(path)
    previous = refresh(path, observed['chapter_pages'])
    confirmed = measured_layout(path)
    if confirmed != observed:
        raise RuntimeError(f'Word chapter pagination changed while caching the TOC: {observed} -> {confirmed}')
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps({
        'page_source': 'microsoft_word',
        'document_sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
        'previous_toc': previous,
        **confirmed,
    }, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(confirmed, ensure_ascii=False))
