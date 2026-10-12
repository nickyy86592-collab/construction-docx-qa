import sys
from pathlib import Path
import win32com.client
word=doc=None
try:
    word=win32com.client.DispatchEx('Word.Application')
    word.Visible=False
    word.DisplayAlerts=0
    doc=word.Documents.Open(str(Path(sys.argv[1]).resolve()),ReadOnly=True,AddToRecentFiles=False)
    doc.Repaginate()
    output=Path(sys.argv[2]);output.parent.mkdir(parents=True,exist_ok=True)
    doc.ExportAsFixedFormat(str(output.resolve()),17)
finally:
    if doc is not None:doc.Close(False)
    if word is not None:word.Quit()
