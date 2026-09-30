import sys
from pathlib import Path
import win32com.client
word=None;doc=None
try:
 word=win32com.client.DispatchEx('Word.Application');word.Visible=False;word.DisplayAlerts=0
 doc=word.Documents.Open(str(Path(sys.argv[1]).resolve()),ReadOnly=True,AddToRecentFiles=False)
 doc.Repaginate()
 print('WORD_TOTAL_PAGES='+str(doc.ComputeStatistics(2)))
 doc.ExportAsFixedFormat(str(Path(sys.argv[2]).resolve()),17)
finally:
 if doc is not None:doc.Close(False)
 if word is not None:word.Quit()
