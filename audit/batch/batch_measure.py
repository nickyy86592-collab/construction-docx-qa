"""Actual Word measurement for a delivered batch or an intermediate prefix.
Does not assert the whole-book 40/75 percent or final-delivery gates.
"""
import argparse,hashlib,json,sys
from pathlib import Path
from datetime import datetime
def main():
 p=argparse.ArgumentParser();p.add_argument('--document',required=True);p.add_argument('--outdir',required=True);a=p.parse_args()
 import win32com.client
 src=Path(a.document).resolve();out=Path(a.outdir).resolve();out.mkdir(parents=True,exist_ok=True)
 word=doc=None
 try:
  word=win32com.client.DispatchEx('Word.Application');word.Visible=False;word.DisplayAlerts=0
  doc=word.Documents.Open(str(src),ReadOnly=True,AddToRecentFiles=False)
  doc.Fields.Update();doc.Repaginate();pages=int(doc.ComputeStatistics(2))
  headings=[]
  for para in doc.Paragraphs:
   level=int(para.OutlineLevel)
   if level<=2:
    headings.append({'level':level,'text':str(para.Range.Text).strip(),'physical_page':int(para.Range.Information(3)),'numbered_page':int(para.Range.Information(1))})
  pdf=out/'batch-word.pdf';doc.ExportAsFixedFormat(str(pdf),17)
  report={'audit_type':'batch_word_measurement','status':'MEASURED','page_source':'microsoft_word','measurement_role':'first_batch_or_prefix','document_sha256':hashlib.sha256(src.read_bytes()).hexdigest(),'measured_at':datetime.now().astimezone().isoformat(),'actual':{'word_total_pages':pages,'word_body_pages':pages-1,'front_matter_pages':1,'word_version':str(word.Version)},'headings':headings,'word_pdf_evidence':{'path':pdf.name,'sha256':hashlib.sha256(pdf.read_bytes()).hexdigest()},'limits':'This measurement is not a whole-book 40/75-percent or final delivery gate.'}
  (out/'batch-word.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(report['actual']))
 finally:
  if doc is not None:doc.Close(False)
  if word is not None:word.Quit()
if __name__=='__main__':main()
