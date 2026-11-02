"""Portable table checks; --word adds actual Windows Word/PDF evidence."""
from pathlib import Path
import argparse,copy,importlib.util,json,sys,tempfile,zipfile,os
from dataclasses import asdict
COMMON=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('table_word_audit',COMMON/'Word正式交付硬审（不含页数）.py')
c=importlib.util.module_from_spec(spec);sys.modules[spec.name]=c;spec.loader.exec_module(c)
RESULTS=[]
CFG={'table_layout':{'tables':[{'index':1,'header_rows':1,'column_roles':['short','short','short','long']}]}}
def check(name,ok):
 assert ok,name
 RESULTS.append(name)
def xml_patch(path,fn):
 import xml.etree.ElementTree as ET
 with zipfile.ZipFile(path) as z:parts={k:z.read(k) for k in z.namelist()}
 root=ET.fromstring(parts['word/document.xml']);styles=ET.fromstring(parts['word/styles.xml'])
 fn(root,styles);parts['word/document.xml']=ET.tostring(root,encoding='utf-8',xml_declaration=True);parts['word/styles.xml']=ET.tostring(styles,encoding='utf-8',xml_declaration=True)
 with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
  for k,v in parts.items():z.writestr(k,v)
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def el(name,**attrs):
 from xml.etree.ElementTree import Element
 return Element(W+name,{W+k:str(v) for k,v in attrs.items()})
def fixture(path,mutation=None,many=False):
 from docx import Document
 from docx.shared import Pt,Twips
 from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
 from docx.enum.text import WD_ALIGN_PARAGRAPH
 from docx.oxml import OxmlElement
 from docx.oxml.ns import qn
 d=Document();sec=d.sections[0];sec.page_width=Twips(11906);sec.page_height=Twips(16838);sec.left_margin=sec.right_margin=Twips(1134)
 normal=d.styles['Normal'];normal.font.name='SimSun';normal.font.size=Pt(12);normal.paragraph_format.first_line_indent=Pt(24)
 d.add_heading('第一章 表格排版集成测试',1);d.add_paragraph('按资料核对设备、数量及投入计划，分工序落实检验和交接。')
 t=d.add_table(rows=1,cols=4);t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
 widths=[1900,1400,1800,4500]
 for g,w in zip(t.columns,widths):g.width=Twips(w)
 contents=[['设备名称','拟用数量','投入时段','规格、用途及能力'],['挖掘机','2台','1—30天','按批准参数选型；清淤与基础作业各1台，进场验收合格后使用。'],['自卸运输车','14辆','6—30天','有效容量按核定值投入，根据现场运输循环及作业面条件调整。'],['振动压路机','1台','16—40天','吨位经试验段、承载核定，作业前检查路基及基层状态。'],['小型夯机','4台','10—40天','狭槽边角适配型；灰土与道路边角按工作面分别投入并备用。'],['排水泵软管','3套','1—90天','按来水流量扬程核定配置，服务导排集水，超出能力时增配。']]
 if many:contents+=contents[1:]*10
 for ri,texts in enumerate(contents):
  row=t.rows[0] if ri==0 else t.add_row()
  if ri==0:
   repeat=OxmlElement('w:tblHeader');repeat.set(qn('w:val'),'1');row._tr.get_or_add_trPr().append(repeat)
  for ci,text in enumerate(texts):
   cell=row.cells[ci];cell.width=Twips(widths[ci]);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;cell.text=text
   p=cell.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER if ri==0 or ci<3 else WD_ALIGN_PARAGRAPH.LEFT
   pf=p.paragraph_format;pf.left_indent=pf.right_indent=pf.first_line_indent=Pt(0);pf.space_before=pf.space_after=Pt(2)
   for run in p.runs:run.font.name='SimSun';run.font.size=Pt(12)
 d.save(path)
 if mutation:xml_patch(path,mutation)
 return path

def portable(base):
 def checks(p,cfg=None):return c.check_table_layout_static(p,CFG if cfg is None else cfg)
 def fails(p):return checks(p)[0].status=='FAIL'
 good=fixture(base/'good.docx');check('configured short/long table accepted',checks(good)[0].status=='PASS' and checks(good)[1].status=='PASS')
 # Inheritance: removing direct zero reveals the body-style first-line indent.
 def indent(root,styles):root.find('.//'+W+'tbl//'+W+'pPr/'+W+'ind').clear()
 bad=fixture(base/'inherited-indent.docx',indent);check('inherited body first-line indent blocked',fails(bad))
 def direct(root,styles):root.find('.//'+W+'tbl//'+W+'pPr/'+W+'ind').set(W+'firstLineChars','200')
 check('character indent blocked',fails(fixture(base/'char-indent.docx',direct)))
 def header(root,styles):root.find('.//'+W+'tbl//'+W+'pPr/'+W+'jc').set(W+'val','left')
 check('header left alignment blocked',fails(fixture(base/'header-left.docx',header)))
 def short(root,styles):root.findall('.//'+W+'tbl/'+W+'tr')[1].find('.//'+W+'pPr/'+W+'jc').set(W+'val','left')
 check('short content left alignment blocked',fails(fixture(base/'short-left.docx',short)))
 def long(root,styles):root.findall('.//'+W+'tbl/'+W+'tr')[1].findall(W+'tc')[3].find('.//'+W+'jc').set(W+'val','center')
 check('long description centered blocked',fails(fixture(base/'long-center.docx',long)))
 def vertical(root,styles):root.find('.//'+W+'tbl//'+W+'vAlign').set(W+'val','top')
 check('vertical top blocked',fails(fixture(base/'vertical-top.docx',vertical)))
 def tableleft(root,styles):root.find('.//'+W+'tbl/'+W+'tblPr/'+W+'jc').set(W+'val','left')
 check('whole table left blocked',fails(fixture(base/'table-left.docx',tableleft)))
 def wide(root,styles):root.find('.//'+W+'tbl/'+W+'tblGrid/'+W+'gridCol').set(W+'w','9000')
 check('fixed width overflow blocked',fails(fixture(base/'overflow.docx',wide)))
 check('missing semantic roles warned not guessed',checks(good,{})[0].status=='PASS' and checks(good,{})[1].status=='WARN')
 for tag in ('gridSpan','hMerge','vMerge'):
  def merged(root,styles,tag=tag):root.find('.//'+W+'tbl//'+W+'tcPr').append(el(tag,val='2' if tag=='gridSpan' else 'restart'))
  p=fixture(base/(tag+'.docx'),merged);check('complex '+tag+' warns without false FAIL',checks(p)[0].status=='PASS' and checks(p)[1].status=='WARN')
 def floating(root,styles):
  pr=root.find('.//'+W+'tbl/'+W+'tblPr');pr.find(W+'jc').set(W+'val','left');pr.append(el('tblpPr',vertAnchor='page',horzAnchor='page',tblpX='300'))
 p=fixture(base/'floating.docx',floating);check('floating alignment requires visual review',checks(p)[0].status=='PASS' and checks(p)[1].status=='WARN')
 def overall_indent(root,styles):root.find('.//'+W+'tbl/'+W+'tblPr').append(el('tblInd',w='240',type='dxa'))
 p=fixture(base/'table-indent.docx',overall_indent);check('whole table indent warned',checks(p)[1].status=='WARN')
 def conditional(root,styles):
  st=el('style',type='table',styleId='ConditionalTable');cond=el('tblStylePr',type='firstRow');pp=el('pPr');pp.append(el('jc',val='center'));cond.append(pp);st.append(cond);styles.append(st)
  root.find('.//'+W+'tbl/'+W+'tblPr').append(el('tblStyle',val='ConditionalTable'))
 p=fixture(base/'conditional.docx',conditional);check('conditional table styles not guessed',checks(p)[0].status=='PASS' and checks(p)[1].status=='WARN')
 def inherited_table(root,styles):
  for st in styles.findall(W+'style'):
   if st.get(W+'type')=='table':st.attrib.pop(W+'default',None)
  st=el('style',type='table',styleId='DefaultCenteredTable',default='1');pr=el('tblPr');pr.append(el('jc',val='center'));st.append(pr)
  cp=el('tcPr');cp.append(el('vAlign',val='center'));st.append(cp);styles.append(st)
  table=root.find('.//'+W+'tbl');tp=table.find(W+'tblPr');tp.remove(tp.find(W+'jc'))
  for cp in table.iter(W+'tcPr'):
   node=cp.find(W+'vAlign')
   if node is not None:cp.remove(node)
 p=fixture(base/'default-table-style.docx',inherited_table);check('default table style effective alignment inherited',checks(p)[0].status=='PASS' and checks(p)[1].status=='PASS')
 def breaks(root,styles):root.find('.//'+W+'tbl//'+W+'r').append(el('br'))
 p=fixture(base/'manual-break.docx',breaks);check('manual breaks warned without deleting content',checks(p)[1].status=='WARN')
 from docx import Document
 d=Document();d.add_table(rows=1,cols=1).cell(0,0).text='【实景图占位：洞口】需插入现场照片';p=base/'photo.docx';d.save(p)
 check('one-cell photo placeholder not data table',checks(p,{})[0].status=='PASS' and checks(p,{})[1].status=='PASS')
 d=Document();t=d.add_table(rows=2,cols=2);t.cell(0,0).text='图形布局';t.cell(1,0).add_table(rows=1,cols=1);p=base/'nested.docx';d.save(p)
 check('nested layout warned',checks(p,{})[0].status=='PASS' and checks(p,{})[1].status=='WARN')
 for cfg in ({'table_alignment':'left'}, {'zero_indents':False}, {'tables':[{'index':1,'column_roles':['unknown']}]}, {'tables':[{'index':1,'kind':'layout'}]}, {'tables':[{'index':1},{'index':1}]}):
  try:c.table_layout({'table_layout':cfg});ok=False
  except c.AuditError:ok=True
  check('invalid/unsourced table rule rejected '+str(cfg),ok)
 override=copy.deepcopy(CFG);override['table_layout'].update(table_alignment='left',source_ref='有效补遗第2条')
 check('explicit tender layout exception accepted',not any(x.status=='FAIL' for x in checks(fixture(base/'exception.docx',tableleft),override)))
 before=good.read_bytes();checks(good);check('audit never edits input DOCX',before==good.read_bytes())
 check('no new production script needed',hasattr(c,'check_table_layout_word'))

def word_suite(base):
 if sys.platform!='win32':raise RuntimeError('--word requires actual Windows Word')
 import win32com.client
 native=Path(os.environ.get('WORD_TABLE_EVIDENCE_DIR',str(Path.cwd()/'word-table-evidence'))).resolve();native.mkdir(parents=True,exist_ok=True)
 word=win32com.client.DispatchEx('Word.Application');word.Visible=False;word.DisplayAlerts=0;details=[]
 try:
  for name in ('good','inherited-indent','table-left','vertical-top','long-center','short-left'):
   path=base/(name+'.docx');doc=word.Documents.Open(str(path.resolve()),ReadOnly=True,AddToRecentFiles=False)
   try:
    doc.Repaginate();results=c.check_table_layout_word(doc,CFG);bad=any(x.status=='FAIL' for x in results)
    check('native Word '+name,not bad if name=='good' else bad)
    if name=='good':doc.ExportAsFixedFormat(str(native/'table-good.word.pdf'),17)
    details.append({'fixture':name,'checks':[asdict(x) for x in results]})
   finally:doc.Close(False)
  many=fixture(native/'multipage.docx',many=True);doc=word.Documents.Open(str(many),ReadOnly=True,AddToRecentFiles=False)
  try:
   doc.Repaginate();check('native multipage table accepted',not any(x.status=='FAIL' for x in c.check_table_layout_word(doc,CFG)))
   pages=int(doc.ComputeStatistics(2));check('native table actually spans pages',pages>1)
   doc.ExportAsFixedFormat(str(native/'table-multipage.word.pdf'),17)
   details.append({'fixture':'multipage','pages':pages,'checks':[asdict(x) for x in c.check_table_layout_word(doc,CFG)]})
  finally:doc.Close(False)
 finally:word.Quit()
 (native/'table-word-results.json').write_text(json.dumps(details,ensure_ascii=False,indent=2),encoding='utf-8')
 return str(native)

def main():
 ap=argparse.ArgumentParser();ap.add_argument('--word',action='store_true');ap.add_argument('--report',type=Path);args=ap.parse_args()
 with tempfile.TemporaryDirectory() as tmp:
  base=Path(tmp);portable(base);evidence=word_suite(base) if args.word else None
 result={'status':'PASS','checks':len(RESULTS),'word_requested':args.word,'word_evidence_dir':evidence,'tests':RESULTS}
 if args.report:args.report.parent.mkdir(parents=True,exist_ok=True);args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
 print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
