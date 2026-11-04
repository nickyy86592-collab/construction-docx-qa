"""Portable table checks; --word adds actual Windows Word/PDF evidence."""
from pathlib import Path
import argparse,copy,importlib.util,json,sys,tempfile,zipfile,os
from dataclasses import asdict
COMMON=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(COMMON))
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
def _write_table_docx(path,widths,contents,center_body=False,body_indent=False):
 from xml.sax.saxutils import escape
 # Standard-library OOXML seed: native Word runners need only pywin32.
 def paragraph(text,alignment=None,zero=False,header=False):
  props=('<w:ind w:left="0" w:right="0" w:firstLine="'+('480' if body_indent and not header else '0')+'"/>' if zero else '')
  if alignment:props+='<w:jc w:val="'+alignment+'"/>'
  props+='<w:spacing w:before="40" w:after="40"/>'
  return '<w:p><w:pPr>'+props+'</w:pPr><w:r><w:rPr><w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/><w:sz w:val="24"/>'+('<w:color w:val="FFFFFF"/>' if header else '')+'</w:rPr><w:t>'+escape(text)+'</w:t></w:r></w:p>'
 rows=[]
 for ri,texts in enumerate(contents):
  cells=[]
  for ci,text in enumerate(texts):
   alignment='center' if ri==0 or center_body or (len(widths)>1 and ci<len(widths)-1) else 'left'
   cells.append('<w:tc><w:tcPr><w:tcW w:w="'+str(widths[ci])+'" w:type="dxa"/><w:shd w:val="clear" w:fill="'+('284B65' if ri==0 else 'EAF0F4')+'"/><w:vAlign w:val="center"/></w:tcPr>'+paragraph(text,alignment,True,ri==0)+'</w:tc>')
  rows.append('<w:tr>'+('<w:trPr><w:tblHeader w:val="1"/></w:trPr>' if ri==0 else '')+''.join(cells)+'</w:tr>')
 borders=''.join('<w:'+side+' w:val="single" w:sz="4" w:color="D4DFE7"/>' for side in ('top','left','bottom','right','insideH','insideV'))
 table='<w:tbl><w:tblPr><w:tblW w:w="'+str(sum(widths))+'" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>'+borders+'</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>'+''.join('<w:gridCol w:w="'+str(w)+'"/>' for w in widths)+'</w:tblGrid>'+''.join(rows)+'</w:tbl>'
 body='<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章 表格排版集成测试</w:t></w:r></w:p>'+paragraph('按资料核对设备、数量及投入计划，分工序落实检验和交接。')+table+'<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1134" w:bottom="1440" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
 styles='<w:styles xmlns:w="'+W[1:-1]+'"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:ind w:firstLine="480"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:firstLine="0"/><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style></w:styles>'
 with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
  z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
  z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
  z.writestr('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
  z.writestr('word/styles.xml',styles)
  z.writestr('word/document.xml','<w:document xmlns:w="'+W[1:-1]+'"><w:body>'+body+'</w:body></w:document>')
 return path

def fixture(path,mutation=None,many=False):
 contents=[['设备名称','拟用数量','投入时段','规格、用途及能力'],['挖掘机','2台','1—30天','按批准参数选型；清淤与基础作业各1台，进场验收合格后使用。'],['自卸运输车','14辆','6—30天','有效容量按核定值投入，根据现场运输循环及作业面条件调整。'],['振动压路机','1台','16—40天','吨位经试验段、承载核定，作业前检查路基及基层状态。'],['小型夯机','4台','10—40天','狭槽边角适配型；灰土与道路边角按工作面分别投入并备用。'],['排水泵软管','3套','1—90天','按来水流量扬程核定配置，服务导排集水，超出能力时增配。']]
 if many:contents+=contents[1:]*10
 _write_table_docx(path,[1900,1400,1800,4500],contents)
 if mutation:xml_patch(path,mutation)
 return path

def single_fixture(path,center_body=False,indent=False):
 return _write_table_docx(path,[9600],[['施工要求'],['作业前核对工作面与设计接口，按批准工艺组织实施，检验合格后交接。']],center_body,indent)


def native_fixtures(base):
 fixture(base/'good.docx')
 def indent(root,styles):root.find('.//'+W+'tbl//'+W+'pPr/'+W+'ind').clear()
 fixture(base/'inherited-indent.docx',indent)
 def tableleft(root,styles):root.find('.//'+W+'tbl/'+W+'tblPr/'+W+'jc').set(W+'val','left')
 fixture(base/'table-left.docx',tableleft)
 def vertical(root,styles):root.find('.//'+W+'tbl//'+W+'vAlign').set(W+'val','top')
 fixture(base/'vertical-top.docx',vertical)
 def long(root,styles):root.findall('.//'+W+'tbl/'+W+'tr')[1].findall(W+'tc')[3].find('.//'+W+'jc').set(W+'val','center')
 fixture(base/'long-center.docx',long)
 def short(root,styles):root.findall('.//'+W+'tbl/'+W+'tr')[1].find('.//'+W+'pPr/'+W+'jc').set(W+'val','left')
 fixture(base/'short-left.docx',short)
 single_fixture(base/'single-good.docx');single_fixture(base/'single-indent.docx',indent=True);single_fixture(base/'single-center.docx',center_body=True)

SINGLE_CFG={'table_layout':{'tables':[{'index':1,'kind':'data','header_rows':1,'column_roles':['long']}]}}

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
 single=single_fixture(base/'single-good.docx')
 result=checks(single,{})
 check('unknown single-column table requires review',result[0].status=='PASS' and result[1].status=='WARN')
 result=checks(single,SINGLE_CFG)
 check('confirmed single-column data table actually checked',result[0].status=='PASS' and result[1].status=='PASS' and result[0].details['checked_tables']==1)
 result=checks(single_fixture(base/'single-indent.docx',indent=True),SINGLE_CFG)
 check('single-column data indent blocked',result[0].status=='FAIL')
 result=checks(single_fixture(base/'single-center.docx',center_body=True),SINGLE_CFG)
 check('single-column long content centered blocked',result[0].status=='FAIL')
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
  for name,config,expected_fail,expected_review in (
      ('single-good',SINGLE_CFG,False,True),('single-indent',SINGLE_CFG,True,True),
      ('single-center',SINGLE_CFG,True,True),('single-good',{},False,True)):
   doc=word.Documents.Open(str((base/(name+'.docx')).resolve()),ReadOnly=True,AddToRecentFiles=False)
   try:
    doc.Repaginate();results=c.check_table_layout_word(doc,config)
    check('native Word single-column '+name+(' unclassified' if not config else ''),any(x.status=='FAIL' for x in results)==expected_fail and (not expected_review or any(x.status=='WARN' for x in results)))
    details.append({'fixture':name,'classification':'configured' if config else 'unknown','checks':[asdict(x) for x in results]})
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
  base=Path(tmp)
  if args.word:native_fixtures(base);evidence=word_suite(base)
  else:portable(base);evidence=None
 result={'status':'PASS','checks':len(RESULTS),'word_requested':args.word,'word_evidence_dir':evidence,'tests':RESULTS}
 if args.report:args.report.parent.mkdir(parents=True,exist_ok=True);args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
 print(json.dumps(result,ensure_ascii=False))
if __name__=='__main__':main()
