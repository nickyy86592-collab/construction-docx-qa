"""2026-10-22 regressions; --word additionally requires real Windows Word COM."""
from pathlib import Path
import argparse,copy,importlib.util,json,sys,tempfile,contextlib,io
COMMON=Path(__file__).resolve().parents[1]
def load(name,file):
 spec=importlib.util.spec_from_file_location(name,COMMON/file);m=importlib.util.module_from_spec(spec);sys.modules[name]=m;spec.loader.exec_module(m);return m
c=load('upgrade_word','Word正式交付硬审（不含页数）.py');f=load('upgrade_flow','编制流程门禁.py');score=load('upgrade_score','评分响应核验.py')
RESULTS=[]
def check(name,ok):
 assert ok,name
 RESULTS.append(name)
def statuses(text):return [x.status for x in c.language_checks([{'text':text}])]
def write(p,obj):p.write_text(json.dumps(obj,ensure_ascii=False),encoding='utf-8');return p

def portable(base):
 from docx import Document
 for text in ('差异尚未闭合','本章需进一步完善','建议补充相关措施','本节应补充工艺描述','待办：补写安全章节','本稿已通过审核','审核状态：待确认'):
  check('internal FAIL '+text,statuses(text)[0]=='FAIL')
 for text in ('疑似符号需核对','建议补充接地措施','监理提出建议补充相关措施，项目部整改后复验。','不得将“本章需进一步完善”写入施工记录。','参数尚待核实'):
  check('context WARN '+text,statuses(text)==['PASS','WARN'])
 for text in ('设备进场后检查外观并办理材料报审。','发现疑似裂纹时停止作业，复检确认后按批准方案处置。','监理审查意见落实后组织复验。','焊接缺陷整改后复查合格再移交。','附着力试验合格后进入下一工序。'):
  check('construction PASS '+text,statuses(text)==['PASS','PASS'])
 d=Document();p=d.add_paragraph();p.add_run('本章需进');p.add_run('一步完善');d.add_table(rows=1,cols=1).cell(0,0).text='差异尚未闭合'
 d.sections[0].header.paragraphs[0].text='审核状态：未完成';d.sections[0].footer.paragraphs[0].text='建议补充相关措施'
 d.add_paragraph('疑似符号待核实');doc=base/'language.docx';d.save(doc)
 records=c.document_language_records(doc);hits=c.language_checks(records)
 check('runs/tables/header/footer detected',len(hits[0].details)==4)
 check('paragraph located without table cell concatenation',all('paragraph' in x and 'part' in x for x in hits[0].details))
 # Textboxes and notes are readable visible text, including split runs.
 import zipfile
 with zipfile.ZipFile(doc,'a') as z:
  xml='<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>本章需进一步完善</w:t></w:r></w:p></w:footnote></w:footnotes>'
  z.writestr('word/footnotes.xml',xml)
 check('footnotes scanned',any(x.get('part')=='word/footnotes.xml' for x in c.language_checks(c.document_language_records(doc))[0].details))
 page=base/'pages';page.mkdir();script=page/'任意项目页数硬审.py'
 def rule(front,extra=None):
  script.write_text(f'EXPECTED_BODY_PAGES=50\nPAGE_TOLERANCE=1\nFRONT_MATTER_PAGES={front}\n')
  rp=write(base/'format.json',extra or {});return c.load_rules(rp,page)
 r=rule(1);check('default cover and level1 share page1',r['body_start_page']==2 and r['front_matter']['toc_levels']==[1])
 r=rule(3,{'front_matter':{'cover_required':True,'toc_levels':[1,2,3],'cover_page_range':[1,1],'toc_page_range':[2,3]}})
 check('separate cover multi-page multi-level TOC',r['body_start_page']==4 and r['front_matter']['toc_page_range']==[2,3])
 r0=rule(0,{'front_matter':{'cover_required':False,'toc_levels':[]}});check('no front matter body page1',r0['body_start_page']==1)
 for extra in ({'body_start_page':2},{'front_matter':{'toc_levels':[3,1]}},{'front_matter':{'toc_page_range':[1,5]}},{'heading_patterns':{'1':'['}}):
  try:rule(3,extra);ok=False
  except c.AuditError:ok=True
  check('invalid format conflict blocked '+str(extra),ok)
 r=rule(3,{'heading_patterns':{'1':r'^([0-9]+) +.+$','2':r'^([0-9]+)\\. +.+$','3':r'^\\(([0-9]+)\\) +.+$'}})
 check('custom tender heading parsed',c.classify_heading('1 总体部署',r)==(1,1))
 # Front checks with page-aware records; Windows suite below tests COM extraction.
 def rec(text,page,start,outline=10):return {'text':text,'page':page,'start':start,'end':start+1,'outline':outline}
 heads=[rec('第一章 施工部署',4,100,1)|{'level':1},rec('第一节 准备',4,110,2)|{'level':2},rec('一、检查',4,120,3)|{'level':3}]
 prefix=[rec('项目名称：测试项目',1,1),rec('第一章 施工部署 4',2,10),rec('第一节 准备 4',2,20),rec('一、检查 4',3,30)]
 rules=rule(3,{'front_matter':{'cover_required':True,'toc_levels':[1,2,3],'cover_page_range':[1,1],'toc_page_range':[2,3],'cover_keywords':['项目名称']}})
 class FakeDoc:pass
 checks=c.check_front_page_and_body(FakeDoc(),prefix+heads+[rec('施工准备和检查要求。',4,130)],heads,rules,4)
 check('multilevel prefix names/order/navigation accepted',not [x for x in checks if x.status=='FAIL'])
 bad=copy.deepcopy(prefix);bad[2]['text']='第一节 其他名称 4'
 check('TOC mismatch blocked',any(x.status=='FAIL' for x in c.check_front_page_and_body(FakeDoc(),bad+heads+[rec('施工措施',4,130)],heads,rules,4)))
 bad=copy.deepcopy(prefix);bad[1]['outline']=1
 check('front navigation pollution blocked',any(x.name=='前置页_导航隔离' and x.status=='FAIL' for x in c.check_front_page_and_body(FakeDoc(),bad+heads+[rec('施工措施',4,130)],heads,rules,4)))
 checks=c.check_front_page_and_body(FakeDoc(),prefix[1:]+heads+[rec('施工措施',4,130)],heads,rules,4)
 check('missing cover cannot be replaced by TOC',any(x.name=='前置封面信息' and x.status=='FAIL' for x in checks))
 src=base/'tender.txt';src.write_text('评分原文：措施针对工程且技术深入，资源和控制完整。',encoding='utf-8')
 ev={'path':str(src),'sha256':f.sha256(src),'locator':'评标办法第1项'}
 req={'id':'R1','scope':src.read_text(),'scoring':True,'evidence':[ev]};mod={'id':'M1','requirement_ids':['R1']};state={'plan':{'requirements':[req],'modules':[mod,{'id':'M2','requirement_ids':[]}]}}
 text='先核验工作面与设计接口，按批准工艺逐段实施，检查合格并记录后移交下道工序。'
 condition={'requirement_id':'R1','criterion_quote':req['scope'],'response_quote':text,'evidence':[ev],'targeted':True,'technical_depth':True,'project_basis':True,'substantive_response':True,'status':'satisfied','conclusion':'当前模块实际措施满足对应子条件；合成判断'}
 review={'construction_language_review':{'reviewed':True,'unresolved_count':0,'conclusion':'施工语言核查'},'scoring_review':{'conditions':[condition],'conclusion':'本模块责任复核'}}
 check('scoring module review no forced score',not f.module_review_errors(state,mod,text,review,base))
 for name,mut in [('missing criterion',lambda x:x['scoring_review'].update(conditions=[])),('insufficient depth',lambda x:x['scoring_review']['conditions'][0].update(technical_depth=False)),('invented response',lambda x:x['scoring_review']['conditions'][0].update(response_quote='这段话不存在于当前实际模块中，不能作为对应技术措施的证明')),('missing source',lambda x:x['scoring_review']['conditions'][0]['evidence'][0].update(sha256='old')),('unfinished language',lambda x:x['construction_language_review'].update(unresolved_count=1)),('forced score',lambda x:x['scoring_review']['conditions'][0].update(self_score=9))]:
  bad=copy.deepcopy(review);mut(bad);check(name+' blocked',bool(f.module_review_errors(state,mod,text,bad,base)))
 deferred=copy.deepcopy(review);deferred['scoring_review']['conditions'][0].update(status='deferred',remaining_condition='与下游资源接口联合复核',responsible_module_id='M2')
 check('cross-module remainder recorded without scoring',not f.module_review_errors(state,mod,text,deferred,base))
 # Existing state is read only: no init/migration or historical reviews rewritten.
 path=write(base/'old-state.json',{'schema_version':'3.0','plan':{'modules':[{'id':'OLD'}]},'module_registry':{'OLD':{'status':'complete','source_path':str(src),'source_sha256':f.sha256(src),'review_path':str(src),'review_sha256':f.sha256(src)}}})
 before=path.read_bytes();check('old registry remains valid without fabricated new review',not f.registry_errors(f.load_state(path),base));check('old state byte-preserved',before==path.read_bytes())
 new=state|{'module_registry':{'M1':{'status':'complete','source_path':str(src),'source_sha256':f.sha256(src),'review_path':str(src),'review_sha256':f.sha256(src),'review_contract':f.REVIEW_CONTRACT,'requirements_sha256':'stale'}}}
 check('changed responsibility invalidates new review',bool(f.registry_errors(new,base,['M1'])))
 # Existing score calculator keeps partial items unscored and final completeness mandatory.
 item={'id':'S','name':'施工方案','source_ref':'第1项','criteria':['完整'], 'max_score':10,'excellent_range':[9,10],'completion':'partial'}
 data={'tender_source':'fixture','items':[item]}
 check('partial scoring condition no forced stage score',score.check(data,'40')['completed_items_score']=='0')
 check('final partial item blocked',score.check(data,'final')['status']=='FAIL')
 check('40/75 nodes preserved',f.STAGES==(40,75))
 check('long sample remains mandatory',f.planning_policy(395)['pilot_required'])

def word_suite(base):
 if sys.platform!='win32':raise RuntimeError('--word requires Windows; no emulation or estimated PASS')
 import win32com.client
 word=win32com.client.DispatchEx('Word.Application');word.Visible=False;word.DisplayAlerts=0
 try:
  for kind in ('default','multi','none'):
   front={'default':1,'multi':3,'none':0}[kind];page=base/kind;page.mkdir()
   (page/'项目页数硬审.py').write_text(f'EXPECTED_BODY_PAGES=50\nPAGE_TOLERANCE=1\nFRONT_MATTER_PAGES={front}\n',encoding='utf-8')
   fm={'cover_required':front>0,'toc_levels':[1,2,3] if kind=='multi' else ([1] if front else [])}
   if kind=='multi':fm.update(cover_page_range=[1,1],toc_page_range=[2,3])
   rules=c.load_rules(write(page/'format.json',{'front_matter':fm}),page)
   from xml.sax.saxutils import escape
   import zipfile
   paragraphs=[]
   def para(text,level=None):
    props='<w:pPr><w:outlineLvl w:val="'+str(level-1)+'"/></w:pPr>' if level else ''
    paragraphs.append('<w:p>'+props+'<w:r><w:t xml:space="preserve">'+escape(text)+'</w:t></w:r></w:p>')
   def newpage():paragraphs.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
   if front:
    para('项目名称：格式适配合成测试')
    if kind=='multi':newpage()
    para('第一章 施工部署 '+str(front+1))
    if kind=='multi':
     para('第一节 施工准备 4');newpage();para('一、工序检查 4')
    newpage()
   para('第一章 施工部署',1);para('第一节 施工准备',2);para('一、工序检查',3)
   para('完成材料报审及工作面核验后组织作业，检查合格后移交。'*5)
   path=page/'fixture.docx'
   with zipfile.ZipFile(path,'w') as z:
    z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+''.join(paragraphs)+'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>')
   doc=word.Documents.Open(str(path.resolve()),ReadOnly=True,AddToRecentFiles=False)
   try:
    doc.Repaginate();records=c.get_paragraph_records(doc);headings,hchecks=c.extract_heading_structure(records,front+1,rules)
    checks=c.check_front_page_and_body(doc,records,headings,rules,int(doc.ComputeStatistics(2)))+hchecks
    check('Word actual layout '+kind,not [x for x in checks if x.status=='FAIL'])
    doc.ExportAsFixedFormat(str((page/'fixture.word.pdf').resolve()),17)
   finally:doc.Close(False)
 finally:word.Quit()

if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--word',action='store_true');p.add_argument('--report',type=Path);a=p.parse_args()
 with tempfile.TemporaryDirectory() as tmp:
  base=Path(tmp)
  if a.word:word_suite(base)
  else:portable(base)
 report={'version':'2026-10-22','passed':len(RESULTS),'windows_word_executed':a.word,'tests':RESULTS}
 print(json.dumps(report,ensure_ascii=False,indent=2))
 if a.report:a.report.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
