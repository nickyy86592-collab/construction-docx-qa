"""2026-10-22 regressions; --word additionally requires real Windows Word COM."""
from pathlib import Path
import argparse,copy,importlib.util,json,sys,tempfile,contextlib,io
COMMON=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(COMMON))
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
 review={'technical_review':{'version':1,'checks':{k:{'status':'reviewed','conclusion':'fixture'} for k in f.TECHNICAL_REVIEW_ITEMS}},'construction_language_review':{'reviewed':True,'unresolved_count':0,'conclusion':'施工语言核查'},'scoring_review':{'conditions':[condition],'conclusion':'本模块责任复核'}}
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
    props='<w:pPr><w:outlineLvl w:val="'+str(level-1)+'"/></w:pPr>' if level else '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9638"/></w:tabs></w:pPr>'
    paragraphs.append('<w:p>'+props+'<w:r><w:t xml:space="preserve">'+escape(text).replace('\t','</w:t><w:tab/><w:t xml:space="preserve">')+'</w:t></w:r></w:p>')
   def newpage():paragraphs.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
   if front:
    para('项目名称：格式适配合成测试')
    if kind=='multi':newpage()
    para('第一章 施工部署\t'+str(front+1))
    if kind=='multi':
     para('第一节 施工准备\t4');newpage();para('一、工序检查\t4')
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
    toc_checks=c.check_toc_layout_word(doc,records,headings,rules)
    check('Word actual TOC alignment '+kind,not [x for x in toc_checks if x.status in ('FAIL','WARN')])
    doc.ExportAsFixedFormat(str((page/'fixture.word.pdf').resolve()),17)
   finally:doc.Close(False)
  # Real Word integration fixture: never executed by the portable branch.
  # Keep evidence outside TemporaryDirectory so exported pages remain reviewable.
  import os
  from dataclasses import asdict
  native=Path(os.environ.get('WORD_TOC_EVIDENCE_DIR',str(Path.cwd()/'word-toc-evidence'))).resolve()
  native.mkdir(parents=True,exist_ok=True)
  seed=native/'native-toc.seed.docx';saved=native/'native-toc.saved.docx'
  chapter_titles=[];native_paragraphs=[]
  section_xml='<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr>'
  def native_para(text,level=None):
   props=('<w:pPr><w:outlineLvl w:val="'+str(level-1)+'"/></w:pPr>') if level else ''
   native_paragraphs.append('<w:p>'+props+'<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>'+escape(text)+'</w:t></w:r></w:p>')
  native_para('项目名称：原生自动目录真实Word集成样例');native_para('目录')
  marker='NATIVE_TOC_INSERT_HERE';native_para(marker)
  # Standard-library OOXML seed avoids adding runtime dependencies to Word runner.
  for chapter in range(1,21):
   if chapter<=3:native_paragraphs.append('<w:p><w:pPr>'+section_xml+'</w:pPr></w:p>')
   title=f'第{chapter}章 '+('复杂条件下施工工序与资源接口的统筹安排及质量安全检验控制要求'*3)
   chapter_titles.append(title);native_para(title,1)
   native_para('完成材料报审及工作面核验后组织作业，检查合格后移交。')
   native_para('第一节 施工准备及跨专业资源接口检查',2)
   native_para('结合现场工作面条件组织工序交接。')
   native_para('一、过程检查与验收记录',3)
   native_para('施工完成后按批准方案进行复查并形成验收记录。')
  with zipfile.ZipFile(seed,'w',zipfile.ZIP_DEFLATED) as z:
   z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
   z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
   z.writestr('word/document.xml','<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+''.join(native_paragraphs)+section_xml+'</w:body></w:document>')
  doc=word.Documents.Open(str(seed),ReadOnly=False,AddToRecentFiles=False)
  try:
   # Bind actual built-in heading styles before creating native TOC.
   for paragraph_index in range(1,int(doc.Paragraphs.Count)+1):
    paragraph=doc.Paragraphs(paragraph_index);level=int(paragraph.OutlineLevel)
    if level in (1,2,3):
     paragraph.Range.Style=doc.Styles(-level-1);paragraph.Range.Font.Size=12
   # Restart displayed page numbers in body sections: physically short fixture,
   # genuine one-, two-, and three-digit PAGEREF results (2, 12, 123).
   for section_index,start_number in ((2,2),(3,12),(4,123)):
    footer=doc.Sections(section_index).Footers(1);footer.LinkToPrevious=False
    numbers=footer.PageNumbers;numbers.RestartNumberingAtSection=True;numbers.StartingNumber=start_number
    if not int(numbers.Count):numbers.Add(PageNumberAlignment=2,FirstPage=True)
   anchor=doc.Content.Duplicate
   if not anchor.Find.Execute(FindText=marker,Forward=True,Wrap=0,MatchWildcards=False,Format=False):
    raise AssertionError('native TOC insertion marker absent')
   anchor.Text='';anchor.Collapse(1)
   # Positional arguments follow Word's complete Add signature (late-bound COM).
   toc=doc.TablesOfContents.Add(anchor,True,1,3,False,'',True,True,'',True,False,True)
   toc.TabLeader=1
   setup=doc.Sections(1).PageSetup
   width=float(setup.PageWidth)-float(setup.LeftMargin)-float(setup.RightMargin)
   for level in (1,2,3):
    style=doc.Styles(-19-level)  # Word built-in wdStyleTOC1/2/3, locale-independent.
    style.Font.Size=11
    fmt=style.ParagraphFormat;fmt.OutlineLevel=10;fmt.CharacterUnitLeftIndent=0;fmt.LeftIndent=(level-1)*14
    fmt.FirstLineIndent=0;fmt.RightIndent=0;fmt.KeepWithNext=False;fmt.KeepTogether=True
    fmt.LineSpacingRule=0;fmt.SpaceBefore=0;fmt.SpaceAfter=3
    fmt.TabStops.ClearAll();fmt.TabStops.Add(Position=width,Alignment=2,Leader=1)
   previous=None;stable=False
   for attempt in range(6):
    doc.Fields.Update();doc.Repaginate();toc.Update();doc.Repaginate()
    first_body=doc.Range(int(doc.Sections(2).Range.Start),int(doc.Sections(2).Range.Start))
    signature=(int(first_body.Information(3)),str(toc.Range.Text))
    if signature==previous:stable=True;break
    previous=signature
   if not stable:raise AssertionError('native TOC pagination did not stabilize after six updates')
   doc.SaveAs2(str(saved),FileFormat=12)
  finally:doc.Close(False)
  # Reopen the actual saved delivery copy; do not update fields in this audit.
  doc=word.Documents.Open(str(saved),ReadOnly=True,AddToRecentFiles=False)
  try:
   doc.Repaginate();doc.ActiveWindow.View.Type=3
   toc=doc.TablesOfContents(1);records=c.get_paragraph_records(doc)
   body_start_pos=int(doc.Sections(2).Range.Start)
   first_body=next(r for r in records if r['start']>=body_start_pos and r['outline']==1)
   body_start=int(first_body['page']);front=body_start-1
   toc_entries=[r for r in records if r['end']>int(toc.Range.Start) and r['start']<int(toc.Range.End)
                and c.classify_heading(c.normalize_toc_title(r['text']))[0] in (1,2,3)]
   if not toc_entries:raise AssertionError('Word-created TOC has no recognized entries')
   toc_span=[min(r['page'] for r in toc_entries),max(r['page'] for r in toc_entries)]
   page_rule=native/'项目页数硬审.py'
   page_rule.write_text(f'EXPECTED_BODY_PAGES=50\nPAGE_TOLERANCE=1\nFRONT_MATTER_PAGES={front}\n',encoding='utf-8')
   fm={'cover_required':True,'toc_levels':[1,2,3],'cover_page_range':[1,1],'toc_page_range':toc_span}
   rules=c.load_rules(write(native/'format.json',{'front_matter':fm}),native)
   headings,hchecks=c.extract_heading_structure(records,body_start,rules)
   front_checks=c.check_front_page_and_body(doc,records,headings,rules,int(doc.ComputeStatistics(2)))+hchecks
   static_checks=c.check_toc_layout_static(saved,rules)
   toc_checks=c.check_toc_layout_word(doc,records,headings,rules)
   wrap_rows=[];indent_rows=[]
   for rec in toc_entries:
    p=doc.Paragraphs(rec['index']);entry=c.paragraph_text_range(doc,rec['index']);raw=c.clean_text(entry.Text)
    token=c.TOC_PAGE_RE.search(raw)
    if not token:raise AssertionError('native TOC entry has no page-number token')
    number=entry.Duplicate
    if not number.Find.Execute(token[1],True,False,False,False,False,False,0,False):
     raise AssertionError('native TOC page token cannot be measured')
    doc.ActiveWindow.ScrollIntoView(entry,True)
    first=doc.Range(int(entry.Start),int(entry.Start));last=doc.Range(int(number.End),int(number.End))
    level=c.classify_heading(c.normalize_toc_title(rec['text']))[0]
    wrap_rows.append({'paragraph':rec['index'],'start_y':float(first.Information(6)),
                     'number_y':float(number.Information(6)),'line_count':int(entry.ComputeStatistics(1))})
    indent_rows.append({'level':level,'left_indent':float(p.Format.LeftIndent),'right_x':float(last.Information(5))})
   chapter_pages=[]
   for title in chapter_titles[:3]:
    h=next(h for h in headings if h['text']==title)
    chapter_pages.append(int(doc.Paragraphs(h['index']).Range.Information(1)))
   # Export before assertions: a failed geometry check still leaves visual evidence.
   doc.ExportAsFixedFormat(str(native/'native-toc.word.pdf'),17)
   evidence={'engine':'Microsoft Word','reopened_saved_copy':str(saved),'body_start_physical_page':body_start,
             'toc_page_range':toc_span,'chapter_displayed_pages':chapter_pages,
             'wrap_measurements':wrap_rows,'indent_measurements':indent_rows,
             'checks':[asdict(x) for x in front_checks+static_checks+toc_checks]}
   (native/'native-toc.measurements.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2),encoding='utf-8')
   check('Word native TOC persisted after Add Update SaveAs reopen',int(doc.TablesOfContents.Count)==1 and len(toc_entries)==60)
   check('Word native TOC current front matter and navigation integration',not [x for x in front_checks if x.status=='FAIL'])
   check('Word native TOC static properties from saved cache',not [x for x in static_checks if x.status in ('FAIL','WARN')])
   # Ambiguous field-boundary coordinates are WARN, never a native visual PASS.
   unresolved=[detail for row in toc_checks if row.status=='WARN' for detail in (row.details or []) if '嵌套域光标坐标' not in detail.get('reason','')]
   check('Word native TOC geometry branches and current page numbers',not [x for x in toc_checks if x.status=='FAIL'] and not unresolved)
   check('Word native TOC spans multiple real pages',toc_span[1]>toc_span[0])
   check('Word native TOC natural long-title wraps',any(r['line_count']>=2 and r['number_y']>r['start_y']+6 for r in wrap_rows))
   check('Word native TOC one two three digit displayed pages',chapter_pages==[2,12,123])
   lefts={level:{round(r['left_indent'],1) for r in indent_rows if r['level']==level} for level in (1,2,3)}
   check('Word native TOC three increasing left indents',all(len(lefts[level])==1 for level in lefts) and next(iter(lefts[1]))<next(iter(lefts[2]))<next(iter(lefts[3])))
   rights=[r['right_x'] for r in indent_rows]
   field_warnings=[detail for row in toc_checks if row.status=='WARN' for detail in (row.details or []) if '嵌套域光标坐标' in detail.get('reason','')]
   check('Word native TOC right edge or explicit render-review warning',all(x>=0 for x in rights) and (max(rights)-min(rights)<=3 or bool(field_warnings)))
  finally:doc.Close(False)
 finally:word.Quit()


def toc_portable_cases(base):
 from docx import Document
 from docx.enum.text import WD_TAB_ALIGNMENT,WD_TAB_LEADER
 from docx.enum.style import WD_STYLE_TYPE
 from docx.shared import Pt
 from docx.oxml import OxmlElement
 from docx.oxml.ns import qn
 import zipfile
 work=base/'toc';work.mkdir()
 rules={'front_matter':{'toc_levels':[1,2,3]}}
 def make(name,mode='valid',entries=None):
  entries=entries or [(1,'第一章 施工部署','2')]
  d=Document();d.add_paragraph('测试项目施工组织设计');d.add_paragraph('目录')
  if mode=='gutter':d.sections[0].gutter=Pt(10)
  width=(d.sections[0].page_width-d.sections[0].left_margin-d.sections[0].right_margin)/12700
  parent=d.styles.add_style('TOC-base-test',WD_STYLE_TYPE.PARAGRAPH)
  parent.paragraph_format.tab_stops.add_tab_stop(Pt(width),WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.DOTS)
  child=d.styles.add_style('TOC-child-test',WD_STYLE_TYPE.PARAGRAPH);child.base_style=parent
  for i,(level,title,page) in enumerate(entries):
   if mode=='manual_dots':text=title+'.'*20+page
   elif mode=='spaces':text=title+'     '+page
   else:text=title+'\t'+page
   p=d.add_paragraph(text,style=child if mode in ('inherit','clear','multi_tabs') else None)
   p.paragraph_format.left_indent=Pt((level-1)*12)
   if mode=='manual_plus_tab':p.text=title+'.'*20+'\t'+page
   if mode not in ('manual_dots','spaces','no_stop','inherit','clear','multi_tabs'):
    p.paragraph_format.tab_stops.add_tab_stop(Pt(width-12 if mode=='wrong_position' else width),WD_TAB_ALIGNMENT.LEFT if mode=='left' else WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.SPACES if mode=='no_leader' else WD_TAB_LEADER.DOTS)
   if mode=='multi_tabs':p.paragraph_format.tab_stops.add_tab_stop(Pt(18),WD_TAB_ALIGNMENT.LEFT)
   if mode=='clear':
    tabs=OxmlElement('w:tabs');t=OxmlElement('w:tab');t.set(qn('w:val'),'clear');t.set(qn('w:pos'),str(int(round(width*20))));tabs.append(t);p._p.get_or_add_pPr().append(tabs)
   if mode=='across_pages' and i==0:d.add_page_break()
  for level,title,page in entries:d.add_heading(title,level=level);d.add_paragraph('实际工序检查与材料报审要求。')
  path=work/(name+'.docx');d.save(path);return path
 def checkdoc(name,path,expected,rr=rules):
  results=c.check_toc_layout_static(path,rr)
  status='FAIL' if any(x.status=='FAIL' for x in results) else 'WARN' if any(x.status=='WARN' for x in results) else 'PASS'
  check('TOC '+name,status==expected)
 for mode in ('valid','inherit'):
  checkdoc(mode,make(mode,mode),'PASS')
 for mode in ('manual_dots','spaces','manual_plus_tab','no_stop','left','no_leader','wrong_position','clear'):
  checkdoc(mode,make(mode,mode),'FAIL')
 entries=[(1,'第一章 施工部署','2'),(2,'第一节 施工准备','12'),(3,'一、检查与验收','123')]
 checkdoc('three levels and 1/2/3-digit pages',make('three_levels',entries=entries),'PASS')
 checkdoc('multi-page entries',make('multi_page','across_pages',entries),'PASS')
 long=[(1,'第一章 '+('针对复杂条件的施工工序、资源接口与质量检验控制要求'*8),'123')]
 checkdoc('long title retains real tab',make('long',entries=long),'PASS')
 checkdoc('gutter requires actual layout',make('gutter','gutter'),'WARN')
 checkdoc('extra stops need Word disposition',make('extra','multi_tabs'),'WARN')
 no_body=Document();no_body.add_paragraph('目录');no_body.add_paragraph('第一章 施工部署 2');path=work/'unlocatable.docx';no_body.save(path)
 checkdoc('unlocatable boundary cannot pass',path,'WARN')
 checkdoc('no TOC exception',path,'PASS',{'front_matter':{'toc_levels':[]}})
 exception={'front_matter':{'toc_levels':[1],'toc_layout':{'leader':'none','source_ref':'有效补遗第2条：目录不使用引导符'}}}
 checkdoc('tender leader exception honored',make('no_leader_exception','no_leader'),'PASS',exception)
 for cfg in ({'leader':'none'},{'show_page_numbers':False},{'right_align':False},{'right_inset_pt':12}):
  try:c.toc_layout({'front_matter':{'toc_layout':cfg}});ok=False
  except c.AuditError:ok=True
  check('TOC override requires source '+str(cfg),ok)
 for cfg in ({'leader':'unknown'},{'right_inset_pt':float('nan')},{'right_inset_pt':True},{'right_align':'yes'}):
  try:c.toc_layout({'front_matter':{'toc_layout':cfg}});ok=False
  except c.AuditError:ok=True
  check('TOC invalid rule blocked '+str(cfg),ok)
 # Native TOC cached results remain detectable through hyperlinks and field code runs.
 auto=make('auto_cache');parts={}
 with zipfile.ZipFile(auto) as z:parts={name:z.read(name) for name in z.namelist()}
 import xml.etree.ElementTree as ET
 W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
 root=ET.fromstring(parts['word/document.xml']);body=root.find(W+'body');p=next(p for p in body.findall(W+'p') if any(t.text=='第一章 施工部署' for t in p.iter(W+'t')))
 start=ET.Element(W+'r');begin=ET.SubElement(start,W+'fldChar');begin.set(W+'fldCharType','begin');p.insert(1,start)
 instr=ET.Element(W+'r');ET.SubElement(instr,W+'instrText').text=' TOC \\o "1-3" \\h ';p.insert(2,instr)
 sep=ET.Element(W+'r');ET.SubElement(sep,W+'fldChar').set(W+'fldCharType','separate');p.insert(3,sep)
 end=ET.SubElement(p,W+'r');ET.SubElement(end,W+'fldChar').set(W+'fldCharType','end')
 parts['word/document.xml']=ET.tostring(root)
 with zipfile.ZipFile(auto,'w',zipfile.ZIP_DEFLATED) as z:
  for name,data in parts.items():z.writestr(name,data)
 checkdoc('native TOC field cache with runs',auto,'PASS')
 # Missing section dimensions must be reviewed rather than guessed.
 root=ET.fromstring(parts['word/document.xml']);sec=root.find('.//'+W+'sectPr');sec.remove(sec.find(W+'pgSz'));parts['word/document.xml']=ET.tostring(root)
 unknown=work/'unknown_width.docx'
 with zipfile.ZipFile(unknown,'w',zipfile.ZIP_DEFLATED) as z:
  for name,data in parts.items():z.writestr(name,data)
 checkdoc('unknown width WARN',unknown,'WARN')

 # Simulate Word interface responses to test failure handling; NOT actual Word layout.
 from types import SimpleNamespace as NS
 def word_case(x=550,number_y=100,title_y=100,actual=2,alignment=2,leader=1,position=500,gutter=0):
  class Range:
   def __init__(self,start=0,end=20,role='toc'):
    self.Start=start;self.End=end;self.role=role
    self.Sections=lambda n:NS(PageSetup=NS(PageWidth=600,LeftMargin=50,RightMargin=50,Gutter=gutter,TextColumns=NS(Count=1)))
   @property
   def Text(self):return '第一章 施工部署\t2\r'
   @property
   def Duplicate(self):return Range(self.Start,self.End)
   @property
   def Find(self):return NS(Execute=self.find)
   def find(self,FindText,*args,**kw):
    self.Start,self.End=(12,13) if FindText=='2' else (7,8);self.role='number' if FindText=='2' else 'title';return True
   def Information(self,n):
    return {1:actual,3:1,5:x,6:number_y if self.role=='number' else title_y}.get(n,0)
  class Tabs:
   Count=1
   def __call__(self,n):return NS(Position=position,Alignment=alignment,Leader=leader)
  p=NS(Range=Range(),Format=NS(TabStops=Tabs(),LineSpacing=18))
  doc=NS(Paragraphs=lambda n:p,Range=lambda start,end:Range(start,end,'caret' if start==end else 'title'))
  rec={'index':1,'page':1,'text':'第一章 施工部署 2'};head={'index':2,'page':2,'text':'第一章 施工部署','level':1}
  return c.check_toc_layout_word(doc,[rec],[head],{'body_start_page':2,'front_matter':{'toc_levels':[1]}})
 def result(rows):return 'FAIL' if any(r.status=='FAIL' for r in rows) else 'WARN' if any(r.status=='WARN' for r in rows) else 'PASS'
 check('TOC Word-interface valid response',result(word_case())=='PASS')
 for name,kw in [('wrong actual right edge',{'x':530}),('outside page',{'x':620}),('page number orphan line',{'number_y':125}),('stale displayed page',{'actual':3}),('left stop',{'alignment':0}),('wrong leader',{'leader':0}),('wrong stop position',{'position':480})]:
  check('TOC Word-interface '+name+' rejected',result(word_case(**kw))=='FAIL')
 check('TOC Word-interface gutter needs review',result(word_case(gutter=10))=='WARN')
 check('TOC Word-interface unavailable coordinate WARN',result(word_case(x=-1))=='WARN')


def toc_boundary_regressions(base):
 from docx import Document
 from docx.enum.text import WD_TAB_ALIGNMENT,WD_TAB_LEADER
 from docx.shared import Pt
 from types import SimpleNamespace as NS
 def status(rows):return 'FAIL' if any(r.status=='FAIL' for r in rows) else 'WARN' if any(r.status=='WARN' for r in rows) else 'PASS'
 def fixture(name,text,stops=None):
  d=Document();d.add_paragraph('测试项目');p=d.add_paragraph(text)
  width=(d.sections[0].page_width-d.sections[0].left_margin-d.sections[0].right_margin)/12700
  for pos,a,l in stops or [(width,WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.DOTS)]:p.paragraph_format.tab_stops.add_tab_stop(Pt(pos),a,l)
  d.add_heading('第一章 施工部署',1);path=base/(name+'.docx');d.save(path);return path
 normal={'front_matter':{'toc_levels':[1]}}
 hidden={'front_matter':{'toc_levels':[1],'toc_layout':{'show_page_numbers':False,'source_ref':'补遗：隐藏目录页码'}}}
 left={'front_matter':{'toc_levels':[1],'toc_layout':{'right_align':False,'source_ref':'补遗：目录页码左对齐'}}}
 check('TOC hidden numbers still present rejected',status(c.check_toc_layout_static(fixture('hidden_bad','第一章 施工部署\t2'),hidden))=='FAIL')
 check('TOC hidden numbers absent accepted',status(c.check_toc_layout_static(fixture('hidden_ok','第一章 施工部署'),hidden))=='PASS')
 stops=[(18,WD_TAB_ALIGNMENT.LEFT,WD_TAB_LEADER.SPACES),(400,WD_TAB_ALIGNMENT.LEFT,WD_TAB_LEADER.DOTS)]
 check('TOC unused left stop cannot cause false FAIL',status(c.check_toc_layout_static(fixture('left_stops','第一章 施工部署\t2',stops),left))=='WARN')
 def front(toc,body,rr=normal):
  records=[{'index':1,'page':1,'text':'测试项目','outline':10,'start':0,'end':5},
   {'index':2,'page':1,'text':toc,'outline':10,'start':6,'end':25},
   {'index':3,'page':2,'text':body,'outline':1,'start':26,'end':40,'level':1},
   {'index':4,'page':2,'text':'施工正文内容。','outline':10,'start':41,'end':50}]
  return c.check_front_page_and_body(NS(),records,[records[2]],dict(rr,front_page=1,body_start_page=2),2)
 check('TOC Roman page actual front-title integration',status(front('第一章 施工部署 iv','第一章 施工部署'))=='PASS')
 check('TOC numeric title suffix preserved in body',status(front('第一章 质量管理体系 ISO 9001 2','第一章 质量管理体系 ISO 9001'))=='PASS')
 check('TOC hidden numeric title suffix preserved',status(front('第一章 ISO 9001','第一章 ISO 9001',hidden))=='PASS')
 check('TOC Roman title suffix preserved in body',status(front('第一章 IV 2','第一章 IV'))=='PASS')
 check('TOC required body title numeric suffix preserved',c.check_required_level1([{'text':'第一章 ISO 9001','level':1}],{'required_level1_titles':['第一章 ISO 9001']}).status=='PASS')
 # First native TOC paragraph may otherwise include the entire outer field.
 class OuterFields:
  Count=1
  def __call__(self,n):return NS(Type=13,Result=NS(Start=45))
 original=NS(Start=10,End=999,Fields=OuterFields())
 bounded=NS(Paragraphs=lambda n:NS(Range=original if n==1 else NS(Start=160)),Range=lambda a,b:NS(Start=a,End=b))
 rr=c.paragraph_text_range(bounded,1)
 check('TOC first native paragraph excludes outer code and later entries',(rr.Start,rr.End)==(45,160))
 # Controlled COM-interface models exercise control flow, not Word pagination.
 def wordcase(text='第一章 施工部署\t2',stops=((500,2,1),),mirror=False,field_count=0,indent=0,x=550,rr=normal):
  calls=[]
  class Range:
   def __init__(self,start=0,end=20,role='toc'):self.Start=start;self.End=end;self.role=role;self.Fields=NS(Count=field_count);self.Sections=lambda n:NS(PageSetup=NS(PageWidth=600,LeftMargin=50,RightMargin=50,Gutter=0,MirrorMargins=mirror,TextColumns=NS(Count=1)))
   @property
   def Text(self):return text+'\r'
   @property
   def Duplicate(self):return Range(self.Start,self.End)
   @property
   def Find(self):return NS(Execute=self.find)
   def find(self,FindText,*args,**kw):
    calls.append(dict(kw,positional=args));self.Start,self.End=(12,13) if FindText=='2' else (7,8);self.role='number' if FindText=='2' else 'title';return True
   def Information(self,n):return {1:2,3:1,5:x,6:100}.get(n,0)
  class Tabs:
   Count=len(stops)
   def __call__(self,n):pos,a,l=stops[n-1];return NS(Position=pos,Alignment=a,Leader=l)
  para=NS(Range=Range(),Format=NS(TabStops=Tabs(),LineSpacing=18,LeftIndent=indent))
  doc=NS(Paragraphs=lambda n:para,Range=lambda start,end:Range(start,end))
  rows=c.check_toc_layout_word(doc,[{'index':1,'page':1,'text':text}],[{'index':2,'page':2,'text':'第一章 施工部署','level':1}],dict(rr,body_start_page=2))
  return rows,calls
 check('TOC Word hidden visible numbers rejected',status(wordcase(rr=hidden)[0])=='FAIL')
 rows,_=wordcase(text='第一章 施工部署',rr=hidden)
 check('TOC Word hidden no false measurement claim',status(rows)=='PASS' and '实际测量完成' not in rows[-1].message)
 check('TOC Word unused left stop needs review only',status(wordcase(stops=((18,0,0),(400,0,1)),rr=left)[0])=='WARN')
 check('TOC Word mirror margins require actual review',status(wordcase(mirror=True)[0])=='WARN')
 check('TOC Word unknown page label WARN',status(wordcase(text='第一章 施工部署\t甲')[0])=='WARN')
 check('TOC Word native indent-shifted field coordinate WARN',status(wordcase(field_count=1,indent=14,x=536)[0])=='WARN')
 check('TOC Word custom actual wrong edge still FAIL',status(wordcase(indent=14,x=536)[0])=='FAIL')
 check('TOC Word native unrelated wrong edge still FAIL',status(wordcase(field_count=1,indent=14,x=520)[0])=='FAIL')
 check('TOC Word Find explicitly searches backward and stops in range',all(k['positional'][5] is False and k['positional'][6]==0 for k in wordcase()[1]))
 check('TOC Word Find ignores inherited wildcard/format',all(k['positional'][2] is False and k['positional'][7] is False for k in wordcase()[1]))

if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--word',action='store_true');p.add_argument('--report',type=Path);a=p.parse_args()
 with tempfile.TemporaryDirectory() as tmp:
  base=Path(tmp)
  if a.word:word_suite(base)
  else:
   portable(base)
   toc_portable_cases(base)
   toc_boundary_regressions(base)
 report={'version':'2026-10-22-r6','passed':len(RESULTS),'windows_word_executed':a.word,'tests':RESULTS}
 print(json.dumps(report,ensure_ascii=False,indent=2))
 if a.report:a.report.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
