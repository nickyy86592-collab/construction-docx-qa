from pathlib import Path
import importlib.util,sys,hashlib,json,copy,subprocess
from PIL import Image,ImageDraw
from docx import Document
from docx.shared import Inches,Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
COMMON=Path(__file__).resolve().parents[1]
import tempfile,shutil
_temporary=tempfile.TemporaryDirectory()
ROOT=Path(_temporary.name)/'skill';(ROOT/'scripts').mkdir(parents=True)
for name in ('Word正式交付硬审（不含页数）.py','编制流程门禁.py','评分响应核验.py','visual_policy.py'):
    shutil.copy2(COMMON/name,ROOT/'scripts'/name)
active_pages=list((COMMON.parents[1]/'input/project').glob('*页数硬审*.py'))
assert len(active_pages)==1,'One active project page script is required'
shutil.copy2(active_pages[0],ROOT/'scripts'/active_pages[0].name)
OUT=Path(_temporary.name)/'tests';OUT.mkdir()
sys.path.insert(0,str(ROOT/'scripts'))
def load(name,file):
 spec=importlib.util.spec_from_file_location(name,ROOT/'scripts'/file);m=importlib.util.module_from_spec(spec);sys.modules[name]=m;spec.loader.exec_module(m);return m
c=load('common','Word正式交付硬审（不含页数）.py');f=load('flow','编制流程门禁.py')
names=c.FIGURE_NAMES
TARGET=f.discover_page_script()[1]['EXPECTED_BODY_PAGES']
TOLERANCE=f.discover_page_script()[1]['PAGE_TOLERANCE']
FRONT=f.discover_page_script()[1]['FRONT_MATTER_PAGES']
policy={'version':1,
        'real_photos':{'decision':'required','reason':'Synthetic fixture photo','replacement_figures':[]},
        'color_figures':{'mode':'color'},
        'additional_figures':{'names':[],'reason':'Isolate the three legacy diagrams',
                              'override':{'source_type':'prompt','source_ref':'synthetic fixture','quote':'Only test the base diagrams'}}}
rules={'visual_policy':policy,'required_figures':[{'name':n,'keywords':[n],'hard':True,'color_required':True} for n in names],'required_real_photo_slots':[{'name':'现场沟塘'}],'photo_delivery_stage':'draft','font_rules':{},'required_level1_titles':[]}
def make_doc(path, gray=False,photo=True,editorial=False,same=False,border=True,frame_text=True):
 d=Document();d.add_paragraph('测试项目施工组织设计');toc=d.add_paragraph('第一章 施工部署\t2')
 from docx.enum.text import WD_TAB_ALIGNMENT,WD_TAB_LEADER
 toc.paragraph_format.tab_stops.add_tab_stop(Inches(6),WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.DOTS)
 d.add_heading('第一章 施工部署',level=1)
 d.add_heading('第一节 材料进场与工序准备',level=2)
 d.add_heading('一、路面材料报审',level=3)
 d.add_paragraph('路面强度按4.5MPa进行材料报审。')
 for i,n in enumerate(names):
  im=Image.new('RGB',(600,900),'white');dr=ImageDraw.Draw(im)
  for j,col in enumerate(['#228bcc','#d94235','#38a76c']):dr.rectangle((70,70+j*200,530,220+j*200), fill='gray' if gray else col)
  dr.text((50,800),'test diagram '+str(0 if same else i),fill='black')
  p=OUT/('gray' if gray else 'color');p=p.with_name(p.name+str(0 if same else i)+'.png');im.save(p)
  d.add_picture(str(p),width=Inches(4));d.add_paragraph('图'+str(i+1)+' '+n)
 if photo:
  t=d.add_table(rows=1,cols=1);t.rows[0].height=Pt(100)
  if border:
   borders=OxmlElement('w:tblBorders')
   for edge in ['top','bottom','left','right']:
    e=OxmlElement('w:'+edge);e.set(qn('w:val'),'single');borders.append(e)
   t._tbl.tblPr.append(borders)
  t.cell(0,0).text='【实景图占位：现场沟塘】'+('需插入现场照片' if frame_text else '')
  d.add_paragraph('图4 现场沟塘照片')
 if editorial:d.add_paragraph('本章应写清楚得分点，引用知识库。')
 d.save(path)
def statuses(path,rr=rules):return [x.name for x in c.portable_checks(path,c.load_rules(None)|rr) if x.status=='FAIL']
results=[]
def expect(label,yes):
 assert yes,label
 results.append(label)
passdoc=OUT/'valid.docx';make_doc(passdoc)
expect('彩图三图及带框照片占位通过',not statuses(passdoc))
for label,kwargs in [('黑灰三图被拦截',{'gray':True}),('照片位置缺失被拦截',{'photo':False}),('编制语言被拦截',{'editorial':True}),('三图共用同一图被拦截',{'same':True}),('占位无边框被拦截',{'border':False}),('占位缺提示被拦截',{'frame_text':False})]:
 p=OUT/(str(len(results))+'.docx');make_doc(p,**kwargs);expect(label,bool(statuses(p)))
expect('正式稿残留照片占位被拦截',bool(statuses(passdoc,rules|{'photo_delivery_stage':'final'})))
expect('空照片规则被拦截',bool(statuses(passdoc,rules|{'required_real_photo_slots':[]})))
flow_doc=OUT/'body-flow.docx'
body=Document(passdoc)
from docx.enum.style import WD_STYLE_TYPE
parent=body.styles.add_style('Locked prose',WD_STYLE_TYPE.PARAGRAPH)
parent.paragraph_format.keep_together=True
child=body.styles.add_style('Inherited lock',WD_STYLE_TYPE.PARAGRAPH)
child.base_style=parent
paragraph=body.add_paragraph('施工面检查合格后办理工序交接。',style=child)
body.save(flow_doc)
expect('继承样式的普通正文锁段被拦截','正文连续排版_分页设置' in statuses(flow_doc))
paragraph.paragraph_format.keep_together=False
body.save(flow_doc)
expect('直接关闭样式锁段后恢复自然排版','正文连续排版_分页设置' not in statuses(flow_doc))
paragraph.paragraph_format.keep_with_next=True
next_paragraph=body.add_paragraph('下道工序复核交接记录后开始。')
body.save(flow_doc)
expect('普通正文串联到下段被拦截','正文连续排版_分页设置' in statuses(flow_doc))
paragraph.paragraph_format.keep_with_next=False
next_paragraph.paragraph_format.page_break_before=True
body.save(flow_doc)
expect('普通正文段前分页被拦截','正文连续排版_分页设置' in statuses(flow_doc))
next_paragraph.paragraph_format.page_break_before=False
section=next(p for p in body.paragraphs if p.style.name=='Heading 2')
section.text='第一节 工期目标明确，计划可行'
body.save(flow_doc)
expect('自拟二级标题串句被拦截','二级标题_串句标点' in statuses(flow_doc))
locked={'chapter':'第一章 施工部署','section':section.text,'source_ref':'合成招标样例第1节'}
expect('有依据的指定节名可逐字保留','二级标题_串句标点' not in statuses(flow_doc,rules|{'locked_level2_titles':[locked]}))
expect('其他章节的锁定登记不能豁免当前节名','二级标题_串句标点' in statuses(flow_doc,rules|{'locked_level2_titles':[locked|{'chapter':'第二章 施工部署'}]}))
# CLI exit code and serialized report, not just internal function.
rules_path=OUT/'format_rules.json';rules_path.write_text(json.dumps(rules,ensure_ascii=False))
rc=subprocess.run([sys.executable,str(ROOT/'scripts/Word正式交付硬审（不含页数）.py'),str(passdoc),'--rules',str(rules_path),'--static-only','--report',str(OUT/'static.json')],capture_output=True,text=True)
expect('静态CLI及文件哈希报告通过',rc.returncode==0 and json.loads((OUT/'static.json').read_text())['document_sha256']==f.sha256(passdoc))
statepath=OUT/'state.json'
source=OUT/'source.txt';source.write_text('Synthetic test source: general 4.5 MPa, detail 4.0 MPa. Not project evidence.')
review={'document_sha256':f.sha256(passdoc),'source_checks':[{'item':n,'status':'verified','conclusion':'synthetic','evidence':[{'path':str(source),'sha256':f.sha256(source),'locator':'synthetic lines'}]} for n in f.REVIEW_ITEMS], 'latest_confirmation':{'confirmed_at':'test','reply_source':'test','conclusion':'test'},'parameters':[{'name':'强度','scope':'路面','cross_source_search':'说明和结构图','occurrences':[{'path':str(source),'sha256':f.sha256(source),'locator':'test1','value':4.5,'unit':'MPa','scope':'路面'},{'path':str(source),'sha256':f.sha256(source),'locator':'test2','value':4.0,'unit':'MPa','scope':'入户道路'}],'status':'consistent','decision':'test','body_locations':[{'location':'第一章','quote':'路面强度按4.5MPa进行材料报审。','value':4.5,'unit':'MPa','scope':'路面'}]}], 'conflicts':[], 'construction_language_review':{'all_paragraphs_reviewed':True,'conclusion':'test'},'visual_review':{'three_color_figures_reviewed':True,'caption_photo_reviewed':True,'conclusion':'test'}}
review['section_heading_review']=[{'chapter':'第一章 施工部署','section':'第一节 材料进场与工序准备',
                                   'criterion':'材料报审重点','evidence_quote':'路面强度按4.5MPa进行材料报审。'}]
review['parameters'][0]['kind']='project_fact'
review['technical_review']={'version':1,'checks':{k:{'status':'reviewed','conclusion':'synthetic'}
                                                      for k in f.TECHNICAL_REVIEW_ITEMS},
                            'reference_targets':{k:{'status':'adapted','conclusion':'synthetic'}
                                                 for k in f.REFERENCE_TARGET_ITEMS}}
rp=OUT/'review.json';rp.write_text(json.dumps(review));expect('4.5及4.0差异未解释被拦截',bool(f.validate_content_review(passdoc,rp)[0]))
review['parameters'][0].update(status='resolved',resolution_evidence='Synthetic scenario: differing scopes reviewed');rp.write_text(json.dumps(review));expect('有部位解释的溯源结构通过',not f.validate_content_review(passdoc,rp)[0])
bad=copy.deepcopy(review);bad['parameters'][0]['body_locations'][0]['quote']='正文不存在的证据';rp.write_text(json.dumps(bad));expect('虚构正文参数引用被拦截',bool(f.validate_content_review(passdoc,rp)[0]))

from argparse import Namespace

# Restore valid review after the prior failure case.
rp.write_text(json.dumps(review))
rr=c.load_rules(None)|rules
staticpath=OUT/'delivery-static.json';c.static_audit(passdoc,rules_path,staticpath)
geometry={'layout_source':'microsoft_word_page_geometry','bottom_blank_ratio':.1999}
last_page_pass={'name':'末页留白','status':'PASS','details':geometry|{'document_page':FRONT+2,'body_page':2}}
pre_final_pass={'name':c.PRE_FINAL_PAGE_CHECK,'status':'PASS','details':{
 'review_ratio':c.PRE_FINAL_PAGE_BLANK_REVIEW_RATIO,'review_pages':[],
 'pages':[geometry|{'document_page':FRONT+1,'body_page':1}]}}
word={'audit_version':c.AUDIT_VERSION,'delivery_status':'PASS','document_sha256':f.sha256(passdoc),'layout_engine':'Microsoft Word','rules':rr,'rules_sha256':c.rules_digest(rr),'photo_delivery_stage':'draft','checks':[last_page_pass,pre_final_pass]}
# Synthetic reports test schema only. They are never represented as Word execution.
wp=OUT/'synthetic-word.json';wp.write_text(json.dumps(word))
pg=OUT/'synthetic-pages.json';pg.write_text(json.dumps({'status':'PASS','page_source':'microsoft_word','document_sha256':f.sha256(passdoc),'rule':{'body_pages_target':TARGET,'page_tolerance':TOLERANCE,'front_matter_pages_excluded':FRONT}}))
sc=OUT/'synthetic-score.json';sc.write_text(json.dumps({'status':'PASS','phase':'final','document_sha256':f.sha256(passdoc),'checks':[]}))
wr=OUT/'warnings.json';wr.write_text(json.dumps({'document_sha256':f.sha256(passdoc),'items':[]}))
a=Namespace(document=passdoc,state=statepath,mode='local-review',review=rp,static_report=staticpath,word_report=wp,page_report=pg,score_report=sc,warning_review=wr)
expect('交付汇总schema正例允许编制稿',f.cmd_delivery(a)==0)
wp.write_text(json.dumps(word|{'checks':[]}))
expect('缺末页实测结果的旧版报告被拒绝',f.cmd_delivery(a)==2)
wp.write_text(json.dumps(word|{'checks':[last_page_pass|{'details':last_page_pass['details']|{'bottom_blank_ratio':.2}},pre_final_pass]}))
expect('留白恰等于20%被拒绝',f.cmd_delivery(a)==2)
wp.write_text(json.dumps(word|{'checks':[last_page_pass]}))
expect('末页达标但缺前页实测仍被拒绝',f.cmd_delivery(a)==2)
wp.write_text(json.dumps(word|{'audit_version':'1.12.0'}))
expect('r12 Word 报告不能冒充新审核',f.cmd_delivery(a)==2)
wp.write_text(json.dumps(word))
bad_heading=copy.deepcopy(review)
bad_heading['section_heading_review'][0]['evidence_quote']='正文不存在的措施'
expect('节名证据必须来自本节',bool(c.section_heading_review_errors(passdoc,bad_heading['section_heading_review'])))
expect('不能跳过逐节复核',bool(c.section_heading_review_errors(passdoc,[])))
# Rechecked language WARN must be closed even when a former report omitted that check.
saved_doc=passdoc.read_bytes();saved_reports={p:p.read_bytes() for p in (rp,staticpath,wp,pg,sc,wr)}
d=Document(passdoc);d.add_paragraph('疑似符号需结合原图核对。');d.save(passdoc)
for path in saved_reports:
 data=json.loads(path.read_text());data['document_sha256']=f.sha256(passdoc);path.write_text(json.dumps(data))
expect('新增语境WARN不能借旧静态报告绕过处置',f.cmd_delivery(a)==2)
wr.write_text(json.dumps({'document_sha256':f.sha256(passdoc),'items':[{'report':'static','check':'正文异常用语_语境复核','status':'reviewed','evidence':'Synthetic exercise: wording belongs to internal review and must be fixed before production','conclusion':'fixture only; no actual delivery'}]}))
expect('语境WARN沿用已有处置机制',f.cmd_delivery(a)==0)
passdoc.write_bytes(saved_doc)
for path,raw in saved_reports.items():path.write_bytes(raw)

wp.write_text(json.dumps(word|{'document_sha256':'old version'}));expect('交付汇总拒绝旧版本Word报告',f.cmd_delivery(a)==2);wp.write_text(json.dumps(word))
wp.write_text(json.dumps(word|{'rules_sha256':'altered'}));expect('交付汇总拒绝不同格式规则指纹',f.cmd_delivery(a)==2);wp.write_text(json.dumps(word))
pre_final_warn=copy.deepcopy(pre_final_pass)
pre_final_warn['status']='WARN'
pre_final_warn['details']['pages'][0]['bottom_blank_ratio']=.26
pre_final_warn['details']['review_pages']=[FRONT+1]
wp.write_text(json.dumps(word|{'checks':[last_page_pass,pre_final_warn]}));expect('前页留白WARN未处置不允许交付',f.cmd_delivery(a)==2)
wr.write_text(json.dumps({'document_sha256':f.sha256(passdoc),'items':[{'report':'word','check':c.PRE_FINAL_PAGE_CHECK,'status':'reviewed','evidence':'synthetic QA only','conclusion':'test'}]}));expect('处理WARN记录结构允许交付',f.cmd_delivery(a)==0)
# Force final stage while leaving placeholder. Both rule hashes match but embedded content blocks it.
finalrules=rr|{'photo_delivery_stage':'final'}
wp.write_text(json.dumps(word|{'rules':finalrules,'rules_sha256':c.rules_digest(finalrules)}))
fakestatic=json.loads(staticpath.read_text());fakestatic.update(rules_sha256=c.rules_digest(finalrules),photo_delivery_stage='final');fakep=OUT/'edited-static.json';fakep.write_text(json.dumps(fakestatic));a.static_report=fakep
expect('伪改静态报告仍不能让占位冒充正式照片',f.cmd_delivery(a)==2)

source.write_text(source.read_text()+' version changed');expect('来源文件改变使旧内容核验失效',bool(f.validate_content_review(passdoc,rp)[0]))
# No-page deployment must still block absent startup.
np=OUT/'no-page';np.mkdir(exist_ok=True);(np/'编制流程门禁.py').write_text((ROOT/'scripts/编制流程门禁.py').read_text());shutil.copy2(ROOT/'scripts/visual_policy.py',np/'visual_policy.py')
rc=subprocess.run([sys.executable,str(np/'编制流程门禁.py'),'--state',str(np/'state.json'),'init'],capture_output=True);expect('无页数脚本仍初始化启动门禁',rc.returncode==2 and (np/'state.json').exists())
rc=subprocess.run([sys.executable,str(np/'编制流程门禁.py'),'--state',str(np/'state.json'),'preflight'],capture_output=True);expect('无页数脚本不能绕过启动质询',rc.returncode==2)
print(json.dumps({'passed':len(results),'tests':results,'limits':['Synthetic fixtures only; Windows Word not executed.']},ensure_ascii=False,indent=2))
(OUT/'test-summary.json').write_text(json.dumps({'passed':len(results),'tests':results,'limits':['Synthetic fixtures only; Windows Word not executed.']},ensure_ascii=False,indent=2))

# Replace only the page script; no generic code edits for a different target/front matter.
common_hashes={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT/'scripts').glob('*.py') if '页数硬审' not in p.name}
oldpage=next((ROOT/'scripts').glob('*页数硬审*.py'));oldpage.unlink()
newpage=ROOT/'scripts'/'替换项目页数硬审.py'
newpage.write_text("EXPECTED_BODY_PAGES = 80\nPAGE_TOLERANCE = 3\nFRONT_MATTER_PAGES = 3\n",encoding='utf-8')
script,vals=f.discover_page_script()
expect('Replacing only page script changes target and exclusion',vals=={'EXPECTED_BODY_PAGES':80,'PAGE_TOLERANCE':3,'FRONT_MATTER_PAGES':3})
expect('Generic scripts stay byte-identical across project swap',all(hashlib.sha256((ROOT/'scripts'/n).read_bytes()).hexdigest()==h for n,h in common_hashes.items()))
expect('Common Word audit derives body start from replacement script',c.load_rules(None)['body_start_page']==4)
conflicting=OUT/'conflicting-format.json';conflicting.write_text(json.dumps({'body_start_page':2}))
try:c.load_rules(conflicting);ok=False
except c.AuditError:ok=True
expect('Conflicting front matter rules blocked',ok)
second=ROOT/'scripts'/'第二份项目页数硬审.py';shutil.copy2(newpage,second)
try:f.discover_page_script();ok=False
except RuntimeError:ok=True
expect('Multiple project page scripts blocked',ok)
second.unlink()
# Stage measurement is generic; new page scripts need not implement a --stage switch.
rc=subprocess.run([sys.executable,str(ROOT/'scripts/编制流程门禁.py'),'page-rule'],capture_output=True,text=True)
expect('Automatic page-rule CLI sees renamed replacement',rc.returncode==0 and json.loads(rc.stdout)['body_pages_target']==80)
rc=subprocess.run([sys.executable,str(ROOT/'scripts/编制流程门禁.py'),'run-page-audit','--document',str(passdoc),'--stage','40','--report',str(OUT/'unavailable-word.json')],capture_output=True,text=True)
expect('Non-Windows cannot fake actual Word stage measurement',rc.returncode==3)
# Structural regressions use actual DOCX blocks and visible caption caches.
def structural_checks(doc, name):
 path=OUT/(name+'.docx')
 doc.save(path)
 # Isolate structure checks so these cases do not rerun unrelated pixel analysis.
 check_rules=c.load_rules(None)|rules|{'visual_policy':None}
 return {x.name:x for x in c.portable_checks(path,check_rules)
         if x.name in ('静态三级标题体系_层级完整性','静态图号连续性','静态表号连续性')}

def numbered_tables(first='表1 进场材料',second='表2 使用计划'):
 doc=Document(passdoc)
 doc.add_heading('第二节 材料供应安排',level=2)
 doc.add_heading('一、材料进场批次',level=3)
 doc.add_heading('二、材料使用计划',level=3)
 for title in (first,second):
  if title is not None:doc.add_paragraph(title,style='Caption')
  doc.add_table(rows=1,cols=2).cell(0,0).text='材料'
 return doc

good=structural_checks(numbered_tables(),'numbered-tables')
expect('逐章逐节覆盖，三级标题数量可因节而异',
       good['静态三级标题体系_层级完整性'].status=='PASS')
expect('图表独立顺排，照片占位只计图号',
       good['静态图号连续性'].details['count']==4 and good['静态表号连续性'].details['count']==2)
for name,first,second in [('表题缺失','表1 进场材料',None),('表号重号','表1 进场材料','表1 使用计划'),
                          ('表号跳号','表1 进场材料','表3 使用计划'),('表号倒序','表2 进场材料','表1 使用计划'),
                          ('表号为占位符','表1 进场材料','表X 使用计划')]:
 expect(name,structural_checks(numbered_tables(first,second),name)['静态表号连续性'].status=='FAIL')

doc=numbered_tables()
doc.add_heading('第三节 材料验收',level=2)
doc.add_heading('第二章 施工进度',level=1)
result=structural_checks(doc,'missing-parent-children')['静态三级标题体系_层级完整性']
expect('全文已有三级标题仍逐项拒绝缺子级的章和节',
       result.status=='FAIL' and result.details['chapters_without_sections']==['第二章 施工进度']
       and result.details['sections_without_subheadings']==[{'chapter':'第一章 施工部署','section':'第三节 材料验收'}])
for name,caption in [('额外图缺题',None),('图号重号','图4 材料运输'),('图号跳号','图6 材料运输')]:
 doc=Document(passdoc)
 doc.add_picture(str(OUT/'color0.png'),width=Inches(2))
 if caption:doc.add_paragraph(caption)
 expect(name,structural_checks(doc,name)['静态图号连续性'].status=='FAIL')

doc=Document(passdoc)
paragraph=doc.add_paragraph()
for image in ('color0.png','color1.png'):
 paragraph.add_run().add_picture(str(OUT/image),width=Inches(2))
doc.add_paragraph('图5 材料运输')
expect('两幅独立图共用图题被拦截',structural_checks(doc,'shared-figure-caption')['静态图号连续性'].status=='FAIL')

manifest=json.loads((COMMON/'skill_version.json').read_text())
expect('Generic scripts match exact synced skill bytes',all(hashlib.sha256((COMMON/n).read_bytes()).hexdigest()==h for n,h in manifest['common_files_sha256'].items()))
print(json.dumps({'total_passed':len(results),'scope':'portable and interface tests; no Windows Word execution'},ensure_ascii=False))

