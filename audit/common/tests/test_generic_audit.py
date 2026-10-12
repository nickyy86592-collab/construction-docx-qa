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
for name in ('Word正式交付硬审（不含页数）.py','编制流程门禁.py','评分响应核验.py'):
    shutil.copy2(COMMON/name,ROOT/'scripts'/name)
active_pages=list((COMMON.parents[1]/'input/project').glob('*页数硬审*.py'))
assert len(active_pages)==1,'One active project page script is required'
shutil.copy2(active_pages[0],ROOT/'scripts'/active_pages[0].name)
OUT=Path(_temporary.name)/'tests';OUT.mkdir()
def load(name,file):
 spec=importlib.util.spec_from_file_location(name,ROOT/'scripts'/file);m=importlib.util.module_from_spec(spec);sys.modules[name]=m;spec.loader.exec_module(m);return m
c=load('common','Word正式交付硬审（不含页数）.py');f=load('flow','编制流程门禁.py')
names=c.FIGURE_NAMES
TARGET=f.discover_page_script()[1]['EXPECTED_BODY_PAGES']
TOLERANCE=f.discover_page_script()[1]['PAGE_TOLERANCE']
FRONT=f.discover_page_script()[1]['FRONT_MATTER_PAGES']
rules={'required_figures':[{'name':n,'keywords':[n],'hard':True,'color_required':True} for n in names],'required_real_photo_slots':[{'name':'现场沟塘'}],'photo_delivery_stage':'draft','font_rules':{},'required_level1_titles':[]}
def make_doc(path, gray=False,photo=True,editorial=False,same=False,border=True,frame_text=True):
 d=Document();d.add_paragraph('第一章 施工部署');d.add_paragraph('路面强度按4.5MPa进行材料报审。')
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
# CLI exit code and serialized report, not just internal function.
rules_path=OUT/'format_rules.json';rules_path.write_text(json.dumps(rules,ensure_ascii=False))
rc=subprocess.run([sys.executable,str(ROOT/'scripts/Word正式交付硬审（不含页数）.py'),str(passdoc),'--rules',str(rules_path),'--static-only','--report',str(OUT/'static.json')],capture_output=True,text=True)
expect('静态CLI及文件哈希报告通过',rc.returncode==0 and json.loads((OUT/'static.json').read_text())['document_sha256']==f.sha256(passdoc))
statepath=OUT/'state.json';statepath.unlink(missing_ok=True);f.cmd_init(statepath,False);state=f.load_state(statepath)
state['plan']={'planned_total_body_pages':TARGET,'chapter_plan':[{'chapter':'施工方法','basis':'技术复杂度','page_range':[TARGET*.45,TARGET*.6]},{'chapter':'保障措施','basis':'评分权重','page_range':[TARGET*.38,TARGET*.55]}],'figure_table_pages_estimate':5,'reserve_pages':1,'planning_notes':'按评分权重和技术复杂度计算总体容量及机动量','checkpoint_expected_page_ranges':{'40':[TARGET*.35,TARGET*.48],'75':[TARGET*.68,TARGET*.82]}}
f.save_state(statepath,state)
expect('未答最新文件且启动证据缺失被拦截',f.cmd_preflight(statepath)==2)
source=OUT/'source.txt';source.write_text('Synthetic test source: general 4.5 MPa, detail 4.0 MPa. Not project evidence.')
state['startup']={'sources':[{'path':str(source),'sha256':f.sha256(source),'role':'tender'}],'latest_updates':{'status':'confirmed_none','asked_at':'2026-10-12','reply_text':'测试用答复','reply_source':'测试会话记录'},'critical_questions':[],'questions_review':'测试资料范围完整','batch_plan':{'count':1,'reason':'一次完成','batches':['全文']},'word_channel':'unavailable','word_capability_evidence':'Linux不能调用Word','real_photo_slots':[{'name':'现场沟塘','location':'第一章','basis':'现场技术约束'}]}
f.save_state(statepath,state);expect('真实结构启动记录通过',f.cmd_preflight(statepath)==0)
state=f.load_state(statepath);bad=copy.deepcopy(state);bad['startup']['batch_plan']={'count':4,'reason':'分批','batches':[1,2,3,4]};f.save_state(statepath,bad);expect('超过三批被拦截',f.cmd_preflight(statepath)==2)
f.save_state(statepath,state)
expect('无阶段记录禁止最终硬审',f.cmd_final_ready(statepath)==2)
# Word report fixtures are synthetic schema tests, not actual Word measurements.
for stage,actual in [(40,round(TARGET*.4)),(75,round(TARGET*.75))]:
 rpt={'page_source':'microsoft_word','status':'MEASURED','stage_percent':stage,'document_sha256':f.sha256(passdoc),'actual':{'word_body_pages':actual,'word_version':'synthetic'},'project_rule':{'target_body_pages':TARGET,'page_tolerance':TOLERANCE,'front_matter_pages_excluded':FRONT}}
 rp=OUT/f'synthetic-{stage}.json';rp.write_text(json.dumps(rpt))
 if stage==40:
  badrp=OUT/'wronghash.json';badrp.write_text(json.dumps(rpt|{'document_sha256':'wrong'}))
  try:f.cmd_checkpoint(statepath,40,badrp,None,'',False,'',passdoc);yes=False
  except RuntimeError:yes=True
  expect('阶段报告错误哈希被拦截',yes)
 expect('阶段报告结构校验'+str(stage),f.cmd_checkpoint(statepath,stage,rp,None,'',False,'测试schema',passdoc)==0)
expect('双阶段结构记录允许最终硬审',f.cmd_final_ready(statepath)==0)
review={'document_sha256':f.sha256(passdoc),'source_checks':[{'item':n,'status':'verified','conclusion':'synthetic','evidence':[{'path':str(source),'sha256':f.sha256(source),'locator':'synthetic lines'}]} for n in f.REVIEW_ITEMS], 'latest_confirmation':{'confirmed_at':'test','reply_source':'test','conclusion':'test'},'parameters':[{'name':'强度','scope':'路面','cross_source_search':'说明和结构图','occurrences':[{'path':str(source),'sha256':f.sha256(source),'locator':'test1','value':4.5,'unit':'MPa','scope':'路面'},{'path':str(source),'sha256':f.sha256(source),'locator':'test2','value':4.0,'unit':'MPa','scope':'入户道路'}],'status':'consistent','decision':'test','body_locations':[{'location':'第一章','quote':'路面强度按4.5MPa进行材料报审。','value':4.5,'unit':'MPa','scope':'路面'}]}], 'conflicts':[], 'construction_language_review':{'all_paragraphs_reviewed':True,'conclusion':'test'},'visual_review':{'three_color_figures_reviewed':True,'caption_photo_reviewed':True,'conclusion':'test'}}
rp=OUT/'review.json';rp.write_text(json.dumps(review));expect('4.5及4.0差异未解释被拦截',bool(f.validate_content_review(passdoc,rp)[0]))
review['parameters'][0].update(status='resolved',resolution_evidence='Synthetic scenario: differing scopes reviewed');rp.write_text(json.dumps(review));expect('有部位解释的溯源结构通过',not f.validate_content_review(passdoc,rp)[0])
bad=copy.deepcopy(review);bad['parameters'][0]['body_locations'][0]['quote']='正文不存在的证据';rp.write_text(json.dumps(bad));expect('虚构正文参数引用被拦截',bool(f.validate_content_review(passdoc,rp)[0]))

from argparse import Namespace

# Restore valid review after the prior failure case.
rp.write_text(json.dumps(review))
rr=c.load_rules(None)|rules
staticpath=OUT/'delivery-static.json';c.static_audit(passdoc,rules_path,staticpath)
word={'delivery_status':'PASS','document_sha256':f.sha256(passdoc),'layout_engine':'Microsoft Word','rules':rr,'rules_sha256':c.rules_digest(rr),'photo_delivery_stage':'draft','checks':[]}
# Synthetic reports test schema only. They are never represented as Word execution.
wp=OUT/'synthetic-word.json';wp.write_text(json.dumps(word))
pg=OUT/'synthetic-pages.json';pg.write_text(json.dumps({'status':'PASS','page_source':'microsoft_word','document_sha256':f.sha256(passdoc),'rule':{'body_pages_target':TARGET,'page_tolerance':TOLERANCE,'front_matter_pages_excluded':FRONT}}))
sc=OUT/'synthetic-score.json';sc.write_text(json.dumps({'status':'PASS','phase':'final','document_sha256':f.sha256(passdoc),'checks':[]}))
wr=OUT/'warnings.json';wr.write_text(json.dumps({'document_sha256':f.sha256(passdoc),'items':[]}))
a=Namespace(document=passdoc,state=statepath,mode='full',review=rp,static_report=staticpath,word_report=wp,page_report=pg,score_report=sc,warning_review=wr)
expect('交付汇总schema正例允许编制稿',f.cmd_delivery(a)==0)
wp.write_text(json.dumps(word|{'document_sha256':'old version'}));expect('交付汇总拒绝旧版本Word报告',f.cmd_delivery(a)==2);wp.write_text(json.dumps(word))
wp.write_text(json.dumps(word|{'rules_sha256':'altered'}));expect('交付汇总拒绝不同格式规则指纹',f.cmd_delivery(a)==2);wp.write_text(json.dumps(word))
wp.write_text(json.dumps(word|{'checks':[{'name':'留白人工复核','status':'WARN'}]}));expect('未关闭留白WARN不允许交付',f.cmd_delivery(a)==2)
wr.write_text(json.dumps({'document_sha256':f.sha256(passdoc),'items':[{'report':'word','check':'留白人工复核','status':'reviewed','evidence':'synthetic QA only','conclusion':'test'}]}));expect('处理WARN记录结构允许交付',f.cmd_delivery(a)==0)
# Force final stage while leaving placeholder. Both rule hashes match but embedded content blocks it.
finalrules=rr|{'photo_delivery_stage':'final'}
wp.write_text(json.dumps(word|{'rules':finalrules,'rules_sha256':c.rules_digest(finalrules)}))
fakestatic=json.loads(staticpath.read_text());fakestatic.update(rules_sha256=c.rules_digest(finalrules),photo_delivery_stage='final');fakep=OUT/'edited-static.json';fakep.write_text(json.dumps(fakestatic));a.static_report=fakep
expect('伪改静态报告仍不能让占位冒充正式照片',f.cmd_delivery(a)==2)

source.write_text(source.read_text()+' version changed');expect('来源文件改变使旧内容核验失效',bool(f.validate_content_review(passdoc,rp)[0]))
# No-page deployment must still block absent startup.
np=OUT/'no-page';np.mkdir(exist_ok=True);(np/'编制流程门禁.py').write_text((ROOT/'scripts/编制流程门禁.py').read_text())
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
manifest=json.loads((COMMON/'skill_version.json').read_text())
expect('Generic scripts match exact synced skill bytes',all(hashlib.sha256((COMMON/n).read_bytes()).hexdigest()==h for n,h in manifest['common_files_sha256'].items()))
print(json.dumps({'total_passed':len(results),'scope':'portable and interface tests; no Windows Word execution'},ensure_ascii=False))
