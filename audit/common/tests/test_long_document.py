"""Portable regressions using synthetic reports; NO actual Windows Word execution."""
from pathlib import Path
import importlib.util,sys,tempfile,shutil,json,math,copy,subprocess,hashlib
from docx import Document
from docx.shared import Inches,Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image,ImageDraw

# Configurable root is resolved from the installed skill when run locally, or generic GH layout.
if len(sys.argv)>1:ORIGIN=Path(sys.argv[1]);COMMON=ORIGIN/'scripts'
else:COMMON=Path(__file__).resolve().parents[1]
TEMPDIR=tempfile.TemporaryDirectory();BASE=Path(TEMPDIR.name)
ROOT=BASE/'scripts';ROOT.mkdir()
for name in ('编制流程门禁.py','Word正式交付硬审（不含页数）.py','评分响应核验.py','visual_policy.py','source_coverage.py','project_response.py','boq.py','schedule.py','hard_checks.py','bid_requirements.py','submission_identity.py'):shutil.copy2(COMMON/name,ROOT/name)
PAGE=ROOT/'测试项目页数硬审.py'
PAGE.write_text('EXPECTED_BODY_PAGES = 395\nPAGE_TOLERANCE = 5\nFRONT_MATTER_PAGES = 1\n')
sys.path.insert(0,str(ROOT))
def load(name,path):
 sp=importlib.util.spec_from_file_location(name,path);m=importlib.util.module_from_spec(sp);sys.modules[name]=m;sp.loader.exec_module(m);return m
f=load('flow',ROOT/'编制流程门禁.py');c=load('word',ROOT/'Word正式交付硬审（不含页数）.py')
RESULTS=[]
def check(label,expr):
 assert expr,label
 RESULTS.append(label)
def write(path,data):path.write_text(json.dumps(data,ensure_ascii=False,indent=2));return path
def evidence(p,locator='synthetic test only'):return {'path':str(p),'sha256':f.sha256(p),'locator':locator}
TERMS=['拆除', '加固', '支护', '降水', '清表', '开挖', '钻孔', '灌注', '排水', '运输', '吊装', '测量', '保护', '养护', '验收', '试压', '回填', '导流', '围堰', '临电']
CATALOG=' '.join(TERMS)+' 满分1分'
SOURCE=BASE/'tender-fixture.txt';SOURCE.write_text(CATALOG+'\nfixture scope R01')
POLICY={'version':1,'real_photos':{'decision':'required','reason':'fixture','replacement_figures':[]},
        'color_figures':{'mode':'color'},
        'additional_figures':{'names':[],'reason':'fixture baseline',
                              'override':{'source_type':'prompt','source_ref':'synthetic','quote':'only base figures'}}}
FORMAT=BASE/'format.json';write(FORMAT,{'visual_policy':POLICY,'font_rules':{},'required_level1_titles':[], 'required_figures':[{'name':n,'hard':True,'color_required':True} for n in c.base_names(POLICY)], 'required_real_photo_slots':[{'name':'现场现状'}],'photo_delivery_stage':'draft'})
PNG=BASE/'pilot.png';im=Image.new('RGB',(400,600),'white');dr=ImageDraw.Draw(im)
for i,col in enumerate(('red','green','blue')):dr.rectangle((30,40+i*160,370,150+i*160),fill=col)
im.save(PNG)
PILOT=BASE/'pilot.docx';d=Document();d.add_paragraph('试排正文仅供控制逻辑测试。'*30);d.add_table(rows=1,cols=1).cell(0,0).text='表格样例';d.add_picture(str(PNG),width=Inches(3));d.save(PILOT)
PDF=BASE/'word-pilot-export-fixture.pdf';PDF.write_bytes(b'%PDF-1.4\n% Synthetic evidence metadata only, not actual Word export\n')

def start_case(target=395,tol=5,front=1):
 PAGE.write_text(f'EXPECTED_BODY_PAGES = {target}\nPAGE_TOLERANCE = {tol}\nFRONT_MATTER_PAGES = {front}\n')
 state_path=BASE/f'state-{target}.json';state_path.unlink(missing_ok=True);f.cmd_init(state_path,False)
 s=f.load_state(state_path)
 s['startup']={'visual_policy':POLICY,'sources':[dict(evidence(SOURCE),role='tender')], 'latest_updates':{'status':'confirmed_none','asked_at':'test','reply_text':'synthetic answer','reply_source':'synthetic record'},'critical_questions':[],'questions_review':'fixture review','batch_plan':{'count':1,'reason':'fixture','batches':['whole']},'word_channel':'unavailable','word_capability_evidence':'tests run without Word','real_photo_slots':[{'name':'现场现状','location':'第一章','basis':'fixture'}]}
 reserve=5 if target>=200 else 2
 n=math.ceil((target-reserve)/f.planning_policy(target)['module_page_cap']);per=(target-reserve)/n
 chapters=['第一章 技术方案','第二章 保障措施','第三章 施工部署']
 mods=[];reqs=[]
 for i in range(n):
  rid=f'R{i+1:02}';mid=f'M{i+1:02}';chapter=chapters[min(i*3//n,2)]
  reqs.append({'id':rid,'scope':'fixture scope '+rid,'scoring':False,'evidence':[evidence(SOURCE)]})
  mods.append({'id':mid,'chapter':chapter,'section':'第一节 '+mid,'scope':'fixture content scope '+mid,'basis':'fixture scoring responsibility','requirement_ids':[rid],'primary_for':[rid], 'evidence':[evidence(SOURCE)],'text_pages':per-1,'table_pages':.5,'figure_pages':.5,'planned_pages':per,'estimated_normal_chars':(per-1)*100})
 count40=max(1,round(n*.4));count75=max(count40+1,min(n-1,round(n*.75)))
 boundaries={'40':[m['id'] for m in mods[:count40]],'75':[m['id'] for m in mods[:count75]]}
 ranges={}
 for stage,ids in boundaries.items():
  center=sum(m['planned_pages'] for m in mods if m['id'] in ids);slack=max(1,math.ceil(center*.05));ranges[stage]=[round(max(1,center-slack),2),round(center+slack,2)]
 s['plan'].update(revision=1,planned_total_body_pages=target,chapter_plan=[{'chapter':ch,'basis':'fixture weighting','page_range':[0,target]} for ch in chapters],figure_table_pages_estimate=n,reserve_pages=reserve,planning_notes='按真实内容范围、评分责任和版面容量进行预算（fixture）',requirements=reqs,modules=mods,pilot_plan={'module_ids':[mods[0]['id'],mods[-1]['id']],'coverage':['narrative','table','figure'],'basis':'different representative content types'},capacity_review={'all_requirements_mapped':True,'scope_supports_volume':True,'no_padding_or_repetition':True,'conclusion':'fixture only','unresolved_gaps':[]},checkpoint_module_ids=boundaries,checkpoint_completion_basis={'40':'fixture completed responsibility','75':'fixture completed responsibility'},checkpoint_expected_page_ranges=ranges)
 import source_coverage as trace
 s['startup']['source_scope']={'roots':[str(SOURCE)],'excluded':[],'reviewed':True,'conclusion':'synthetic source inventory'}
 rows,texts=trace.scan(s,BASE)
 sid=s['startup']['sources'][0]['id']; units=rows[0]['units'];first=units[0]['id'];second=units[1]['id']
 for i,r in enumerate(reqs):
  r.update(source_refs=[{'source_id':sid,'unit_id':second if i==0 else first,'quote':r['scope'] if i==0 else CATALOG}],response_requirements=['synthetic substantive response'],score_refs=[{'item_id':'S1','criterion_index':0}])
  if i==0:r['scoring']=True
 s['startup']['source_coverage'].update(dispositions=[{'unit_ids':[second],'status':'reviewed','disposition':'requirements','requirement_ids':[reqs[0]['id']],'conclusion':'synthetic scoring source'},{'unit_ids':[first],'status':'reviewed','disposition':'requirements','requirement_ids':[r['id'] for r in reqs[1:]],'conclusion':'synthetic scope source'}],extraction_review={k:{'reviewed':True,'conclusion':'synthetic source pass'} for k in ('source_pass','scoring_pass','scope_reconciliation')})
 scores=write(BASE/f'scores-{target}.json',{'tender_source':'fixture line 2','items':[{'id':'S1','name':'fixture','source_ref':'line 2','criteria':[reqs[0]['scope']],'max_score':1}]})
 s['plan']['scoring_basis']={'path':str(scores),'basis_sha256':trace.digest(trace.scoring_basis(json.loads(scores.read_text())))}
 s['plan']['project_response'].update(keywords=[{'term':t,'requirement':reqs[1]['id']} for t in TERMS],headings={'2':TERMS[:15],'3':TERMS},priorities=[{'chapters':[ch],'basis':'synthetic scope ranking'} for ch in chapters])
 # Stage views in this fixture precede the final planning module.
 last=mods[-1]['id']
 s['plan']['hard_checks']={'bills':[], 'gaps':[],
   'figures':[{'name':name,'kind':kind,'basis':'fixture figure obligation','module':last,
               'elements':sorted(f.hard_checks.ELEMENTS[kind]) or ['界面分区']}
              for name,kind in f.hard_checks.required_figures(s).items()],
   'measures':[{'id':'交接','module':last,'requirement':reqs[-1]['id'],'basis':'synthetic interface',
                'completion':'handover inspected'}],
   'schedule':{'tasks':[{'id':'A','name':'场地交接','unit':'段','module':last,'crews':['crew'],
        'equipment':[],'materials':[],'equipment_basis':'人工交接','materials_basis':'现场核查',
        'ready':{},'after':[],'quantity':1,'rate':1,'start':1,'finish':1,'windows':[[1,8,9]],'basis':'fixture sample'}]}}
 from bid_fixture import project_profile,fixture_capacity
 s['plan']['bid_requirements']=project_profile({'source_id':sid,'unit_id':first,'quote':CATALOG},last,reqs[-1]['id'],[reqs[0]['id']],['S1'],['crew'])
 s['plan']['hard_checks']['schedule']['tasks'][0].update(productivity_basis='核验接口工效',capacity=fixture_capacity({'source_id':sid,'unit_id':first,'quote':CATALOG},['crew']))
 f.save_state(state_path,s)
 return state_path,s

STATE,s=start_case()
check('395-page detailed plan admitted',f.cmd_preflight(STATE)==0)
check('Long plan admits sample only, not mass drafting',f.load_state(STATE)['gate']['draft_ready']=='BLOCKED')
for badtotal in (363,389,401,427):
 bad=copy.deepcopy(s);bad['plan']['planned_total_body_pages']=badtotal;f.save_state(STATE,bad)
 check(f'Planning total {badtotal} blocked',f.cmd_preflight(STATE)==2)
for label,mutate in [
 ('Missing module map blocked',lambda x:x['plan'].update(modules=[])),
 ('Uncovered requirement blocked',lambda x:x['plan']['requirements'].append({'id':'UNCOVERED','scope':'fixture','evidence':[evidence(SOURCE)]})),
 ('Missing primary response blocked',lambda x:x['plan']['modules'][0].update(primary_for=[])),
 ('Repeated figure page accounting blocked',lambda x:x['plan']['modules'][0].update(figure_pages=3)),
 ('Module exceeding adaptive cap blocked',lambda x:x['plan']['modules'][0].update(planned_pages=40)),
 ('Missing content-volume judgment blocked',lambda x:x['plan']['capacity_review'].update(scope_supports_volume=False)),
 ('Unresolved content gap blocked',lambda x:x['plan']['capacity_review'].update(unresolved_gaps=['missing scope'])),
 ('Excessively wide checkpoint range blocked',lambda x:x['plan']['checkpoint_expected_page_ranges'].update({'40':[1,300]})),
 ('Pilot lacks representative figure coverage blocked',lambda x:x['plan']['pilot_plan'].update(coverage=['narrative','table']))]:
 bad=copy.deepcopy(s);mutate(bad);f.save_state(STATE,bad);check(label,f.cmd_preflight(STATE)==2)
f.save_state(STATE,s);f.cmd_preflight(STATE)
check('Draft-ready before actual sample report blocked',f.cmd_draft_ready(STATE)==2)

REPORT=BASE/'pilot-report-fixture.json'
report={'page_source':'microsoft_word','status':'MEASURED','measurement_role':'pilot','document_sha256':f.sha256(PILOT),'project_rule':{'target_body_pages':395,'page_tolerance':5,'front_matter_pages_excluded':1},'actual':{'word_body_pages':10,'word_normal_text_characters':800},'word_pdf_evidence':evidence(PDF)|{'path':PDF.name}}
write(REPORT,report)
ANALYSIS=BASE/'pilot-analysis.json';analysis={'document_sha256':f.sha256(PILOT),'module_ids':s['plan']['pilot_plan']['module_ids'],'format_rules':evidence(FORMAT),'font_and_spacing_reviewed':True,'table_layout_reviewed':True,'figure_readability_reviewed':True,'photo_frame_layout_reviewed':True,'visual_evidence':[evidence(PDF)|{'path':PDF.name}],'normal_text_pages':8,'table_pages':1,'figure_pages':1}
write(ANALYSIS,analysis)
write(REPORT,report|{'page_source':'libreoffice_estimate'});check('LO sample cannot pass Word sample gate',f.cmd_pilot(STATE,PILOT,REPORT,ANALYSIS)==2)
write(REPORT,report|{'document_sha256':'old'});check('Stale sample report blocked',f.cmd_pilot(STATE,PILOT,REPORT,ANALYSIS)==2)
write(REPORT,report);write(ANALYSIS,analysis|{'figure_readability_reviewed':False});check('Missing visual review blocked',f.cmd_pilot(STATE,PILOT,REPORT,ANALYSIS)==2)
write(ANALYSIS,analysis);check('Pilot report schema and evidence accepted',f.cmd_pilot(STATE,PILOT,REPORT,ANALYSIS)==0)
check('Calibrated budget permits mass drafting',f.cmd_draft_ready(STATE)==0)
base=f.load_state(STATE)
bad=copy.deepcopy(base);bad['plan']['modules'][0]['estimated_normal_chars']*=2;bad['gate']['plan_hash']=f.plan_hash(bad);f.save_state(STATE,bad)
check('Budget ignoring measured density blocked',f.cmd_draft_ready(STATE)==2)
f.save_state(STATE,base)
check('Checkpoint with unsaved modules blocked',f.cmd_checkpoint(STATE,40,None,156,'libreoffice_estimate',True,'fixture',PILOT)==2)

MODULE_SOURCES={};REVIEWS={};QUOTES={}
def module_files(m):
 mid=m['id'];q1=('模块'+mid+'材料报审及设备进场完成后，按已核验的施工流程组织作业，执行工序检查并记录结果。')*2
 q2=('模块'+mid+'发现检查偏差时，按批准方案处理并复查，满足移交条件后进入后续工序，保留相应记录。')*2
 text=q1+'\n'+q2+'\n';text+=('本模块'+mid+'施工执行记录用于控制工序及接口。\n')*max(1,int((m['estimated_normal_chars']-len(f._normal(text)))/24))
 if mid==s['plan']['modules'][-1]['id']:
  text+='试件操作标准 试件第1条 检查接口 班长 核对记录 复查接口后移交 场地受限 分时交接 逐段验收。\n'
 src=BASE/(mid+'.md');src.write_text(text)
 rv=BASE/(mid+'-review.json');write(rv,{'module_id':mid,'source_sha256':f.sha256(src),'covered_requirement_ids':m['requirement_ids'],'fact_sources_reviewed':True,'technical_depth_reviewed':True,'interfaces_reviewed':True,'repetition_reviewed':True,'conclusion':'synthetic review','technical_review':{'version':1,'checks':{k:{'status':'reviewed','conclusion':'fixture'} for k in f.TECHNICAL_REVIEW_ITEMS}},'construction_language_review':{'reviewed':True,'unresolved_count':0,'conclusion':'synthetic language judgment'},'scoring_review':{'conditions':[],'not_applicable_reason':'fixture requirements only; no tender scoring conditions','conclusion':'synthetic responsibility judgment'},'probe_quotes':[q1,q2]})
 data=json.loads(rv.read_text());requirements={r['id']:r for r in s['plan']['requirements']}
 data['response_review']={'checks':[{'requirement_id':rid,'status':'adequate','conclusion':'synthetic response','aspects':[{'expectation':'synthetic substantive response','location':mid,'quote':q1,'adequacy':'adequate','conclusion':'synthetic adequacy'}]} for rid in m['primary_for']]}
 conditions=[]
 for rid in m['requirement_ids']:
  if requirements[rid]['scoring']:
   conditions.append({'requirement_id':rid,'criterion_quote':requirements[rid]['scope'],'response_quote':q1,'evidence':[evidence(SOURCE)],'targeted':True,'technical_depth':True,'project_basis':True,'substantive_response':True,'status':'satisfied','conclusion':'synthetic scoring response'})
 if conditions:data['scoring_review']={'conditions':conditions,'conclusion':'synthetic scoring response'}
 if mid==s['plan']['modules'][-1]['id']:
  loc='L'+str(len(text.splitlines()))
  p=s['plan']['bid_requirements']
  data['bid_evidence']={r['id']:{k:loc for k in r} for r in p['obligations']+p['decisions']}
  data['measures']={'交接':{k:{'loc':'L1','quote':'材料报审及设备进场完成后','finding':'synthetic module evidence'} for k in f.hard_checks.ASPECTS}}
 write(rv,data)
 return src,rv,[q1,q2]
for m in s['plan']['modules']:
 src,rv,qs=module_files(m);MODULE_SOURCES[m['id']]=src;REVIEWS[m['id']]=rv;QUOTES[m['id']]=qs
check('Module review with wrong source version blocked',f.cmd_module_checkpoint(STATE,'M01',MODULE_SOURCES['M02'],REVIEWS['M01'])==2)
check('First module persists actual source and review',f.cmd_module_checkpoint(STATE,'M01',MODULE_SOURCES['M01'],REVIEWS['M01'])==0)
old=MODULE_SOURCES['M01'].read_text();MODULE_SOURCES['M01'].write_text('short')
check('Lost/changed module source prevents registry PASS',bool(f.registry_errors(f.load_state(STATE),BASE,['M01'])))
MODULE_SOURCES['M01'].write_text(old)
for m in s['plan']['modules'][1:8]:check('Persist '+m['id'],f.cmd_module_checkpoint(STATE,m['id'],MODULE_SOURCES[m['id']],REVIEWS[m['id']])==0)
check('Cannot write beyond 40-percent node before actual calibration',f.cmd_module_checkpoint(STATE,'M09',MODULE_SOURCES['M09'],REVIEWS['M09'])==2)

def compose(path,ids):
 d=Document()
 for mid in ids:
  for line in MODULE_SOURCES[mid].read_text().splitlines():d.add_paragraph(line)
 d.save(path)
SNAP40=BASE/'stage-40.docx';SNAP75=BASE/'stage-75.docx';FINAL=BASE/'final.docx'
compose(SNAP40,s['plan']['checkpoint_module_ids']['40']);compose(SNAP75,s['plan']['checkpoint_module_ids']['75']);compose(FINAL,f._ids(s))
for path in (SNAP40,SNAP75):
 review=f.hard_checks.pending_review()
 review['bill_scan']='Checked all fixture paragraphs and table rows'
 write(path.with_suffix('.hard.json'),{'document_sha256':f.sha256(path),'hard_checks':review})
def checkpoint_report(stage,path,actual):
 return {'page_source':'microsoft_word','status':'MEASURED','stage_percent':stage,'document_sha256':f.sha256(path),'project_rule':{'target_body_pages':395,'page_tolerance':5,'front_matter_pages_excluded':1},'actual':{'word_body_pages':actual,'word_version':'synthetic-fixture'}}
CP40=BASE/'cp40-fixture.json';CP75=BASE/'cp75-fixture.json'
def score_report(stage, path):
 return write(BASE/f'score-{stage}-fixture.json',{'status':'PASS','phase':str(stage),
              'document_sha256':f.sha256(path),'scoring_basis_sha256':s['plan']['scoring_basis']['basis_sha256']})
SCORE40=score_report(40,SNAP40)
SCORE75=score_report(75,SNAP75)
write(CP75,checkpoint_report(75,SNAP75,292));check('75 before 40 blocked',f.cmd_checkpoint(STATE,75,CP75,None,'',False,'fixture',SNAP75,SCORE75)==2)
write(CP40,checkpoint_report(40,SNAP40,156)|{'document_sha256':'wrong'})
try:f.cmd_checkpoint(STATE,40,CP40,None,'',False,'fixture',SNAP40,SCORE40);wrong_hash=False
except RuntimeError:wrong_hash=True
check('Checkpoint rejects report from another document',wrong_hash)
write(CP40,checkpoint_report(40,SNAP40,162));check('Stage page range can fit but final forecast too long is blocked',f.cmd_checkpoint(STATE,40,CP40,None,'',False,'fixture',SNAP40,SCORE40)==2)
check('Failed forecast freezes further writing',f.cmd_module_checkpoint(STATE,'M20',MODULE_SOURCES['M20'],REVIEWS['M20'])==2)
failed_state=f.load_state(STATE);oldsource=MODULE_SOURCES['M01'].read_bytes();oldreview=REVIEWS['M01'].read_bytes()
MODULE_SOURCES['M01'].write_bytes(oldsource+'\n返修补充本工序接口控制与检查记录。'.encode());repair=json.loads(oldreview);repair['source_sha256']=f.sha256(MODULE_SOURCES['M01']);write(REVIEWS['M01'],repair)
check('Failed stage permits actual revision of a saved module',f.cmd_module_checkpoint(STATE,'M01',MODULE_SOURCES['M01'],REVIEWS['M01'])==0)
check('Revision makes existing stage snapshot stale',f.load_state(STATE)['checkpoints']['40']['status']=='STALE')
MODULE_SOURCES['M01'].write_bytes(oldsource);REVIEWS['M01'].write_bytes(oldreview);f.save_state(STATE,failed_state)

write(CP40,checkpoint_report(40,SNAP40,156))
check('Word measurement requires score evidence for the same draft',f.cmd_checkpoint(STATE,40,CP40,None,'',False,'fixture',SNAP40)==2
      and f.cmd_checkpoint(STATE,40,CP40,None,'',False,'fixture',SNAP40,SCORE75)==2)
check('Measured 40 schema passes after corrected evidence',f.cmd_checkpoint(STATE,40,CP40,None,'',False,'fixture',SNAP40,SCORE40)==0)
for m in s['plan']['modules'][8:15]:check('Persist '+m['id'],f.cmd_module_checkpoint(STATE,m['id'],MODULE_SOURCES[m['id']],REVIEWS[m['id']])==0)
check('Cannot write beyond 75-percent node before calibration',f.cmd_module_checkpoint(STATE,'M16',MODULE_SOURCES['M16'],REVIEWS['M16'])==2)
check('Final-ready missing second checkpoint blocked',f.cmd_final_ready(STATE,FINAL)==2)
check('Measured 75 schema passes',f.cmd_checkpoint(STATE,75,CP75,None,'',False,'fixture',SNAP75,SCORE75)==0)
for m in s['plan']['modules'][15:]:check('Persist '+m['id'],f.cmd_module_checkpoint(STATE,m['id'],MODULE_SOURCES[m['id']],REVIEWS[m['id']])==0)
check('Full merged source probes permit final audit',f.cmd_final_ready(STATE,FINAL)==0)
PARTIAL=BASE/'only-probes.docx';pd=Document()
for mid in f._ids(s):
 for quote in QUOTES[mid]:pd.add_paragraph(quote)
pd.save(PARTIAL);check('Keeping probes but dropping substantive module content blocked',f.cmd_final_ready(STATE,PARTIAL)==2)
original_report=CP75.read_bytes();CP75.write_bytes(original_report+b'\n')
check('Changed stage Word report invalidates final readiness',f.cmd_final_ready(STATE,FINAL)==2);CP75.write_bytes(original_report)
original_score=SCORE75.read_bytes();SCORE75.write_bytes(original_score+b'\n')
check('Changed stage score report invalidates final readiness',f.cmd_final_ready(STATE,FINAL)==2);SCORE75.write_bytes(original_score)
MISSING=BASE/'missing-module.docx';compose(MISSING,f._ids(s)[:-1]);check('Dropped final module blocked',f.cmd_final_ready(STATE,MISSING)==2)
FORMAT.write_text(FORMAT.read_text()+'\n');check('Changing frozen pilot format invalidates sample',f.cmd_final_ready(STATE,FINAL)==2);write(FORMAT,{'visual_policy':POLICY,'font_rules':{},'required_level1_titles':[], 'required_figures':[{'name':n,'hard':True,'color_required':True} for n in c.base_names(POLICY)], 'required_real_photo_slots':[{'name':'现场现状'}],'photo_delivery_stage':'draft'})
check('Restored identical pilot format returns valid evidence',f.cmd_final_ready(STATE,FINAL)==0)

# Smaller/new project: same scripts, no forced long-document sample.
state50,s50=start_case(50,1,2);check('Smaller target has valid detailed plan',f.cmd_preflight(state50)==0);check('Small project needs no long-document pilot',f.load_state(state50)['gate']['draft_ready']=='PASS')
check('Page script replacement updates source constants',f.discover_page_script()[1]['EXPECTED_BODY_PAGES']==50 and f.discover_page_script()[1]['FRONT_MATTER_PAGES']==2)
check('Common Word location follows current project',c.load_rules(None)['body_start_page']==3)
rc=subprocess.run([sys.executable,str(ROOT/'编制流程门禁.py'),'run-page-audit','--pilot','--document',str(PILOT),'--report',str(BASE/'no-word.json')],capture_output=True,text=True)
check('Linux cannot fake native Word pilot execution',rc.returncode==3)
print(json.dumps({'passed':len(RESULTS),'tests':RESULTS,'limits':'Synthetic report/interface tests only; Windows Word was NOT executed.'},ensure_ascii=False,indent=2))
