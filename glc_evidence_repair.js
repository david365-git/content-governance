/** Google Love evidence repair. Isolated: writes only GL Evidence Repairs. */
var GLC_ER_IDS=['48109','7952','5774','21838','21746','8027','8336','12107','11452','22942','11945'];
var GLC_ER_TAB='GL Evidence Repairs';
function glc_erSheet_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),s=ss.getSheetByName(GLC_ER_TAB);
 if(!s){s=ss.insertSheet(GLC_ER_TAB);s.appendRow(['Post ID','Status','Original comparison row','Source images','Draft images','Revised images','Omitted image URLs','Review notes','Repaired HTML','API cost USD','Updated UTC']);s.setFrozenRows(1);}
 return s;
}
function glc_erImages_(h){
 var m,re=/<img\b[^>]*?\bsrc\s*=\s*(["'])(.*?)\1/gi,a=[];
 while((m=re.exec(String(h||''))))a.push(m[2].replace(/&amp;/g,'&'));
 return a.filter(function(x,i){return a.indexOf(x)===i;});
}
function glc_erSource_(id){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),b=ss.getSheetByName('GL Sol 20 Batch'),t=ss.getSheetByName('GL Model Comparison Tests');
 if(!b||!t)throw Error('Missing completed comparison sheets.');
 var matches=b.getRange(2,1,20,4).getValues().filter(function(v){return String(v[0])===id;});
 if(matches.length!==1||String(matches[0][1])!=='DONE')throw Error('Completed original comparison missing: '+id);
 var row=Number(matches[0][3]),v=t.getRange(row,2,1,8).getValues()[0];
 if(String(v[0])!==id)throw Error('Comparison row/Post ID mismatch.');
 var a=glc_getArticle_(id),draft=String(v[7]||'');
 if(!a.original||!draft)throw Error('Missing source or draft.');
 return {a:a,draft:draft,row:row};
}
function glc_erSave_(id,status,v){
 var sh=glc_erSheet_(),ids=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().map(function(r){return r[0];}):[];
 var n=ids.indexOf(id),row=n<0?sh.getLastRow()+1:n+2;
 sh.getRange(row,1,1,11).setValues([[id,status,v.row||'',v.sourceCount||0,v.draftCount||0,v.revisedCount||0,String(v.omissions||'').slice(0,49000),String(v.notes||'').slice(0,49000),String(v.html||''),Number(v.cost)||0,Utilities.formatDate(new Date(),'Etc/UTC',"yyyy-MM-dd'T'HH:mm:ss'Z'")]]);
 return row;
}
function glc_repairNextCase(){
 var guard=LockService.getScriptLock();
 if(!guard.tryLock(2000))return 'Another repair is running.';
 var id,record;
 try{
  var sh=glc_erSheet_(),done=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,2).getValues():[];
  var pending=GLC_ER_IDS.filter(function(x){return !done.some(function(r){return String(r[0])===x&&String(r[1]).indexOf('REPAIRED')===0;});});
  if(!pending.length)return 'All 11 repaired drafts saved for review.';
  id=pending[0];
  var existing=done.filter(function(r){return String(r[0])===id;});
  if(existing.length && (String(existing[0][1]).indexOf('STOPPED')===0 || String(existing[0][1])==='IN PROGRESS')) throw Error('Prior repair for '+id+' requires review before retry. No API call made.');
  record=glc_erSource_(id);
  glc_erSave_(id,'IN PROGRESS',{row:record.row,notes:'API request in progress — check before rerunning after interruption.'});
 }finally{guard.releaseLock();}
 var orig=glc_erImages_(record.a.original),old=glc_erImages_(record.draft);
 var missing=orig.filter(function(x){return old.indexOf(x)<0;});
 var prompt=[
 'ROLE: UK stone and tile restoration case-study evidence editor.',
 'Repair this already generated article to preserve genuinely useful case-study facts from the original project account. Keep its useful structure. Do not add SEO filler.',
 'Preserve actually documented tasks and sequence, diamond grit grades, equipment, dwell times, repair methods, cleaning/rinsing/extraction, number and kind of sealer coats, stone/brand identification, original client decisions, real limits and actual outcomes.',
 'Never invent or silently convert general technical guidance into work claimed to have happened. If historical treatment is questionable, report what actually happened, while distinguishing it from current advice.',
 'Retain relevant original image URLs, only if each image is distinct and genuinely evidences the project. Do not invent URLs or captions. Do not alter existing valid image src values.',
 'Use owner-approved Technical Matrix as current advice/safeguards; preserve documentary facts separately. Do not use prohibited UK moisture-defect terminology. Acid cleaners are professional-only, condition-dependent, never general homeowner advice.',
 'When urethane/polyurethane identifies sealer chemistry in homeowner text call it film-forming sealer; preserve modifiers and number of coats; do not change topical/impregnating.',
 'Return only full HTML without Markdown. Maintain factual fidelity and do not add repeated legalistic disclaimers.',
 'POST ID: '+id,'TITLE: '+record.a.title,
 'CURRENT APPROVED MATRIX:',record.a.technical,
 'GLOBAL MOISTURE RULE:',glc_moistureRule_(),
 'ORIGINAL CASE STUDY HTML:',record.a.original,
 'EXISTING GOOGLE LOVE HTML:',record.draft,
 'SOURCE IMAGE URLs ABSENT FROM DRAFT (review individually):',JSON.stringify(missing)
 ].join('\n\n');
 var result;
 try{
  result=bc_sendPromptViaOpenAI(prompt,16000,'gpt-6.1-sol');
  if(!result||!result.success||result.truncated||!String(result.text||'').trim())throw Error('Repair API unsuccessful/truncated: '+String(result&&result.message||''));
  var html=String(result.text).replace(/^\s*`{3}(?:html)?\s*/i,'').replace(/\s*`{3}\s*$/,'').trim();
  if(!/<h[12]\b/i.test(html)||!/<p\b/i.test(html)||html.length>49000)throw Error('Repair HTML invalid or too long — not saved.');
  var newer=glc_erImages_(html),omitted=orig.filter(function(x){return newer.indexOf(x)<0;});
  var prohibited=/\brising[\s-]+damp\b/i.test(html.replace(/<[^>]*>/g,' '));
  glc_erSave_(id,prohibited?'REVIEW REQUIRED':'REPAIRED — EDITORIAL REVIEW',{row:record.row,sourceCount:orig.length,draftCount:old.length,revisedCount:newer.length,omissions:JSON.stringify(omitted),notes:'Source-vs-revised editorial fact check still required. Prohibited terminology detected: '+prohibited,cost:Number(result.cost)||0,html:html});
  return {postId:id,omittedImages:omitted.length,cost:Number(result.cost)||0,status:'Saved separate repair draft'};
 }catch(e){
  glc_erSave_(id,'STOPPED — REVIEW REQUIRED',{row:record.row,notes:String(e.message||e),cost:Number(result&&result.cost)||0});
  throw e;
 }
}

/** Optional unattended runner: one case study per minute, stops on any error. */
function glc_evidenceRepairTriggers_(){return ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==='glc_evidenceRepairTick';});}
function glc_stopEvidenceRepairBatch(){
 PropertiesService.getScriptProperties().deleteProperty('GLC_EVIDENCE_BATCH_ACTIVE');
 glc_evidenceRepairTriggers_().forEach(function(t){ScriptApp.deleteTrigger(t);});
 return 'Evidence repair scheduling stopped. A current request might finish.';
}
function glc_startEvidenceRepairBatch(){
 var sh=glc_erSheet_(),all=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,2).getValues():[];
 var blocked=all.filter(function(r){return String(r[1])==='IN PROGRESS'||String(r[1]).indexOf('STOPPED')===0;});
 if(blocked.length)throw Error('Review existing repair status for Post '+blocked[0][0]+' before resuming; no duplicate API calls.');
 glc_stopEvidenceRepairBatch();
 PropertiesService.getScriptProperties().setProperty('GLC_EVIDENCE_BATCH_ACTIVE','YES');
 ScriptApp.newTrigger('glc_evidenceRepairTick').timeBased().everyMinutes(1).create();
 return 'Scheduled the 11-case-study evidence repair. Outputs saved separately in GL Evidence Repairs.';
}
function glc_evidenceRepairTick(){
 if(PropertiesService.getScriptProperties().getProperty('GLC_EVIDENCE_BATCH_ACTIVE')!=='YES')return;
 try{
  var status=glc_repairNextCase();
  if(typeof status==='string' && status.indexOf('All 11')===0)glc_stopEvidenceRepairBatch();
 }catch(e){
  glc_stopEvidenceRepairBatch();
  console.error('Evidence repair stopped: '+String(e.message||e));
 }
}
