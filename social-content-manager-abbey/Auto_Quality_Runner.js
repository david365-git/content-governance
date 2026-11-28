/**
 * Autonomous GBP drafting controller.
 * No live publication is performed; external publication requires a separately
 * authorised and tested GBP adapter. Automatic AI approval is distinct from publishing.
 */
function scm_autoGBPAdvance() {
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000))return {status:'BUSY'};
  const result={status:'RUNNING',stages:[],advanced:0,held:0};
  const runStage=(name,fn)=>{
    for(let attempt=1;attempt<=2;attempt++) {
      try{const output=fn();result.stages.push({name:name,result:output});return output;}
      catch(e){
        if(attempt===2){
          const safe=String(e.message||e).slice(0,220);
          result.stages.push({name:name,error:safe});
          scm_autoLog_('STAGE HELD',name,safe);
          return null;
        }
        Utilities.sleep(1500*attempt);
      }
    }
  };
  try {
    runStage('DISCOVERY',()=>scm_scanAllArticles());
    runStage('BATCH PLAN',()=>scm_batchPlanGBP());
    const props=PropertiesService.getScriptProperties();
    if(props.getProperty('SCM_OPENAI_API_KEY')) {
      const ai=runStage('AI WRITING + REVIEW',()=>scm_generateReviewedAIPreviews());
      if(ai) {
        const promotions=runStage('AUTO QUALITY PROMOTION',()=>scm_autoPromoteAIPassed_());
        if(promotions){result.advanced=promotions.advanced;result.held=promotions.held;}
      }
    } else {
      scm_autoLog_('CONFIGURATION','AI WRITING','API key missing. No paid call made; drafts retained safely.');
      result.stages.push({name:'AI WRITING',status:'WAITING FOR CONFIGURATION'});
    }
    runStage('IMAGE PREPARATION',()=>scm_autoPrepareQueuedImages());
    runStage('SCHEDULE SYNC',()=>scm_syncPublishingSchedule());
    result.status='DONE — NO LIVE PUBLISHING';
    return result;
  } finally {lock.releaseLock();}
}
function scm_autoPromoteAIPassed_() {
  const review=scm_reviewRows_(),queue=scm_rows_(SCM.queue);
  const queueById=new Map(queue.rows.map(r=>[String(r.item['Social Post ID']),r]));
  let advanced=0,held=0;
  const put=(row,name,value)=>{
    const i=queue.keys.indexOf(name);
    if(i<0)throw new Error('Missing queue column: '+name);
    queue.sheet.getRange(row.number,i+1).setValue(value);
    row.item[name]=value;
  };
  for(const row of review.rows) {
    if(String(row.item.Status)!=='AI REVIEW PASSED')continue;
    const q=queueById.get(String(row.item['Social Post ID']));
    if(!q||String(q.item.Approval)!=='DRAFT')continue;
    const notes=String(q.item.Notes||'');
    // Preserve any draft previously edited by an operator.
    if(!/Requires human review|Evidence-based draft from nearby|Automatically selected/i.test(notes)){
      scm_autoLog_('PRESERVED','DRAFT','Existing edited text not replaced: '+row.item['Social Post ID']);
      held++;continue;
    }
    if(String(q.item['Draft Text'])!==String(row.item['Current Draft'])){
      scm_autoLog_('PRESERVED','DRAFT','Text changed since preview: '+row.item['Social Post ID']);
      held++;continue;
    }
    const text=String(row.item['Suggested Draft']||'').trim();
    if(!scm_aiSafeText_(text,String(q.item['Article URL']||''))){
      scm_autoLog_('QUALITY HOLD','DRAFT','Invalid AI proposal: '+row.item['Social Post ID']);
      held++;continue;
    }
    put(q,'Draft Text',text);
    put(q,'Revision',(Number(q.item.Revision)||1)+1);
    put(q,'Notes','AI writer + independent AI review passed. Automatic editorial promotion; not yet published.');
    // A separate control permits only quality-passed drafts to be auto-approved.
    const approvalSetting=PropertiesService.getScriptProperties().getProperty('SCM_AUTO_AI_APPROVE');
    if(approvalSetting==='YES')put(q,'Approval','APPROVED');
    scm_reviewPut_(review,row,{
      'Status':'APPLIED',
      'Notes':'Automatically promoted after two-stage quality check; original copy retained in Current Draft.'
    });
    advanced++;
  }
  return {advanced:advanced,held:held};
}
function scm_autoLog_(action,stage,details) {
  scm_append_(SCM.log,{
    'Event ID':scm_id_('EVT'), 'Platform':'GBP','Action':action,
    'Attempt UTC':scm_now_(),'Result':stage,
    'Details':String(details).slice(0,500),
    'Request Fingerprint':'auto|'+action+'|'+stage+'|'+new Date().toISOString()
  });
}

/** Optional standalone background trigger. Explicit opt-in via Script Properties. */
function scm_autoGBPBackgroundTick() {
  const enabled=PropertiesService.getScriptProperties().getProperty('SCM_AUTOMATION_ENABLED');
  if(enabled!=='YES')return {status:'DISABLED'};
  return scm_autoGBPAdvance();
}
function scm_installDraftAutomationTrigger() {
  const enabled=PropertiesService.getScriptProperties().getProperty('SCM_AUTOMATION_ENABLED');
  if(enabled!=='YES')throw new Error('Set SCM_AUTOMATION_ENABLED=YES before installing a paid automated runner.');
  const current=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='scm_autoGBPBackgroundTick');
  if(current.length)return {installed:false,existing:current.length};
  ScriptApp.newTrigger('scm_autoGBPBackgroundTick').timeBased().everyHours(6).create();
  return {installed:true,frequency:'every 6 hours',publishing:'DISABLED'};
}
function scm_removeDraftAutomationTrigger() {
  const current=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='scm_autoGBPBackgroundTick');
  current.forEach(t=>ScriptApp.deleteTrigger(t));
  return {removed:current.length};
}
