/**
 * Google Love persistent batch scheduler.
 * Runs one CI / Sol / W5 / W5C stage per time-driven trigger invocation.
 * The sidebar is monitoring only; switching tabs cannot pause the queue.
 * An interrupted or uncertain paid request is NEVER replayed automatically.
 */
var GLB_BG_KEY='GLB_BACKGROUND_BATCH_V1';
var GLB_BG_HANDLER='glb_backgroundStep';

function glb_backgroundRead_(){
  var value=PropertiesService.getScriptProperties().getProperty(GLB_BG_KEY);
  if(!value)return null;
  try{return JSON.parse(value);}catch(e){throw Error('Invalid saved Google Love batch state; manual review required');}
}
function glb_backgroundWrite_(state){
  state.updatedUtc=glb_utc_();
  PropertiesService.getScriptProperties().setProperty(GLB_BG_KEY,JSON.stringify(state));
}
function glb_backgroundRemoveTriggers_(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()===GLB_BG_HANDLER)ScriptApp.deleteTrigger(t);
  });
}
function glb_backgroundFinish_(job,status,message){
  job.state=status;job.message=String(message||'').slice(0,2000);
  job.inflight=null;job.inflightSince=null;
  glb_backgroundWrite_(job);
  glb_backgroundRemoveTriggers_();
}
function glb_backgroundStart(idsText,mode){
  mode=mode==='solOnly'?'solOnly':'full';
  var ids=String(idsText||'').split(/[\s,;]+/).map(function(x){return x.trim();}).filter(Boolean);
  if(!ids.length||ids.length>50||ids.some(function(x){return !/^\d+$/.test(x);}))
    throw Error('Enter 1–50 valid numeric Post IDs');
  var lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw Error('Another Google Love operation is starting');
  try{
    var previous=glb_backgroundRead_();
    if(previous&&(previous.state==='RUNNING'||previous.state==='REVIEW REQUIRED'||previous.inflight))
      throw Error('A batch is running or has an unresolved stage. Inspect it before starting another; no automatic retry is permitted.');
    // Existing preparation retains row mapping and validation.
    var prepared=glb_prepare(ids.join('\n'));
    if(!prepared||!prepared.ids||!prepared.ids.length)throw Error('No Post IDs prepared');
    var job={version:1,ids:prepared.ids,index:0,phase:mode==='solOnly'?'SOL':'CI',
      mode:mode,state:'RUNNING',message:'Waiting for Apps Script scheduler',
      stopRequested:false,inflight:null,inflightSince:null,startedUtc:glb_utc_(),updatedUtc:glb_utc_()};
    // Persist before scheduling; fail closed if scheduling is unsuccessful.
    glb_backgroundWrite_(job);
    try {
      glb_backgroundRemoveTriggers_();
      ScriptApp.newTrigger(GLB_BG_HANDLER).timeBased().everyMinutes(1).create();
    }catch(e){
      glb_backgroundFinish_(job,'STOPPED','Trigger creation failed: '+String(e.message||e));
      throw e;
    }
    return glb_backgroundStatus();
  }finally{lock.releaseLock();}
}
function glb_backgroundStop(){
  var lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw Error('Google Love scheduler busy; try Stop again');
  try{
    var job=glb_backgroundRead_();
    if(!job)return {state:'IDLE',message:'No scheduled batch'};
    if(job.state!=='RUNNING')return glb_backgroundStatus();
    job.stopRequested=true;
    if(!job.inflight)glb_backgroundFinish_(job,'STOPPED','Stopped before next stage by operator');
    else{job.message='Stopping after current stage finishes';glb_backgroundWrite_(job);}
    return glb_backgroundStatus();
  }finally{lock.releaseLock();}
}
function glb_backgroundStatus(){
  var job=glb_backgroundRead_();
  if(!job)return {state:'IDLE',message:'No scheduled batch',ids:[],index:0,total:0,phase:'',currentId:''};
  var stale=!!(job.inflight&&Date.now()-Number(job.inflightSince||0)>15*60*1000);
  var canRecover=job.state==='REVIEW REQUIRED'&&job.phase==='SOL'&&
    !job.inflight&&/W1\.5D_Sol/.test(job.message||'');
  var canRecoverCI=job.state==='REVIEW REQUIRED'&&job.phase==='CI'&&
    !job.inflight&&/Post [0-9]+ CI failed: STOPPED at W0/.test(job.message||'');
  return {canRecover:canRecover,canRecoverCI:canRecoverCI,state:job.state,message:job.message,ids:job.ids||[],index:job.index||0,
    total:(job.ids||[]).length,phase:job.phase,mode:job.mode,stopRequested:!!job.stopRequested,
    currentId:job.ids[job.index]||'',inflight:job.inflight||'',stale:stale,updatedUtc:job.updatedUtc};
}
function glb_backgroundStep(){
  var lock=LockService.getScriptLock();
  if(!lock.tryLock(1000))return;
  var job,id,phase;
  try{
    job=glb_backgroundRead_();
    if(!job||job.state!=='RUNNING'){glb_backgroundRemoveTriggers_();return;}
    if(job.inflight){
      // Another invocation may still be running. Never launch the same paid stage twice.
      if(Date.now()-Number(job.inflightSince||0)>15*60*1000){
        job.state='REVIEW REQUIRED';
        job.message='Stage '+job.inflight+' is unresolved or timed out. Inspect Sheets and API usage; do not automatically retry.';
        glb_backgroundWrite_(job);
        glb_backgroundRemoveTriggers_();
      }
      return;
    }
    if(job.stopRequested){glb_backgroundFinish_(job,'STOPPED','Stopped before next stage');return;}
    if(job.index>=job.ids.length){glb_backgroundFinish_(job,'COMPLETE','All Post IDs completed through W5C');return;}
    id=job.ids[job.index];phase=job.phase;
    // Reserve the stage BEFORE executing any paid or sheet-mutating work.
    job.inflight=id+':'+phase;job.inflightSince=Date.now();
    job.message='Running '+id+' — '+phase;
    glb_backgroundWrite_(job);
  }finally{lock.releaseLock();}
  var result=null,error='';
  try{
    if(phase==='SOL'&&job.mode==='solOnly')result=glb_runSolOnly(id);
    else if(phase==='CI'||phase==='SOL')result=glb_runDirectPost(id,phase);
    else if(phase==='W5'||phase==='W5C')result=glb_runSeoStage(id,phase);
    else throw Error('Unexpected stage '+phase);
    if(!result||!result.success)error=String(result&&result.message||'Stage failed');
  }catch(e){error=String(e&&e.message||e);}
  // A script timeout before this point leaves the reservation intact: no replay.
  lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw Error('Cannot record Google Love stage outcome; reservation retained for review');
  try{
    var latest=glb_backgroundRead_();
    if(!latest||latest.inflight!==id+':'+phase)throw Error('Google Love stage reservation changed; review required');
    if(error){
      glb_backgroundFinish_(latest,'REVIEW REQUIRED','Post '+id+' '+phase+' failed: '+error+'. No automatic retry.');
      return;
    }
    latest.inflight=null;latest.inflightSince=null;
    if(phase==='CI')latest.phase='SOL';
    else if(phase==='SOL')latest.phase='W5';
    else if(phase==='W5')latest.phase='W5C';
    else {latest.index++;latest.phase=latest.mode==='solOnly'?'SOL':'CI';}
    latest.message='Completed '+id+' — '+phase;
    if(latest.stopRequested)glb_backgroundFinish_(latest,'STOPPED','Stopped after '+id+' — '+phase);
    else if(latest.index>=latest.ids.length)
      glb_backgroundFinish_(latest,'COMPLETE','All '+latest.ids.length+' Post IDs completed through W5C');
    else glb_backgroundWrite_(latest);
  }finally{lock.releaseLock();}
}

/**
 * Operator-attested recovery of a CI run completed manually after W0 input repair.
 * No CI/API work is repeated. Accepts only the exact stopped CI batch item.
 */
function glb_resumeAfterManualCI(postId,confirmed){
 var id=String(postId||'').trim();
 if(!/^\d+$/.test(id)||confirmed!==true)throw Error('Confirm successful manual Content Intelligence for the exact Post ID.');
 var lock=LockService.getScriptLock();
 if(!lock.tryLock(15000))throw Error('Google Love recovery is busy');
 try{
  var job=glb_backgroundRead_();
  if(!job||job.state!=='REVIEW REQUIRED'||job.phase!=='CI'||
     String(job.ids[job.index])!==id||job.inflight||
     !/Post [0-9]+ CI failed: STOPPED at W0/.test(job.message||''))
   throw Error('This is not the current stopped W0 Content Intelligence item.');
  var queue=glb_queueState_(id);
  if(String(queue.values[4])!=='CI REVIEW'||String(queue.values[5]||'').trim()||
     String(queue.values[6]||'').trim())
   throw Error('Queue is not an uncompleted CI REVIEW without a prior CI or Sol completion.');
  var article=glb_lookup_(id);
  if(Number(queue.values[1])!==article.row)throw Error('Posts row changed since batch started.');
  var posts=glb_posts_(),headers=posts.getRange(1,1,1,posts.getLastColumn()).getDisplayValues()[0];
  function field(name){
   var col=headers.indexOf(name);
   if(col<0)throw Error('Missing posts column '+name);
   return String(posts.getRange(article.row,col+1).getDisplayValue()||'').trim();
  }
  var locality=field('Locality'),parent=field('Parent Area'),context=field('Location Context'),
      term=field('Primary Search Term');
  if(article.type==='Case Study'&&(!locality||!parent||!context))
   throw Error('Case Study still needs Locality, Parent Area and Location Context.');
  if(!term||locality&&term.toLowerCase().indexOf(locality.toLowerCase())<0)
   throw Error('Primary Search Term must contain the full Locality.');
  if(posts.getRange(article.row,4).getBackground().toLowerCase()!=='#2196f3')
   throw Error('Successful Content Intelligence completion marker (blue Post ID cell) was not found.');
  if(article.type==='Case Study'&&field('Evidence Status').indexOf('SOURCE-QUOTED | sha256:')!==0)
   throw Error('Completed case-study source-quoted evidence was not found.');
  // Preflight Sol prompt and its prerequisites before changing any queue state.
  glb_buildSolPrompt_(article);
  // Ensure the background trigger is installed before clearing the review state.
  glb_backgroundRemoveTriggers_();
  try{ScriptApp.newTrigger(GLB_BG_HANDLER).timeBased().everyMinutes(1).create();}
  catch(e){throw Error('Could not schedule SOL; queue review retained: '+String(e.message||e));}
  var stamp=glb_utc_();
  queue.sh.getRange(queue.row,6).setValue(stamp);
  glb_mark_(id,'CI DONE','Operator-confirmed manual CI success after W0 recovery; verified source fields, marker and Sol prerequisites. Original CI cost not recharged.');
  job.phase='SOL';job.state='RUNNING';job.stopRequested=false;
  job.message='Manual CI acknowledged for Post '+id+'; scheduling SOL without rerunning CI.';
  glb_backgroundWrite_(job);
  return {success:true,message:job.message};
 }finally{lock.releaseLock();}
}
