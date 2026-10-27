/**
 * Isolated Google Love: 20-article GPT-6.1 Sol batch.
 * No posts, technical, WordPress or production workflow writes.
 * Install with clasp; run glc_startSol20Batch once in Apps Script editor.
 * Time-driven trigger continues one article per execution.
 */
var GLC_SOL20_IDS = [
  '47185','48034','48032','48029','48109','48152','48150','48148','48146','48143',
  '7952','5774','21838','21746','8027','8336','12107','11452','22942','11945'
];
var GLC_SOL20_SHEET = 'GL Sol 20 Batch';
var GLC_SOL20_RUNNER = 'glc_processSol20Batch';
var GLC_SOL20_CAP_USD = 10;

function glc_sol20Sheet_() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName(GLC_SOL20_SHEET);
  if (!sh) {
    sh=ss.insertSheet(GLC_SOL20_SHEET);
    sh.getRange(1,1,1,7).setValues([['Post ID','State','Frozen prompt row','Comparison result row','API cost USD','Updated UTC','Message']]);
    sh.setFrozenRows(1);
    sh.getRange(2,1,GLC_SOL20_IDS.length,1).setValues(GLC_SOL20_IDS.map(function(id){return [id];}));
  }
  var existing=sh.getRange(2,1,GLC_SOL20_IDS.length,1).getDisplayValues().map(function(r){return r[0];});
  if (existing.join('|')!==GLC_SOL20_IDS.join('|')) throw new Error('Batch Post IDs no longer match the configured list. No action taken.');
  return sh;
}
function glc_sol20Clock_() {return Utilities.formatDate(new Date(),'Etc/UTC',"yyyy-MM-dd'T'HH:mm:ss'Z'");}
function glc_sol20Triggers_() {
  return ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()===GLC_SOL20_RUNNER;});
}
function glc_sol20RemoveTriggers_(){glc_sol20Triggers_().forEach(function(t){ScriptApp.deleteTrigger(t);});}
function glc_sol20Mark_(sh,i,state,message,extra) {
  var row=i+2;
  sh.getRange(row,2).setValue(state);
  if (extra && extra.promptRow) sh.getRange(row,3).setValue(extra.promptRow);
  if (extra && extra.resultRow) sh.getRange(row,4).setValue(extra.resultRow);
  if (extra && typeof extra.cost==='number') sh.getRange(row,5).setValue(extra.cost);
  sh.getRange(row,6).setValue(glc_sol20Clock_());
  sh.getRange(row,7).setValue(String(message||'').slice(0,1000));
}
function glc_startSol20Batch() {
  var lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error('Another batch operation is running.');
  try {
    var sh=glc_sol20Sheet_(), props=PropertiesService.getScriptProperties();
    if(props.getProperty('GLC_SOL20_INFLIGHT')) throw new Error('A previous API request is unresolved. Check the sheet; do not restart and risk double billing.');
    var data=sh.getRange(2,2,GLC_SOL20_IDS.length,1).getDisplayValues();
    var outstanding=data.filter(function(r){return r[0]!=='DONE';}).length;
    if(!outstanding) return 'All 20 article drafts are already complete.';
    props.setProperty('GLC_SOL20_ACTIVE','YES');
    glc_sol20RemoveTriggers_();
    ScriptApp.newTrigger(GLC_SOL20_RUNNER).timeBased().everyMinutes(1).create();
    return 'Scheduled GPT-6.1 Sol batch. '+outstanding+' articles outstanding. Check '+GLC_SOL20_SHEET+'. Run glc_stopSol20Batch to stop.';
  } finally {lock.releaseLock();}
}
function glc_stopSol20Batch() {
  PropertiesService.getScriptProperties().deleteProperty('GLC_SOL20_ACTIVE');
  glc_sol20RemoveTriggers_();
  return 'Stopped scheduling further API calls. Any currently executing call may finish.';
}
function glc_processSol20Batch() {
  var lock=LockService.getScriptLock();
  if(!lock.tryLock(1500)) return;
  var sh, props=PropertiesService.getScriptProperties(), id='',i=-1,promptRow=0;
  try {
    if(props.getProperty('GLC_SOL20_ACTIVE')!=='YES') {glc_sol20RemoveTriggers_();return;}
    sh=glc_sol20Sheet_();
    var inflight=props.getProperty('GLC_SOL20_INFLIGHT');
    if(inflight) {
      // Never replay an unresolved request; the previous execution may have incurred a charge.
      var age=Date.now()-Number(props.getProperty('GLC_SOL20_INFLIGHT_MS')||0);
      if(age>12*60*1000) {
        props.deleteProperty('GLC_SOL20_ACTIVE');
        glc_sol20RemoveTriggers_();
        sh.getRange('H1').setValue('STOPPED: unresolved API request for '+inflight+'. Inspect comparison sheet before any restart.');
      }
      return;
    }
    var costValues=sh.getRange(2,5,GLC_SOL20_IDS.length,1).getValues();
    var spent=costValues.reduce(function(sum,x){return sum+(Number(x[0])||0);},0);
    if(spent>=GLC_SOL20_CAP_USD) {
      props.deleteProperty('GLC_SOL20_ACTIVE'); glc_sol20RemoveTriggers_();
      sh.getRange('H1').setValue('STOPPED: $'+GLC_SOL20_CAP_USD+' recorded cost limit reached.');
      return;
    }
    var states=sh.getRange(2,2,GLC_SOL20_IDS.length,2).getValues();
    for(var j=0;j<GLC_SOL20_IDS.length;j++){
      if(String(states[j][0])!=='DONE') {i=j;promptRow=Number(states[j][1])||0;break;}
    }
    if(i<0) {
      props.deleteProperty('GLC_SOL20_ACTIVE'); glc_sol20RemoveTriggers_();
      sh.getRange('H1').setValue('COMPLETE: '+GLC_SOL20_IDS.length+' drafts saved.');
      return;
    }
    id=GLC_SOL20_IDS[i];
    // Reserve before the API request. A lost execution must never silently repeat it.
    props.setProperty('GLC_SOL20_INFLIGHT',id);
    props.setProperty('GLC_SOL20_INFLIGHT_MS',String(Date.now()));
    glc_sol20Mark_(sh,i,'IN PROGRESS','Preparing prompt or calling GPT-6.1 Sol.');
  } finally {lock.releaseLock();}

  try {
    if(!promptRow) {
      var prepared=glc_preparePrompt(id);
      promptRow=prepared.row;
      glc_sol20Mark_(sh,i,'PROMPT READY','SHA-256 '+prepared.hash,{promptRow:promptRow});
    }
    // Reuse the isolated comparison runner, with exactly the frozen prompt.
    var res=glc_apiRun(id,'API — GPT-6.1 Sol','1.2',promptRow);
    glc_sol20Mark_(sh,i,'DONE','Draft saved. Self-check not included in batch.',{
      promptRow:promptRow,resultRow:res.row,cost:Number(res.cost||0)
    });
  } catch(e) {
    // Stop on first error: this may be an API, timeout, unsupported-model,
    // prompt-length or cell-size issue. Never retry automatically.
    glc_sol20Mark_(sh,i,'STOPPED — REVIEW REQUIRED',String(e&&e.message||e),{promptRow:promptRow});
    props.deleteProperty('GLC_SOL20_ACTIVE');
    glc_sol20RemoveTriggers_();
  } finally {
    // Only clear after we have recorded a definitive result or failure.
    props.deleteProperty('GLC_SOL20_INFLIGHT');
    props.deleteProperty('GLC_SOL20_INFLIGHT_MS');
  }
}
function glc_sol20Status() {
  var sh=glc_sol20Sheet_(),props=PropertiesService.getScriptProperties();
  var vals=sh.getRange(2,2,GLC_SOL20_IDS.length,4).getValues();
  var done=vals.filter(function(r){return r[0]==='DONE';}).length;
  var spent=vals.reduce(function(v,r){return v+(Number(r[3])||0);},0);
  return {active:props.getProperty('GLC_SOL20_ACTIVE')==='YES',inflight:props.getProperty('GLC_SOL20_INFLIGHT')||null,done:done,total:GLC_SOL20_IDS.length,costUSD:spent,statusSheet:GLC_SOL20_SHEET};
}
