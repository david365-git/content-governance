/**
 * Recover W1.5D_Sol internal links from the last saved Sol draft.
 * Never invokes Content Intelligence or the full-price Sol model.
 * Only offered for a stopped background batch at its failed SOL phase.
 */
function glb_recoverLinksOnly(postId){
 var id=String(postId||'').trim();
 if(!/^\d+$/.test(id))throw Error('Invalid Post ID');
 var lock=LockService.getScriptLock();
 if(!lock.tryLock(15000))throw Error('Google Love recovery is busy');
 var a,draft,previous,job,htmlColumn;
 try{
  job=glb_backgroundRead_();
  if(!job||job.state!=='REVIEW REQUIRED'||job.phase!=='SOL'||
      job.ids[job.index]!==id||job.inflight||!/W1\.5D_Sol/.test(job.message||''))
   throw Error('Recovery allowed only for the current stopped W1.5D_Sol failure. No Sol regeneration performed.');
  var queue=glb_queueState_(id);
  if(String(queue.values[4])!=='SOL REVIEW')throw Error('Expected SOL REVIEW queue status for '+id);
  a=glb_lookup_(id);
  if(Number(queue.values[1])!==a.row)throw Error('Post row changed; recovery cancelled');
  var ss=SpreadsheetApp.getActiveSpreadsheet(),out=ss.getSheetByName('Google Love Production Drafts');
  if(!out||out.getLastRow()<2)throw Error('No saved Google Love Production Drafts');
  var rows=out.getRange(2,1,out.getLastRow()-1,8).getValues(),latest=null;
  // The most recent draft for this exact Post ID and posts row wins.
  for(var i=rows.length-1;i>=0;i--){
   if(String(rows[i][0]).trim()===id&&Number(rows[i][1])===a.row){latest=rows[i];break;}
  }
  if(!latest||!String(latest[3]||'').trim())throw Error('No matching saved Sol draft for '+id);
  draft=String(latest[3]);
  if(draft.length>48000||!/<h2\b/i.test(draft)||!/<p\b/i.test(draft))
   throw Error('Saved draft failed HTML validation');
  var target=glb_posts_();htmlColumn=glb_col_(target,'New HTML',false);
  previous=String(target.getRange(a.row,htmlColumn).getValue()||'');
  if(previous!==String(latest[7]||''))
   throw Error('New HTML has changed since the saved Sol draft; will not overwrite newer work.');
  // Reserve before any API request; failed or interrupted recovery never retries itself.
  job.inflight=id+':LINK_RECOVERY';job.inflightSince=Date.now();
  job.message='W1.5D_Sol link recovery running for Post '+id;
  glb_backgroundWrite_(job);
 }finally{lock.releaseLock();}
 var linked,error='';
 try{
  linked=glb_runW15DSol_(a,draft);
  if(!linked||!String(linked.html||'').trim())throw Error('No linked HTML returned');
  linked.html=glb_canonicalAbbeyUrls_(linked.html);
  if(linked.html.length>48000||!/<h2\b/i.test(linked.html)||!/<p\b/i.test(linked.html))
   throw Error('Linked HTML failed validation');
 }catch(e){error=String(e&&e.message||e);}
 lock=LockService.getScriptLock();
 if(!lock.tryLock(15000))throw Error('Could not finalize link recovery. Reserved request requires review.');
 try{
  job=glb_backgroundRead_();
  if(!job||job.inflight!==id+':LINK_RECOVERY')
   throw Error('Recovery reservation changed; manual inspection required.');
  if(error){
   job.state='REVIEW REQUIRED';job.inflight=null;job.inflightSince=null;
   job.message='W1.5D_Sol recovery failed for Post '+id+': '+error+'; saved Sol draft retained, no automatic retry.';
   glb_backgroundWrite_(job);
   return {success:false,message:job.message};
  }
  var target=glb_posts_();
  if(String(target.getRange(a.row,htmlColumn).getValue()||'')!==previous)
   throw Error('New HTML changed during link recovery. Reservation retained for review; do not overwrite.');
  // Commit only the verified linked draft. Sol API charges are never repeated or reset.
  target.getRange(a.row,htmlColumn).setValue(linked.html);
  var stamp=glb_stamp_(a.row,GLB_SOL_HEADER),queue=glb_queueState_(id);
  queue.sh.getRange(queue.row,7).setValue(stamp);
  glb_mark_(id,'SOL DONE','Recovered verified W1.5D_Sol internal links from the saved Sol draft; no Sol regeneration.');
  job.inflight=null;job.inflightSince=null;job.phase='W5';job.stopRequested=false;
  job.message='Links recovered for Post '+id+'; resuming at W5.';
  try{
   glb_backgroundRemoveTriggers_();
   ScriptApp.newTrigger(GLB_BG_HANDLER).timeBased().everyMinutes(1).create();
   job.state='RUNNING';
  }catch(e){
   job.state='REVIEW REQUIRED';
   job.message='Links saved successfully for '+id+', but background trigger failed: '+String(e.message||e)+'. W5 not scheduled.';
  }
  glb_backgroundWrite_(job);
  return {success:job.state==='RUNNING',message:job.message,
   linkCount:(linked.links||[]).length,phase:job.phase};
 }finally{lock.releaseLock();}
}
