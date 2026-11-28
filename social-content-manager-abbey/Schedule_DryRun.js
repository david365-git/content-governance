/** Social Content Manager: schedule validation and dry-run only.
 * Compatible with Code.js; do not enable a live publisher until separately reviewed.
 */
const SCM_SCHEDULE = Object.freeze({
  queue:'Social Posts Queue',log:'Publication Log',settings:'Settings',
  timezone:'Europe/London'
});
function scm_scheduleConfig_(){
  const settings=scm_rows_(SCM_SCHEDULE.settings).rows;
  const values={};settings.forEach(r=>values[String(r.item.Setting)]=String(r.item.Value));
  if(values['Auto Publish']!=='NO') throw new Error('Dry run safety stop: Auto Publish must be NO.');
  return values;
}
function scm_validDate_(value){
  if(value instanceof Date && !isNaN(value.valueOf()))
    return Utilities.formatDate(value,SCM_SCHEDULE.timezone,'yyyy-MM-dd');
  const s=String(value||'').trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
  if(/^\d{2}\/\d{2}\/\d{4}$/.test(s)){const p=s.split('/');return p[2]+'-'+p[1]+'-'+p[0];}
  return '';
}
function scm_validTime_(value){
  if(value instanceof Date && !isNaN(value.valueOf()))
    return Utilities.formatDate(value,SCM_SCHEDULE.timezone,'HH:mm');
  const s=String(value||'').trim();
  const m=s.match(/^([01]?\d|2[0-3]):([0-5]\d)(?::\d\d)?$/);
  return m?m[1].padStart(2,'0')+':'+m[2]:'';
}
function scm_validateQueuedPost_(p){
  const reasons=[];
  if(p.Platform!=='GBP')reasons.push('Only GBP is supported in this phase');
  if(p.Approval!=='APPROVED')reasons.push('Post is not APPROVED');
  if(!String(p['Social Post ID']||'').trim())reasons.push('Social Post ID is missing');
  if(!String(p['Draft Text']||'').trim())reasons.push('Post text is empty');
  if(!/^https:\/\/www\.abbeyfloorcare\.co\.uk\//.test(String(p['Article URL']||'')))
    reasons.push('Valid www Abbey article URL required');
  if(!/^https:\/\//.test(String(p['Prepared JPG URL']||'')) ||
     !/\.jpe?g(?:[?#]|$)/i.test(String(p['Prepared JPG URL']||'')))
    reasons.push('Prepared JPG URL required (image processing is not yet active)');
  if(!scm_validDate_(p['Scheduled Date']))reasons.push('Scheduled Date is invalid or absent');
  if(!scm_validTime_(p['Scheduled Time (UK)']))reasons.push('Scheduled Time (UK) is invalid or absent');
  if(String(p['Publication Status'])==='PUBLISHED')reasons.push('Already published');
  if(String(p['Publication Status'])==='PUBLISHING')reasons.push('Already being published');
  return reasons;
}
function scm_scheduledDue_(p, now){
  const d=scm_validDate_(p['Scheduled Date']),t=scm_validTime_(p['Scheduled Time (UK)']);
  if(!d||!t)return false;
  // Treat scheduled date/time as UK local time, including seasonal clock changes.
  // Search a small UTC offset window to avoid naive parsing in script timezone.
  const targetLocal=d+' '+t;
  const currentLocal=Utilities.formatDate(now,SCM_SCHEDULE.timezone,'yyyy-MM-dd HH:mm');
  return targetLocal<=currentLocal;
}
function scm_previewDueGBP(){
  scm_scheduleConfig_();
  scm_syncPublishingSchedule();
  const now=new Date();
  const queue=scm_rows_(SCM_SCHEDULE.queue).rows;
  const result={due:[],blocked:[],future:[]};
  queue.forEach(row=>{
    const p=row.item, id=String(p['Social Post ID']||'');
    if(!id)return;
    const problems=scm_validateQueuedPost_(p);
    if(problems.length){result.blocked.push({id:id,issues:problems});return;}
    if(scm_scheduledDue_(p,now)) result.due.push({id:id,postId:String(p['Article Post ID']),row:row.number});
    else result.future.push({id:id,row:row.number});
  });
  scm_notice_(
    'Dry-run: '+result.due.length+' due, '+result.future.length+' future, '+result.blocked.length+' blocked. No posts sent.'
  );
  return result;
}
/** Emits a recorded preview, never invokes a social API. Repeated dry runs log again intentionally. */
function scm_logDryRunGBP(){
  const report=scm_previewDueGBP(), at=new Date().toISOString();
  report.due.forEach(post=>scm_append_(SCM_SCHEDULE.log,{
    'Event ID':'EVT-'+Utilities.getUuid().slice(0,12),
    'Social Post ID':post.id,'Article Post ID':post.postId,
    'Platform':'GBP','Action':'DRY RUN','Attempt UTC':at,
    'Result':'ELIGIBLE — NOT PUBLISHED',
    'Details':'Validated scheduling and approval only; external publishing disabled',
    'Request Fingerprint':post.id+'|DRY-RUN|'+at
  }));
  return {eligible:report.due.length,blocked:report.blocked.length,future:report.future.length};
}

/** Reflect queue approvals and date changes in the calendar before any schedule check. */
function scm_syncPublishingSchedule() {
  scm_scheduleConfig_();
  const queue=scm_rows_(SCM_SCHEDULE.queue);
  const cal=scm_rows_('Publishing Schedule');
  const indexed=new Map(cal.rows.map(r=>[String(r.item['Social Post ID']),r]));
  let updated=0,added=0;
  const put=(table,row,name,value)=>{
    const column=table.keys.indexOf(name);
    if(column<0)throw new Error('Missing column '+name);
    if(String(row.item[name]||'')===String(value||''))return;
    table.sheet.getRange(row.number,column+1).setValue(value);
    row.item[name]=value;
  };
  for(const r of queue.rows) {
    const p=r.item;
    if(p.Platform!=='GBP'||!scm_validDate_(p['Scheduled Date'])||!scm_validTime_(p['Scheduled Time (UK)']))continue;
    const approval=String(p.Approval||'DRAFT');
    const prior=String(p['Publication Status']||'');
    const status=prior==='PUBLISHED'||prior==='PUBLISHING'||prior==='REVIEW REQUIRED'
      ?prior:(approval==='APPROVED'?'SCHEDULED':'NOT SCHEDULED');
    put(queue,r,'Publication Status',status);
    const fields={
      'Social Post ID':p['Social Post ID'],'Article Post ID':p['Article Post ID'],
      'Image Opportunity ID':p['Opportunity ID'],'Platform':'GBP',
      'Scheduled Date':scm_validDate_(p['Scheduled Date']),
      'Scheduled Time (UK)':scm_validTime_(p['Scheduled Time (UK)']),
      'Approval':approval,'Publication Status':status,
      'Last Checked UTC':scm_now_()
    };
    const calendarRow=indexed.get(String(p['Social Post ID']));
    if(calendarRow){
      ['Scheduled Date','Scheduled Time (UK)','Approval','Publication Status','Last Checked UTC']
        .forEach(name=>put(cal,calendarRow,name,fields[name]));
      updated++;
    }else{scm_append_('Publishing Schedule',fields);added++;}
  }
  return {updated:updated,added:added};
}
