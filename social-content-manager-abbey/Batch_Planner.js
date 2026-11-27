/**
 * Social Content Manager Abbey — batch planning (GBP only).
 * Safe for a standalone Apps Script project. Never publishes externally.
 * Run: scm_batchPlanGBP, then review generated drafts, dates and image choices.
 */
const SCM_BATCH = Object.freeze({
  daysBetween: 7,
  scheduledTime: '10:00',
  startDelayDays: 3,
  maxNewPostsPerRun: 60
});
function scm_captionForImage_(html, url) {
  const parts = String(html || '').match(/<figure\b[\s\S]*?<\/figure>/gi) || [];
  for (let i=0;i<parts.length;i++) {
    const figure=parts[i];
    if (figure.indexOf(url)<0) continue;
    const cap=(figure.match(/<figcaption\b[^>]*>([\s\S]*?)<\/figcaption>/i)||[])[1] || '';
    return scm_plainText_(cap);
  }
  return '';
}
function scm_plainText_(html) {
  return String(html||'').replace(/<[^>]*>/g,' ').replace(/&nbsp;/gi,' ')
    .replace(/&amp;/gi,'&').replace(/&#39;/g,"'")
    .replace(/&quot;/gi,'"').replace(/\s+/g,' ').trim();
}
function scm_isProjectImage_(img, caption) {
  const url=String(img['Original WebP URL']||'');
  const alt=String(img['Original Alt / Caption']||'');
  if(!/^https:\/\/www\.abbeyfloorcare\.co\.uk\/wp-content\/uploads\//i.test(url))return false;
  if(/(?:david[_-]?allen|author|logo|avatar|headshot|portrait)/i.test(url+' '+alt+' '+caption))return false;
  if(!caption || caption.length < 12)return false;
  if(/(?:affiliate|product image|advertisement)/i.test(caption+' '+alt))return false;
  return true;
}
function scm_setField_(table,row,key,value) {
  const col=table.keys.indexOf(key);
  if(col<0)throw new Error('Missing column '+key);
  table.sheet.getRange(row.number,col+1).setValue(value);
  row.item[key]=value;
}
/** Make an editorial first-pass selection, but preserve all explicit GENERATE/SKIP decisions. */
function scm_selectDistinctProjectImages_() {
  const articles=scm_rows_(SCM.articles).rows;
  const articleMap=new Map(articles.map(r=>[String(r.item['Article Post ID']),r.item]));
  const images=scm_rows_(SCM.images);
  let selected=0,skipped=0,untouched=0;
  const seenByArticle=new Map();
  for(const row of images.rows) {
    const item=row.item, decision=String(item.Decision||'').toUpperCase();
    if(decision!=='REVIEW'){untouched++;continue;}
    const article=articleMap.get(String(item['Article Post ID']));
    if(!article){untouched++;continue;}
    const caption=scm_captionForImage_(article['Finished HTML'],String(item['Original WebP URL']));
    if(!scm_isProjectImage_(item,caption)) {
      scm_setField_(images,row,'Decision','SKIP');
      scm_setField_(images,row,'Reason / Notes','Excluded by initial photo relevance review; manual override possible.');
      skipped++;continue;
    }
    const grouping=String(item['Article Post ID']);
    const seen=seenByArticle.get(grouping)||new Set();
    // Caption evidence is used as an educational hook; near-duplicate captions remain in review.
    const key=caption.toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
    if(seen.has(key)){untouched++;continue;}
    seen.add(key);seenByArticle.set(grouping,seen);
    scm_setField_(images,row,'Distinct Post Angle',caption);
    scm_setField_(images,row,'Educational Value','Different documented project moment: '+caption);
    scm_setField_(images,row,'Decision','GENERATE');
    scm_setField_(images,row,'Reason / Notes','Automatically selected from documented article figure caption; review before approval.');
    selected++;
  }
  return {selected:selected,skipped:skipped,untouched:untouched};
}
function scm_nextUKDay_(daysAhead) {
  // Noon UTC avoids DST crossover while producing the target UK calendar date.
  const now=new Date(), utcNoon=new Date(Date.UTC(now.getUTCFullYear(),now.getUTCMonth(),now.getUTCDate()+daysAhead,12));
  return Utilities.formatDate(utcNoon,'Europe/London','yyyy-MM-dd');
}
function scm_dateAdd_(yyyyMMdd, days) {
  const p=yyyyMMdd.split('-').map(Number);
  const d=new Date(Date.UTC(p[0],p[1]-1,p[2]+days,12));
  return Utilities.formatDate(d,'UTC','yyyy-MM-dd');
}
/** Schedule unscheduled GBP drafts without approving them or overwriting existing dates. */
function scm_autoScheduleGBP() {
  scm_scheduleConfig_();
  const queue=scm_rows_(SCM.queue);
  const existing=queue.rows.filter(r=>r.item.Platform==='GBP' && scm_validDate_(r.item['Scheduled Date']));
  const usedDates=new Set(existing.map(r=>scm_validDate_(r.item['Scheduled Date'])));
  const waiting=queue.rows.filter(r=>r.item.Platform==='GBP' &&
    !scm_validDate_(r.item['Scheduled Date']) &&
    !['REJECTED','PAUSED'].includes(String(r.item.Approval||'')) &&
    !['PUBLISHED','PUBLISHING'].includes(String(r.item['Publication Status']||'')));
  const buckets=new Map();
  for(const row of waiting){
    const id=String(row.item['Article Post ID']);
    if(!buckets.has(id))buckets.set(id,[]);
    buckets.get(id).push(row);
  }
  const articleIds=Array.from(buckets.keys());
  let cursor=0, offset=SCM_BATCH.startDelayDays, planned=0;
  const schedule=scm_rows_('Publishing Schedule');
  const scheduledIds=new Set(schedule.rows.map(r=>String(r.item['Social Post ID'])));
  while(articleIds.some(id=>buckets.get(id).length) && planned<SCM_BATCH.maxNewPostsPerRun){
    const id=articleIds[cursor++%articleIds.length], bucket=buckets.get(id);
    if(!bucket.length)continue;
    const row=bucket.shift();
    let date=scm_nextUKDay_(offset);
    while(usedDates.has(date)){offset+=SCM_BATCH.daysBetween;date=scm_nextUKDay_(offset);}
    usedDates.add(date);offset+=SCM_BATCH.daysBetween;
    scm_setField_(queue,row,'Scheduled Date',date);
    scm_setField_(queue,row,'Scheduled Time (UK)',SCM_BATCH.scheduledTime);
    const approved=String(row.item.Approval)==='APPROVED';
    scm_setField_(queue,row,'Publication Status',approved?'SCHEDULED':'NOT SCHEDULED');
    if(!scheduledIds.has(String(row.item['Social Post ID']))) {
      scm_append_('Publishing Schedule',{
        'Social Post ID':row.item['Social Post ID'],
        'Article Post ID':row.item['Article Post ID'],
        'Image Opportunity ID':row.item['Opportunity ID'],
        'Platform':'GBP',
        'Scheduled Date':date,'Scheduled Time (UK)':SCM_BATCH.scheduledTime,
        'Approval':row.item.Approval,
        'Publication Status':approved?'SCHEDULED':'NOT SCHEDULED',
        'Last Checked UTC':scm_now_(),
        'Notes':'Auto-planned; no external posting and approval still required.'
      });
      scheduledIds.add(String(row.item['Social Post ID']));
    }
    planned++;
  }
  return {planned:planned,remaining:waiting.length-planned};
}
/** One run: safe selection, generate the eligible drafts, schedule them. */
function scm_batchPlanGBP() {
  scm_checkSetup();
  const selection=scm_selectDistinctProjectImages_();
  const generated=scm_generateGBP();
  const scheduled=scm_autoScheduleGBP();
  scm_book_().toast('GBP plan: '+generated.created+' drafts; '+scheduled.planned+' new dates. Publishing remains OFF.');
  return {selection:selection,generated:generated,scheduled:scheduled};
}
