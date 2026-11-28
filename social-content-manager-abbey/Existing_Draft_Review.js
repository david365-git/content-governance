/**
 * Preview improved GBP wording for existing queue items without overwriting posts.
 * Explicit REPLACE decision is required before applying any rewrite.
 */
const SCM_REWRITE_TAB='Draft Rewrite Review';
function scm_rewriteSheet_() {
  const book=scm_book_();
  let sheet=book.getSheetByName(SCM_REWRITE_TAB);
  if(!sheet){
    sheet=book.insertSheet(SCM_REWRITE_TAB);
    sheet.appendRow([
      'Social Post ID','Article Post ID','Opportunity ID',
      'Current Draft','Suggested Draft','Decision','Status','Generated UTC','Notes'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function scm_reviewRows_() {
  const sheet=scm_rewriteSheet_(),data=sheet.getDataRange().getValues();
  const header=(data.shift()||[]).map(String);
  return {sheet:sheet,header:header,rows:data.map((values,index)=>({
    number:index+2,item:Object.fromEntries(header.map((key,j)=>[key,values[j]]))
  }))};
}
function scm_reviewPut_(review,row,data) {
  review.header.forEach((key,index)=>{
    if(Object.prototype.hasOwnProperty.call(data,key) &&
       String(row.item[key]||'')!==String(data[key]||'')){
      review.sheet.getRange(row.number,index+1).setValue(data[key]);
      row.item[key]=data[key];
    }
  });
}
function scm_previewExistingGBP() {
  scm_checkSetup();
  const articles=scm_rows_(SCM.articles).rows,
    imageRows=scm_rows_(SCM.images).rows, queue=scm_rows_(SCM.queue).rows;
  const aById=new Map(articles.map(r=>[String(r.item['Article Post ID']),r.item]));
  const imageById=new Map(imageRows.map(r=>[String(r.item['Opportunity ID']),r.item]));
  const review=scm_reviewRows_();
  const existing=new Map(review.rows.map(r=>[String(r.item['Social Post ID']),r]));
  let created=0,refreshed=0,skipped=0;
  for(const row of queue){
    const p=row.item;
    if(p.Platform!=='GBP')continue;
    const a=aById.get(String(p['Article Post ID']));
    const im=imageById.get(String(p['Opportunity ID']));
    if(!a||!im){skipped++;continue;}
    const proposed=scm_draftFromEvidence_(a,im);
    if(!proposed.trim()){skipped++;continue;}
    const fields={
      'Social Post ID':p['Social Post ID'],
      'Article Post ID':p['Article Post ID'],
      'Opportunity ID':p['Opportunity ID'],
      'Current Draft':p['Draft Text'],
      'Suggested Draft':proposed,
      'Decision':'REVIEW',
      'Status':'PREVIEW ONLY',
      'Generated UTC':new Date().toISOString(),
      'Notes':'Compare wording before choosing REPLACE or KEEP. No queue edit.'
    };
    const old=existing.get(String(p['Social Post ID']));
    if(old){
      if(String(old.item.Decision)==='REPLACE'||String(old.item.Status)==='APPLIED') {
        skipped++;continue;
      }
      // Refresh content but preserve a prior KEEP decision.
      if(old.item.Decision==='KEEP')delete fields.Decision;
      scm_reviewPut_(review,old,fields);refreshed++;
    } else {
      review.sheet.appendRow(review.header.map(key=>fields[key]===undefined?'':fields[key]));
      created++;
    }
  }
  scm_notice_('Rewrite preview: '+created+' new, '+refreshed+' refreshed. Existing posts unchanged.');
  return {created:created,refreshed:refreshed,skipped:skipped};
}
/** Only explicitly REPLACE-marked drafts may be updated. Never touch approved or published posts. */
function scm_applySelectedGBPRewrites() {
  scm_checkSetup();
  const preview=scm_reviewRows_(),queue=scm_rows_(SCM.queue);
  const posts=new Map(queue.rows.map(r=>[String(r.item['Social Post ID']),r]));
  let applied=0,blocked=0;
  const col=(name)=>{const c=queue.keys.indexOf(name);if(c<0)throw new Error('Missing queue field '+name);return c+1;};
  for(const row of preview.rows) {
    if(String(row.item.Decision)!=='REPLACE'||String(row.item.Status)==='APPLIED')continue;
    const p=posts.get(String(row.item['Social Post ID']));
    const state=p?String(p.item['Publication Status']||''):'';
    if(!p || String(p.item.Approval)!=='DRAFT' ||
       ['PUBLISHED','PUBLISHING','REVIEW REQUIRED'].includes(state)){
      scm_reviewPut_(preview,row,{'Status':'BLOCKED','Notes':'Only unapproved DRAFT posts can be replaced.'});
      blocked++;continue;
    }
    const suggestion=String(row.item['Suggested Draft']||'').trim();
    if(suggestion.length<80||suggestion.length>1500){
      scm_reviewPut_(preview,row,{'Status':'BLOCKED','Notes':'Suggested copy failed length check.'});
      blocked++;continue;
    }
    queue.sheet.getRange(p.number,col('Draft Text')).setValue(suggestion);
    queue.sheet.getRange(p.number,col('Revision')).setValue((Number(p.item.Revision)||1)+1);
    queue.sheet.getRange(p.number,col('Notes')).setValue('Rewritten from article evidence; review again before approval.');
    scm_reviewPut_(preview,row,{'Status':'APPLIED','Notes':'Applied to unapproved DRAFT; previous text retained in Current Draft column.'});
    applied++;
  }
  scm_notice_(applied+' GBP drafts updated; '+blocked+' blocked. No posts published.');
  return {applied:applied,blocked:blocked};
}
