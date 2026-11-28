/** Manual social publishing: GBP, Facebook, X. Never sends a post to a platform. */
const SCM_MANUAL = Object.freeze({
  ready:'Ready to Publish', published:'Published Posts',
  destinations:{GBP:'https://business.google.com/',FACEBOOK:'https://www.facebook.com/',X:'https://x.com/home'},
  days:{GBP:[1,4],FACEBOOK:[2,5],X:[1,3,6]},
  headers:['Scheduled Date','Platform','Article','Image','Social Post ID','Post Wording','Prepared JPG URL','Article URL','Login URL','Status','Article Post ID','Opportunity ID','Published UTC','Published Post URL']
});
function scm_manualTab_(name) {
  const book=scm_book_();
  let sheet=book.getSheetByName(name);
  if(!sheet) {
    sheet=book.insertSheet(name,0);
    sheet.appendRow(SCM_MANUAL.headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,SCM_MANUAL.headers.length).setFontWeight('bold').setBackground('#e8f0e9');
    sheet.setColumnWidth(3,260);
    sheet.setColumnWidth(6,440);
    sheet.setColumnWidth(7,320);
  }
  const header=sheet.getRange(1,1,1,SCM_MANUAL.headers.length).getValues()[0].map(String);
  if(header.join('|')!==SCM_MANUAL.headers.join('|'))
    throw new Error('Unexpected headers in '+name+'. Existing data preserved.');
  return sheet;
}
function scm_manualText_(p,article,image) {
  const source=String(p['Article URL']||'').trim();
  const title=String(article?.Title||'stone restoration').trim();
  const caption=String(image?.['Distinct Post Angle']||image?.['Original Alt / Caption']||'').trim();
  const raw=String(p['Draft Text']||'').trim();
  const paragraphs=raw.split(/\n\s*\n/).filter(Boolean);
  const fact=(paragraphs.find(x=>x.length>65&&!x.includes('https://'))||caption||title).slice(0,330);
  if(String(p.Platform)==='X') {
    const base=(caption||title).slice(0,82)+' — '+fact.slice(0,110);
    // Reserve 24 URL characters plus separators under the standard X limit.
    return base.slice(0,245)+'\n'+source;
  }
  if(String(p.Platform)==='FACEBOOK')return [caption||title,fact,'See the full Abbey Floor Care project:',source].filter(Boolean).join('\n\n');
  return raw;
}
/** Create one Facebook and one X draft for each GBP photo, without duplicating reruns. */
function scm_createFacebookAndXDrafts() {
  scm_checkSetup();
  const queue=scm_rows_(SCM.queue),existing=new Set(queue.rows.map(r=>String(r.item.Platform)+'|'+String(r.item['Opportunity ID'])));
  const articles=new Map(scm_rows_(SCM.articles).rows.map(r=>[String(r.item['Article Post ID']),r.item]));
  const images=new Map(scm_rows_(SCM.images).rows.map(r=>[String(r.item['Opportunity ID']),r.item]));
  let created=0;
  for(const row of queue.rows) {
    const p=row.item;
    if(String(p.Platform)!=='GBP')continue;
    for(const platform of ['FACEBOOK','X']) {
      const key=platform+'|'+String(p['Opportunity ID']);
      if(existing.has(key))continue;
      const newPost=Object.assign({},p,{
        'Social Post ID':scm_id_(platform==='FACEBOOK'?'FB':'X'),
        'Platform':platform,'Draft Text':'','Prepared JPG URL':'',
        'Approval':'DRAFT','Scheduled Date':'','Scheduled Time (UK)':'10:00',
        'Publication Status':'NOT SCHEDULED','Platform Post ID':'',
        'Published UTC':'','Last Error':'','Revision':1,
        'Notes':'Automatically derived from the source article for manual publishing; not posted.'
      });
      newPost['Draft Text']=scm_manualText_(newPost,articles.get(String(p['Article Post ID'])),images.get(String(p['Opportunity ID'])));
      scm_append_(SCM.queue,newPost);
      existing.add(key);created++;
    }
  }
  return {created};
}
function scm_manualNextSlots_(platform,total) {
  const days=SCM_MANUAL.days[platform],results=[],now=new Date();
  const start=new Date(Date.UTC(now.getUTCFullYear(),now.getUTCMonth(),now.getUTCDate()+1,12));
  for(let i=0;i<500&&results.length<total;i++) {
    const date=new Date(start.getTime()+i*86400000);
    if(!days.includes(date.getUTCDay()))continue;
    results.push(Utilities.formatDate(date,'UTC','yyyy-MM-dd'));
  }
  if(results.length<total)throw new Error('Scheduling horizon exceeded for '+platform);
  return results;
}
function scm_manualReschedule_() {
  const queue=scm_rows_(SCM.queue);
  const imgs=new Map(scm_rows_(SCM.images).rows.map(r=>[String(r.item['Opportunity ID']),r.item]));
  let scheduled=0;
  for(const platform of ['GBP','FACEBOOK','X']) {
    const available=queue.rows.filter(r=>String(r.item.Platform)===platform &&
      !['PUBLISHED','PUBLISHING','PAUSED','REJECTED'].includes(String(r.item['Publication Status'])) &&
      !['PAUSED','REJECTED'].includes(String(r.item.Approval)));
    const groups=new Map();
    available.forEach(r=>{
      const article=String(r.item['Article Post ID']);
      if(!groups.has(article))groups.set(article,[]);
      groups.get(article).push(r);
    });
    for(const group of groups.values())group.sort((a,b)=>{
      const ai=Number(imgs.get(String(a.item['Opportunity ID']))?.['Image Number']||0);
      const bi=Number(imgs.get(String(b.item['Opportunity ID']))?.['Image Number']||0);
      return ai-bi;
    });
    const order=[...groups.keys()].sort();
    // Rotate articles, preserving the order of photographs inside each article.
    const arranged=[];
    while(order.some(key=>groups.get(key).length)) {
      for(const key of order)if(groups.get(key).length)arranged.push(groups.get(key).shift());
    }
    const dates=scm_manualNextSlots_(platform,arranged.length);
    arranged.forEach((row,i)=>{
      scm_setField_(queue,row,'Scheduled Date',dates[i]);
      scm_setField_(queue,row,'Scheduled Time (UK)','10:00');
      scm_setField_(queue,row,'Publication Status',String(row.item.Approval)==='APPROVED'?'SCHEDULED':'NOT SCHEDULED');
      scheduled++;
    });
  }
  return {scheduled};
}
function scm_manualRefreshViews_() {
  const readySheet=scm_manualTab_(SCM_MANUAL.ready),publishedSheet=scm_manualTab_(SCM_MANUAL.published);
  const articles=new Map(scm_rows_(SCM.articles).rows.map(r=>[String(r.item['Article Post ID']),r.item]));
  const images=new Map(scm_rows_(SCM.images).rows.map(r=>[String(r.item['Opportunity ID']),r.item]));
  const queue=scm_rows_(SCM.queue).rows;
  const make=r=>{
    const p=r.item,article=articles.get(String(p['Article Post ID'])),image=images.get(String(p['Opportunity ID']));
    return [scm_validDate_(p['Scheduled Date']),String(p.Platform||''),String(article?.Title||''),
      String(image?.['Original Alt / Caption']||''),String(p['Social Post ID']||''),
      String(p['Draft Text']||''),String(p['Prepared JPG URL']||''),String(p['Article URL']||''),
      SCM_MANUAL.destinations[String(p.Platform)]||'',(String(p['Prepared JPG URL']||'') ? String(p['Publication Status']||'SCHEDULED') : 'IMAGE PENDING'),
      String(p['Article Post ID']||''),String(p['Opportunity ID']||''),
      String(p['Published UTC']||''),String(p['Platform Post ID']||'')];
  };
  const ready=queue.filter(r=>!['PUBLISHED','PAUSED','REJECTED'].includes(String(r.item['Publication Status'])) &&
    !['REJECTED','PAUSED'].includes(String(r.item.Approval)) &&
    ['GBP','FACEBOOK','X'].includes(String(r.item.Platform)))
    .sort((a,b)=>scm_validDate_(a.item['Scheduled Date']).localeCompare(scm_validDate_(b.item['Scheduled Date']))).map(make);
  const published=queue.filter(r=>String(r.item['Publication Status'])==='PUBLISHED')
    .sort((a,b)=>String(b.item['Published UTC']||'').localeCompare(String(a.item['Published UTC']||''))).map(make);
  for(const [sheet,values] of [[readySheet,ready],[publishedSheet,published]]) {
    if(sheet.getLastRow()>1)sheet.getRange(2,1,sheet.getLastRow()-1,SCM_MANUAL.headers.length).clearContent();
    if(values.length)sheet.getRange(2,1,values.length,SCM_MANUAL.headers.length).setValues(values);
  }
  return {ready:ready.length,published:published.length};
}
/** Initial setup / repeated sync. Does not publish externally. */
function scm_manualEnablePlatforms_() {
  const sheet=scm_imageSettings_(); // Ensure settings tab exists.
  const table=scm_rows_('Image Settings');
  const col=table.keys.indexOf('Enabled')+1;
  if(col<1)throw new Error('Image Settings: Enabled column missing');
  for(const row of table.rows) {
    if(['FACEBOOK','X'].includes(String(row.item.Platform).toUpperCase()) &&
       String(row.item.Enabled).toUpperCase()!=='YES')
      table.sheet.getRange(row.number,col).setValue('YES');
  }
}
function scm_manualGetPost(postId) {
  const row=scm_rows_(SCM.queue).rows.find(r=>String(r.item['Social Post ID'])===String(postId));
  if(!row)throw new Error('Post not found');
  const p=row.item;
  return {id:postId,platform:String(p.Platform),text:String(p['Draft Text']||''),
    image:String(p['Prepared JPG URL']||''),article:String(p['Article URL']||''),
    status:String(p['Publication Status']),login:SCM_MANUAL.destinations[String(p.Platform)]||''};
}
function scm_openManualPublisher() {
  const sheet=scm_book_().getActiveSheet();
  if(sheet.getName()!==SCM_MANUAL.ready)throw new Error('Select a post row on Ready to Publish first.');
  const row=sheet.getActiveRange().getRow();
  if(row<2)throw new Error('Select a post row, not the heading.');
  const id=String(sheet.getRange(row,5).getValue()||'');
  const template=HtmlService.createTemplateFromFile('Manual_Publish_Sidebar');
  template.postId=id;
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('Publish prepared post'));
}
function scm_buildManualPublishing() {
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000))return {status:'BUSY'};
  try {
    scm_manualEnablePlatforms_();
    const created=scm_createFacebookAndXDrafts();
    const scheduled=scm_manualReschedule_();
    const images=scm_autoPrepareQueuedImages();
    const views=scm_manualRefreshViews_();
    return {status:'PREPARED — MANUAL PUBLICATION ONLY',created,scheduled,images,views};
  } finally {lock.releaseLock();}
}
function scm_manualMarkPublished(postId,postUrl) {
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1500))throw new Error('Publishing sheet is busy; please retry.');
  try {
    const queue=scm_rows_(SCM.queue);
    const matches=queue.rows.filter(r=>String(r.item['Social Post ID'])===String(postId));
    if(matches.length!==1)throw new Error('Post ID not uniquely found.');
    const row=matches[0],p=row.item;
    if(String(p['Publication Status'])==='PUBLISHED')return {status:'ALREADY PUBLISHED'};
    if(!['GBP','FACEBOOK','X'].includes(String(p.Platform)))throw new Error('Invalid platform');
    if(!/^https:\/\/www\.abbeyfloorcare\.co\.uk\//.test(String(p['Article URL']||'')))throw new Error('Invalid article link');
    if(!String(p['Prepared JPG URL']||''))throw new Error('Image not ready');
    const at=scm_now_();
    scm_setField_(queue,row,'Publication Status','PUBLISHED');
    scm_setField_(queue,row,'Published UTC',at);
    if(postUrl)scm_setField_(queue,row,'Platform Post ID',String(postUrl).slice(0,400));
    scm_append_(SCM.log,{'Event ID':scm_id_('EVT'),'Social Post ID':postId,
      'Article Post ID':p['Article Post ID'],'Image Opportunity ID':p['Opportunity ID'],
      'Platform':p.Platform,'Action':'MANUAL PUBLISH CONFIRMED','Attempt UTC':at,
      'Result':'PUBLISHED','Published URL':String(postUrl||''),
      'Details':'Operator confirmed publication; no API call was made.',
      'Request Fingerprint':postId+'|MANUAL|'+at});
    scm_manualRefreshViews_();
    return {status:'PUBLISHED',postId};
  } finally {lock.releaseLock();}
}
