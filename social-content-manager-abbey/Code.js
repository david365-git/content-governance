/** Social Content Manager Abbey — GBP foundation. No live publishing. */
const SCM = Object.freeze({
  id: '1XQbRXDfesWPEKY_EmD8YXj0M5CVtgP-C5OH2bAx3ucE',
  articles: 'Article Library',
  images: 'Image Opportunities',
  queue: 'Social Posts Queue',
  log: 'Publication Log',
  settings: 'Settings'
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Social Content Manager')
    .addItem('Check setup', 'scm_checkSetup')
    .addItem('Find article images', 'scm_scanAllArticles')
    .addItem('Create GBP drafts', 'scm_generateGBP')
    .addItem('Prepare selected JPG images', 'scm_prepareSelectedImages')
    .addItem('Preview scheduled GBP posts', 'scm_previewDueGBP')
    .addItem('Log GBP dry run', 'scm_logDryRunGBP')
    .addToUi();
}

function scm_book_() { return SpreadsheetApp.openById(SCM.id); }
function scm_sheet_(name) {
  const sheet = scm_book_().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}
function scm_rows_(name) {
  const s = scm_sheet_(name), data = s.getDataRange().getValues();
  const keys = (data.shift() || []).map(String);
  return {sheet:s, keys:keys, rows:data.map((values,i) => ({
    number:i+2, values:values, item:Object.fromEntries(keys.map((k,j)=>[k,values[j]]))
  }))};
}
function scm_append_(name, data) {
  const s=scm_sheet_(name), keys=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  s.appendRow(keys.map(k=>data[k]===undefined?'':data[k]));
}
function scm_now_() { return new Date().toISOString(); }
function scm_id_(prefix) { return prefix+'-'+Utilities.getUuid().slice(0,12); }
function scm_checkSetup() {
  [SCM.articles,SCM.images,SCM.queue,SCM.log,SCM.settings].forEach(scm_sheet_);
  const setting=scm_rows_(SCM.settings).rows;
  const auto=setting.find(r=>r.item.Setting==='Auto Publish');
  if (!auto || String(auto.item.Value).toUpperCase()!=='NO') {
    throw new Error('Safety stop: Settings > Auto Publish must be NO during setup.');
  }
  SpreadsheetApp.getActive().toast('Core sheets checked; automatic publishing is OFF.');
}

/** Explicit send only; callable from the Content Engine project via the companion script. */
function scm_receiveArticle(article) {
  if (!article || !/^\d+$/.test(String(article.postId||''))) throw new Error('Valid WordPress Post ID required');
  if (!article.html || !String(article.html).trim()) throw new Error('Finished HTML required');
  const id=String(article.postId), existing=scm_rows_(SCM.articles);
  const row=existing.rows.find(r=>String(r.item['Article Post ID'])===id);
  const fields={
    'Article Post ID':id, 'Title':String(article.title||''),
    'Article Type':String(article.type||''), 'Stone Type':String(article.stone||''),
    'Locality':String(article.locality||''), 'Parent Area':String(article.parentArea||''),
    'Article URL':String(article.url||''), 'Finished HTML':String(article.html),
    'Source Version':String(article.version||''), 'Received UTC':scm_now_(),
    'Status':'RECEIVED'
  };
  if (row) {
    const header=existing.keys;
    // Preserve any notes and all already-created image/post decisions.
    header.forEach((key,j)=>{
      if (Object.prototype.hasOwnProperty.call(fields,key))
        existing.sheet.getRange(row.number,j+1).setValue(fields[key]);
    });
    return {postId:id,action:'UPDATED',row:row.number};
  }
  scm_append_(SCM.articles,fields);
  return {postId:id,action:'CREATED'};
}
function scm_imageTags_(html) {
  const matches=String(html||'').match(/<img\b[^>]*>/gi)||[];
  return matches.map(tag=>{
    const get=n=>{
      const re=new RegExp('\\b'+n+'\\s*=\\s*(?:"([^"]*)"|\'([^\']*)\'|([^\\s>]+))','i');
      const m=tag.match(re);return m?(m[1]||m[2]||m[3]||''):'';
    };
    const src=get('src').replace(/&amp;/g,'&');
    return {url:src,alt:get('alt')};
  }).filter(x=>/^https:\/\/www\.abbeyfloorcare\.co\.uk\/wp-content\/uploads\//i.test(x.url));
}
function scm_scanAllArticles() {
  const articles=scm_rows_(SCM.articles).rows;
  const images=scm_rows_(SCM.images).rows;
  const existing=new Set(images.map(r=>String(r.item['Image Fingerprint'])));
  let created=0;
  articles.forEach(row=>{
    const a=row.item, id=String(a['Article Post ID']||'').trim();
    if (!id || !a['Finished HTML'])return;
    const seenWithinArticle=new Set();
    scm_imageTags_(a['Finished HTML']).forEach((image,index)=>{
      const key=id+'|'+image.url;
      if(seenWithinArticle.has(key)||existing.has(key))return;
      seenWithinArticle.add(key);existing.add(key);
      scm_append_(SCM.images,{
        'Opportunity ID':scm_id_('IMG'),'Article Post ID':id,
        'Image Number':index+1,'Original WebP URL':image.url,
        'Original Alt / Caption':image.alt,
        'Evidence / Context':'Check article and photograph before drafting',
        'Distinct Post Angle':'','Educational Value':'',
        'Decision':'REVIEW','Reason / Notes':'Needs editorial selection',
        'Image Fingerprint':key,'Created UTC':scm_now_()
      });
      created++;
    });
  });
  SpreadsheetApp.getActive().toast(created+' new image opportunities added.');
  return {created:created};
}
/** Makes reviewable GBP draft placeholders only; no invented visual claims. */
function scm_generateGBP(){
  scm_checkSetup();
  const imgs=scm_rows_(SCM.images).rows, articles=scm_rows_(SCM.articles).rows,
    queue=scm_rows_(SCM.queue).rows;
  const articleById=new Map(articles.map(r=>[String(r.item['Article Post ID']),r.item]));
  const queued=new Set(queue.filter(r=>r.item.Platform==='GBP').map(r=>String(r.item['Opportunity ID'])));
  let made=0;
  imgs.forEach(row=>{
    const im=row.item;
    if(String(im.Decision).toUpperCase()!=='GENERATE'||queued.has(String(im['Opportunity ID'])))return;
    const a=articleById.get(String(im['Article Post ID']));
    if(!a)return;
    const title=String(a.Title||'').trim();
    const place=String(a.Locality||'').trim();
    const angle=String(im['Distinct Post Angle']||'').trim();
    if(!angle)return; // No unsupported automated assertions.
    scm_append_(SCM.queue,{
      'Social Post ID':scm_id_('GBP'), 'Article Post ID':im['Article Post ID'],
      'Opportunity ID':im['Opportunity ID'],'Platform':'GBP',
      'Draft Text':angle+'\n\nRead the full article: '+title+(place?' ('+place+')':''),
      'Original Image URL':im['Original WebP URL'],
      'Article URL':a['Article URL'],
      'CTA':'Learn more','Approval':'DRAFT',
      'Publication Status':'NOT SCHEDULED','Revision':1,
      'Notes':'Requires human review, confirmed image suitability and prepared JPG before scheduling'
    });
    made++;queued.add(String(im['Opportunity ID']));
  });
  SpreadsheetApp.getActive().toast(made+' GBP drafts created for review.');
  return {created:made};
}
