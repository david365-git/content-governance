/** Social Content Manager: opt-in temporary JPG preparation.
 * Requires WordPress plugin and Script Properties SCM_WP_USER and SCM_WP_APP_PASSWORD.
 * No third-party publishing occurs here.
 */
const SCM_IMG = Object.freeze({
  endpoint:'https://www.abbeyfloorcare.co.uk/wp-json/abbey-social/v1',
  sheet:'Image Processing'
});
function scm_wpRequest_(action, data) {
  const props=PropertiesService.getScriptProperties();
  const user=props.getProperty('SCM_WP_USER');
  const password=props.getProperty('SCM_WP_APP_PASSWORD');
  if(!user || !password)throw new Error('WordPress credentials not configured. Add SCM_WP_USER and SCM_WP_APP_PASSWORD to Script Properties.');
  if(action!=='prepare' && action!=='release')throw new Error('Invalid processing operation');
  const response=UrlFetchApp.fetch(SCM_IMG.endpoint+'/'+action,{
    method:'post', contentType:'application/json', payload:JSON.stringify(data),
    headers:{Authorization:'Basic '+Utilities.base64Encode(user+':'+password)},
    muteHttpExceptions:true, followRedirects:false
  });
  const status=response.getResponseCode();
  let body={};try{body=JSON.parse(response.getContentText());}catch(e){}
  if(status!==200)throw new Error('WordPress image service returned HTTP '+status+' ('+String(body.code||'request_failed')+')');
  return body;
}
function scm_prepareSelectedImages() {
  scm_checkSetup();
  const opportunities=scm_rows_(SCM.images).rows;
  const jobs=scm_rows_(SCM_IMG.sheet).rows;
  const byOpportunity=new Map();
  jobs.forEach(r=>byOpportunity.set(String(r.item['Opportunity ID']),r));
  const queue=scm_rows_(SCM.queue);
  let ready=0,failed=0,skipped=0;
  opportunities.forEach(row=>{
    if(ready+failed>=10)return; // Conservative limit per manual execution.
    const o=row.item;
    if(String(o.Decision).toUpperCase()!=='GENERATE')return;
    const id=String(o['Opportunity ID']||'');
    const old=byOpportunity.get(id);
    if(old && String(old.item['Processing Status'])==='READY'){skipped++;return;}
    const src=String(o['Original WebP URL']||'');
    if(!/^https:\/\/www\.abbeyfloorcare\.co\.uk\/wp-content\/uploads\//.test(src)){failed++;return;}
    const started=scm_now_();
    try {
      const result=scm_wpRequest_('prepare',{url:src,max_width:1200,max_height:1200,quality:90});
      if(!/^https:\/\//.test(String(result.jpg_url||'')) || !/^[a-f0-9]{64}$/.test(String(result.token||'')))
        throw new Error('Invalid image processor response');
      const job={
        'Image Job ID':old?old.item['Image Job ID']:scm_id_('JOB'),
        'Opportunity ID':id,'Original URL':src,
        'Prepared JPG URL':result.jpg_url,'Processing Status':'READY',
        'Requested UTC':started,'Completed UTC':scm_now_(),
        'Reference Count':queue.rows.filter(r=>String(r.item['Opportunity ID'])===id &&
          String(r.item['Publication Status'])!=='PUBLISHED').length,
        'Last Error':'','Conversion Details':'WordPress token: '+result.token+'; '+result.width+'×'+result.height
      };
      if(old)scm_updateRow_(jobs,old.number,job);else scm_append_(SCM_IMG.sheet,job);
      queue.rows.filter(r=>String(r.item['Opportunity ID'])===id).forEach(r=>{
        const col=queue.keys.indexOf('Prepared JPG URL')+1;
        if(col>0)queue.sheet.getRange(r.number,col).setValue(result.jpg_url);
      });
      ready++;
    } catch(error) {
      const data={
        'Image Job ID':old?old.item['Image Job ID']:scm_id_('JOB'),
        'Opportunity ID':id,'Original URL':src,'Processing Status':'FAILED',
        'Requested UTC':started,'Last Error':String(error.message||error).slice(0,250)
      };
      if(old)scm_updateRow_(jobs,old.number,data);else scm_append_(SCM_IMG.sheet,data);
      failed++;
    }
  });
  SpreadsheetApp.getActive().toast('JPG preparation: '+ready+' ready, '+failed+' failed, '+skipped+' already ready.');
  return {ready:ready,failed:failed,skipped:skipped};
}
function scm_updateRow_(sheetData,rowNumber,fields) {
  sheetData.keys.forEach((key,i)=>{
    if(Object.prototype.hasOwnProperty.call(fields,key))
      sheetData.sheet.getRange(rowNumber,i+1).setValue(fields[key]);
  });
}
/** Release only an explicitly requested job. Call after verified platform ingestion,
 * never immediately after generating or scheduling a post.
 */
function scm_releasePreparedJob(jobId) {
  if(!jobId)throw new Error('Image Job ID required');
  const jobs=scm_rows_(SCM_IMG.sheet);
  const row=jobs.rows.find(r=>String(r.item['Image Job ID'])===String(jobId));
  if(!row || row.item['Processing Status']!=='READY')throw new Error('READY image job not found');
  const linked=scm_rows_(SCM.queue).rows.filter(r=>String(r.item['Opportunity ID'])===String(row.item['Opportunity ID']));
  if(!linked.length || linked.some(r=>String(r.item['Publication Status'])!=='PUBLISHED'))
    throw new Error('Cannot release: at least one linked social post is unpublished or missing');
  const tokenMatch=String(row.item['Conversion Details']||'').match(/WordPress token: ([a-f0-9]{64})/);
  if(!tokenMatch)throw new Error('WordPress image token missing');
  scm_wpRequest_('release',{token:tokenMatch[1]});
  scm_updateRow_(jobs,row.number,{'Processing Status':'EXPIRED','Expires UTC':new Date(Date.now()+7*86400000).toISOString()});
  return {released:true,jobId:String(jobId)};
}
