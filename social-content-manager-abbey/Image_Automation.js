/** Platform-specific image preparation. No live social publishing. */
const SCM_IMAGE_DEFAULTS = [
  ['GBP','YES',1200,900,90,'WIDTH','#FFFFFF'],
  ['FACEBOOK','NO',1200,630,90,'WIDTH','#FFFFFF'],
  ['X','NO',1200,675,90,'WIDTH','#FFFFFF']
];
function scm_imageSettings_() {
  const book=scm_book_();
  let sheet=book.getSheetByName('Image Settings');
  if(!sheet) {
    sheet=book.insertSheet('Image Settings');
    sheet.appendRow(['Platform','Enabled','Width','Height','Quality','Fit','Background']);
    SCM_IMAGE_DEFAULTS.forEach(row=>sheet.appendRow(row));
  }
  // One-time migration of the previously installed default CONTAIN layout.
  // Thereafter the Image Settings tab remains fully user-editable.
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('SCM_FULL_WIDTH_SETTINGS_MIGRATED')!=='YES') {
    const range=sheet.getDataRange();
    const oldValues=range.getValues();
    const headers=(oldValues[0]||[]).map(String);
    const fitIndex=headers.indexOf('Fit');
    const platformIndex=headers.indexOf('Platform');
    if(fitIndex>=0&&platformIndex>=0) {
      for(let i=1;i<oldValues.length;i++) {
        if(['GBP','FACEBOOK','X'].includes(String(oldValues[i][platformIndex]).toUpperCase()) &&
           String(oldValues[i][fitIndex]).toUpperCase()==='CONTAIN')
          sheet.getRange(i+1,fitIndex+1).setValue('WIDTH');
      }
    }
    props.setProperty('SCM_FULL_WIDTH_SETTINGS_MIGRATED','YES');
  }
  const data=sheet.getDataRange().getValues();
  const keys=data.shift().map(String);
  const defaults=new Map(SCM_IMAGE_DEFAULTS.map(r=>[r[0],r]));
  const found=new Map();
  data.forEach(values=>{
    const obj=Object.fromEntries(keys.map((k,i)=>[k,values[i]]));
    const name=String(obj.Platform||'').toUpperCase().trim();
    if(defaults.has(name))found.set(name,obj);
  });
  const settings={};
  for(const [platform,row] of defaults) {
    const x=found.get(platform)||{};
    const width=Number(x.Width||row[2]),height=Number(x.Height||row[3]);
    const quality=Number(x.Quality||row[4]);
    const background=String(x.Background||row[6]).toUpperCase();
    const fit=String(x.Fit||row[5]).toUpperCase();
    if(!Number.isInteger(width)||width<400||width>1600 ||
       !Number.isInteger(height)||height<400||height>1600 ||
       !Number.isInteger(quality)||quality<75||quality>95 ||
       !['WIDTH','CONTAIN','COVER'].includes(fit)||!/^#[0-9A-F]{6}$/.test(background))
      throw new Error('Invalid Image Settings for '+platform);
    settings[platform]={enabled:String(x.Enabled||row[1]).toUpperCase()==='YES',
      width,height,quality,fit,background};
  }
  return settings;
}
function scm_autoPrepareQueuedImages() {
  scm_checkSetup();
  const settings=scm_imageSettings_();
  const props=PropertiesService.getScriptProperties();
  const queue=scm_rows_(SCM.queue);
  const jobs=scm_rows_('Image Processing');
  for(const key of ['Platform','Settings Fingerprint']) {
    if(jobs.keys.includes(key))continue;
    const col=jobs.sheet.getLastColumn()+1;
    jobs.sheet.getRange(1,col).setValue(key);
    jobs.keys.push(key);
  }
  const priorJobs=new Map(jobs.rows.map(r=>[
    String(r.item['Opportunity ID'])+'|'+String(r.item.Platform||'GBP'),r
  ]));
  const put=(row,key,value)=>{
    const col=queue.keys.indexOf(key)+1;
    if(!col)throw new Error('Missing Social Posts Queue column '+key);
    queue.sheet.getRange(row.number,col).setValue(value);
    row.item[key]=value;
  };
  let prepared=0,ready=0,held=0,disabled=0;
  const groups=new Map();
  for(const row of queue.rows) {
    const p=row.item, platform=String(p.Platform||'').toUpperCase();
    if(!settings[platform]||!settings[platform].enabled){disabled++;continue;}
    if(['PUBLISHED','PUBLISHING'].includes(String(p['Publication Status'])) ||
       ['PAUSED','REJECTED'].includes(String(p.Approval)))continue;
    const source=String(p['Original Image URL']||'').trim();
    const id=String(p['Opportunity ID']||'').trim();
    if(!id || !/^https:\/\/www\.abbeyfloorcare\.co\.uk\/wp-content\/uploads\//i.test(source)) {
      held++;continue;
    }
    const opts=settings[platform];
    const fingerprint=[platform,opts.width,opts.height,opts.quality,opts.fit,opts.background,source].join('|');
    const previous=priorJobs.get(id+'|'+platform);
    const validPrepared=previous&&String(previous.item['Processing Status'])==='READY'&&
      String(previous.item['Settings Fingerprint'])===fingerprint&&
      String(previous.item['Prepared JPG URL']||'')===String(p['Prepared JPG URL']||'')&&
      /\/abbey-social-temp\/social-[a-f0-9]{64}\.jpg$/i.test(String(p['Prepared JPG URL']||''));
    if(validPrepared) {ready++;continue;}
    const groupKey=id+'|'+platform;
    if(!groups.has(groupKey))groups.set(groupKey,{source,id,platform,opts,fingerprint,previous,rows:[]});
    groups.get(groupKey).rows.push(row);
  }
  if(!groups.size)return {status:'UP TO DATE',prepared,ready,held,disabled};
  if(!props.getProperty('SCM_WP_USER')||!props.getProperty('SCM_WP_APP_PASSWORD'))
    return {status:'WORDPRESS CONNECTION MISSING',prepared,ready,held,disabled,pending:groups.size};
  const started=Date.now();
  for(const [key,g] of groups) {
    if(prepared>=8||Date.now()-started>180000)break;
    const holdKey='SCM_IMAGE_HOLD_'+Utilities.base64EncodeWebSafe(key).slice(0,170);
    const oldHold=Number(props.getProperty(holdKey)||0);
    if(oldHold&&Date.now()-oldHold<24*3600000){held++;continue;}
    const job={'Image Job ID':g.previous?g.previous.item['Image Job ID']:scm_id_('JOB'),
      'Opportunity ID':g.id,'Platform':g.platform,'Original URL':g.source,
      'Settings Fingerprint':g.fingerprint,'Requested UTC':scm_now_()};
    try {
      const result=scm_wpRequest_('prepare',{
        url:g.source,max_width:g.opts.width,max_height:g.opts.height,
        quality:g.opts.quality,fit:g.opts.fit,background:g.opts.background
      });
      const url=String(result.jpg_url||'');
      if(!/^https:\/\/www\.abbeyfloorcare\.co\.uk\/wp-content\/uploads\/abbey-social-temp\/social-[a-f0-9]{64}\.jpg$/i.test(url) ||
         !/^[a-f0-9]{64}$/.test(String(result.token||'')) ||
         Number(result.width)!==g.opts.width || Number(result.height)<g.opts.height || Number(result.height)>1600)
        throw new Error('Image service did not produce the requested platform dimensions. Update WordPress plugin.');
      Object.assign(job,{'Prepared JPG URL':url,'Processing Status':'READY',
        'Completed UTC':scm_now_(),'Last Error':'',
        'Reference Count':g.rows.length,
        'Conversion Details':'WordPress token: '+result.token+'; '+result.width+'x'+result.height});
      if(g.previous)scm_updateRow_(jobs,g.previous.number,job);
      else scm_append_('Image Processing',job);
      g.rows.forEach(row=>put(row,'Prepared JPG URL',url));
      props.deleteProperty(holdKey);
      prepared++;
    } catch(e) {
      Object.assign(job,{'Processing Status':'FAILED','Last Error':String(e.message||e).slice(0,250)});
      if(g.previous)scm_updateRow_(jobs,g.previous.number,job);
      else scm_append_('Image Processing',job);
      props.setProperty(holdKey,String(Date.now()));
      held++;
    }
  }
  return {status:'NO LIVE PUBLISHING',prepared,ready,held,disabled,pending:Math.max(0,groups.size-prepared-held)};
}
