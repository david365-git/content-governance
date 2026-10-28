/**
 * Google Love structured-evidence comparison — standalone isolated test.
 * One Sol generation per eligible case study. Does not read original article body.
 * Writes only GL Structured Evidence Test; existing drafts and posts untouched.
 */
var GLSE_IDS=['48109','7952','5774','21838','21746','8027','8336','12107','11452','22942','11945'];
var GLSE_FIELDS=['Project Location & Property','Observed Conditions','Client Requirements & Restoration Objectives','Materials & Construction','Inspection & Tests','Restoration Methods','Products & Equipment','Technical Findings','Project Results','Maintenance Handover','Case Study Evidence Package','Evidence Relationships','Evidence Coverage','Evidence Media & Links','Evidence Gaps & Uncertainties','Evidence Status'];
var GLSE_TAB='GL Structured Evidence Media Test';
function glse_sheet_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),s=ss.getSheetByName(GLSE_TAB);
 if(!s){s=ss.insertSheet(GLSE_TAB);s.appendRow(['Post ID','Status','Posts row','Populated fields','Missing fields','HTML','Model','Cost USD','Updated UTC','Message']);s.setFrozenRows(1);}
 return s;
}
function glse_now_(){return Utilities.formatDate(new Date(),'Etc/UTC',"yyyy-MM-dd'T'HH:mm:ss'Z'");}
function glse_rows_(){
 var s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
 if(!s)throw Error('posts sheet missing');
 var heads=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0];
 GLSE_FIELDS.forEach(function(h,i){if(String(heads[193+i]||'').trim()!==h)throw Error('Evidence column header mismatch at position '+(194+i)+': '+h);});
 var ids=s.getRange(2,4,s.getLastRow()-1,1).getDisplayValues();
 var result={};
 ids.forEach(function(v,i){var id=String(v[0]).trim();if(GLSE_IDS.indexOf(id)>=0){if(result[id])throw Error('Duplicate Post ID '+id);result[id]=i+2;}});
 return result;
}
function glse_evidence_(id){
 var rows=glse_rows_(),row=rows[id];
 if(!row)throw Error('Post ID not found: '+id);
 var s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
 var values=s.getRange(row,194,1,16).getDisplayValues()[0];
 var present=[],missing=[];
 values.forEach(function(v,i){if(String(v||'').trim())present.push({label:GLSE_FIELDS[i],value:String(v)});else missing.push(GLSE_FIELDS[i]);});
 return {postId:id,row:row,present:present,missing:missing};
}
function glse_media_(e){
 var fld=e.present.filter(function(x){return x.label==='Evidence Media & Links';})[0];
 if(!fld)throw Error('Evidence Media & Links missing for '+e.postId);
 var parsed;
 try{parsed=JSON.parse(fld.value);}catch(err){throw Error('Invalid media JSON for '+e.postId+': '+err.message);}
 if(!parsed||!Array.isArray(parsed.imagesAndVideos))throw Error('imagesAndVideos array missing for '+e.postId);
 var items=parsed.imagesAndVideos.filter(function(m){return m&&m.type==='image'&&String(m.url||'').trim();});
 var seen={};
 items=items.filter(function(m){var u=String(m.url).replace(/&amp;/g,'&');if(seen[u])return false;seen[u]=true;return true;});
 return {images:items,links:Array.isArray(parsed.links)?parsed.links:[],audit:parsed.mediaAudit||{}};
}
function glse_outputImages_(html){
 var re=/<img\b[^>]*\bsrc\s*=\s*(["'])(.*?)\1/gi,m,urls=[];
 while((m=re.exec(html)))urls.push(m[2].replace(/&amp;/g,'&'));
 return urls;
}
function glse_mediaCheck_(html,media){
 var expected=media.images.map(function(x){return String(x.url).replace(/&amp;/g,'&');});
 var actual=glse_outputImages_(html);
 var missing=expected.filter(function(u){return actual.indexOf(u)<0;});
 var unexpected=actual.filter(function(u){return expected.indexOf(u)<0;});
 return {expected:expected.length,used:actual.length,missing:missing,unexpected:unexpected};
}
function glse_saved_(id){
 var s=glse_sheet_(),n=s.getLastRow();
 if(n<2)return null;
 var v=s.getRange(2,1,n-1,2).getDisplayValues();
 for(var i=0;i<v.length;i++)if(v[i][0]===id)return {row:i+2,status:v[i][1]};
 return null;
}
function glse_write_(id,status,e,html,cost,msg){
 var s=glse_sheet_(),prev=glse_saved_(id),row=prev?prev.row:s.getLastRow()+1;
 s.getRange(row,1,1,10).setValues([[id,status,e.row,e.present.length,e.missing.join(' | '),String(html||''), 'gpt-6.1-sol',cost||0,glse_now_(),String(msg||'').slice(0,1500)]]);
 return row;
}
function glse_overview(){
 var completed=0,eligible=0,blocked=0,items=[];
 GLSE_IDS.forEach(function(id){
  var e=glse_evidence_(id),v=glse_saved_(id),state=v?v.status:(e.present.length===0?'NO EVIDENCE':e.present.length<16?'PARTIAL EVIDENCE':'READY');
  if(e.present.length)eligible++;else blocked++;
  if(state.indexOf('DONE')===0)completed++;
  items.push({id:id,state:state,populated:e.present.length,missing:e.missing,postsRow:e.row});
 });
 var running=PropertiesService.getScriptProperties().getProperty('GLSE_ACTIVE')==='YES';
 return {items:items,completed:completed,total:items.length,eligible:eligible,blocked:blocked,running:running};
}
function glse_generateOne_(id){
 var e=glse_evidence_(id);
 if(!e.present.length){glse_write_(id,'NO EVIDENCE',e,'',0,'No structured evidence supplied. No API call.');return;}
 var existing=glse_saved_(id);
 if(existing && (existing.status.indexOf('DONE')===0||existing.status==='IN PROGRESS'||existing.status==='STOPPED — REVIEW'))return;
 var a=glc_getArticle_(id); // metadata and Technical Matrix; historical HTML deliberately excluded.
 if(a.type!=='Case Study')throw Error('Not classified as Case Study: '+id);
 var media=glse_media_(e);
 var evidence=e.present.filter(function(x){return x.label!=='Evidence Media & Links';})
  .map(function(x){return '=== '+x.label+' ===\n'+x.value;}).join('\n\n');
 var mediaInput=media.images.map(function(m,i){
  return JSON.stringify({imageNumber:i+1,url:m.url,alt:m.alt||'',originalCaption:m.originalCaption||'',
    sourceSection:m.sourceSection||'',nearbySourceText:m.nearbySourceText||'',
    nearbyTextBasis:m.nearbyTextBasis||'',suggestedStage:m.suggestedStage||'uncertain',
    placementConfidence:m.placementConfidence||'unconfirmed',sourceOrder:m.order||i+1});
 }).join('\n');
 var prompt=[
  'Create one outstanding UK homeowner-facing restoration case study, in clean article HTML only.',
  'FACTUAL AUTHORITY: The supplied structured evidence is the sole project-specific factual record. Never use general technical guidance to invent work, observations, results or client quotations.',
  'Clearly distinguish completed work, proposed work, recommendations, uncertainty and information not recorded. Never convert a recommended treatment into work performed.',
  'Give the actual project a compelling, chronological narrative. Explain why decisions mattered. Retain concrete facts: abrasive grades, products, equipment, extraction, timings, counts of coatings, client requirements, what failed and genuine outcomes wherever documented.',
  'Use technical matrix for material behaviour, safe advice and corrections; do not override recorded history or present outdated methods as universal recommendations.',
  'Evidence Package and Evidence Relationships support factual fidelity; Evidence Coverage, Gaps and Status restrict certainty. Do not print internal metadata or evidence schema.',
  'MEDIA PLACEMENT: The separate image inventory below contains actual source HTML URLs and description metadata, NOT independent visual verification. Never infer image contents from a URL alone.',
  'Place every relevant, distinct project image near the paragraph or section supported by its original alt/caption, source section, and nearby text. Use original source order if equally relevant. Use the exact image URL, unchanged, as an HTML img src.',
  'Do not invent captions, image descriptions, image URLs, or work-to-image relationships. If stage or description is uncertain, use a neutral caption, or omit the caption; do not assert that an image proves a specific method or outcome.',
  'Photographs of unrelated branding, authors, repeated views or illustrations may be omitted if they do not serve the homeowner; do not replace them with invented photos. Include relevant source videos only if supported and safe to embed; never invent embed URLs.',
  'Never silently drop a useful before/during/after photograph for brevity. Place captions conservatively, based on supplied source wording, and do not imply a photo was independently inspected.',
  'Direct answer, British English, practical and readable, no fluff, padding, generic town history, SEO stuffing or lengthy disclaimers.',
  'When urethane/polyurethane means a type of sealer in homeowner text, use film-forming sealer. Preserve distinct topical and impregnating terminology, sealer coat counts, and actual finishes.',
  'Never mention or imply the prohibited UK building defect moisture diagnosis. Describe actual moisture and salt movement only when evidence supports it.',
  'Return only full HTML.',
  'POST ID '+id,'TITLE '+a.title,'MATERIAL '+a.material,'TYPE '+a.type,
  '=== GLOBAL TECHNICAL SAFEGUARD ===',glc_moistureRule_(),
  '=== APPROVED TECHNICAL MATRIX N:S ===',a.technical,
  '=== STRUCTURED PROJECT EVIDENCE ===',evidence,
  '=== SOURCE IMAGE INVENTORY (URLS + CONTEXT, NOT VISUALLY VERIFIED) ===',mediaInput||'(no indexed project images)',
  '=== SOURCE MEDIA COMPLETENESS ===',JSON.stringify(media.audit),
  '=== UNPOPULATED FIELDS; DO NOT GUESS ===',e.missing.join('; ')||'(none)'
 ].join('\n\n');
 // This test does not trim evidence. If the request is too large for the adapter,
 // stop for review rather than silently lose factual content.
 if(prompt.length>170000)throw Error('Structured prompt unusually large: review source instead of truncating.');
 glse_write_(id,'IN PROGRESS',e,'',0,'API request underway. If interrupted, review billing before resetting.');
 var result;
 try{
  result=bc_sendPromptViaOpenAI(prompt,16000,'gpt-6.1-sol');
  if(!result||!result.success||result.truncated||!String(result.text||'').trim())throw Error('API returned no complete article: '+String(result&&result.message||''));
  var html=String(result.text).replace(/^\s*```(?:html)?\s*/i,'').replace(/\s*```\s*$/,'').trim();
  if(html.length>48000||!/<h[12]\b/i.test(html)||!/<p\b/i.test(html))throw Error('Output too long or not valid article HTML.');
  if(/\brising[\s-]+damp\b/i.test(html.replace(/<[^>]*>/g,' ')))throw Error('Generated prohibited moisture terminology; no article saved.');
  var check=glse_mediaCheck_(html,media);
  var review=check.missing.length>0||check.unexpected.length>0;
  var report='Single-pass media-aware draft; '+check.used+'/'+check.expected+' image references. '+
    (check.missing.length?'Omitted source URLs: '+check.missing.join(' | ')+'. ':'')+
    (check.unexpected.length?'Unapproved image URLs: '+check.unexpected.join(' | ')+'. ':'')+
    'Independent editorial review pending.';
  glse_write_(id,review?'DONE — MEDIA REVIEW':'DONE',e,html,Number(result.cost)||0,report);
 }catch(err){
  glse_write_(id,'STOPPED — REVIEW',e,'',Number(result&&result.cost)||0,String(err.message||err));
  throw err;
 }
}
function glse_triggers_(){return ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==='glse_tick';});}
function glse_stop(){
 PropertiesService.getScriptProperties().deleteProperty('GLSE_ACTIVE');
 glse_triggers_().forEach(function(t){ScriptApp.deleteTrigger(t);});
 return 'Structured evidence test stopped.';
}
function glse_start(){
 var ov=glse_overview();
 var s=glse_sheet_();
 if(s.getLastRow()>1){
  var states=s.getRange(2,1,s.getLastRow()-1,2).getDisplayValues();
  var risky=states.filter(function(r){return r[1]==='IN PROGRESS'||r[1]==='STOPPED — REVIEW';});
  if(risky.length)throw Error('Review '+risky[0][0]+' before restarting. No repeat paid call.');
 }
 glse_stop();
 PropertiesService.getScriptProperties().setProperty('GLSE_ACTIVE','YES');
 ScriptApp.newTrigger('glse_tick').timeBased().everyMinutes(1).create();
 return 'Scheduled eligible case studies, skipping missing evidence. Current completed: '+ov.completed+'/'+ov.total;
}
function glse_tick(){
 if(PropertiesService.getScriptProperties().getProperty('GLSE_ACTIVE')!=='YES')return;
 var lock=LockService.getScriptLock();if(!lock.tryLock(1500))return;
 try{
  var next=GLSE_IDS.filter(function(id){var s=glse_saved_(id);return !s;})[0];
  if(!next){glse_stop();return;}
  glse_generateOne_(next);
 }catch(e){glse_stop();console.error('Structured evidence run stopped: '+String(e.message||e));}
 finally{lock.releaseLock();}
}
function glse_openSidebar(){SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('glse_Progress').setTitle('Google Love — Evidence Test'));}
