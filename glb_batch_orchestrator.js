/**
 * Independent Post ID -> full Content Intelligence -> single-pass Sol batch.
 * Leaves original Content Intelligence and final publishing modals unchanged.
 * Queue and two completion timestamps are saved to Sheets.
 */
var GLB_SHEET='Google Love Batch Queue';
var GLB_CI_HEADER='Content Intelligence Completed — Date';
var GLB_SOL_HEADER='Google Love Completed — Date';
var GLB_QUEUE_HEADERS=['Post ID','Posts row','Title','Article Type','Status','CI completed UTC','Sol completed UTC','CI cost USD','Sol cost USD','Message','Updated UTC'];
function glb_utc_(){return Utilities.formatDate(new Date(),'Etc/UTC',"yyyy-MM-dd'T'HH:mm:ss'Z'");}
function glb_posts_(){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');if(!sh)throw Error('posts worksheet missing');return sh;}
function glb_col_(sh,label,create){
 var h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],i=h.indexOf(label);
 if(i>=0)return i+1;
 if(!create)throw Error('Column missing: '+label);
 var next=sh.getLastColumn()+1;
 sh.getRange(1,next).setValue(label);
 sh.setColumnWidth(next,185);
 return next;
}
function glb_setup_(){
 // Keep sidebar polling read-only; initialise posts tracking headings only in glb_prepare.

 var ss=SpreadsheetApp.getActiveSpreadsheet(),s=ss.getSheetByName(GLB_SHEET);
 if(!s){s=ss.insertSheet(GLB_SHEET);}
 if(s.getLastRow()<1){s.getRange(1,1,1,GLB_QUEUE_HEADERS.length).setValues([GLB_QUEUE_HEADERS]);s.setFrozenRows(1);}
 var h=s.getRange(1,1,1,GLB_QUEUE_HEADERS.length).getDisplayValues()[0];
 GLB_QUEUE_HEADERS.forEach(function(x,i){if(h[i]!==x)throw Error('Unexpected batch queue heading '+(i+1));});
 return s;
}
function glb_lookup_(postId){
 var sh=glb_posts_(),id=String(postId||'').trim();
 var ids=sh.getRange(2,4,sh.getLastRow()-1,1).getDisplayValues(),match=[];
 ids.forEach(function(v,i){if(String(v[0]).trim()===id)match.push(i+2);});
 if(match.length!==1)throw Error('Expected exactly one posts row for '+id+'; found '+match.length);
 var row=match[0];
 return {id:id,row:row,title:String(sh.getRange(row,2).getDisplayValue()),material:String(sh.getRange(row,7).getDisplayValue()),type:String(sh.getRange(row,8).getDisplayValue())};
}
function glb_findQueueRow_(sh,id){
 var n=sh.getLastRow();if(n<2)return -1;
 var values=sh.getRange(2,1,n-1,1).getDisplayValues();
 var found=-1;values.forEach(function(v,i){if(v[0]===id){if(found>=0)throw Error('Duplicate queue Post ID '+id);found=i+2;}});
 return found;
}
function glb_progress(){
 var s=glb_setup_(),n=s.getLastRow();
 if(n<2)return [];
 var raw=s.getRange(2,1,n-1,11).getDisplayValues();
 return raw.filter(function(r){return r[0];}).map(function(r){return {postId:r[0],postsRow:r[1],title:r[2],type:r[3],status:r[4],ciDate:r[5],solDate:r[6],ciCost:r[7],solCost:r[8],message:r[9],updated:r[10]};});
}
function glb_prepare(postIds){
 var text=String(postIds||''),ids=text.split(/[\s,;]+/).map(function(x){return x.trim();}).filter(Boolean);
 if(!ids.length)throw Error('Enter at least one Post ID');
 if(ids.length>50)throw Error('Maximum 50 Post IDs per batch');
 var seen={};ids=ids.filter(function(id){if(!/^\d+$/.test(id))throw Error('Invalid Post ID '+id);if(seen[id])return false;seen[id]=true;return true;});
 var posts=glb_posts_();
 var data=posts.getRange(2,2,Math.max(1,posts.getLastRow()-1),7).getDisplayValues();
 var byId={};data.forEach(function(r,i){var id=String(r[2]||'').trim();if(!id)return;if(byId[id])byId[id].duplicate=true;else byId[id]={id:id,row:i+2,title:r[0],material:r[5],type:r[6]};});
 var objects=ids.map(function(id){var x=byId[id];if(!x||x.duplicate)throw Error('Post ID missing or duplicated in posts: '+id);return x;});
 glb_col_(posts,GLB_CI_HEADER,true);
 glb_col_(posts,GLB_SOL_HEADER,true);
 var q=glb_setup_();
 var lock=LockService.getScriptLock();if(!lock.tryLock(15000))throw Error('Batch queue busy');
 try{
  var last=q.getLastRow(),existing=last>1?q.getRange(2,1,last-1,1).getDisplayValues():[];
  var present={};existing.forEach(function(v){if(v[0])present[String(v[0])]=true;});
  var added=objects.filter(function(x){return !present[x.id];}).map(function(x){return [x.id,x.row,x.title,x.type,'QUEUED','','',0,0,'',glb_utc_()];});
  if(added.length)q.getRange(last+1,1,added.length,11).setValues(added);
 }finally{lock.releaseLock();}
 return {ids:ids,rows:glb_progress().filter(function(x){return seen[x.postId];})};
}
function glb_queueState_(id){
 var q=glb_setup_(),r=glb_findQueueRow_(q,String(id));
 if(r<2)throw Error('Post ID not queued: '+id);
 return {sh:q,row:r,values:q.getRange(r,1,1,11).getValues()[0]};
}
function glb_stamp_(postsRow,header){
 var sh=glb_posts_(),c=glb_col_(sh,header,false);
 var d=new Date();
 sh.getRange(postsRow,c).setValue(d).setNumberFormat('dd/MM/yyyy HH:mm');
 return glb_utc_();
}
function glb_mark_(id,status,message,cost,phase){
 var x=glb_queueState_(id),r=x.row,s=x.sh;
 s.getRange(r,5).setValue(status);
 if(phase==='CI'&&cost!==undefined)s.getRange(r,8).setValue(cost);
 if(phase==='SOL'&&cost!==undefined)s.getRange(r,9).setValue(cost);
 s.getRange(r,10).setValue(String(message||'').slice(0,3000));
 s.getRange(r,11).setValue(glb_utc_());
}
function glb_runCI(postId){
 var id=String(postId||''),x=glb_queueState_(id),state=String(x.values[4]||'');
 if(state==='CI DONE'||state==='SOL DONE'||state==='SOL REVIEW')return {success:true,skipped:true};
 if(state!=='QUEUED')throw Error('CI cannot be automatically retried from '+state+'; inspect the row first.');
 var a=glb_lookup_(id);
 if(Number(x.values[1])!==a.row)throw Error('Posts row changed for '+id);
 glb_mark_(id,'CI RUNNING','Running the existing Content Intelligence chain.',undefined);
 try{
  var result=runMasterContentIntelligenceBatchRow(a.row,a.material);
  if(!result||!result.success)throw Error((result&&result.message)||'Content Intelligence failed');
  // This existing chain already invokes bc_caseEvidenceExtract_ for case studies,
  // including the latest source-HTML media enrichment on unchanged evidence.
  if(a.type==='Case Study'){
   var post=glb_posts_(),h=post.getRange(1,1,1,post.getLastColumn()).getDisplayValues()[0];
   var idx=h.indexOf('Evidence Status');
   if(idx<0||String(post.getRange(a.row,idx+1).getDisplayValue()).indexOf('SOURCE-QUOTED | sha256:')!==0)
    throw Error('Case-study extraction did not leave valid source-quoted evidence');
  }
  var date=glb_stamp_(a.row,GLB_CI_HEADER);
  x.sh.getRange(x.row,6).setValue(date);
  glb_mark_(id,'CI DONE','Full Content Intelligence successful, including case-study media extraction where applicable.',Number(result.cost)||0,'CI');
  return {success:true,cost:Number(result.cost)||0};
 }catch(err){
  glb_mark_(id,'CI REVIEW',String(err.message||err),undefined);
  return {success:false,message:String(err.message||err)};
 }
}
/** Explicit operator reset after correcting CI input. Never resets a Sol attempt. */
function glb_retryCI(postId){
 var id=String(postId||''),q=glb_queueState_(id),state=String(q.values[4]||'');
 if(state!=='CI REVIEW')throw Error('Only CI REVIEW can be retried; current state '+state);
 var current=glb_lookup_(id);
 if(current.row!==Number(q.values[1]))throw Error('Post row has moved; cannot reset');
 if(q.values[6])throw Error('Sol completion already recorded; cannot reset CI');
 q.sh.getRange(q.row,5).setValue('QUEUED');
 q.sh.getRange(q.row,10).setValue('Operator requested a new full CI attempt after correcting source data. Earlier CI attempts may have incurred costs.');
 q.sh.getRange(q.row,11).setValue(glb_utc_());
 return {success:true,message:'Post '+id+' queued for another full Content Intelligence run.'};
}
/** Run exactly the requested Post ID: full existing CI then one Sol call.
 * No queue state gate; the sidebar controls sequence. A failed stage returns
 * an error and the sidebar moves to the next explicitly supplied ID.
 */
function glb_runRequestedPost(postId){
 var id=String(postId||'').trim(),a=glb_lookup_(id);
 var q=glb_setup_(),qr=glb_findQueueRow_(q,id);
 if(qr<2){q.appendRow([id,a.row,a.title,a.type,'QUEUED','','',0,0,'',glb_utc_()]);}
 else{
  // A new user-initiated batch intentionally resets the listed item.
  q.getRange(qr,2,1,10).setValues([[a.row,a.title,a.type,'QUEUED','','',0,0,'Requested again by operator',glb_utc_()]]);
 }
 return {success:true,postId:id};
}
function glb_runRequestedCI(postId){
 var id=String(postId||''),a=glb_lookup_(id),q=glb_queueState_(id);
 glb_mark_(id,'CI RUNNING','Content Intelligence started for requested Post ID.');
 try{
  var result=runMasterContentIntelligenceBatchRow(a.row,a.material);
  if(!result||!result.success)throw Error(result&&result.message||'Content Intelligence failed');
  var date=glb_stamp_(a.row,GLB_CI_HEADER);
  q.sh.getRange(q.row,6).setValue(date);
  glb_mark_(id,'CI DONE','Full Content Intelligence finished.',Number(result.cost)||0,'CI');
  return {success:true,postId:id,cost:Number(result.cost)||0};
 }catch(e){glb_mark_(id,'CI FAILED',String(e.message||e));return {success:false,postId:id,message:String(e.message||e)};}
}
function glb_runRequestedSol(postId){
 // Use existing tested Sol logic, which requires the preceding CI DONE mark.
 return glb_runSol(postId);
}
function glb_buildSolPrompt_(a){
 if(typeof glc_getArticle_!=='function')throw Error('Missing Google Love article reader');
 var source=glc_getArticle_(a.id);
 if(!String(source.technical||'').trim()||String(source.technicalSource||'').indexOf('technical!')!==0)
  throw Error('Approved Technical Matrix N:S missing for '+a.id);
 var rule=glc_moistureRule_();
 var common=[
  'Write one outstanding standalone UK homeowner-facing floor restoration article. Return clean HTML only.',
  'Use British English, direct helpful explanations, evidence-led practical expertise, excellent structure and no empty padding, town history or excessive SEO repetition.',
  'Use the approved Technical Matrix to explain materials, limitations and safe practical methods; never invent observed conditions, completed work, clients, results or testing.',
  'When urethane or polyurethane means a sealer, use film-forming sealer in homeowner language. Retain evidence-backed topical and impregnating distinctions.',
  'Do not use or imply the prohibited UK property-defect moisture diagnosis.',
  'Avoid invented photos, captions, product claims, business guarantees and links.',
  'POST ID '+a.id,'TITLE '+source.title,'TYPE '+source.type,'STONE '+source.material,
  '=== MATERIAL RULE ===',rule,'=== APPROVED TECHNICAL MATRIX N:S ===',source.technical
 ];
 if(a.type!=='Case Study'){
  if(!String(source.original||'').trim())throw Error('Original source content absent for non-case article '+a.id);
  common.push('This is an article, not a case study. The historical text is working reference material, not proof of on-site work. Improve it as a complete standalone article; retain genuinely useful explanations, existing relevant images and source links without inventing facts.');
  common.push('=== ORIGINAL ARTICLE SOURCE ===',source.original);
 }else{
  var sh=glb_posts_(),v=sh.getRange(a.row,194,1,16).getDisplayValues()[0];
  var fields=typeof GLSE_FIELDS!=='undefined'?GLSE_FIELDS:null;
  if(!fields||fields.length!==16)throw Error('Structured evidence schema missing');
  // The unchanged Content Intelligence extractor writes the first ten
  // source-quoted categories, media GY and Evidence Status HA. Remaining
  // auxiliary evidence fields are optional, not mandatory for new posts.
  var required=fields.slice(0,10).concat([fields[13],fields[15]]);
  var missing=required.filter(function(h){return !String(v[fields.indexOf(h)]||'').trim();});
  if(missing.length)throw Error('Required extracted evidence incomplete: '+missing.join(', '));
  var media=JSON.parse(String(v[13]||'{}'));
  if(!Array.isArray(media.imagesAndVideos))throw Error('Case-study media inventory invalid');
  var images=media.imagesAndVideos.filter(function(x){return x&&x.type==='image'&&x.url;});
  // Reuse the tested single-pass media-aware case-study instructions.
  common=[
   'Create one outstanding UK homeowner-facing restoration case study, in clean article HTML only.',
   'FACTUAL AUTHORITY: The supplied structured evidence is the sole project-specific factual record. Never use general technical guidance to invent work, observations, results or client quotations.',
   'Clearly distinguish completed work, proposed work, recommendations, uncertainty and information not recorded. Never convert a recommended treatment into work performed.',
   'Give the actual project a compelling, chronological narrative. Explain why decisions mattered. Retain concrete facts: abrasive grades, products, equipment, extraction, timings, counts of coatings, client requirements, what failed and genuine outcomes wherever documented.',
   'Use technical matrix for material behaviour, safe advice and corrections; do not override recorded history or present outdated methods as universal recommendations.',
   'Evidence Package and Evidence Relationships support factual fidelity when present; Evidence Coverage, Gaps and Status restrict certainty. Do not print internal metadata or evidence schema.',
   'MEDIA PLACEMENT: The separate image inventory below contains actual source HTML URLs and description metadata, NOT independent visual verification. Never infer image contents from a URL alone.',
   'Place every relevant, distinct project image near the paragraph or section supported by its original alt/caption, source section, and nearby text. Use original source order if equally relevant. Use the exact image URL, unchanged, as an HTML img src.',
   'Do not invent captions, image descriptions, image URLs, or work-to-image relationships. If stage or description is uncertain, use a neutral caption, or omit the caption; do not assert that an image proves a specific method or outcome.',
   'Photographs of unrelated branding, authors, repeated views or illustrations may be omitted if they do not serve the homeowner; do not replace them with invented photos. Include relevant source videos only if supported and safe to embed; never invent embed URLs.',
   'Never silently drop a useful before/during/after photograph for brevity. Place captions conservatively, based on supplied source wording, and do not imply a photo was independently inspected.',
   'Direct answer, British English, practical and readable, no fluff, padding, generic town history, SEO stuffing or lengthy disclaimers.',
   'When urethane/polyurethane means a type of sealer in homeowner text, use film-forming sealer. Preserve distinct topical and impregnating terminology, sealer coat counts, and actual finishes.',
   'Never mention or imply the prohibited UK building defect moisture diagnosis. Describe actual moisture and salt movement only when evidence supports it.',
   'Return only full HTML.',
   'POST ID '+a.id,'TITLE '+source.title,'MATERIAL '+source.material,'TYPE '+source.type,
   '=== GLOBAL TECHNICAL SAFEGUARD ===',rule,
   '=== APPROVED TECHNICAL MATRIX N:S ===',source.technical,
   '=== STRUCTURED PROJECT EVIDENCE ===',fields.filter(function(x,i){return i!==13&&String(v[i]||'').trim();}).map(function(x){return '=== '+x+' ===\n'+v[fields.indexOf(x)];}).join('\n\n'),
   '=== SOURCE IMAGE INVENTORY (URLS + CONTEXT, NOT VISUALLY VERIFIED) ===',
   images.map(function(m,i){return JSON.stringify({imageNumber:i+1,url:m.url,alt:m.alt||'',originalCaption:m.originalCaption||'',sourceSection:m.sourceSection||'',nearbySourceText:m.nearbySourceText||'',nearbyTextBasis:m.nearbyTextBasis||'',suggestedStage:m.suggestedStage||'uncertain',placementConfidence:m.placementConfidence||'unconfirmed',sourceOrder:m.order||i+1});}).join('\n')||'(no indexed project images)',
   '=== SOURCE MEDIA COMPLETENESS ===',JSON.stringify(media.mediaAudit||{}),
   '=== UNPOPULATED FIELDS; DO NOT GUESS ===',fields.filter(function(x,i){return !String(v[i]||'').trim();}).join('; ')||'(none)'
  ];
 }
 var prompt=common.join('\n\n');
 if(prompt.length>170000)throw Error('Prompt too large; stopped without truncating');
 return {prompt:prompt,media:a.type==='Case Study'?JSON.parse(String(glb_posts_().getRange(a.row,207).getDisplayValue()||'{}')):null};
}
function glb_runSol(postId){
 var id=String(postId||''),q=glb_queueState_(id),state=String(q.values[4]||'');
 if(state==='SOL DONE'||state==='SOL REVIEW')return {success:true,skipped:true,status:state};
 if(state!=='CI DONE')throw Error('Sol requires successful CI; current status '+state);
 var a=glb_lookup_(id);
 if(Number(q.values[1])!==a.row)throw Error('Posts row changed for '+id);
 var built=glb_buildSolPrompt_(a); // preflight before creating an uncertain paid request
 glb_mark_(id,'SOL RUNNING','Single GPT-6.1 Sol request reserved; do not automatically retry if interrupted.',undefined);
 var res;
 try{
  res=bc_sendPromptViaOpenAI(built.prompt,16000,'gpt-6.1-sol');
  if(!res||!res.success||res.truncated||!String(res.text||'').trim())throw Error('Sol returned no complete article');
  var html=String(res.text).replace(/^\s*```(?:html)?\s*/i,'').replace(/\s*```\s*$/,'').trim();
  if(html.length>48000||!/<h[12]\b/i.test(html)||!/<p\b/i.test(html))throw Error('Sol returned invalid or oversized HTML');
  if(/\brising[\s-]+damp\b/i.test(html.replace(/<[^>]*>/g,' ')))throw Error('Prohibited moisture wording detected');
  var audit=a.type==='Case Study'?glse_mediaCheck_(html,{images:built.media.imagesAndVideos.filter(function(x){return x&&x.type==='image'&&x.url;})}):null;
  var issues=audit&&(audit.missingProjectEvidence.length||audit.unexpected.length);
  // Save a recoverable draft before replacing the production New HTML cell.
  var ss=SpreadsheetApp.getActiveSpreadsheet(),out=ss.getSheetByName('Google Love Production Drafts');
  if(!out){out=ss.insertSheet('Google Love Production Drafts');out.appendRow(['Post ID','Posts row','UTC','HTML','Model','Cost USD','Media audit JSON','Prior New HTML']);out.setFrozenRows(1);}
  var target=glb_posts_(),col=glb_col_(target,'New HTML',false);
  var prior=String(target.getRange(a.row,col).getValue()||'');
  out.appendRow([id,a.row,glb_utc_(),html,'gpt-6.1-sol',Number(res.cost)||0,audit?JSON.stringify(audit):'',prior]);
  target.getRange(a.row,col).setValue(html);
  var date=glb_stamp_(a.row,GLB_SOL_HEADER);
  q.sh.getRange(q.row,7).setValue(date);
  glb_mark_(id,issues?'SOL REVIEW':'SOL DONE',issues?'Article generated; media retention warning recorded in production drafts.':'Article generated and saved to posts New HTML.',Number(res.cost)||0,'SOL');
  return {success:true,status:issues?'SOL REVIEW':'SOL DONE',cost:Number(res.cost)||0};
 }catch(err){
  glb_mark_(id,'SOL REVIEW',String(err.message||err)+' — do not auto-repeat paid request.',Number(res&&res.cost)||0,'SOL');
  return {success:false,message:String(err.message||err)};
 }
}
function glb_openSidebar(){
 glb_setup_();
 SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('glb_BatchSidebar').setTitle('Content Intelligence → Google Love Sol'));
}

/** Direct operator-selected batch: no queue setup, retry decisions or state gating.
 * Processes each supplied ID once per click. Existing stages are not changed.
 */
function glb_runDirectPost_(postId,phase){
 var a=glb_lookup_(String(postId));
 if(phase==='CI'){
  glb_mark_(a.id,'CI RUNNING','Content Intelligence running');
  var ci=runMasterContentIntelligenceBatchRow(a.row,a.material);
  if(!ci||!ci.success){glb_mark_(a.id,'CI REVIEW',String(ci&&ci.message||'CI failed'));return {success:false,message:String(ci&&ci.message||'CI failed')};}
  var sh=glb_posts_();
  if(a.type==='Case Study'){
   var h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],ix=h.indexOf('Evidence Status');
   if(ix<0||String(sh.getRange(a.row,ix+1).getDisplayValue()).indexOf('SOURCE-QUOTED | sha256:')!==0)
    return {success:false,message:'Case-study evidence extraction was not completed'};
  }
  var stamp=glb_stamp_(a.row,GLB_CI_HEADER),q=glb_queueState_(a.id);
  q.sh.getRange(q.row,6).setValue(stamp);
  glb_mark_(a.id,'CI DONE','Content Intelligence completed',Number(ci.cost||0),'CI');
  return {success:true,stage:'CI DONE'};
 }
 if(phase==='SOL'){
  // Allow user-directed execution after CI completes, irrespective of older queue status.
  var state=glb_queueState_(a.id),current=String(state.values[4]||'');
  if(current!=='CI DONE')return {success:false,message:'Content Intelligence has not completed for this run'};
  return glb_runSol(a.id);
 }
 return {success:false,message:'Unknown phase'};
}
