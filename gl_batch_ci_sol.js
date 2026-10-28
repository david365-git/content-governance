/**
 * Content Intelligence -> single-pass Google Love Sol batch.
 * Existing CI engine, image creation, meta and schema stages are unchanged.
 * The sidebar runs one server call per stage. Interrupted RUNNING entries
 * stop for manual reconciliation rather than risking duplicate paid calls.
 */
var GLB_Q='GL CI + Sol Batch',GLB_ARCHIVE='GL Sol HTML Archive';
var GLB_CI_DATE='Content Intelligence Completed — Date';
var GLB_SOL_DATE='Google Love Completed — Date';
function glb_ss_(){return SpreadsheetApp.getActiveSpreadsheet();}
function glb_posts_(){var s=glb_ss_().getSheetByName('posts');if(!s)throw Error('posts missing');return s;}
function glb_headers_(){return glb_posts_().getRange(1,1,1,glb_posts_().getLastColumn()).getDisplayValues()[0].map(function(x){return String(x).trim();});}
function glb_col_(heads,name){var i=heads.indexOf(name);if(i<0)throw Error('Missing posts column: '+name);return i+1;}
function glb_setup_(){
 var s=glb_posts_(),h=glb_headers_();
 [GLB_CI_DATE,GLB_SOL_DATE].forEach(function(name){if(h.indexOf(name)<0){s.getRange(1,s.getLastColumn()+1).setValue(name);h.push(name);}});
 var q=glb_ss_().getSheetByName(GLB_Q);
 if(!q){q=glb_ss_().insertSheet(GLB_Q);q.appendRow(['Post ID','Posts row','Title','Article Type','Status','CI completed','Sol completed','Sol cost USD','Detail','Updated']);q.setFrozenRows(1);}
 return q;
}
function glb_time_(){return Utilities.formatDate(new Date(),'Europe/London','yyyy-MM-dd HH:mm:ss');}
function glb_index_(){
 var sh=glb_posts_(),h=glb_headers_(),data=sh.getRange(2,1,Math.max(1,sh.getLastRow()-1),8).getDisplayValues(),o={};
 data.forEach(function(v,i){var id=String(v[3]||'').trim();if(!id)return;if(o[id])throw Error('Duplicate Post ID in posts: '+id);o[id]={row:i+2,title:v[1],type:v[7],material:v[6]};});
 return o;
}
function glb_parseIds_(text){
 var ids=String(text||'').split(/[\s,;]+/).map(function(x){return x.trim();}).filter(Boolean);
 if(!ids.length)throw Error('Enter at least one Post ID');
 var seen={};return ids.filter(function(id){if(!/^\d+$/.test(id))throw Error('Invalid Post ID '+id);if(seen[id])return false;seen[id]=true;return true;});
}
function glb_addPostIds(idsText){
 var q=glb_setup_(),ids=glb_parseIds_(idsText),idx=glb_index_(),existing={};
 if(q.getLastRow()>1)q.getRange(2,1,q.getLastRow()-1,1).getDisplayValues().forEach(function(v){existing[String(v[0])]=true;});
 var absent=ids.filter(function(id){return !idx[id];});if(absent.length)throw Error('Unknown Post IDs: '+absent.join(', '));
 var added=[];
 ids.forEach(function(id){if(existing[id])return;var x=idx[id];q.appendRow([id,x.row,x.title,x.type,'QUEUED','','',0,'','']);added.push(id);});
 return {added:added,alreadyQueued:ids.filter(function(id){return existing[id];}),queue:glb_queue()};
}
function glb_queue(){
 var q=glb_setup_(),n=q.getLastRow();if(n<2)return [];
 return q.getRange(2,1,n-1,10).getDisplayValues().map(function(v,i){return {queueRow:i+2,id:String(v[0]),postsRow:Number(v[1]),title:v[2],type:v[3],status:v[4],ciDate:v[5],solDate:v[6],cost:v[7],detail:v[8]};});
}
function glb_set_(r,status,message){var sh=glb_setup_();sh.getRange(r,5).setValue(status);sh.getRange(r,9).setValue(String(message||'').slice(0,2500));sh.getRange(r,10).setValue(glb_time_());}
function glb_stop(){PropertiesService.getScriptProperties().setProperty('GLB_STOP','YES');return glb_queue();}
function glb_start(){
 var running=glb_queue().filter(function(x){return /RUNNING/.test(x.status);});
 if(running.length)throw Error('Interrupted stage needs manual reconciliation for Post ID '+running[0].id+'. No paid call repeated.');
 PropertiesService.getScriptProperties().deleteProperty('GLB_STOP');
 return glb_queue();
}
function glb_ciContext_(row,headers){
 var sh=glb_posts_();
 var labels=['Confirmed Primary Intent','Secondary Intent Decisions','Page Rewrite Brief','Rewrite Role Lock','Strategic Reasoning','Topical Intent','Matrix Role','Key Decisions Explained','Analysis Brief','Authority Brief','Problem Angle','Article Angle','ac_entry_condition','ac_homeowner_perception','ac_material_behaviour','ac_constraint','ac_process_emphasis','ac_result_type','ac_narrative_archetype'];
 var lines=[];
 labels.forEach(function(k){var i=headers.indexOf(k);if(i<0)return;var t=String(sh.getRange(row,i+1).getDisplayValue()||'').trim();if(t){if(t.length>12000)throw Error('Content Intelligence field too long: '+k);lines.push('=== '+k+' ===\n'+t);}});
 return lines.join('\n\n');
}
function glb_caseEvidence_(row,headers){
 var sh=glb_posts_(),fields=typeof GLSE_FIELDS!=='undefined'?GLSE_FIELDS:[];
 if(fields.length!==16)throw Error('Structured evidence configuration unavailable');
 var vals=sh.getRange(row,glb_col_(headers,fields[0]),1,16).getDisplayValues()[0];
 var missing=fields.filter(function(x,i){return !String(vals[i]||'').trim();});
 if(missing.length)throw Error('Case-study evidence incomplete. Required before Sol: '+missing.join(', '));
 var entries=fields.map(function(k,i){return {label:k,value:String(vals[i])};});
 var media=glse_media_({postId:String(sh.getRange(row,glb_col_(headers,'Post ID')).getDisplayValue()),present:entries});
 var input=media.images.map(function(m,i){return JSON.stringify({order:i+1,url:m.url,alt:m.alt||'',originalCaption:m.originalCaption||'',sourceSection:m.sourceSection||'',nearbySourceText:m.nearbySourceText||'',suggestedStage:m.suggestedStage||'',placementConfidence:m.placementConfidence||''});}).join('\n');
 return {body:entries.filter(function(x){return x.label!=='Evidence Media & Links';}).map(function(x){return '=== '+x.label+' ===\n'+x.value;}).join('\n\n'),media:media,input:input};
}
function glb_solPrompt_(id,row){
 var a=glc_getArticle_(id),heads=glb_headers_(),ci=glb_ciContext_(row,heads);
 if(!a.technical||a.technicalSource.indexOf('technical!')!==0)throw Error('Approved Technical Matrix N:S missing for '+id);
 var isCase=a.type==='Case Study',ce=isCase?glb_caseEvidence_(row,heads):null;
 if(!isCase&&!String(a.original||'').trim())throw Error('Original article source missing for non-case '+id);
 var rules=[
 'Write one outstanding standalone UK homeowner article as clean HTML, with compelling logical sections and concise practical explanations.',
 'Single-pass complete article, with no subsequent AI rewrite. Answer the title and homeowner intent directly; natural UK English; never keyword-stuff.',
 'Only use project-specific facts actually supported by provided evidence. Never invent observations, tests, methods, prices, timings, client remarks or outcomes.',
 'Use technical matrix for professional technical correctness, not to invent work carried out. Distinguish recommendations from completed actions.',
 'Explain material behaviour and practical causes where relevant. Avoid academic chemistry, filler, repeated caveats and marketing fluff.',
 'Sealer references: homeowner copy says film-forming sealer when specifying urethane/polyurethane; retain topical, impregnating and recorded coat counts.',
 'Never use or imply the prohibited UK property-defect moisture diagnosis. Mention moisture/salt movement only where supported.',
 'Treat Content Intelligence as intent/structure guidance, not proof of project facts.',
 'No meta description, schema, publication instruction or invented supporting images.',
 'RETURN ONLY FULL ARTICLE HTML.',
 'POST ID '+id,'TITLE '+a.title,'MATERIAL '+a.material,'ARTICLE TYPE '+a.type,
 '=== GLOBAL TECHNICAL RULE ===',glc_moistureRule_(),
 '=== APPROVED TECHNICAL MATRIX ===',a.technical,
 '=== CONTENT INTELLIGENCE ===',ci||'(no extra brief recorded)'
 ];
 if(isCase){
  rules.push('CASE STUDY: follow actual restoration chronology, preserve documented decisions and results, and make the account useful on its own.',
   'Place relevant project photos by using exact original URLs and original caption/nearby context. Do not invent what an image depicts. Unclassified images may be omitted when unrelated.',
   'Do not silently drop useful before/during/after project photographs. Author/product/illustrative images are optional.',
   '=== STRUCTURED CASE STUDY EVIDENCE ===',ce.body,
   '=== SOURCE MEDIA INVENTORY ===',ce.input||'(no photos recorded)');
 }else{
  rules.push('=== HISTORICAL ARTICLE SOURCE (reference, not unquestioned authority) ===',a.original);
 }
 var prompt=rules.join('\n\n');if(prompt.length>170000)throw Error('Prompt too large; no source trimmed');
 return {prompt:prompt,media:ce&&ce.media};
}
function glb_archive_(id,oldHtml,newHtml){
 var ss=glb_ss_(),s=ss.getSheetByName(GLB_ARCHIVE);
 if(!s){s=ss.insertSheet(GLB_ARCHIVE);s.appendRow(['Saved local time','Post ID','Previous New HTML','New Google Love HTML']);s.setFrozenRows(1);}
 s.appendRow([glb_time_(),id,String(oldHtml||''),String(newHtml||'')]);
}
function glb_solRun_(id,row){
 var sh=glb_posts_(),h=glb_headers_(),p=glb_solPrompt_(id,row),col=glb_col_(h,'New HTML');
 var result=bc_sendPromptViaOpenAI(p.prompt,16000,'gpt-6.1-sol');
 if(!result||!result.success||result.truncated||!String(result.text||'').trim())throw Error('Sol did not return a complete article: '+String(result&&result.message||''));
 var html=String(result.text).replace(/^\s*`(?:html)?\s*/i,'').replace(/\s*`\s*$/,'').trim();
 if(html.length>49000||!/<h[12]\b/i.test(html)||!/<p\b/i.test(html))throw Error('Sol HTML failed basic completeness checks');
 if(/\brising[\s-]+damp\b/i.test(html.replace(/<[^>]*>/g,' ')))throw Error('Prohibited terminology found; article not saved');
 var check=p.media?glse_mediaCheck_(html,p.media):null;
 if(check&&(check.missingProjectEvidence.length||check.unexpected.length))throw Error('Project image retention requires review; draft not written to New HTML. Missing: '+check.missingProjectEvidence.join(' | '));
 var old=sh.getRange(row,col).getValue();
 glb_archive_(id,old,html);
 sh.getRange(row,col).setValue(html);
 return {cost:Number(result.cost||0),images:check?check.used:0};
}
function glb_next(){
 if(PropertiesService.getScriptProperties().getProperty('GLB_STOP')==='YES')return {done:true,stopped:true,queue:glb_queue()};
 var queue=glb_queue(),next=queue.filter(function(x){return x.status!=='DONE';})[0];
 if(!next)return {done:true,queue:queue};
 if(/RUNNING|REVIEW|FAILED/.test(next.status))return {done:true,blocked:true,message:'Review Post ID '+next.id+': '+next.status,queue:queue};
 var idx=glb_index_(),current=idx[next.id];if(!current||current.row!==next.postsRow)throw Error('Post row moved for '+next.id+'. Requeue or inspect before running');
 var s=glb_setup_(),h=glb_headers_(),ciCol=glb_col_(h,GLB_CI_DATE),solCol=glb_col_(h,GLB_SOL_DATE);
 if(next.status==='QUEUED'){
  glb_set_(next.queueRow,'CI RUNNING','Full existing Content Intelligence process called.');
  try{
   var out=runMasterContentIntelligenceBatchRow(next.postsRow,current.material);
   if(!out||!out.success)throw Error(String(out&&out.message||'CI failed'));
   var date=glb_time_();glb_posts_().getRange(next.postsRow,ciCol).setValue(date);s.getRange(next.queueRow,6).setValue(date);
   glb_set_(next.queueRow,'CI COMPLETE','Full Content Intelligence complete; Sol pending.');
   return {done:false,stage:'CI',id:next.id,queue:glb_queue()};
  }catch(e){glb_set_(next.queueRow,'CI — REVIEW',String(e.message||e));return {done:true,blocked:true,message:String(e.message||e),queue:glb_queue()};}
 }
 if(next.status==='CI COMPLETE'){
  // Validate all Sol inputs before reserving a paid request.
  try{glb_solPrompt_(next.id,next.postsRow);}catch(e){glb_set_(next.queueRow,'SOL PREFLIGHT — REVIEW',String(e.message||e));return {done:true,blocked:true,message:String(e.message||e),queue:glb_queue()};}
  glb_set_(next.queueRow,'SOL RUNNING','API request reserved. Check billing before retry after interruption.');
  try{
   var res=glb_solRun_(next.id,next.postsRow),date2=glb_time_();
   glb_posts_().getRange(next.postsRow,solCol).setValue(date2);
   s.getRange(next.queueRow,7).setValue(date2);s.getRange(next.queueRow,8).setValue(res.cost);
   glb_set_(next.queueRow,'DONE','Sol article saved to New HTML. '+res.images+' approved source images retained. Existing final modal stages unchanged.');
   return {done:false,stage:'SOL',id:next.id,queue:glb_queue()};
  }catch(e){glb_set_(next.queueRow,'SOL — REVIEW',String(e.message||e));return {done:true,blocked:true,message:String(e.message||e),queue:glb_queue()};}
 }
 return {done:true,blocked:true,message:'Unexpected batch status '+next.status,queue:glb_queue()};
}
function glb_openSidebar(){glb_setup_();SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('gl_batch_ci_sol').setTitle('Content Intelligence → Google Love Sol'));}
