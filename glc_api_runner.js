/**
 * Isolated, actual API comparison runner. NEVER writes posts/WordPress.
 * The name "Luna" is a UI test label: MODEL_CHEAP is the real API model.
 * Strong model can be set in Script Properties: GLC_STRONG_MODEL.
 */
function glc_apiRun(postId,route,promptVersion,preparedRow) {
  if (route!=='API — Luna' && route!=='API — stronger model') throw new Error('Choose an API route.');
  if(typeof bc_sendPromptViaOpenAI!=='function') throw new Error('Existing OpenAI API adapter is unavailable.');
  var a=glc_getArticle_(String(postId||'').trim());
  if (!a.original.trim()) throw new Error('No original content found in site-export. No API call made.');
  if (!a.technical.trim() || a.technicalSource.indexOf('technical!')!==0) throw new Error('Approved N:S Technical Matrix not found. No API call made.');
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var globalRule=glc_moistureRule_();
  var model=route==='API — Luna'
    ? (typeof MODEL_CHEAP!=='undefined'?MODEL_CHEAP:null)
    : (PropertiesService.getScriptProperties().getProperty('GLC_STRONG_MODEL') || (typeof MODEL_STRONG!=='undefined'?MODEL_STRONG:null));
  if (!model) throw new Error('Model not configured for '+route+'. Set GLC_STRONG_MODEL for stronger model.');
  var prompt=glc_buildPrompt_(a,globalRule);
  if(!preparedRow)throw new Error('Prepare and freeze a v1.2 prompt before running an API model.');
  var prepared=glc_readPreparedPrompt_(preparedRow,a.postId);
  if(prepared.version!=='1.2')throw new Error('Prepared prompt is not version 1.2.');
  prompt=prepared.prompt;
  var auditPrompt=[
    'You are an independent UK floor-restoration article auditor. Do not rewrite the article.',
    'Assess against the article title, approved technical reference and source HTML.',
    'Flag fabricated first-person experience, unsupported case facts, misleading images, incorrect material-specific advice, lost relevant figures, bad intent fit and unsupported product insertions.',
    'Report a provisional result, not owner approval: PASS, FAIL or BLOCKED. Be exact about evidence and uncertainty.',
    'Return JSON with keys verdict, important_findings (array), notes (string). No Markdown fences.',
    'TITLE: '+a.title,'MATERIAL: '+a.material,'APPROVED N:S: '+a.technical,'GLOBAL RULE: '+globalRule
  ].join('\n\n');
  // Use the same existing adapter as the simplified workflow.
  if(prompt.length>49000) throw new Error('Prompt exceeds single-cell limit.');
  var stamp=Utilities.formatDate(new Date(),'Etc/UTC',"yyyy-MM-dd'T'HH:mm:ss'Z'");
  var sheet=ss.getSheetByName('GL Model Comparison Tests');
  if(!sheet){
    sheet=ss.insertSheet('GL Model Comparison Tests');
    sheet.appendRow(glc_headings_());
    sheet.setFrozenRows(1);
  }
  var lock=LockService.getDocumentLock() || LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error('Another comparison is reserving a test row. Retry.');
  var row;
  try {
    sheet.appendRow([stamp,a.postId,a.title,a.material,a.type,route,String(promptVersion||'pilot-v1'),'RUNNING']);
    row=sheet.getLastRow();
  } finally {lock.releaseLock();}
  sheet.getRange(row,11,1,6).setValues([[a.originalSource,a.original.length,a.technicalSource,a.technical.length,(a.original.match(/<img\b/gi)||[]).length,a.url]]);
  sheet.getRange(row,17).setValue(String(model));
  sheet.getRange(row,21).setValue(glc_hash_(prompt));
  sheet.getRange(row,20).setValue(prompt);
  SpreadsheetApp.flush();
  var cost=0;
  try{
    var draft=bc_sendPromptViaOpenAI(prompt,16000,model);
    cost+=Number(draft&&draft.cost||0);
    if(!draft||!draft.success||!String(draft.text||'').trim()) throw new Error('API draft failed: '+String(draft&&draft.message||'empty response'));
    if(draft.truncated){sheet.getRange(row,22).setValue('TRUE');throw new Error('Model output truncated; draft rejected.');}
    var html=String(draft.text).replace(/^\s*```(?:html)?\s*/i,'').replace(/\s*```\s*$/,'').trim();
    if(html.length>49000)throw new Error('Draft too long for single-cell test storage.');
    if(!/<(?:p|section|h2)\b/i.test(html))throw new Error('Model did not return article HTML.');
    sheet.getRange(row,9).setValue(html);
    sheet.getRange(row,18).setValue(cost);
    sheet.getRange(row,8).setValue('DRAFT READY — SELF-CHECK SEPARATE');
    return {row:row,model:String(model),cost:cost,status:'Draft ready — run separate self-check',audit:'Saved with prompt SHA-256.'};
  }catch(e){
    sheet.getRange(row,8).setValue('FAILED');
    sheet.getRange(row,10).setValue(String(e.message||e).slice(0,8000));
    sheet.getRange(row,18).setValue(cost);
    throw new Error('Test row '+row+': '+e.message);
  }
}

function glc_runSelfCheck(testRow){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GL Model Comparison Tests');
  var row=Number(testRow);
  if(!sh||!Number.isInteger(row)||row<2||row>sh.getLastRow())throw new Error('Invalid comparison test row');
  var html=String(sh.getRange(row,9).getValue()||'');
  if(!html)throw new Error('No saved HTML for this test.');
  var prompt=String(sh.getRange(row,20).getValue()||'');
  var model=String(typeof MODEL_CHEAP!=='undefined'?MODEL_CHEAP:'');
  if(!model)throw new Error('Missing recorded API model.');
  var auditPrompt=['Common fixed-model comparison check (same checker for both API writers). Provisional; not an independent human audit. Do not rewrite.', 'Provisional SELF-CHECK only: assess the article for unsupported experience, misleading images, technical errors and whether it answers the title. Never claim independence or owner approval. Return concise JSON with verdict and findings.','SOURCE PROMPT:',prompt,'GENERATED HTML:',html].join('\n\n');
  var response=bc_sendPromptViaOpenAI(auditPrompt,3500,model);
  var oldCost=Number(sh.getRange(row,18).getValue()||0);
  sh.getRange(row,18).setValue(oldCost+Number(response&&response.cost||0));
  if(!response||!response.success){sh.getRange(row,8).setValue('SELF-CHECK FAILED');throw new Error(String(response&&response.message||'Self-check returned no result'));}
  sh.getRange(row,19).setValue(String(response.text||'').slice(0,49000));
  sh.getRange(row,8).setValue('SELF-CHECK RECORDED — NOT INDEPENDENT');
  return {row:row,status:'Self-check recorded (not independent)',cost:Number(response.cost||0),audit:String(response.text||'').slice(0,600)};
}

/** The only global rule is the complete first C1 block, stopping at the legacy marker. */
function glc_moistureRule_(){
 var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('technical');
 if(!sh)throw new Error('technical sheet missing');
 var c1=String(sh.getRange('C1').getDisplayValue()||'');
 var marker='MatrixTECHNICAL GUARDRAILS FOR CONTENT GENERATION';
 var i=c1.indexOf(marker);
 if(i<0)throw new Error('C1 legacy boundary not found; refusing to send material-specific legacy guidance as global.');
 var rule=c1.substring(0,i).trim();
 if(rule.indexOf('GLOBAL FLOOR-MOISTURE TERMINOLOGY RULE')!==0 || rule.length<150)
  throw new Error('C1 global moisture restriction is malformed; generation blocked.');
 return rule;
}
/** Freeze exact model input before either model runs. Only the isolated test sheet is written. */
function glc_preparePrompt(postId){
 var a=glc_getArticle_(String(postId||'').trim());
 if(!a.original.trim()||!a.technical.trim()||a.technicalSource.indexOf('technical!')!==0)
  throw new Error('Article source or approved N:S matrix missing.');
 var globalRule=glc_moistureRule_();
 var prompt=glc_buildPrompt_(a,globalRule);
 if(prompt.length>49000)throw new Error('Prompt exceeds Sheets single-cell limit; preparation blocked without truncation.');
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('GL Comparison Prompts');
 if(!sh){
  sh=ss.insertSheet('GL Comparison Prompts');
  sh.appendRow(['Created UTC','Post ID','Version','SHA-256','Exact prompt','Technical source','Historical source']);
  sh.setFrozenRows(1);
 }
 var lock=LockService.getDocumentLock()||LockService.getScriptLock();
 if(!lock.tryLock(30000))throw new Error('Prompt sheet busy.');
 var row;
 try{
  sh.appendRow([Utilities.formatDate(new Date(),'Etc/UTC',"yyyy-MM-dd'T'HH:mm:ss'Z'"),a.postId,'1.2',glc_hash_(prompt),prompt,a.technicalSource,a.originalSource]);
  row=sh.getLastRow();
 }finally{lock.releaseLock();}
 return {row:row,hash:glc_hash_(prompt),length:prompt.length,postId:a.postId,version:'1.2'};
}
function glc_buildPrompt_(a,globalRule){
 return [
 'ROLE: Skilled UK stone and tile floor restoration editor. Create a genuinely useful homeowner article.',
 'This is an ISOLATED COMPARISON PILOT, not a claim of owner approval or final publishability.',
 'Use British English, clear headings, concise homeowner-friendly language, and material-specific practical reasoning.',
 'Answer the article title directly. Preserve documented experience, but do not invent project details, personal anecdotes or results.',
 'TECHNICAL PRECEDENCE: (1) Technical Restrictions & Safeguards in S and the global moisture rule; (2) specific qualified matrix N–R statements including owner experience and frequency; (3) summaries, takeaways and headings; (4) the historical article. For conflicting claims use the more qualified statement or omit it. Never turn a heading, summary, or possibility into a required procedure.',
 'Describe treatments as options selected after assessment and testing. Use a fixed sequence, must, always, or a treatment superlative only when the applicable matrix body supports it. Include optional or specialist treatments only when relevant to the question; preserve stated frequencies.',
 'From the historical article retain relevant genuine knowledge, images and internal links consistent with the approved matrix. Correct or omit conflicting technical claims. Never manufacture project evidence.',
 'Never infer specific work carried out from general technical guidance. Distinguish likely causes from actual observations.',
 'No product recommendations or brand-name product insertion. No unsupported guarantees.',
 'Preserve relevant existing image src attributes and honest captions. Do not portray illustrations as project photographs.',
 'Preserve relevant internal links from the source only if they can be checked later; do not invent links.',
 'Do not include banned UK property-defect moisture terminology or imply that diagnosis.',
 'Return only the article HTML, without markdown.',
 'POST ID: '+a.postId,'TITLE: '+a.title,'ARTICLE TYPE: '+a.type,'MATERIAL: '+a.material,
 '=== GLOBAL TECHNICAL RULE ===',globalRule,
 '=== CURRENT APPROVED TECHNICAL MATRIX (N:S) ===',a.technical,
 '=== HISTORICAL SOURCE HTML (not verified project facts) ===',a.original
 ].join('\n\n');
}
function glc_readPreparedPrompt_(row,postId){
 var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GL Comparison Prompts');
 row=Number(row);
 if(!sh||!Number.isInteger(row)||row<2||row>sh.getLastRow())throw new Error('Invalid frozen prompt row.');
 var v=sh.getRange(row,2,1,4).getValues()[0];
 if(String(v[0])!==String(postId))throw new Error('Frozen prompt belongs to a different Post ID.');
 var prompt=String(v[3]||''),hash=glc_hash_(prompt);
 if(hash!==String(v[2]))throw new Error('Frozen prompt SHA-256 mismatch.');
 return {version:String(v[1]),prompt:prompt,hash:hash};
}
