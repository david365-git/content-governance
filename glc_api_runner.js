/**
 * Isolated, actual API comparison runner. NEVER writes posts/WordPress.
 * The name "Luna" is a UI test label: MODEL_CHEAP is the real API model.
 * Strong model can be set in Script Properties: GLC_STRONG_MODEL.
 */
function glc_apiRun(postId,route,promptVersion) {
  if (route!=='API — Luna' && route!=='API — stronger model') throw new Error('Choose an API route.');
  if(typeof bc_sendPromptViaOpenAI!=='function') throw new Error('Existing OpenAI API adapter is unavailable.');
  var a=glc_getArticle_(String(postId||'').trim());
  if (!a.original.trim()) throw new Error('No original content found in site-export. No API call made.');
  if (!a.technical.trim() || a.technicalSource.indexOf('technical!')!==0) throw new Error('Approved N:S Technical Matrix not found. No API call made.');
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var globalRule=String(ss.getSheetByName('technical').getRange('C1').getDisplayValue()||'');
  var model=route==='API — Luna'
    ? (typeof MODEL_CHEAP!=='undefined'?MODEL_CHEAP:null)
    : (PropertiesService.getScriptProperties().getProperty('GLC_STRONG_MODEL') || (typeof MODEL_STRONG!=='undefined'?MODEL_STRONG:null));
  if (!model) throw new Error('Model not configured for '+route+'. Set GLC_STRONG_MODEL for stronger model.');
  var prompt=[
    'ROLE: Skilled UK stone and tile floor restoration editor. Create a genuinely useful homeowner article.',
    'This is an ISOLATED COMPARISON PILOT, not a claim of owner approval or final publishability.',
    'Use British English, clear headings, concise homeowner-friendly language, and material-specific practical reasoning.',
    'Answer the article title directly. Preserve documented experience, but do not invent project details, personal anecdotes or results.',
    'The CURRENT owner-approved Technical Matrix N:S is authoritative; the original article is historical text, NOT independently proven job evidence.',
    'Never infer specific work carried out from general technical guidance. Distinguish likely causes from actual observations.',
    'No product recommendations or brand-name product insertion. No unsupported guarantees.',
    'Preserve relevant existing image src attributes and honest captions. Do not portray illustrations as project photographs.',
    'Preserve relevant internal links from the source only if they can be checked later; do not invent links.',
    'Do not include banned UK property-defect moisture terminology or imply that diagnosis.',
    'Return only the article HTML, without markdown.',
    'POST ID: '+a.postId,
    'TITLE: '+a.title,
    'ARTICLE TYPE: '+a.type,
    'MATERIAL: '+a.material,
    '=== GLOBAL TECHNICAL RULE ===',globalRule,
    '=== CURRENT APPROVED TECHNICAL MATRIX (N:S) ===',a.technical,
    '=== HISTORICAL SOURCE HTML (not verified project facts) ===',a.original
  ].join('\n\n');
  var auditPrompt=[
    'You are an independent UK floor-restoration article auditor. Do not rewrite the article.',
    'Assess against the article title, approved technical reference and source HTML.',
    'Flag fabricated first-person experience, unsupported case facts, misleading images, incorrect material-specific advice, lost relevant figures, bad intent fit and unsupported product insertions.',
    'Report a provisional result, not owner approval: PASS, FAIL or BLOCKED. Be exact about evidence and uncertainty.',
    'Return JSON with keys verdict, important_findings (array), notes (string). No Markdown fences.',
    'TITLE: '+a.title,'MATERIAL: '+a.material,'APPROVED N:S: '+a.technical,'GLOBAL RULE: '+globalRule
  ].join('\n\n');
  // Use the same existing adapter as the simplified workflow.
  if(prompt.length>49000) throw new Error('Full prompt exceeds one Sheets cell; cannot perform fair comparison without separate complete prompt storage.');
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
  var model=String(sh.getRange(row,17).getValue()||'');
  if(!model)throw new Error('Missing recorded API model.');
  var auditPrompt=['Provisional SELF-CHECK only: assess the article for unsupported experience, misleading images, technical errors and whether it answers the title. Never claim independence or owner approval. Return concise JSON with verdict and findings.','SOURCE PROMPT:',prompt,'GENERATED HTML:',html].join('\n\n');
  var response=bc_sendPromptViaOpenAI(auditPrompt,3500,model);
  var oldCost=Number(sh.getRange(row,18).getValue()||0);
  sh.getRange(row,18).setValue(oldCost+Number(response&&response.cost||0));
  if(!response||!response.success){sh.getRange(row,8).setValue('SELF-CHECK FAILED');throw new Error(String(response&&response.message||'Self-check returned no result'));}
  sh.getRange(row,19).setValue(String(response.text||'').slice(0,49000));
  sh.getRange(row,8).setValue('SELF-CHECK RECORDED — NOT INDEPENDENT');
  return {row:row,status:'Self-check recorded (not independent)',cost:Number(response.cost||0),audit:String(response.text||'').slice(0,600)};
}
