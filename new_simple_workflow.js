/**
 * NEW: Simplified independent Case Study workflow.
 * Reuses existing OpenAI adapter and verified evidence helper.
 * Writes solely to "NEW Workflow Test". No WordPress publishing or posts edits.
 */
function new_openWorkflowSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('new_workflow_sidebar').setTitle('NEW — Simplified Workflow')
  );
}
function new_runWorkflowForSelectedRow() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName('posts');
  if(!sheet || ss.getActiveSheet().getName()!=='posts') throw new Error('Select a row in the posts tab.');
  var row=sheet.getActiveRange().getRow();
  if(row<2)throw new Error('Select a data row.');
  var headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();});
  function get(name) {var i=headers.indexOf(name);if(i<0)throw new Error('Missing column '+name);return sheet.getRange(row,i+1);}
  var postId=String(get('Post ID').getDisplayValue()).trim();
  var articleType=String(get('Article Type').getValue()).trim();
  if(articleType!=='Case Study')throw new Error('Pilot workflow supports Case Studies only.');
  var title=String(get('Title').getValue()||'');
  var material=String(get('Stone Type').getValue()||'');
  var exportSheet=ss.getSheetByName('site-export');
  if(!exportSheet)throw new Error('site-export missing.');
  var count=exportSheet.getLastRow()-1;
  var ids=count>0?exportSheet.getRange(2,1,count,1).getDisplayValues():[];
  var sourceRows=[];
  ids.forEach(function(v,i){if(String(v[0]).trim()===postId)sourceRows.push(i+2);});
  if(sourceRows.length!==1)throw new Error('Expected one original article for Post ID '+postId+', found '+sourceRows.length);
  var original=String(exportSheet.getRange(sourceRows[0],15).getValue()||'');
  if(!original.trim())throw new Error('Original article is empty in site-export column O.');
  // This checks the existing source fingerprint before making any paid API calls.
  var evidence=bc_caseEvidencePromptBlock_(row);
  var model=MODEL_MID;
  var resultSheet=ss.getSheetByName('NEW Workflow Test');
  if(!resultSheet){
    resultSheet=ss.insertSheet('NEW Workflow Test');
    resultSheet.appendRow(['Date','Posts Row','Post ID','Title','Material','Source Fingerprint','NEW Rewrite HTML','NEW Independent Audit','Status','API Cost USD','Model']);
    resultSheet.setFrozenRows(1);
  }
  var outRow=resultSheet.getLastRow()+1;
  resultSheet.getRange(outRow,1,1,6).setValues([[new Date(),row,postId,title,material,bc_caseEvidenceDigest_(original)]]);
  resultSheet.getRange(outRow,9).setValue('RUNNING');
  SpreadsheetApp.flush();
  var cost=0;
  try{
    var rewritePrompt=[
      'ROLE: UK floor-restoration case-study editor.',
      'TASK: Rewrite one real project into useful, concise, accurate HTML for nontechnical UK homeowners.',
      'Case Studies use first-person I for the author; do not invent attendance, activities, measurements or customer reactions.',
      'Documented source-quoted evidence controls all specific claims about the project.',
      'The original article provides context, but generic technical benefits must never become claimed observed results.',
      'Outcome requirements must reflect documented visual results.',
      'Respect exact floor location (kitchen versus hallway), construction era, repairs, machinery, finish level, sealer, and handover.',
      'Keep source images with original src URLs and meaningful captions ONLY when image content can be supported. Do not reuse images as documentation of another procedure.',
      'Keep useful original internal link URLs; do not invent links.',
      'Avoid generic tutorials, fabricated lighting comparisons, assumed tests, additional work, promises and repetitive headings.',
      'Use an introduction, varied H2 sections and a conclusion. WordPress provides the H1, so do not include an H1.',
      'Write British English. Return only raw HTML (no Markdown fences).',
      'TITLE: '+title,'MATERIAL: '+material,
      '=== PROJECT EVIDENCE ===',evidence,
      '=== ORIGINAL HTML ===',original
    ].join('\n\n');
    var rewritten=bc_sendPromptViaOpenAI(rewritePrompt,8500,model);
    cost+=Number(rewritten&&rewritten.cost||0);
    if(!rewritten||!rewritten.success||!String(rewritten.text||'').trim())throw new Error('Rewrite failed: '+(rewritten&&rewritten.message||'empty result'));
    var html=String(rewritten.text).trim().replace(/^`(x)`/,''); // placeholder
    if(!/<(?:header|section|p)\b/i.test(html))throw new Error('Model did not produce article HTML.');
    if(html.length>49000)throw new Error('HTML exceeds a Sheets cell limit.');
    resultSheet.getRange(outRow,7).setValue(html);
    resultSheet.getRange(outRow,10).setValue(cost);
    SpreadsheetApp.flush();
    var auditPrompt=[
      'ROLE: Independent senior UK floor restoration technical and factual auditor.',
      'Audit the rewrite; do NOT rewrite it.',
      'Check every project-specific observation, step, product, comparison, outcome, client statement and alleged test against quoted evidence and original HTML.',
      'Treat original as possibly imperfect: flag technically doubtful claims separately from source-faithfulness.',
      'Never insist on numerical results; outcome requirements must reflect documented visual results.',
      'First-person I is the preferred house style; do not fail it simply because the original used we.',
      'Check accuracy of finish level, tools, crack repairs, treatment chronology, aftercare, images/alt/captions, and links.',
      'First line must be exactly VERDICT: PASS or VERDICT: FAIL.',
      'Then include each material concern with exact quote, reason, and recommended correction, or NO MATERIAL ISSUES.',
      '=== VERIFIED PROJECT EVIDENCE ===',evidence,
      '=== ORIGINAL ARTICLE ===',original,
      '=== REWRITE ===',html
    ].join('\n\n');
    var audit=bc_sendPromptViaOpenAI(auditPrompt,4200,model);
    cost+=Number(audit&&audit.cost||0);
    if(!audit||!audit.success||!String(audit.text||'').trim())throw new Error('Audit failed: '+(audit&&audit.message||'empty result'));
    var report=String(audit.text).trim();
    var status=/^VERDICT:\s*PASS\b/i.test(report)?'PASS':/^VERDICT:\s*FAIL\b/i.test(report)?'FAIL':'REVIEW';
    if(report.length>49000)throw new Error('Audit exceeds Sheets cell limit.');
    resultSheet.getRange(outRow,8).setValue(report);
    resultSheet.getRange(outRow,9).setValue(status);
    resultSheet.getRange(outRow,10).setValue(cost);
    resultSheet.getRange(outRow,11).setValue(model);
    return {success:true,status:status,postId:postId,resultRow:outRow,cost:cost,message:'NEW Workflow Test row '+outRow+' — '+status};
  }catch(err){
    resultSheet.getRange(outRow,9).setValue('ERROR: '+String(err.message||err).slice(0,250));
    resultSheet.getRange(outRow,10).setValue(cost);
    throw err;
  }
}