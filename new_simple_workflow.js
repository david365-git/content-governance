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
  if(!ss.getSheetByName('posts') || ss.getActiveSheet().getName()!=='posts') throw new Error('Select a row in the posts tab.');
  return new_runWorkflowForRow_(ss.getActiveRange().getRow());
}
function new_runWorkflowForRow_(row) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName('posts');
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
      'Preserve the original author\'s professional diagnostic judgements, including assessments of likely or principal deterioration causes made from inspection and experience. Do not require laboratory or independent testing unless the article expressly claims such testing occurred.',
      'Clearly distinguish observed conditions, professional opinions about cause, and laboratory-confirmed findings; never upgrade opinion to tested fact or remove a documented professional assessment merely because tests were not recorded.',
      'Present documented work in technically correct order: preparation and repairs as appropriate, honing/refinement before final polishing, and sealing after mechanical finishing where those stages actually occurred. Do not invent omitted operations.',
      'The original article provides context, but generic technical benefits must never become claimed observed results.',
      'Outcome requirements must reflect documented visual results.',
      'Respect exact floor location (kitchen versus hallway), construction era, repairs, machinery, finish level, sealer, and handover.',
      'Preserve all original article images (exact src, in original order) unless clearly unrelated; do not invent photo descriptions or use an image as proof of a different operation.',
      'Preserve existing relevant internal links with exact href addresses; do not invent links. If original has no images or links, do not invent any.',
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
    var html=String(rewritten.text).trim();
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
      'An experienced restoration specialist\'s stated assessment of a cause is admissible source evidence of professional judgement, not an invented fact. Do NOT fail solely because there is no independent chemical test or laboratory confirmation. Flag only a cause newly invented by the rewrite, an unsupported claim of testing, or a technical contradiction; where necessary distinguish assessment from confirmed test findings.',
      'Where the original source attributes deterioration chiefly to unsuitable cleaning chemicals on the basis of the author\'s experience, preserve that assessment as professional judgement rather than requiring independent tests. Do not apply that cause to any project whose source does not support it.',
      'Audit the order of documented stages: where honing/refinement and sealing both occurred, sealing before honing is a sequencing error. Require mechanical honing and polishing before final sealing; correct narrative order without inventing any additional operation.',
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
    if(status==='FAIL') {
      var firstAudit=report;
      var correctionPrompt=[
        'ROLE: Conservative HTML copy editor. Correct ONLY the factual issues explicitly identified below.',
        'Keep all other paragraphs, layout, links, images, headings and documented facts unchanged.',
        'Do not add new observations, tests, measurements, processes or outcomes.',
        'Preserve any source-supported professional assessment that unsuitable cleaning chemicals were the principal cause of deterioration; do not delete it solely because independent testing was not done. Do not insert this cause where the source does not support it, and never present it as a laboratory result.',
        'When the audit identifies sealing before honing, reorder only the documented stages so honing/refinement and polishing precede final sealing. Do not remove the sealer or add new work.',
        '=== VERIFIED PROJECT EVIDENCE ===',evidence,
        '=== ORIGINAL ARTICLE ===',original,
        'Return ONLY complete corrected HTML. Do not include commentary or code fences.',
        '=== FIRST AUDIT ===',firstAudit,
        '=== EXISTING HTML ===',html
      ].join('\n\n');
      var correction=bc_sendPromptViaOpenAI(correctionPrompt,8500,model);
      cost+=Number(correction&&correction.cost||0);
      if(!correction||!correction.success||!String(correction.text||'').trim())throw new Error('Targeted correction failed: '+(correction&&correction.message||'empty result'));
      var candidate=String(correction.text).trim();
      if(!/<(?:header|section|p)\b/i.test(candidate)||candidate.length>49000)throw new Error('Targeted correction returned invalid HTML.');
      // Never silently lose original URLs, including images and links, in a correction.
      function urls(text){return [...String(text).matchAll(/<(?:img|a)\b[^>]*\b(?:src|href)=["']([^"']+)["']/gi)].map(function(m){return m[1];});}
      var beforeUrls=urls(html),afterUrls=urls(candidate);
      var missingUrls=beforeUrls.filter(function(u){return afterUrls.indexOf(u)<0;});
      if(missingUrls.length){
        // Reject the unsafe correction while retaining the valid rewrite and audit.
        report='INITIAL AUDIT:\\n'+firstAudit+'\\n\\nCORRECTION REJECTED: The proposed correction removed existing image/link URLs. Original rewrite retained.\\nMissing URLs:\\n'+missingUrls.join('\\n');
        status='FAIL';
      } else {
      var followUp=auditPrompt.replace('=== REWRITE ===\n\n'+html,'=== REWRITE ===\n\n'+candidate);
      var recheck=bc_sendPromptViaOpenAI(followUp,4200,model);
      cost+=Number(recheck&&recheck.cost||0);
      if(!recheck||!recheck.success||!String(recheck.text||'').trim())throw new Error('Correction verification failed: '+(recheck&&recheck.message||'empty result'));
      html=candidate;
      report='INITIAL AUDIT:\n'+firstAudit+'\n\nPOST-CORRECTION INDEPENDENT AUDIT:\n'+String(recheck.text).trim();
      status=/^VERDICT:\s*PASS\b/i.test(String(recheck.text).trim())?'PASS':/^VERDICT:\s*FAIL\b/i.test(String(recheck.text).trim())?'FAIL':'REVIEW';
      } // End accepted correction.
    }
    resultSheet.getRange(outRow,7).setValue(html);
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
/**
 * A fresh filtered-row snapshot. Only currently visible Case Study rows.
 * Keep execution per-row to avoid one Apps Script timeout for a large batch.
 */
function new_getFilteredCaseStudyRows() {
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if(!sh)throw new Error('posts sheet missing');
  var last=sh.getLastRow();
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();});
  var typeCol=headers.indexOf('Article Type')+1;
  if(!typeCol)throw new Error('Article Type heading missing');
  var types=last>1?sh.getRange(2,typeCol,last-1,1).getValues():[];
  var rows=[];
  for(var i=0;i<types.length;i++){
    var r=i+2;
    if(String(types[i][0]).trim()==='Case Study' && !sh.isRowHiddenByFilter(r) && !sh.isRowHiddenByUser(r))rows.push(r);
  }
  return rows;
}
/**
 * Called once per visible Case Study by sidebar; evidence extraction is conditional.
 * Only mark column D when BOTH the evidence check and the NEW workflow completed.
 */
function new_runFilteredCaseStudyRow(row) {
  row=Number(row);
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if(!sh||!Number.isInteger(row)||row<2||row>sh.getLastRow())throw new Error('Invalid posts row');
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();});
  function col(name){var c=headers.indexOf(name)+1;if(!c)throw new Error('Missing heading '+name);return c;}
  if(String(sh.getRange(row,col('Article Type')).getValue()).trim()!=='Case Study')throw new Error('Row '+row+' is not a Case Study');
  if(sh.isRowHiddenByFilter(row)||sh.isRowHiddenByUser(row))return {success:true,skipped:true,message:'Row hidden from filter',row:row,cost:0};
  // The fingerprint plus JSON evidence validity determines whether re-extraction is needed.
  var fresh=true;
  var evidenceLabels=['Project Location & Property','Observed Conditions','Client Requirements','Materials & Construction','Inspection & Tests','Restoration Methods','Products & Equipment','Technical Findings','Project Results','Maintenance Handover'];
  try{
    bc_caseEvidencePromptBlock_(row);
    for(var i=0;i<evidenceLabels.length;i++){
      var val=String(sh.getRange(row,col(evidenceLabels[i])).getValue()||'').trim();
      if(!val||!Array.isArray(JSON.parse(val))){fresh=false;break;}
    }
  }catch(e){fresh=false;}
  var extractionCost=0,extractionRan=false;
  if(!fresh){
    // If a fingerprint exists but an evidence cell is blank, force the extractor
    // to rebuild rather than return its "unchanged" shortcut.
    var statusCell=sh.getRange(row,col('Evidence Status'));
    var savedStatus=String(statusCell.getValue()||'');
    if(savedStatus.indexOf('SOURCE-QUOTED | sha256:')===0){
      statusCell.clearContent();
    }
    var ext=bc_caseEvidenceExtract_(row);
    extractionCost=Number(ext&&ext.cost||0);
    extractionRan=!ext.skipped;
    if(!ext||!ext.success)throw new Error('Evidence extraction failed on row '+row+': '+(ext&&ext.message||'unknown error'));
    bc_caseEvidencePromptBlock_(row);
  }
  var result=new_runWorkflowForRow_(row);
  // Status PASS or FAIL means workflow finished, whereas exceptions remain unmarked.
  sh.getRange(row,4).setBackground('#fff2cc').setFontColor('#0000ff');
  SpreadsheetApp.flush();
  result.row=row;
  result.evidenceExtracted=extractionRan;
  result.extractionCost=extractionCost;
  result.cost=Number(result.cost||0)+extractionCost;
  return result;
}
