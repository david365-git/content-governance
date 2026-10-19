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
/**
 * Standard author credibility panel and honest next-step link for Case Studies.
 * Add in code (not the AI rewrite) so the panel cannot be silently omitted.
 */
function new_addAuthorFooter_(html,material) {
  html=String(html||'');
  var escaped=String(material||'stone').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  var authorImage='https://www.abbeyfloorcare.co.uk/wp-content/uploads/David_Allen.jpg';
  var isVictorian=/victorian|minton|encaustic/i.test(escaped);
  var bio=isVictorian
    ? 'David Allen has more than 30 years of hands-on experience restoring Victorian and encaustic tile floors through Abbey Floor Care. This case study documents his first-hand assessment, restoration work and practical decisions on this floor.'
    : 'David Allen has more than 30 years of experience in hard-floor cleaning and restoration through Abbey Floor Care. In this case study, he shares first-hand experience of assessing and restoring '+escaped+' floors, including the methods used and practical maintenance considerations.';
  // Idempotent: model corrections may return existing standard footer.
  // Remove all copies before rebuilding exactly one author panel and one CTA.
  html=html.replace(/<section\b[^>]*class=["'][^"']*\bnew-author-bio\b[^"']*["'][^>]*>[\s\S]*?<\/section>/gi,'');
  html=html.replace(/<p\b[^>]*class=["'][^"']*\bnew-case-study-cta\b[^"']*["'][^>]*>[\s\S]*?<\/p>/gi,'');
  html=html.trim();
  html+='\n<section class="new-author-bio" aria-label="About the author" style="display:flex;gap:20px;align-items:flex-start;padding:20px;border:1px solid #e7d7a9;border-radius:10px;background:#fff6d9;margin:28px 0;">'
    +'<img src="'+authorImage+'" alt="David Allen, stone floor restoration specialist" width="100" height="105" style="width:100px;height:105px;object-fit:cover;border-radius:7px;">'
    +'<div><p><strong>David Allen — <a href="https://www.abbeyfloorcare.co.uk/">Abbey Floor Care</a></strong></p>'
    +'<p>'+bio+'</p></div></section>';
  html+='\n<p class="new-case-study-cta">For advice about a '+escaped+' floor, <a href="https://www.abbeyfloorcare.co.uk/">visit Abbey Floor Care to explore our floor restoration advice and services</a>.</p>';
  return html;
}

/**
 * Build a small, governed set of published/internal link candidates for one case study.
 * Use actual posts URLs only; do not trust Feeds Hub if it points to a non-hub page.
 */
function new_caseStudyLinks_(sh,row,title,material,ownUrl){
  var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();});
  function c(n){return h.indexOf(n);}
  var fields=['Title','URL','Stone Type','Article Type','Feeds Hub'];
  if(fields.some(function(n){return c(n)<0;}))throw new Error('Missing governed linking columns');
  var data=sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),Math.max.apply(null,fields.map(function(n){return c(n)+1;}))).getValues();
  function canonical(v){return String(v||'').trim().replace(/^http:/i,'https:').replace(/^https:\/\/www\./i,'https://').replace(/\/+$/,'').toLowerCase();}
  var own=canonical(ownUrl),mat=String(material||'').trim().toLowerCase();
  var stop={the:1,and:1,for:1,with:1,from:1,need:1,floor:1,floors:1,terrazzo:1,restoration:1,restoring:1,polishing:1,cleaning:1};
  function words(v){return String(v||'').toLowerCase().split(/[^a-z0-9]+/).filter(function(x){return x.length>3&&!stop[x];});}
  var targets=words(title);
  var list=data.map(function(v,i){return {row:i+2,title:String(v[c('Title')]||''),url:String(v[c('URL')]||'').trim(),material:String(v[c('Stone Type')]||'').trim().toLowerCase(),type:String(v[c('Article Type')]||'').trim()};})
    .filter(function(v){return v.url&&/^https:\/\/(www\.)?abbeyfloorcare\.co\.uk\//i.test(v.url)&&canonical(v.url)!==own&&v.material===mat;});
  var ownRow=data[row-2]||[];
  var fed=canonical(ownRow[c('Feeds Hub')]);
  var hubs=list.filter(function(v){return v.type==='Hub Page';});
  var hub=hubs.find(function(v){return canonical(v.url)===fed;});
  if(!hub&&hubs.length===1)hub=hubs[0];
  // If multiple candidate hubs exist and Feeds Hub is invalid, do not guess.
  var other=list.filter(function(v){return v.type==='Method Guide'||v.type==='Diagnostic Guide'||v.type==='Educational Guide'||v.type==='Buyer Guide'||v.type==='Service Page'||v.type==='Case Study';})
    .map(function(v){var w=words(v.title);v.score=w.filter(function(x){return targets.indexOf(x)>=0;}).length+(v.type==='Case Study'?0:1);return v;})
    .sort(function(a,b){return b.score-a.score||a.row-b.row;}).slice(0,12);
  return {hub:hub||null,related:other,invalidHub:fed&&!hub?String(ownRow[c('Feeds Hub')]||''):''};
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
  var linkPlan=new_caseStudyLinks_(sheet,row,title,material,String(get('URL').getValue()||''));
  var approvedLinks=(linkPlan.hub?[linkPlan.hub]:[]).concat(linkPlan.related);
  if(!original.trim())throw new Error('Original article is empty in site-export column O.');
  // This checks the existing source fingerprint before making any paid API calls.
  var evidence=bc_caseEvidencePromptBlock_(row);
  // Use the approved live M:S reference; never use historic technical columns C or L.
  // The relevant M:R excerpt is bounded, and column S safeguards are always included.
  var approvedTechnicalReference=bc_getRelevantTechnicalMatrixExcerpt_(material,original,3200,articleType);
  if(!approvedTechnicalReference)throw new Error('Approved Technical Matrix M:S reference is empty for '+material);
  var model=MODEL_CHEAP;
  // Column H: publication/technical PASS gate. Column M: separate original-vs-rewrite assessment.
  var resultSheet=ss.getSheetByName('NEW Workflow Test');
  if(!resultSheet){
    resultSheet=ss.insertSheet('NEW Workflow Test');
    resultSheet.appendRow(['Date','Posts Row','Post ID','Title','Material','Source Fingerprint','NEW Rewrite HTML','Publication & Technical Quality Audit','Status','API Cost USD','Model','Original html','Original vs Rewrite — AI Comparison Report']);
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
      'TASK: Write an outstanding search-focused real-project case study for UK homeowners: lead with the problem and result, answer likely search intent, give original first-hand experience and useful decisions, and use clear, engaging, concise British English. Never keyword-stuff.',
      'Case Studies use first-person I for the author; do not invent attendance, activities, measurements or customer reactions.',
      'Documented source-quoted evidence controls all specific claims about the project.',
      'Preserve the original author\'s professional diagnostic judgements, including assessments of likely or principal deterioration causes made from inspection and experience. Do not require laboratory or independent testing unless the article expressly claims such testing occurred.',
      'Clearly distinguish observed conditions, professional opinions about cause, and laboratory-confirmed findings; never upgrade opinion to tested fact or remove a documented professional assessment merely because tests were not recorded.',
      'Present documented work in technically correct order: preparation and repairs as appropriate, honing/refinement before final polishing, and sealing after mechanical finishing where those stages actually occurred. Do not invent omitted operations.',
      'E-E-A-T: demonstrate EXPERIENCE through specific first-hand observations, constraints, decisions, tools and outcomes actually recorded; EXPERTISE through clear material-specific reasoning that helps a homeowner understand why methods were selected; AUTHORITATIVENESS through an authentic specialist voice and relevant original project photographs; TRUST through accuracy, honest limitations, visible authorship where supported, and no fabricated proof or credentials.',
      'Google helpful people-first content: make this worth reading even without search traffic. Answer the underlying homeowner question, explain what was unusual about THIS job and what a homeowner can learn from it. Avoid templated filler, recycled generic advice, sales pressure, keyword repetition and claims not supported by the source.',
      'Show the who, how and why naturally: who carried out the restoration when supported, how and why key decisions were made, and what the result can and cannot demonstrate. Do not invent tests or assurances.',
      'The original article provides context, but generic technical benefits must never become claimed observed results.',
      'Outcome requirements must reflect documented visual results.',
      'Respect exact floor location (kitchen versus hallway), construction era, repairs, machinery, finish level, sealer, and handover.',
      'Preserve all original article images (exact src, in original order) unless clearly unrelated; do not invent photo descriptions or use an image as proof of a different operation.',
      'Retain every relevant existing original internal link with its exact href (including links to comparable restoration case studies and diagnostic guides), incorporating it naturally where its subject is discussed. Do not drop original links to meet the limit on NEW links. Do not invent links. The application adds the one standard author bio and one CTA: do not include either in your output.',
      'CONTEXTUAL LINK GOVERNANCE: Link to the verified material hub when supplied, naturally within a relevant explanatory paragraph. Select up to 2 additional genuinely relevant supporting destinations from the APPROVED LINKS list; never invent destinations or change supplied URLs. Do not add links if relevance is weak. Preserve existing relevant links from the original.',
      'Integrate links in meaningful paragraphs at the point the subject is discussed, with descriptive natural anchor text. Avoid generic "click here", standalone signposting, overlinking and repetitive exact-match anchor text.',
      'Avoid generic tutorials, fabricated lighting comparisons, assumed tests, additional work, promises and repetitive headings.',
      'Use an introduction, varied H2 sections and a conclusion. WordPress provides the H1, so do not include an H1.',
      'Write British English. Return only raw HTML (no Markdown fences).',
      'TITLE: '+title,'MATERIAL: '+material,
      '=== APPROVED GOVERNED LINKS ===',JSON.stringify({verifiedHub:linkPlan.hub,related:linkPlan.related}),
      '=== APPROVED MATERIAL TECHNICAL REFERENCE (M:S) ===',approvedTechnicalReference,
      'Use this reference for general material-specific technical guidance, but NEVER treat it as evidence of work performed on this project.',
      '=== PROJECT EVIDENCE ===',evidence,
      '=== ORIGINAL HTML ===',original
    ].join('\n\n');
    var rewritten=bc_sendPromptViaOpenAI(rewritePrompt,8500,model);
    cost+=Number(rewritten&&rewritten.cost||0);
    if(!rewritten||!rewritten.success||!String(rewritten.text||'').trim())throw new Error('Rewrite failed: '+(rewritten&&rewritten.message||'empty result'));
    var html=new_addAuthorFooter_(String(rewritten.text).trim(),material);
    // Reject invented destinations in a rewrite, while allowing old source links and our footer.
    function htmlLinks(v){return [...String(v).matchAll(/<a\b[^>]*\bhref=["']([^"']+)["']/gi)].map(function(m){return m[1];});}
    var accepted=htmlLinks(original).concat(approvedLinks.map(function(v){return v.url;}));
    accepted.push('https://www.abbeyfloorcare.co.uk/');
    function normLink(v){return String(v||'').replace(/^http:/i,'https:').replace(/^https:\/\/www\./i,'https://').replace(/\/+$/,'').toLowerCase();}
    var allowedSet=accepted.map(normLink);
    var unknown=htmlLinks(html).filter(function(u){return allowedSet.indexOf(normLink(u))<0;});
    if(unknown.length)throw new Error('Rewrite introduced unapproved internal links: '+unknown.slice(0,3).join(', '));
    if(!/<(?:header|section|p)\b/i.test(html))throw new Error('Model did not produce article HTML.');
    if(html.length>49000)throw new Error('HTML exceeds a Sheets cell limit.');
    resultSheet.getRange(outRow,7).setValue(html);
    resultSheet.getRange(outRow,10).setValue(cost);
    SpreadsheetApp.flush();
    var auditPrompt=[
      'ROLE: Senior UK stone-restoration editor reviewing publication quality, SEO usefulness and material accuracy.',
      'Judge the article as a whole. The goal is a helpful, distinctive, trustworthy case study that satisfies the homeowner search intent. Google rankings are not guaranteed.',
      'Apply Google people-first helpful content and E-E-A-T principles holistically: first-hand Experience, explained material-specific Expertise, authentic specialist Authority, and above all Trust. E-E-A-T is not a standalone ranking factor.',
      'A strong case study should offer original project insight and photos where available, clearly identify the real problem, explain key choices and limitations, show the actual outcome, and leave the homeowner better informed than generic restoration advice.',
      'Assess whether the content is satisfying and genuinely useful to the intended reader, with clear headings, authentic tone, natural local relevance, descriptive media and honest authorship. Do not require every E-E-A-T element as a separate checkbox or invent author qualifications.',
      'FAIL for materially thin, templated, generic, misleading or unhelpful content, not merely because a minor detail might be phrased differently. Focus on meaningful improvement, not arbitrary length or keyword targets.',
      'PASS when the article is publishable and has no MATERIAL issue, even if minor detail, paraphrase, equipment-list wording or nonessential alt text could be polished.',
      'FAIL ONLY for a significant problem: misleading or invented core project facts or work; major technical error or unsafe advice; incorrect core outcome; major loss of original relevant photos or links; serious search-intent mismatch, thin/generic content, keyword stuffing or unusable structure.',
      'Do NOT do forensic word-by-word comparisons or fail over minor chronology omissions, harmless phrasing, subjective descriptions, nonessential captions/alt wording, or listed-versus-used equipment when meaning and credibility are not materially affected.',
      'Preserve source-supported professional assessments of deterioration causes; do not demand laboratory evidence or reframe them as independent testing.',
      'Check that the homeowner problem, what was actually done, relevant constraints and results are prominent; evaluate headings, readability, helpfulness and natural use of relevant location/material terms.',
      'Consider source-quoted evidence and original article as ground truth for substantive case facts. Correct major process-sequencing errors such as presenting sealing before honing when that is only a writing mistake; do not invent operations.',
      'Images and relevant internal links matter. Flag missing original URLs as a material issue; do not fail merely for small alt-text differences.',
      'Evaluate newly inserted contextual links against APPROVED GOVERNED LINKS. Preserve existing original article links even when their published URLs differ from current governed URL data. Do not flag an original link as broken or demand replacement merely because the path differs; require actual evidence of a broken URL. The verified material hub should appear naturally in the body when available. Up to two relevant related links are useful, not a quota; never reward irrelevant links.',
      'Confirm the standard David Allen author box with portrait and a useful closing next-step link are present. A homepage link is acceptable when clearly described as a website link rather than a direct contact form. Treat absence as a material omission, but do not otherwise micromanage their wording.',
      'First line must be exactly VERDICT: PASS or VERDICT: FAIL.',
      'If FAIL, state at most three material blockers with a concise exact quote and focused fix. If PASS, write NO MATERIAL ISSUES and optionally one short nonblocking improvement suggestion.',
      '=== APPROVED GOVERNED LINKS ===',JSON.stringify({verifiedHub:linkPlan.hub,related:linkPlan.related}),
      '=== APPROVED MATERIAL TECHNICAL REFERENCE (M:S) ===',approvedTechnicalReference,
      'Check material-specific technical claims against this reference. Project-specific facts must still be supported by the original article or verified evidence.',
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
        'ROLE: Publication-quality HTML editor. Correct ONLY the material blockers explicitly identified by the audit. Prioritise reader clarity and search intent; do not chase trivial discrepancies.',
        'Preserve source-grounded project facts and professional judgement. Retain all images and valid original source links. A link URL may be corrected only where its replacement is explicitly documented in APPROVED GOVERNED LINKS or verified as a working destination; do not invent or guess link destinations. Make minimal coherent edits to resolve material blockers without flattening the article.',
        'Do not add new observations, tests, measurements, processes or outcomes.',
        'Do not generate an author box, author portrait, or closing CTA. The application will add exactly one of each after correction.',
        'Restore relevant original internal links missing from the rewrite, keeping the original hrefs; do not remove other valid links.',
        'Preserve any source-supported professional assessment that unsuitable cleaning chemicals were the principal cause of deterioration; do not delete it solely because independent testing was not done. Do not insert this cause where the source does not support it, and never present it as a laboratory result.',
        'When the audit identifies sealing before honing, reorder only the documented stages so honing/refinement and polishing precede final sealing. Do not remove the sealer or add new work.',
        '=== APPROVED MATERIAL TECHNICAL REFERENCE (M:S) ===',approvedTechnicalReference,
        '=== VERIFIED PROJECT EVIDENCE ===',evidence,
        '=== ORIGINAL ARTICLE ===',original,
        'Return ONLY complete corrected HTML. Do not include commentary or code fences.',
        '=== FIRST AUDIT ===',firstAudit,
        '=== EXISTING HTML ===',html
      ].join('\n\n');
      var correction=bc_sendPromptViaOpenAI(correctionPrompt,8500,model);
      cost+=Number(correction&&correction.cost||0);
      if(!correction||!correction.success||!String(correction.text||'').trim())throw new Error('Targeted correction failed: '+(correction&&correction.message||'empty result'));
      var candidate=new_addAuthorFooter_(String(correction.text).trim(),material);
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
    // Keep this comparison completely independent of the established
    // ce_compareOriginalToCtById() stage and the Rewrite Comparison Audit sheet.
    if(status==='PASS'){
      resultSheet.getRange(outRow,12).setValue(original); // L: source HTML
      try {
        var comparisonPrompt=[
          'You are independently auditing a UK stone/tile floor restoration webpage rewrite.',
          'Compare full ORIGINAL HTML with REVISED HTML. ORIGINAL is evidence for documented project events.',
          'Do not treat absence of general technical education from the original as a factual error; flag invented project-specific facts.',
          'Assess: Helpful Content, E-E-A-T/first-hand experience, factual fidelity, readability for nontechnical UK homeowners, intent alignment, technical relevance, and likely search usefulness (not ranking predictions).',
          'For Case Studies, the author deliberately prefers first-person singular I, including when original editorial copy says we. Do NOT score or flag we-to-I changes as a regression by themselves. Flag only concrete unsupported claims about who attended, performed or supervised particular work; retain I as the default unless the project explicitly documents a team effort.',
          'Distinguish documented completed outcomes from generic intended benefits; flag unsupported site observations, lighting comparisons, measurements, customer reactions, and unverified performance as factual regressions.',
          'Detect diluted client details, invented outcomes, SEO keyword stuffing, templated phrasing, image changes, heading changes and useful improvements.',
          'For each important observation provide short concrete quotations or clear examples from both versions.',
          'Return only plain text with these EXACT headings: VERDICT; SCORES (original/revised for Helpful, Experience, Factual, Readability, Intent, Technical, Search); IMPROVEMENTS; REGRESSIONS; DOCUMENTED PROJECT FACTS AT RISK; RECURRING SCRIPT-RULE CANDIDATES; SPECIFIC W2B RULE RECOMMENDATIONS.',
          'Never assume longer is better. Label unverified concerns as needing verification.',
          'POST ID: '+postId,'TITLE: '+title,'ARTICLE TYPE: '+articleType,
          '=== ORIGINAL HTML ('+postId+') ===',original,
          '=== REVISED HTML ('+postId+') ===',html
        ].join('\\n\\n');
        var comparison=bc_sendPromptViaOpenAI(comparisonPrompt,5500,MODEL_CHEAP);
        cost+=Number(comparison&&comparison.cost||0);
        if(!comparison||!comparison.success||!String(comparison.text||'').trim())
          throw new Error('AI comparison failed: '+(comparison&&comparison.message||'no response'));
        var comparisonText=String(comparison.text).trim();
        if(comparisonText.length>49000)throw new Error('Comparison report exceeds Sheets cell limit.');
        resultSheet.getRange(outRow,13).setValue(comparisonText); // M: analysis
      }catch(comparisonError){
        // A comparison failure never turns a passed publication audit into FAIL.
        resultSheet.getRange(outRow,13).setValue('COMPARISON ERROR: '+String(comparisonError.message||comparisonError).slice(0,450));
      }
      resultSheet.getRange(outRow,10).setValue(cost);
    }
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
function new_runFilteredCaseStudyRow(row, includeHidden) {
  row=Number(row);
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if(!sh||!Number.isInteger(row)||row<2||row>sh.getLastRow())throw new Error('Invalid posts row');
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(x){return String(x).trim();});
  function col(name){var c=headers.indexOf(name)+1;if(!c)throw new Error('Missing heading '+name);return c;}
  if(String(sh.getRange(row,col('Article Type')).getValue()).trim()!=='Case Study')throw new Error('Row '+row+' is not a Case Study');
  if(!includeHidden && (sh.isRowHiddenByFilter(row)||sh.isRowHiddenByUser(row)))return {success:true,skipped:true,message:'Row hidden from filter',row:row,cost:0};
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

/**
 * Convert pasted Post IDs to their exact rows on posts.
 * Explicit ID requests run regardless of the current sheet filter.
 * Reject unknown IDs, duplicates in the sheet, and non-Case Study entries.
 */
function new_getCaseStudyRowsForPostIds(postIdsText) {
  var ids=String(postIdsText||'').split(/[\s,;]+/).map(function(x){return x.trim();}).filter(Boolean);
  if(!ids.length)throw new Error('Enter at least one Post ID.');
  if(ids.length>250)throw new Error('Maximum 250 Post IDs per batch.');
  var seen={},requested=[];
  ids.forEach(function(id){if(!seen[id]){seen[id]=true;requested.push(id);}});
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if(!sh)throw new Error('posts sheet missing');
  var h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(function(x){return String(x).trim();});
  var idCol=h.indexOf('Post ID')+1,typeCol=h.indexOf('Article Type')+1;
  if(!idCol||!typeCol)throw new Error('Post ID or Article Type heading missing.');
  var n=sh.getLastRow()-1;
  var idValues=n>0?sh.getRange(2,idCol,n,1).getDisplayValues():[];
  var types=n>0?sh.getRange(2,typeCol,n,1).getDisplayValues():[];
  var byId={};
  idValues.forEach(function(v,i){var id=String(v[0]).trim();if(!id)return;(byId[id]||(byId[id]=[])).push({row:i+2,type:String(types[i][0]).trim()});});
  var errors=[],rows=[];
  requested.forEach(function(id){
    var matches=byId[id]||[];
    if(!matches.length)errors.push('Post ID '+id+' not found');
    else if(matches.length!==1)errors.push('Post ID '+id+' appears in multiple posts rows');
    else if(matches[0].type!=='Case Study')errors.push('Post ID '+id+' is not a Case Study');
    else rows.push({postId:id,row:matches[0].row});
  });
  if(errors.length)throw new Error(errors.join('; '));
  return rows;
}
