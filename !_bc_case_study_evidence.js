/**
 * Content Intelligence: source-quoted case-study evidence, GL:GV.
 * Evidence quotes must match original HTML in site-export Column O, by Post ID.
 */
function bc_caseEvidenceDigest_(source) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source, Utilities.Charset.UTF_8);
  return bytes.map(function(b){ return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function bc_caseEvidenceText_(html) {
  return String(html || '').replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ').replace(/&nbsp;|&#160;/gi, ' ')
    .replace(/&amp;/gi, '&').replace(/&quot;/gi, '"')
    .replace(/&apos;|&#39;/gi, "'").replace(/\s+/g, ' ').trim();
}
/**
 * Recover a quotation only when the source contains the same contiguous words.
 * Accept harmless case differences and an extra leading "the/a/an"; never
 * fabricate missing words or use a semantic/fuzzy match as source evidence.
 */
function bc_caseEvidenceRepairQuote_(quote, sourceText) {
  var clean=function(v){return String(v||'').replace(/[’‘]/g,"'").replace(/[“”]/g,'"').replace(/\s+/g,' ').trim();};
  var src=clean(sourceText), q=clean(quote);
  var candidates=[q];
  var withoutArticle=q.replace(/^(?:the|a|an)\s+/i,'');
  if(withoutArticle!==q && withoutArticle.split(/\s+/).length>=3)candidates.push(withoutArticle);
  for(var i=0;i<candidates.length;i++){
    var part=candidates[i];
    if(part.length<12)continue;
    var at=src.toLowerCase().indexOf(part.toLowerCase());
    if(at>=0)return src.slice(at,at+part.length);
  }
  return null;
}
var BC_CASE_EVIDENCE_VERSION_ = 'v3';
function bc_caseEvidenceStatus_(digest) { return 'SOURCE-QUOTED-' + BC_CASE_EVIDENCE_VERSION_ + ' | sha256:' + digest; }
function bc_caseEvidenceExtract_(row) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if (!sh) throw new Error('posts sheet missing');
  row = Number(row);
  if (!row || row < 2) throw new Error('Valid posts row required');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  function col(name) { var i = headers.indexOf(name); if (i < 0) throw new Error('Missing heading: ' + name); return i + 1; }
  var type = String(sh.getRange(row, col('Article Type')).getValue() || '').trim();
  if (type !== 'Case Study') return { success:true, skipped:true, message:'Not a Case Study', cost:0 };
  var postId = String(sh.getRange(row, col('Post ID')).getDisplayValue() || '').trim();
  if (!postId) throw new Error('Post ID missing in posts row ' + row);
  var exportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('site-export');
  if (!exportSheet) throw new Error('site-export sheet missing');
  var exportLastRow = exportSheet.getLastRow();
  var ids = exportLastRow > 1 ? exportSheet.getRange(2, 1, exportLastRow - 1, 1).getDisplayValues() : [];
  var matchedRow = -1;
  for (var k = 0; k < ids.length; k++) {
    if (String(ids[k][0] || '').trim() === postId) {
      if (matchedRow !== -1) throw new Error('Duplicate Post ID ' + postId + ' in site-export');
      matchedRow = k + 2;
    }
  }
  if (matchedRow === -1) throw new Error('Post ID ' + postId + ' not found in site-export');
  var source = String(exportSheet.getRange(matchedRow, 15).getValue() || ''); // site-export column O: Full Post HTML
  if (!source.trim()) {
    sh.getRange(row, col('Evidence Status')).setValue('SOURCE REQUIRED — site-export O is empty');
    return {success:false, cost:0, message:'SOURCE REQUIRED — site-export Full Post HTML missing'};
  }
  var digest = bc_caseEvidenceDigest_(source);
  var statusCell = sh.getRange(row, col('Evidence Status'));
  var oldStatus = String(statusCell.getValue() || '');
  if (oldStatus === bc_caseEvidenceStatus_(digest)) {
    return {success:true, skipped:true, cost:0, message:'Source-quoted evidence unchanged'};
  }
  var labels = ['Project Location & Property','Observed Conditions','Client Requirements',
    'Materials & Construction','Inspection & Tests','Restoration Methods',
    'Products & Equipment','Technical Findings','Project Results','Maintenance Handover'];
  var prompt = [
    'Extract ONLY documented facts from this original case-study article.',
    'Return ONLY a valid JSON object with exactly these keys: ' + JSON.stringify(labels) + '.',
    'Each key maps to an array of objects {"fact":"concise source-supported fact","quote":"exact contiguous quotation from the visible original article"}',
    'Every quote MUST be copied exactly from the source visible text. Never infer a performed test, cause, property detail, client conversation, or project outcome.',
    'If a category has no documented facts, return an empty array. Inspection & Tests must not infer test results from general professional practice.',
    'Distinguish original construction materials from repair products, and maintain documented chronological order for methods.',
    'Treat statements of general advice as advice, not work performed.',
    'CLASSIFICATION AND CERTAINTY — MANDATORY:',
    'Project Location & Property: actual property facts only; exclude generic council, postcode and regional geography that does not describe this property.',
    'Observed Conditions: what was actually seen or recorded at this site; do not place treatment decisions here.',
    'Client Requirements: ONLY requests or preferences explicitly attributed to the client. Do not turn an author-described restoration aim into a client request. Use [] when none is documented.',
    'Materials & Construction: ONLY documented installed materials and construction; exclude hypothetical moisture paths, restoration systems, future sealers and generic technical principles.',
    'Inspection & Tests: separate what was inspected or CONSIDERED from findings actually confirmed; never imply a test occurred merely because it would be normal practice.',
    'Restoration Methods: actions actually performed, in documented order; conditional treatments must remain conditional, never imply they always occurred.',
    'Products & Equipment: products/tools documented as used, not items merely recommended or hypothetical.',
    'Technical Findings: actual on-site findings, meaningful negative findings and explicitly supported professional decisions; exclude unverified mechanisms and general explanations.',
    'Project Results: ONLY documented results and limitations; distinguish observed outcome from generic benefits.',
    'Maintenance Handover: include documented maintenance recommendations as RECOMMENDATIONS; only say a personal handover occurred when explicitly recorded.',
    'Preserve meaningful negative results (e.g. no major adhesive field), declined or unnecessary treatment, and differences between considered possibilities and confirmed findings.',
    'DISTINCT METHOD RETENTION: Extract each separately documented completed restoration step, including old sealer or coating stripping, even if a broader cleaning or sealing summary is already recorded. Do not collapse stripping, cleaning, rinsing, repairs and sealing into one fact if doing so loses a distinct actual operation.',
    'REPAIR OPTIONS CONSIDERED: Retain documented repair alternatives and assessment decisions, including replacement-tile matching considered but not performed. Label options as considered, not completed. Distinguish local resetting from replacement and preserve why a more extensive repair was unnecessary where documented.',
    'FACT–QUOTE ENTAILMENT: every fact must be directly supported by its own quote in meaning, certainty, attribution and action. An exact copied quote is NOT sufficient when the fact changes its meaning.',
    'Do not repeat generic advice or hypothetical examples as project findings. Do not invent a fact to fill any category. Use [] for categories with no actual applicable evidence.',
    'Each category must be accurate and concise; retain material distinct findings without keyword repetition.',
    'ORIGINAL ARTICLE:\n' + bc_caseEvidenceText_(source)
  ].join('\n');
  var normalizedSource = bc_caseEvidenceText_(source).replace(/[’‘]/g,"'").replace(/[“”]/g,'"');
  var output = null;
  var totalCost = 0;
  var lastError = '';
  for (var attempt = 0; attempt < 2; attempt++) {
    var retryPrompt = attempt === 0 ? prompt : prompt + '\nVALIDATION CORRECTION: Your previous response contained this quotation which is not an exact substring of ORIGINAL ARTICLE: ' + lastError + '\nRegenerate the complete JSON. Copy quotations character-for-character from ORIGINAL ARTICLE. Do not paraphrase, omit words, or change punctuation inside quotes.';
    var result = bc_sendPromptViaOpenAI(retryPrompt, 5000, MODEL_CHEAP);
    if (!result || !result.success) throw new Error('Evidence extraction API failed: ' + (result && result.message || 'unknown error'));
    totalCost += Number(result.cost || 0);
    var raw = String(result.text || '').trim().replace(/^\x60\x60\x60(?:json)?\s*/i,'').replace(/\s*\x60\x60\x60$/,'');
    var parsed;
    try { parsed = JSON.parse(raw); } catch(e) {
      lastError = 'Invalid JSON: ' + e.message;
      continue;
    }
    try {
      var proposed = [];
      labels.forEach(function(label) {
        var items = parsed[label];
        if (!Array.isArray(items)) throw new Error('Missing evidence array: ' + label);
        var checked = [];
        items.forEach(function(item) {
          var fact = String(item && item.fact || '').trim();
          var quote = String(item && item.quote || '').replace(/\s+/g,' ').trim();
          var normalizedQuote = quote.replace(/[’‘]/g,"'").replace(/[“”]/g,'"');
          if (!fact || !quote) {
            throw new Error('Incomplete evidence in ' + label);
          }
          // Preserve exact matches; only repair case or an extra leading article
          // by copying the real contiguous phrase back from the source.
          if (normalizedSource.indexOf(normalizedQuote) < 0) {
            var repaired=bc_caseEvidenceRepairQuote_(quote,normalizedSource);
            if (!repaired)throw new Error('Unverifiable quote in ' + label + ': ' + quote.slice(0,120));
            quote=repaired;
          }
          checked.push({fact:fact,quote:quote});
        });
        proposed.push(JSON.stringify(checked));
      });
      output = proposed;
      break;
    } catch(e) { lastError = e.message; }
  }
  if (!output) throw new Error('Evidence extraction failed after 2 attempts: ' + lastError);
  // Second-pass semantic validation within extraction, not an extra pipeline stage.
  // Exact quote matching alone cannot establish correct category or certainty.
  var reviewPrompt = [
    'You are an independent case-study evidence verifier. Inspect the SOURCE and the EXTRACTED JSON.',
    'Return ONLY a JSON object with exactly these same ten category keys and arrays of {fact,quote}.',
    'For every item, verify: (1) its quote directly entails the fact, (2) its category is appropriate,',
    '(3) its certainty and attribution have not been strengthened, and (4) it concerns this job rather than general advice.',
    'Drop unsupported, generic geographical padding and misleading inferred client requirements.',
    'Keep negative findings and distinct material job details. Separate considered conditions from confirmed discoveries.',
    'Do not remove unique documented completed steps while deduplicating broader method summaries: preserve old sealer stripping separately from cleaning and applying new sealer.',
    'Retain documented repair options considered (such as replacement-tile matching), clearly as considerations rather than completed repairs.',
    'Only client-attributed requests belong in Client Requirements.',
    'General maintenance RECOMMENDATIONS may appear in Maintenance Handover, but do not imply a personal handover.',
    'Materials & Construction must not include theoretical moisture paths or future protective treatments.',
    'Keep conditional methods conditional. Never infer tests, chemicals used, owner statements or outcomes.',
    'You may move an existing fact to a better category or reword it to match its original quote,',
    'but NEVER invent quotes or facts. Preserve exact contiguous source quotations.',
    'If a category has no supported items, use [].',
    'ORIGINAL ARTICLE:\n' + bc_caseEvidenceText_(source),
    'EXTRACTED JSON:\n' + JSON.stringify(JSON.parse('{' + labels.map(function(label,i){
      return JSON.stringify(label)+':'+output[i];
    }).join(',') + '}'))
  ].join('\n');
  var validatedOutput = null;
  var reviewError = '';
  for (var reviewAttempt = 0; reviewAttempt < 2; reviewAttempt++) {
    var reviewApi = bc_sendPromptViaOpenAI(reviewPrompt + (reviewAttempt ? '\nPRIOR OUTPUT INVALID: ' + reviewError + '. Return valid, source-supported JSON only.' : ''), 5000, MODEL_CHEAP);
    if (!reviewApi || !reviewApi.success) throw new Error('Evidence semantic validation API failed: ' + (reviewApi && reviewApi.message || 'unknown'));
    totalCost += Number(reviewApi.cost || 0);
    try {
      var reviewRaw=String(reviewApi.text||'').trim().replace(/^\x60\x60\x60(?:json)?\s*/i,'').replace(/\s*\x60\x60\x60$/,'');
      var reviewParsed=JSON.parse(reviewRaw);
      var checkedReview=[];
      labels.forEach(function(label) {
        if (!Array.isArray(reviewParsed[label])) throw new Error('Invalid reviewed category: '+label);
        var items=reviewParsed[label].map(function(item) {
          var fact=String(item && item.fact || '').trim();
          var quote=String(item && item.quote || '').replace(/\s+/g,' ').trim();
          if (!fact||!quote) throw new Error('Incomplete reviewed fact in '+label);
          var nq=quote.replace(/[’‘]/g,"'").replace(/[“”]/g,'"');
          if (normalizedSource.indexOf(nq)<0) {
            var repaired=bc_caseEvidenceRepairQuote_(quote,normalizedSource);
            if (!repaired) throw new Error('Unsupported reviewed quote in '+label+': '+quote.slice(0,90));
            quote=repaired;
          }
          return {fact:fact,quote:quote};
        });
        checkedReview.push(JSON.stringify(items));
      });
      validatedOutput=checkedReview;
      break;
    } catch(e) { reviewError=e.message; }
  }
  if (!validatedOutput) throw new Error('Evidence semantic validation failed: '+reviewError);
  output=validatedOutput;
  // Fail closed: no partial evidence is written on model error or quote mismatch.
  for (var j=0;j<labels.length;j++) col(labels[j]);
  sh.getRange(row, col(labels[0]), 1, labels.length).setValues([output]);
  statusCell.setValue(bc_caseEvidenceStatus_(digest));
  return {success:true, skipped:false, cost:totalCost, message:'Evidence quotations checked and categories independently reviewed against source'};
}
function bc_caseEvidenceActiveRow() {
  return bc_caseEvidenceExtract_(SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getRow());
}

/**
 * Shared read-only evidence block for W1.5 and W2B.
 * Do not trust a prior extraction unless its original HTML fingerprint matches.
 */
function bc_caseEvidencePromptBlock_(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('posts');
  row = Number(row);
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  function c(name) { var idx=headers.indexOf(name); if(idx<0) throw new Error('Evidence column missing: '+name); return idx+1; }
  if (String(sh.getRange(row,c('Article Type')).getValue()).trim() !== 'Case Study') return '';
  var id=String(sh.getRange(row,c('Post ID')).getDisplayValue()).trim();
  var ex=ss.getSheetByName('site-export');
  var ids=ex.getRange(2,1,Math.max(ex.getLastRow()-1,1),1).getDisplayValues();
  var match=-1;
  ids.forEach(function(v,i){if(String(v[0]).trim()===id){if(match>=0)throw new Error('Duplicate site-export Post ID '+id);match=i+2;}});
  if(match<0)throw new Error('No site-export source for Post ID '+id);
  var original=String(ex.getRange(match,15).getValue()||'');
  if(!original.trim())throw new Error('Original case study HTML is empty in site-export O');
  var status=String(sh.getRange(row,c('Evidence Status')).getValue()||'');
  if(status!==bc_caseEvidenceStatus_(bc_caseEvidenceDigest_(original))){
    throw new Error('Case Study Evidence is missing or out of date. Run Extract Case Study Evidence in Content Intelligence first.');
  }
  var labels=['Project Location & Property','Observed Conditions','Client Requirements','Materials & Construction',
    'Inspection & Tests','Restoration Methods','Products & Equipment','Technical Findings','Project Results','Maintenance Handover'];
  var lines=[];
  labels.forEach(function(label){
    var raw=String(sh.getRange(row,c(label)).getValue()||'[]');
    var items=JSON.parse(raw);
    if(!Array.isArray(items))throw new Error('Invalid case-study evidence in '+label);
    var entries=items.map(function(x){
      if(!x||!x.fact||!x.quote)throw new Error('Incomplete case-study evidence in '+label);
      return '- '+String(x.fact)+' [SOURCE: '+String(x.quote)+']';
    });
    lines.push(label+':\n'+(entries.length?entries.join('\n'):'(Not documented)'));
  });
  return '\nCASE STUDY PROJECT EVIDENCE (source-quoted, not independently fact-checked):\n'
    +'All project-specific claims must stay within the evidence below and confirmed owner corrections. Do not turn general technical notes into observations or performed tests. An empty category does not prove an activity never happened. Governed headings and SEO cannot add project facts.\n'
    +lines.join('\n')+'\nEND CASE STUDY EVIDENCE\n';
}


/**
 * W0 filtered Stone Type batch. Only visible Case Studies for one stone type
 * are returned; completed evidence is not scheduled again.
 * The row extraction remains source-fingerprint aware for safety.
 */
function bc_caseEvidenceFilteredRows() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if (!sh) return {success:false,message:'posts sheet missing'};
  // Supports both standard Sheets filters and the custom Stone Type filter that hides rows.
  var last = sh.getLastRow();
  if (last < 2) return {success:true,stoneType:'Filtered selection',rows:[],alreadyCompleted:0};
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  function col(name) { var idx=headers.indexOf(name); if(idx<0) throw new Error('Missing heading: '+name); return idx+1; }
  var typeCol=col('Article Type'), stoneCol=col('Stone Type'), statusCol=col('Evidence Status');
  var types=sh.getRange(2,typeCol,last-1,1).getDisplayValues();
  var stones=sh.getRange(2,stoneCol,last-1,1).getDisplayValues();
  var statuses=sh.getRange(2,statusCol,last-1,1).getDisplayValues();
  var rows=[], completed=0, stoneTypes={}, hidden=0;
  for (var i=0;i<last-1;i++) {
    var row=i+2;
    if (sh.isRowHiddenByFilter(row) || sh.isRowHiddenByUser(row)) { hidden++; continue; }
    var stone=String(stones[i][0]||'').trim();
    if (stone) stoneTypes[stone]=true;
    if (String(types[i][0]||'').trim()!=='Case Study') continue;
    // Include completed rows: extraction checks the source hash and safely skips unchanged evidence.
    rows.push(row);
  }
  var names=Object.keys(stoneTypes);
  if (!hidden || names.length!==1) return {success:false,message:'Filter the posts sheet to exactly one Stone Type first. Visible types: '+(names.join(', ')||'none')};
  return {success:true,stoneType:names[0],rows:rows,alreadyCompleted:completed};
}
/** Public endpoint for client-side sequential batching. */
function bc_caseEvidenceRunRowBatch(row) {
  return bc_caseEvidenceExtract_(row);
}

/** Direct Row Analysis menu command: processes visible Case Studies, skipping unchanged evidence. */
function bc_caseEvidenceFilteredBatch() {
  var info = bc_caseEvidenceFilteredRows();
  var ui = SpreadsheetApp.getUi();
  if (!info.success) { ui.alert(info.message); return; }
  var started = Date.now(), done = 0, skipped = 0, failed = [], cost = 0, remaining = 0;
  for (var i=0; i<info.rows.length; i++) {
    if (Date.now()-started > 240000) { remaining = info.rows.length-i; break; }
    var row = info.rows[i];
    try {
      var result = bc_caseEvidenceExtract_(row);
      cost += Number(result.cost || 0);
      if (result.skipped) skipped++;
      else if (result.success) done++;
      else failed.push('Row '+row+': '+(result.message||'Unsuccessful'));
    } catch (e) { failed.push('Row '+row+': '+(e.message||e)); }
  }
  ui.alert('Case Study Evidence — '+info.stoneType+'\n'
    +'Extracted: '+done+'\nAlready complete: '+skipped
    +'\nFailed: '+failed.length+'\nRemaining: '+remaining
    +'\nAPI cost: $'+cost.toFixed(4)
    +(failed.length?'\n'+failed.slice(0,5).join('\n'):'')
    +'\nRerun to continue; unchanged evidence is skipped.');
}

/** Open a live, modeless batch-progress window for the current Stone Type. */
function bc_caseEvidenceOpenBatchProgress() {
  var html = HtmlService.createHtmlOutputFromFile('bc_CaseEvidenceBatchProgress')
    .setWidth(450).setHeight(340);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Case Study Evidence — Batch');
}
