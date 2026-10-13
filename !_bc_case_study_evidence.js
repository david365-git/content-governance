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
  if (oldStatus === 'SOURCE-QUOTED | sha256:' + digest) {
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
          if (!fact || !quote || normalizedSource.indexOf(normalizedQuote) < 0) {
            throw new Error('Unverifiable quote in ' + label + ': ' + quote.slice(0,120));
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
  // Fail closed: no partial evidence is written on model error or quote mismatch.
  for (var j=0;j<labels.length;j++) col(labels[j]);
  sh.getRange(row, col(labels[0]), 1, labels.length).setValues([output]);
  statusCell.setValue('SOURCE-QUOTED | sha256:' + digest);
  return {success:true, skipped:false, cost:totalCost, message:'Evidence quotations matched to source (not independently fact-checked)'};
}
function bc_caseEvidenceActiveRow() {
  return bc_caseEvidenceExtract_(SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getRow());
}
