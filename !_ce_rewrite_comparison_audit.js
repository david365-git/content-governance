/**
 * Original vs CT Audit v1. One Post ID per Apps Script request.
 * site-export!O vs posts!CT. Read-only sources; reports to Rewrite Comparison Audit.
 */
/**
 * Capture the content of CT immediately after W8E finishes, before later image insertion.
 * One latest snapshot per Post ID; no mutation of posts or original export.
 */
function ce_saveW8EHtmlSnapshot_(postsSheet, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = String(postsSheet.getRange(row,4).getDisplayValue() || '').trim();
  var html = String(postsSheet.getRange(row,98).getValue() || '');
  if (!/^\d+$/.test(id) || !html.trim()) throw new Error('W8E snapshot requires a Post ID and nonempty CT HTML.');
  var sh = ss.getSheetByName('W8E HTML Snapshots');
  if (!sh) {
    sh = ss.insertSheet('W8E HTML Snapshots');
    sh.getRange(1,1,1,3).setValues([['Post ID','W8E completed at','W8E CT HTML snapshot']]);
    sh.setFrozenRows(1);
  }
  var hit = sh.getLastRow() > 1
    ? sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext()
    : null;
  var targetRow = hit ? hit.getRow() : sh.getLastRow()+1;
  sh.getRange(targetRow,1,1,3).setValues([[id,new Date(),html]]);
}

function ce_compareOriginalToCtById(postId, compareColumn) {
  var selectedColumn = String(compareColumn || 'CT').trim().toUpperCase();
  if (selectedColumn !== 'CT' && selectedColumn !== 'W8E') throw new Error('Choose CT or W8E snapshot for the comparison.');
  var id = String(postId || '').trim();
  if (!/^\d+$/.test(id)) throw new Error('Enter a numeric Post ID.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName('site-export'), posts = ss.getSheetByName('posts');
  if (!source || !posts) throw new Error('Missing site-export or posts sheet.');
  function getRow(sh, column) {
    if (sh.getLastRow() < 2) return 0;
    var hit = sh.getRange(2, column, sh.getLastRow()-1, 1)
      .createTextFinder(id).matchEntireCell(true).findNext();
    return hit ? hit.getRow() : 0;
  }
  var sourceRow = getRow(source, 1), postsRow = getRow(posts, 4);
  if (!sourceRow || !postsRow) throw new Error('Post ' + id + ' missing from ' + (!sourceRow ? 'site-export' : 'posts') + '.');
  var original = String(source.getRange(sourceRow,15).getValue() || '');
  var revised = String(posts.getRange(postsRow,98).getValue() || '');
  if (selectedColumn === 'W8E') {
    var archive = ss.getSheetByName('W8E HTML Snapshots');
    if (!archive || archive.getLastRow() < 2) throw new Error('No W8E HTML snapshot is saved yet. Run W8E with the updated script first.');
    var match = archive.getRange(2,1,archive.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();
    if (!match) throw new Error('Post '+id+' has no saved W8E snapshot. Run W8E with the updated script first.');
    revised = String(archive.getRange(match.getRow(),3).getValue() || '');
  }
  if (!original.trim() || !revised.trim()) throw new Error(id + ': missing original O or revised '+selectedColumn+'.');
  if (original.length > 45000 || revised.length > 45000) throw new Error(id + ': HTML exceeds v1 archive cell limit.');
  var title = String(posts.getRange(postsRow,2).getDisplayValue() || '');
  var articleType = String(posts.getRange(postsRow,8).getDisplayValue() || '');
  // Read live article-type coverage and global rules from the NEW matrix.
  // Fail closed rather than silently applying the retired editorial framework.
  var coverage = getTierStructuralRequirements(
    String(posts.getRange(postsRow,7).getDisplayValue() || ''), '', articleType
  );
  if (!coverage || !coverage.length) throw new Error('No NEW matrix coverage for article type: '+articleType);
  var globals = ce_getNewTierGlobalGovernanceBlock_();
  var coverageText = coverage.map(function(r) {
    return 'Order '+r.order+' — '+r.name+' | Priority: '+r.priority+
      ' | When applicable: '+r.whenApplicable+' | Depth: '+r.depth+
      '\\nGuidance: '+r.guidance+'\\nEvidence: '+r.evidence+
      '\\nSafeguards: '+r.safeguards+'\\nLinking: '+r.linking;
  }).join('\\n\\n');
  var prompt = [
    'You are independently auditing a UK stone/tile floor restoration webpage rewrite.',
    'Compare full ORIGINAL HTML with REVISED HTML. ORIGINAL is evidence for documented project events.',
    'NEW MATRIX GOVERNANCE OVERRIDES OLD EDITORIAL ASSUMPTIONS. Assess the chosen rewrite against the actual Article Type requirements below. Required functions must appear; conditional functions apply only when supported. Assess depth by reader need and available evidence, NOT word counts, equal section lengths or predetermined quotas.',
    'For UK homeowners, prefer clear practical floor restoration guidance over unnecessary scientific or chemical complexity. Do not demand detail that would not change diagnosis, treatment, safety or outcome.',
    'Do not penalise an article for using the NEW matrix content order or headings or for retaining essential detail that might appear repetitive. Identify only meaningful, harmful repetition.',
    'Preserve real project chronology, methods, observed conditions, client decisions, outcomes and useful media and links. Report omissions or unsupported additions as evidence regressions. Do not invent facts to satisfy coverage functions.',
    globals,
    'NEW MATRIX ARTICLE TYPE COVERAGE:',
    coverageText,
    'Do not treat absence of general technical education from the original as a factual error; flag invented project-specific facts.',
    'Assess: Helpful Content, E-E-A-T/first-hand experience, factual fidelity, readability for nontechnical UK homeowners, intent alignment, technical relevance, and likely search usefulness (not ranking predictions).',
    'For Case Studies, the author deliberately prefers first-person singular I, including when original editorial copy says we. Do NOT score or flag we-to-I changes as a regression by themselves. Flag only concrete unsupported claims about who attended, performed or supervised particular work; retain I as the default unless the project explicitly documents a team effort.',
    'Distinguish documented completed outcomes from generic intended benefits; flag unsupported site observations, lighting comparisons, measurements, customer reactions, and unverified performance as factual regressions.',
    'Detect diluted client details, invented outcomes, SEO keyword stuffing, templated phrasing, image changes, heading changes and useful improvements.',
    'For each important observation provide short concrete quotations or clear examples from both versions.',
    'Return only plain text with these EXACT headings: VERDICT; SCORES (original/revised for Helpful, Experience, Factual, Readability, Intent, Technical, Search); IMPROVEMENTS; REGRESSIONS; DOCUMENTED PROJECT FACTS AT RISK; RECURRING SCRIPT-RULE CANDIDATES; SPECIFIC W2B RULE RECOMMENDATIONS.',
    'Never assume longer is better. Label unverified concerns as needing verification.',
    'POST ID: '+id, 'TITLE: '+title, 'ARTICLE TYPE: '+articleType, 'REWRITE SOURCE COLUMN: '+selectedColumn,
    '=== ORIGINAL HTML ('+id+') ===', original,
    '=== REVISED HTML ('+id+') ===', revised
  ].join('\n\n');
  var answer = bc_sendPromptViaOpenAI(prompt, 5500, MODEL_CHEAP);
  if (!answer || !answer.success || !String(answer.text || '').trim())
    throw new Error('AI comparison failed: '+(answer && answer.message || 'no response'));
  var report = String(answer.text).trim();
  var target = ss.getSheetByName('Rewrite Comparison Audit');
  if (!target) {
    target = ss.insertSheet('Rewrite Comparison Audit');
    target.appendRow(['Timestamp','Post ID','Article Type','Title','Record','HTML / Report','AI Cost']);
    target.setFrozenRows(1);
    target.setColumnWidth(5,200);
    target.setColumnWidth(6,700);
  }
  var when = new Date(), cost=Number(answer.cost||0);
  target.getRange(target.getLastRow()+1,1,3,7).setValues([
    [when,id,articleType,title,id+' - Original html',original,''],
    [when,id,articleType,title,id+' - Revised '+selectedColumn+' html',revised,''],
    [when,id,articleType,title,id+' - AI Comparison '+selectedColumn+' Report',report,cost]
  ]);
  return {success:true,id:id,title:title,column:selectedColumn,cost:cost,report:report.slice(0,11500)};
}
