/**
 * Original vs CT Audit v1. One Post ID per Apps Script request.
 * site-export!O vs posts!CT. Read-only sources; reports to Rewrite Comparison Audit.
 */
function ce_compareOriginalToCtById(postId) {
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
  if (!original.trim() || !revised.trim()) throw new Error(id + ': missing original O or revised CT.');
  if (original.length > 45000 || revised.length > 45000) throw new Error(id + ': HTML exceeds v1 archive cell limit.');
  var title = String(posts.getRange(postsRow,2).getDisplayValue() || '');
  var articleType = String(posts.getRange(postsRow,8).getDisplayValue() || '');
  var prompt = [
    'You are independently auditing a UK stone/tile floor restoration webpage rewrite.',
    'Compare full ORIGINAL HTML with REVISED HTML. ORIGINAL is evidence for documented project events.',
    'Do not treat absence of general technical education from the original as a factual error; flag invented project-specific facts.',
    'Assess: Helpful Content, E-E-A-T/first-hand experience, factual fidelity, readability for nontechnical UK homeowners, intent alignment, technical relevance, and likely search usefulness (not ranking predictions).',
    'Detect lost first-person accounts, diluted client details, invented outcomes, SEO keyword stuffing, templated phrasing, image changes, heading changes and useful improvements.',
    'For each important observation provide short concrete quotations or clear examples from both versions.',
    'Return only plain text with these EXACT headings: VERDICT; SCORES (original/revised for Helpful, Experience, Factual, Readability, Intent, Technical, Search); IMPROVEMENTS; REGRESSIONS; DOCUMENTED PROJECT FACTS AT RISK; RECURRING SCRIPT-RULE CANDIDATES; SPECIFIC W2B RULE RECOMMENDATIONS.',
    'Never assume longer is better. Label unverified concerns as needing verification.',
    'POST ID: '+id, 'TITLE: '+title, 'ARTICLE TYPE: '+articleType,
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
    [when,id,articleType,title,id+' - Revised html',revised,''],
    [when,id,articleType,title,id+' - AI Comparison Report',report,cost]
  ]);
  return {success:true,id:id,title:title,cost:cost,report:report.slice(0,11500)};
}
