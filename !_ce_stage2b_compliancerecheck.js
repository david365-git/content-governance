/* ============================================================
   ce_Stage2B_ComplianceRecheck.gs
   Scoped re-check logic for W2B Rewrite Brief Compliance Check.
   After a fix cycle, only re-audits items that previously FAILED,
   instead of re-auditing the entire article from scratch.
============================================================ */

function _parseComplianceItems_(rawText) {
  var items = [];
  var body = String(rawText || "").replace(/^(PASS|FAIL)\s*—\s*/, '');
  body = body.replace(/\n*CHECK 2B-BRIEF:\s*(PASS|FAIL)\s*$/i, '');
  var blocks = body.split(/\n(?=RE-ANCHOR ITEM:|DRIFT ITEM:)/);

  blocks.forEach(function(block) {
    var isReAnchor = /^RE-ANCHOR ITEM:/.test(block.trim());
    var isDrift    = /^DRIFT ITEM:/.test(block.trim());
    if (!isReAnchor && !isDrift) return;

    var nameMatch = block.match(/^(?:RE-ANCHOR ITEM|DRIFT ITEM):\s*(.+)/);
    var name = nameMatch ? nameMatch[1].trim() : "";
    if (!name) return;

    var passed;
      if (isReAnchor) {
        passed = /COVERED:\s*YES/i.test(block);
      } else {
        passed = /STATUS:\s*PASS\s*[—-]\s*(ABSENT|CONTEXTUAL|ROUTED)/i.test(block);
      }

    items.push({
      type: isReAnchor ? 'RE-ANCHOR' : 'DRIFT',
      name: name,
      passed: passed,
      block: block.trim()
    });
  });

  return items;
}

function buildRewriteBriefComplianceRecheckPromptW2B_(savedResult) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();

  var html = String(sheet.getRange(row, 189).getValue() || "").trim(); // GG = a previous fix already applied
  if (!html) html = String(sheet.getRange(row, 188).getValue() || "").trim(); // fallback: GF, first run
  if (!html) html = String(sheet.getRange(row, 98).getValue() || "").trim(); // fallback: New HTML
  if (!html) return 'ERROR: No article HTML found.';

  const d = getActiveRowDataMap();
  const rewriteBrief = String(d["Page Rewrite Brief"] || "").trim();

  var allItems = _parseComplianceItems_(savedResult);
  var failingItems = allItems.filter(function(item) { return !item.passed; });

  if (failingItems.length === 0) {
    // Nothing to recheck — treat as full re-run fallback
    return null;
  }

  // Extract section ids referenced by failing items
  var sectionIds = [];
  failingItems.forEach(function(item) {
    var secMatch = item.block.match(/SECTION:\s*(section-[\w-]+)/i);
    if (secMatch) {
      var sid = secMatch[1].trim();
      if (sectionIds.indexOf(sid) === -1) sectionIds.push(sid);
    }
  });

  // Fallback: if a failing item has no section, default to last section
  if (sectionIds.length === 0) {
    var allSectionIds = [];
    var sidRe = /<section[^>]*\bid=["'](section-[\w-]+)["']/gi;
    var sidM;
    while ((sidM = sidRe.exec(html)) !== null) {
      if (allSectionIds.indexOf(sidM[1]) === -1) allSectionIds.push(sidM[1]);
    }
    if (allSectionIds.length > 0) sectionIds.push(allSectionIds[allSectionIds.length - 1]);
  }

  var extractedSections = [];
  sectionIds.forEach(function(sid) {
    var re = new RegExp('<section[^>]*\\bid=["\']' + sid + '["\'][^>]*>[\\s\\S]*?<\\/section>', 'i');
    var m = html.match(re);
    if (m) extractedSections.push(m[0]);
  });

  if (extractedSections.length === 0) {
    return 'ERROR: Could not locate the previously failing section(s) in the current article HTML.';
  }

  var failingItemNames = failingItems.map(function(item) {
    return (item.type === 'RE-ANCHOR' ? 'RE-ANCHOR ITEM: ' : 'DRIFT ITEM: ') + item.name;
  }).join('\n');

  const prompt = `
STAGE 2B — REWRITE BRIEF COMPLIANCE RECHECK (SCOPED — PREVIOUSLY FAILING ITEMS ONLY)
ROLE: Senior UK SEO & Editorial Quality Auditor

You previously audited this article against its Rewrite Brief. The items below FAILED that audit.
The section(s) shown have since been corrected. Re-check ONLY these specific items against
ONLY the section(s) provided — do not re-audit any other part of the article, and do not
introduce new RE-ANCHOR or DRIFT items not listed below.

--- REWRITE BRIEF (for reference only) ---
${rewriteBrief}

--- ITEMS TO RECHECK ---
${failingItemNames}

--- SECTION(S) TO RECHECK AGAINST ---
${extractedSections.join('\n\n')}

--- TASK ---
For each item listed above, state whether the corrected section(s) now resolve the failure.

COVERAGE STANDARD: Coverage may be conceptual or mechanistic, not just verbatim. If the section
explains the mechanism a RE-ANCHOR entity name describes — even without using the governed term
itself — mark it COVERED: YES and quote the sentence that demonstrates the mechanism. Only mark
COVERED: NO where the underlying mechanism is genuinely absent, not merely unnamed.

OUTPUT FORMAT (exactly, one block per item, same order as listed above):
RE-ANCHOR ITEM: [entity/topic text — copy exactly from the list above]
COVERED: YES / NO
EVIDENCE: [short quote or section reference, or "none"]
SECTION: [section id, e.g. section-3]

DRIFT ITEM: [drift term/framing text — copy exactly from the list above]
STATUS: PASS — ABSENT / PASS — CONTEXTUAL / PASS — ROUTED / FAIL — DEVELOPED
EVIDENCE: [short quote or routing/link evidence, or "none"]
SECTION: [section id, e.g. section-2, or "none" if STATUS is PASS — ABSENT]

Last line: "CHECK 2B-BRIEF: PASS" only if every RE-ANCHOR item is COVERED: YES and every DRIFT item is PASS — ABSENT, PASS — CONTEXTUAL, or PASS — ROUTED. Otherwise "CHECK 2B-BRIEF: FAIL".
`.trim();

  return prompt;
}

function saveRewriteBriefComplianceRecheckResultW2B(raw) {
  if (!raw || !raw.trim()) {
    return { success: false, message: 'No response provided.' };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row   = sheet.getActiveRange().getRow();
  if (row < 2) return { success: false, message: 'Select a data row first.' };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h) { return String(h).trim(); });
  var colIdx = headers.indexOf('Rewrite Brief Compliance Check W2B');
  if (colIdx === -1) {
    return { success: false, message: "Column 'Rewrite Brief Compliance Check W2B' not found in posts sheet." };
  }

  var previousResult = String(sheet.getRange(row, colIdx + 1).getValue() || "").trim();
  if (!previousResult) {
    return { success: false, message: 'No previous compliance result found to merge with — run the full check first.' };
  }

  var previousItems = _parseComplianceItems_(previousResult);
  var newItems      = _parseComplianceItems_(raw);

  // Merge: replace previous items with matching new ones; keep the rest unchanged
  var mergedItems = previousItems.map(function(oldItem) {
    var match = newItems.find(function(n) { return n.type === oldItem.type && n.name === oldItem.name; });
    return match || oldItem;
  });

  // Append any genuinely new items not previously tracked (shouldn't normally happen)
  newItems.forEach(function(n) {
    var exists = mergedItems.some(function(m) { return m.type === n.type && m.name === n.name; });
    if (!exists) mergedItems.push(n);
  });

  var allPass = mergedItems.every(function(item) { return item.passed; });

      var failingItemsOnly = mergedItems.filter(function(item) {
        return !item.passed;
      });

      var mergedText = allPass
        ? 'No compliance failures found.'
        : failingItemsOnly.map(function(item) {
            return item.block;
          }).join('\n\n');

      var summary = (allPass ? 'PASS' : 'FAIL') + ' — ' + mergedText +
        '\n\nCHECK 2B-BRIEF: ' + (allPass ? 'PASS' : 'FAIL');

  var cell = sheet.getRange(row, colIdx + 1);
  cell.setNumberFormat('@');
  cell.setValue(summary);

  return {
    success: true,
    passed: allPass,
    message: allPass
      ? 'All previously failing items now pass — article body complies with Rewrite Brief. Proceed to W3 humanisation.'
      : 'Some items still fail — saved for review. Re-run fix on remaining failures.',
    rawResult: summary
  };
}