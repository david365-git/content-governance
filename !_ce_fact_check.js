/**
 * ce_FactCheck.gs - W2B.05 FACT CHECK AGAINST SOURCE ARTICLE
 * Compares W2B/New HTML output against the original site-export article
 * to catch fabricated facts and unsupported embellishments introduced
 * during the AI rewrite pipeline.
 */

function buildFactCheckPrompt(sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const requestedCol = (sourceColumn === 'CT') ? 'CT' : 'EU';

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  const colPostId = headers.indexOf('Post ID');
  const colTitle = 1; // column B

  if (colPostId === -1) {
    return { success: false, message: "ERROR: Post ID column not found." };
  }

  const rowData = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const postId = String(rowData[colPostId] || '').trim().replace(/\.0$/, '');
  const title = String(rowData[colTitle] || '').trim();

  const rowMap = {};
  headers.forEach(function(h, i) {
    rowMap[h] = rowData[i];
  });

  const governedMaterial = String(
    rowMap["Stone Type"] ||
    rowMap["Material"] ||
    ""
  ).trim();

  const governedLocality = String(
    rowMap["Locality"] ||
    rowMap["Location"] ||
    ""
  ).trim();

  const articleType = String(
    rowMap["Article Type"] ||
    ""
  ).trim();

  const rewriteBrief = String(
    rowMap["Page Rewrite Brief"] ||
    ""
  ).trim();

  const finalSectionPlan = String(
    rowMap["W1.5D Final Plan"] ||
    rowMap["Final Plan"] ||
    ""
  ).trim();

  const authorityBrief = String(
    rowMap["Authority Brief"] ||
    ""
  ).trim();

  const feedsHub = String(
    rowMap["Feeds Hub"] ||
    ""
  ).trim();

  const safeHandoffPages = String(
    rowMap["Safe Handoff Pages"] ||
    ""
  ).trim();

  // Automatically recheck a previously saved fix if one exists — GE holds
  // the W2B.05 Fixed HTML from an earlier fact-check fix pass. Checking GE
  // first (regardless of which source column was originally audited) means
  // a second "Generate Fact-Check Prompt" click always verifies the actual
  // fix, not the untouched original — no manual column-switching needed.
  const geHtml = String(sh.getRange(row, 187).getValue() || '').trim();
  const targetCol = geHtml ? 'GE' : requestedCol;
  const generatedHtml = geHtml || String(sh.getRange(requestedCol + row).getValue() || '').trim();

  if (!postId) {
    return { success: false, message: "ERROR: Post ID is empty for this row." };
  }
  if (!generatedHtml) {
    return { success: false, message: "ERROR: Column " + targetCol + " is empty for this row." };
  }

  // Look up original source HTML from site-export
  let originalHtml = '';
  try {
    const exportSheet = ss.getSheetByName('site-export');
    if (!exportSheet) {
      return { success: false, message: "ERROR: 'site-export' sheet not found." };
    }
    const exportData = exportSheet.getDataRange().getValues();
    const exportHeaders = exportData[0].map(function(h) { return String(h).trim(); });
    const exportIdCol = exportHeaders.indexOf('ID');
    const exportHtmlCol = exportHeaders.indexOf('Full Post HTML');
    if (exportIdCol === -1 || exportHtmlCol === -1) {
      return { success: false, message: "ERROR: Could not find ID or Full Post HTML columns in site-export." };
    }
    for (let i = 1; i < exportData.length; i++) {
      if (String(exportData[i][exportIdCol]).trim().replace(/\.0$/, '') === postId) {
        originalHtml = String(exportData[i][exportHtmlCol] || '').trim();
        break;
      }
    }
  } catch (e) {
    return { success: false, message: "ERROR reading site-export: " + e.message };
  }

  if (!originalHtml) {
    return { success: false, message: "ERROR: No original article found in site-export for Post ID: " + postId };
  }

  const prompt =
`You are performing W2B.05 PROJECT FACT CHECK.

PURPOSE:
Check whether the rewritten case study has invented, contradicted or materially exaggerated facts about the ACTUAL documented project.

This stage is NOT a general technical fact-check, SEO check, internal-link check, heading check or governance-compliance check. Other stages handle those jobs.

PAGE CONTEXT:
MATERIAL: ${governedMaterial || "Not supplied"}
ARTICLE TYPE: ${articleType || "Not supplied"}
LOCALITY: ${governedLocality || "Not supplied"}

SHORT REWRITE SCOPE:
${rewriteBrief || "No active Rewrite Brief supplied."}

ORIGINAL SOURCE ARTICLE — GROUND TRUTH FOR WHAT ACTUALLY HAPPENED:
${originalHtml}

REWRITTEN ARTICLE — CHECK PROJECT CLAIMS IN THIS VERSION:
${generatedHtml}

CORE RULE — HARD LOCK:
The ORIGINAL SOURCE ARTICLE is authoritative for historical project facts, including:
- what the client/homeowner did or said;
- what condition was actually found;
- what work was actually carried out;
- specific tools, products, chemicals, quantities, grit sizes or timings actually used;
- what the restorer actually advised the client;
- what result was actually observed.

Judge meaning, NOT exact wording.

PASS reasonable semantic equivalents.
Examples:
- "cleaned the floor" -> "scrubbed the floor"
- "removed dirt" -> "lifted embedded soil"
- "polished the floor" -> "refined the finish"

FAIL only when the rewrite introduces a genuinely new or contradictory PROJECT fact.
Examples:
- source says "cleaned" but rewrite says "steam cleaned";
- source says "honed progressively" but rewrite invents exact grit sizes;
- source says guidance was given, but rewrite invents specific advice the client was told;
- rewrite invents a customer action, property condition, cause, duration, measurement or observed result.

DO NOT FLAG:
- ordinary paraphrasing that preserves the same factual meaning;
- general technical explanation that is clearly presented as general explanation rather than something that happened on this job;
- headings, internal links, CTA wording, SEO wording or general maintenance signposting merely because they were not in the old article;
- approved explanatory material that does not turn into a claim that it happened on this project.

IMPORTANT DISTINCTION:
"Terrazzo can be damaged by unsuitable chemicals." = general explanation; do not fact-check here.
"I warned the client not to use unsuitable chemicals." = project event; must be supported by the original source.

Likewise:
"Routine dust control can help reduce abrasive soil." = general guidance; do not fact-check here.
"I told the client to remove grit before every damp mop." = project advice; must be supported by the source.

CLASSIFY ONLY:
1. FACTUAL CONTRADICTION — the rewrite states a project fact that conflicts with the source.
2. UNSUPPORTED EMBELLISHMENT — the rewrite presents a new project-specific fact, severity, event, customer action, treatment detail or result that the source does not support.

Before flagging anything, ask:
"Is this sentence claiming that something actually happened on this Northampton project?"
If NO, do not flag it in W2B.05 unless it directly contradicts the documented project.
If YES, compare its meaning with the original source.

OUTPUT FORMAT — HARD LOCK:
Return ONLY valid JSON. No markdown fences and no commentary outside the JSON.

If there are no issues:
{
  "status": "PASS",
  "summary": "No unsupported project facts found.",
  "fixes": []
}

If issues are found:
{
  "status": "FAIL",
  "summary": "Brief summary.",
  "fixes": [
    {
      "fixLabel": "short label",
      "issueType": "FACTUAL CONTRADICTION or UNSUPPORTED EMBELLISHMENT",
      "targetText": "smallest exact visible-text fragment identifying the unsupported project claim",
      "contextText": "nearby exact visible text only if needed to distinguish duplicate targetText; otherwise empty string",
      "newText": "minimum corrected visible wording supported by the original source",
      "sourceEvidence": "brief source evidence"
    }
  ]
}

ATOMIC FIX RULES:
1. targetText and contextText are visible text, not HTML.
2. Prefer a short, unique targetText.
3. If targetText occurs more than once, provide contextText from the same local sentence/paragraph.
4. newText must make the minimum correction necessary.
5. If deletion is safest, use an empty newText.
6. Do not replace one unsupported project fact with another.
7. Preserve general explanatory material that is not itself a false project claim.

LINK PRESERVATION:
8. targetText may span visible anchor text when necessary.
9. If targetText crosses an existing link, put [[LINK_1]], [[LINK_2]], etc. in newText wherever the original link must remain.
10. Do not reproduce URLs or HTML. The application code restores the original exact anchor HTML.
11. Every crossed link must have a matching placeholder in the same order.

Return the JSON only.`;
  return { success: true, prompt: prompt, postId: postId, title: title };
}
function buildFactCheckFixPrompt(findingsText, sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const requestedCol = (sourceColumn === 'CT') ? 'CT' : 'EU';

  // Same auto-detect as buildFactCheckPrompt — always build the fix against
  // the latest saved state (GE if a previous fix exists), never the
  // untouched original, so consecutive fix passes compound correctly.
  const geHtml = String(sh.getRange(row, 187).getValue() || '').trim();
  const targetCol = geHtml ? 'GE' : requestedCol;
  const html = geHtml || String(sh.getRange(requestedCol + row).getValue() || '').trim();

  if (!html) {
    return { success: false, message: "ERROR: Column " + targetCol + " is empty for this row." };
  }
  if (!findingsText || !findingsText.trim()) {
    return { success: false, message: "ERROR: No fact-check findings provided." };
  }

  const prompt =
`You are correcting specific fabricated or unsupported claims in an article, identified by a fact-check against the original source project.

FINDINGS TO ADDRESS (only fix items marked FACTUAL CONTRADICTION or UNSUPPORTED EMBELLISHMENT):
${findingsText}

RULES:
1. Only rewrite the specific fragments flagged above. Leave everything else in the article completely unchanged.
2. Remove or correct the fabricated/embellished claim so it accurately reflects only what the source evidence supports. Do not replace one fabrication with another — if the source says nothing on a point, remove the claim rather than inventing a different one.
3. Keep the surrounding sentence natural and grammatically correct after the fix.
4. Do NOT rewrite, wrap, or alter any <a href="..."> link or its surrounding anchor tag unless the finding explicitly flags that link's text. Never convert HTML links into markdown-style syntax.
5. Return ONLY a JSON array, no markdown fences, no preamble, in this exact format:
[
  {"fixLabel": "short label for this fix", "old": "exact original fragment from the article below", "new": "exact replacement fragment"}
]
6. Every double quote character inside the old/new string values MUST be escaped as \\" so the result is valid JSON. Use straight quotes only.
7. The "old" fragment must be copied verbatim from the article HTML below — exact characters — so it can be located by an automated find-and-replace. Include enough surrounding text to make each "old" fragment unique within the article.
8. REDUNDANCY CHECK (Hard Lock): Before finalising each replacement, check the surrounding paragraphs in the ARTICLE HTML below (not just the flagged fragment) for another sentence that already describes the same technical step or claim. If the article elsewhere already covers the same point, do not reinsert it — tighten or shorten your replacement so the corrected article does not end up stating the same fact twice in nearby paragraphs.
9. ${ce_getConcreteLanguageStandard()}
10. ${ce_getFixBatchIntegrityCheck()}
11. ${ce_getContentPreservationRules()}
12. ${ce_getSourceEvidencePriorityLock()}
13. ${ce_getCustomerFacingCopyOnlyLock()}
14. ${ce_getLocationVarietyWithoutFabricationLock()}

ARTICLE HTML:
${html}`;

  return { success: true, prompt: prompt };
}

function applyFactCheckFix(rawJson, sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const requestedCol = (sourceColumn === 'CT') ? 'CT' : 'EU';

  // Work on the latest GE if a previous fix exists.
  const geHtmlCheck =
    String(sh.getRange(row, 187).getValue() || '').trim();

  const targetCol =
    geHtmlCheck ? 'GE' : requestedCol;

  if (!rawJson || !rawJson.trim()) {
    return {
      success: false,
      message: "No response provided."
    };
  }

  let fixes;

  let cleaned = rawJson
    .trim()
    .replace(/^```json/i, '')
    .replace(/^```/, '')
    .replace(/```$/, '')
    .trim();

  try {
    fixes = JSON.parse(cleaned);

  } catch (e) {

    try {
      const repaired = cleaned.replace(
        /"(old|new|fixLabel)":\s*"([\s\S]*?)"(?=\s*[,}])/g,
        function(match, key, value) {

          const fixedValue =
            value.replace(/(?<!\\)"/g, '\\"');

          return '"' + key + '": "' + fixedValue + '"';
        }
      );

      fixes = JSON.parse(repaired);

    } catch (e2) {

      return {
        success: false,
        message:
          "Could not parse response as JSON: " +
          e.message +
          " (repair attempt also failed: " +
          e2.message +
          ")"
      };
    }
  }

  if (!Array.isArray(fixes) || fixes.length === 0) {
    return {
      success: false,
      message: "No fixes found in parsed response."
    };
  }

  const html =
    (targetCol === 'GE')
      ? String(sh.getRange(row, 187).getValue() || '')
      : String(sh.getRange(targetCol + row).getValue() || '');

  if (!html) {
    return {
      success: false,
      message:
        "Column " +
        targetCol +
        " is empty for this row."
    };
  }

  const brokenLinkPattern =
    /<a\s+href="\[https?:\/\/[^\]]+\]\(/i;

  const validationErrors = [];
  const seenOld = {};

  // ---------------------------------------------------------
  // VALIDATE EVERY FIX BEFORE CHANGING GE
  // ---------------------------------------------------------

  fixes.forEach(function(fix, index) {

    const label =
      String(
        fix.fixLabel ||
        ('Fix ' + (index + 1))
      );

    if (
      typeof fix.old !== 'string' ||
      fix.old.length === 0
    ) {
      validationErrors.push(
        label + ' — OLD text is missing.'
      );
      return;
    }

    if (typeof fix.new !== 'string') {
      validationErrors.push(
        label +
        ' — NEW must be a string, including "" for deletion.'
      );
      return;
    }

    if (
      brokenLinkPattern.test(fix.old) ||
      brokenLinkPattern.test(fix.new)
    ) {
      validationErrors.push(
        label +
        ' — contains malformed markdown-in-HTML link syntax.'
      );
      return;
    }

    if (seenOld[fix.old]) {
      validationErrors.push(
        label +
        ' — duplicate OLD target in same batch.'
      );
      return;
    }

    seenOld[fix.old] = true;

    const occurrences =
      html.split(fix.old).length - 1;

    if (occurrences === 0) {
      validationErrors.push(
        label +
        ' — OLD fragment not found in article.'
      );
      return;
    }

    if (occurrences > 1) {
      validationErrors.push(
        label +
        ' — OLD fragment matches ' +
        occurrences +
        ' places.'
      );
      return;
    }
  });

  // Reject overlapping OLD targets.
  for (let i = 0; i < fixes.length; i++) {
    for (let j = i + 1; j < fixes.length; j++) {

      const oldA =
        String(fixes[i].old || '');

      const oldB =
        String(fixes[j].old || '');

      if (
        oldA &&
        oldB &&
        (
          oldA.indexOf(oldB) !== -1 ||
          oldB.indexOf(oldA) !== -1
        )
      ) {
        validationErrors.push(
          'Overlapping OLD targets detected between "' +
          (fixes[i].fixLabel || ('Fix ' + (i + 1))) +
          '" and "' +
          (fixes[j].fixLabel || ('Fix ' + (j + 1))) +
          '".'
        );
      }
    }
  }

  // Atomic rule: one bad fix means NOTHING changes.
  if (validationErrors.length > 0) {
    return {
      success: false,
      message:
        'Fact-check fix batch rejected — no changes applied. ' +
        validationErrors.join(' | ')
    };
  }

  // ---------------------------------------------------------
  // APPLY ONLY AFTER THE WHOLE BATCH PASSES
  // ---------------------------------------------------------

  let updatedHtml = html;
  const applied = [];

  fixes.forEach(function(fix, index) {

    updatedHtml =
      updatedHtml.replace(
        fix.old,
        fix.new
      );

    applied.push(
      String(
        fix.fixLabel ||
        ('Fix ' + (index + 1))
      )
    );
  });

  // Fact-check final/working HTML is GE.
  sh.getRange(row, 187).setValue(updatedHtml);

  if (typeof logPipelineResume === 'function') {
    logPipelineResume(
      "W2B.05 — Fact-Check Fix Applied Atomically (GE)",
      ""
    );
  }

  return {
    success: true,
    message:
      fixes.length +
      '/' +
      fixes.length +
      ' fact-check fix(es) applied atomically to GE. Applied: ' +
      applied.join(', '),
    appliedCount: fixes.length
  };
}

function passThroughFactCheckToGE() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const html = String(sh.getRange(row, 151).getValue() || '').trim(); // EU = column 151
  if (!html) {
    return { success: false, message: "ERROR: Column EU (W2B Raw HTML) is empty for this row." };
  }

  sh.getRange(row, 187).setValue(html); // GE = column 187
  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B.05 — Fact-Check PASS, copied EU to GE", "");
  }

  return { success: true, message: "No issues found. EU copied unchanged to column GE (W2B.05 Fixed HTML)." };
}

function ce_checkMaterialCategoryIntegrity_(html) {

  var d = getActiveRowDataMap();

  var material = String(
    d["Stone Type"] ||
    d["Material"] ||
    ""
  ).trim();

  if (!material || !html) {
    return { passed: true };
  }

  var materialLower = material.toLowerCase();

  // These governed materials are NOT natural stone.
  // They may legitimately be described as tiles, but not as stone.
  var nonStoneMaterialTests = [
    { test: /ceramic/i, articleTest: /\bceramic\b/i },
    { test: /porcelain/i, articleTest: /\bporcelain\b/i },
    { test: /terracotta/i, articleTest: /\bterracotta\b/i },
    { test: /quarry/i, articleTest: /\bquarry\b/i },
    { test: /victorian/i, articleTest: /\bvictorian\b/i },
    { test: /edwardian/i, articleTest: /\bedwardian\b/i },
    { test: /encaustic/i, articleTest: /\bencaustic\b/i }
  ];

  var governedTest = null;

  for (var i = 0; i < nonStoneMaterialTests.length; i++) {
    if (nonStoneMaterialTests[i].test.test(materialLower)) {
      governedTest = nonStoneMaterialTests[i].articleTest;
      break;
    }
  }

  // Natural stone and other materials are not governed by this hard check.
  if (!governedTest) {
    return { passed: true };
  }

  var plainText = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  var sentences =
    plainText.match(/[^.!?]+[.!?]+|[^.!?]+$/g) || [];

  var directStoneReference =
    /\b(each stone|the stone|this stone|that stone|stone behaves|stone surface|stone itself|natural stone)\b/i;

  var comparisonLanguage =
    /\b(differs? from|different from|unlike|compared with|compared to|rather than|as opposed to|is not natural stone|isn't natural stone|not a natural stone)\b/i;

  for (var s = 0; s < sentences.length; s++) {

    var sentence = String(sentences[s]).trim();

    if (
      governedTest.test(sentence) &&
      directStoneReference.test(sentence) &&
      !comparisonLanguage.test(sentence)
    ) {

      return {
        passed: false,
        finding:
          'ISSUE 1\n' +
          'TYPE: FACTUAL CONTRADICTION\n' +
          'CLAIM: "' + sentence + '"\n' +
          'WHY IT FAILS: The governed material is "' +
          material +
          '", but this sentence directly describes that material as stone. ' +
          'The governed material may legitimately be described as tile, but it must not be reassigned to the natural-stone material category.\n' +
          'SOURCE EVIDENCE: Governed Material = "' +
          material +
          '".'
      };
    }
  }

  return { passed: true };
}