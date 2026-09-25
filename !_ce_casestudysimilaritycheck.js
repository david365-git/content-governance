/**
 * ce_CaseStudySimilarityCheck.gs - CROSS-ARTICLE PHRASE & STRUCTURE SIMILARITY
 * Compares the active row's New HTML against all other same-material Case
 * Study articles (with New HTML populated), flagging repeated phrases and
 * structural patterns appearing in 2 or more articles.
 */

function buildCaseStudySimilarityPrompt(sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const targetCol = (sourceColumn === 'EU') ? 'GE' : 'CT';
  const comparisonCol = 'CT'; // always compare against other rows' finished New HTML, not their in-progress GE draft

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  const colArticleType = headers.indexOf('Article Type');
  const colStoneType = headers.indexOf('Stone Type');
  const colPostId = headers.indexOf('Post ID');
  const colTitle = 1; // column B

  if (colArticleType === -1 || colStoneType === -1) {
    return { success: false, message: "ERROR: Could not find Article Type or Stone Type columns." };
  }

  const rowData = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const thisArticleType = String(rowData[colArticleType] || '').trim();
  const thisStoneType = String(rowData[colStoneType] || '').trim();
  const thisHtml = String(sh.getRange(targetCol + row).getValue() || '').trim();
  const thisPostId = String(rowData[colPostId] || '').trim();
  const thisTitle = String(rowData[colTitle] || '').trim();

  if (thisArticleType !== 'Case Study') {
    return { success: false, message: "ERROR: Active row's Article Type is not 'Case Study'." };
  }
  if (!thisStoneType) {
    return { success: false, message: "ERROR: Stone Type is empty for this row." };
  }
  if (!thisHtml) {
    return { success: false, message: "ERROR: Column " + targetCol + " is empty for this row." };
  }

  const allData = sh.getDataRange().getValues();
  const comparisonColHtmlValues = sh.getRange(comparisonCol + '1:' + comparisonCol + sh.getLastRow()).getValues();
  const comparisons = [];

  for (let i = 1; i < allData.length; i++) {
    if (i + 1 === row) continue; // skip the active row itself
    const r = allData[i];
    const articleType = String(r[colArticleType] || '').trim();
    const stoneType = String(r[colStoneType] || '').trim();
    const html = String(comparisonColHtmlValues[i][0] || '').trim();
    if (articleType !== 'Case Study') continue;
    if (stoneType !== thisStoneType) continue;
    if (!html) continue;

    comparisons.push({
      postId: String(r[colPostId] || '').trim(),
      title: String(r[colTitle] || '').trim(),
      html: html
    });
  }

  if (comparisons.length === 0) {
    return { success: false, message: "No other " + thisStoneType + " Case Studies with New HTML found to compare against." };
  }

  let comparisonBlock = '';
  comparisons.forEach(function(c, idx) {
    comparisonBlock += '\n\n=== COMPARISON ARTICLE ' + (idx + 1) + ' — Post ID ' + c.postId + ' — "' + c.title + '" ===\n' + c.html;
  });

  const prompt =
`You are auditing content for repeated phrases and structural patterns across a set of Case Study articles for the same stone material, to catch content-footprint risk before publication.

TARGET ARTICLE (Post ID ${thisPostId} — "${thisTitle}"):
${thisHtml}

COMPARISON ARTICLES (same material — ${thisStoneType}):
${comparisonBlock}

TASK:
Compare the TARGET ARTICLE against each COMPARISON ARTICLE and identify:

1. REPEATED PHRASES — any distinctive diagnostic, descriptive, or narrative phrase (roughly 5+ words) that appears in the TARGET ARTICLE and also appears in at least one COMPARISON ARTICLE, even if reworded slightly. Only flag phrases in 2 or more articles total (the target plus at least one comparison) — do not flag phrases unique to a single article. Quote the phrase as it appears in each article, and name which article(s) it appears in.

Example of what to catch: "catches a fingernail" or "catches the tip of a fingernail" appearing in multiple articles as a stock diagnostic description, even with different surrounding wording.

2. STRUCTURAL PATTERN REPETITION — note if the TARGET ARTICLE's opening paragraph formula, section/H3 pacing, or closing "honesty/expectations" section follows the same shape as a COMPARISON ARTICLE, even if the words differ.

3. SEVERITY — for each finding, state whether it is a MINOR overlap (a natural shared technical term) or a GENUINE RISK (a stock phrase or structural template that a reader or search engine could recognise as repeated content).

OUTPUT FORMAT:
For each finding, output:
FINDING: [brief description]
SEVERITY: MINOR or GENUINE RISK
FOUND IN: [Post ID / title of each article where it appears, including the target]
QUOTE FROM TARGET: "[exact phrase from target article]"
QUOTE FROM COMPARISON: "[exact phrase from comparison article]"

If no genuine repeated phrases or structural patterns are found, state: "No cross-article repetition found — article reads as structurally and phrasally distinct within this material."

Do not flag material-specific technical terms that must reasonably repeat (e.g. "diamond honing", "colour-matched filler") — only flag phrasing and structure that reflects templated writing rather than necessary shared vocabulary.

Do NOT flag as GENUINE RISK, and preferably omit entirely: standard footer/CTA text (e.g. "Contact us to arrange a no-obligation ... assessment"), product recommendation box headings (e.g. "Pro Tip: We recommend these products for daily ... maintenance cleaning"), bio box boilerplate about years of experience, or any other site-wide template component that is intentionally identical across all articles by design. These are not editorial prose and repeating them is correct, not a footprint risk.`;

  return { success: true, prompt: prompt, comparisonCount: comparisons.length };
}

function buildSimilarityFixPrompt(findingsText, sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const targetCol = (sourceColumn === 'EU') ? 'GE' : 'CT';

  const html = String(sh.getRange(targetCol + row).getValue() || '');
  if (!html) {
    return { success: false, message: "ERROR: Column " + targetCol + " is empty for this row." };
  }
  if (!findingsText || !findingsText.trim()) {
    return { success: false, message: "ERROR: No similarity check findings provided." };
  }

  const prompt =
  `You are rewriting specific flagged phrases in an article to remove repetition with other same-material articles, without changing anything else.

  FINDINGS TO ADDRESS (only rewrite what is flagged here — ignore anything marked MINOR):
  ${findingsText}

  RULES:
  1. Only rewrite phrases/sections explicitly flagged as GENUINE RISK in the findings above. Leave everything else in the article completely unchanged.
  2. Do NOT create a fix for any product recommendation box heading (e.g. "Pro Tip: We recommend these products..."), standard footer/CTA text, or bio box boilerplate — these are intentional site-wide template components, not editorial prose, even if they appear in the findings above. Skip them entirely.
  2b. NEVER reword, replace, or restructure the existing wording of any <h2> heading text — it is governed and validated in an earlier pipeline stage. If a finding flags a repeated heading PATTERN or STRUCTURE (not verbatim duplicate text) across articles, you MAY append a short distinguishing detail (a locality name, or a specific technical detail unique to this project, 2-5 words) onto the END of the existing heading text, joined by " — ". Do NOT delete, shorten, or reword any part of the original heading — only append. If the finding flags truly identical or near-identical heading TEXT (not just a shared structural pattern), skip it entirely rather than appending, since that requires a full rewrite which is out of scope here.
  3. Each rewrite must preserve the same factual meaning and technical accuracy — only the wording/phrasing changes, not the underlying claim.
  4. Do not introduce a new stock phrase that could itself become repetitive — vary the wording naturally, as a human writer would.
  5. NEVER rewrite, wrap, or alter any <a href="..."> link or its surrounding anchor tag unless the finding explicitly flags that specific link's text as repeated. Do not convert HTML links into markdown-style syntax under any circumstances — links must remain in their original valid HTML form, character-for-character, unless untouched entirely.
  6. Return ONLY a JSON array, no markdown fences, no preamble, in this exact format:
  [
    {"fixLabel": "short label for this fix", "old": "exact original fragment from the article below", "new": "exact replacement fragment"}
  ]
  7. Every double quote character inside the old/new string values MUST be escaped as \\" so the result is valid JSON. Use straight quotes only.
  8. The "old" fragment must be copied verbatim from the article HTML below — exact characters — so it can be located by an automated find-and-replace. Include enough surrounding text to make each "old" fragment unique within the article.
  9. REDUNDANCY CHECK (Hard Lock): Before finalising each replacement, check the surrounding paragraphs in the ARTICLE HTML below (not just the flagged fragment) for another sentence that already describes the same technical step or claim. If the article elsewhere already covers the same point, do not reinsert it — tighten or shorten your replacement so the corrected article does not end up stating the same fact twice in nearby paragraphs.
  10. ${ce_getConcreteLanguageStandard()}
  11. ${ce_getFixBatchIntegrityCheck()}
  12. ${ce_getContentPreservationRules()}
  13. ${ce_getSourceEvidencePriorityLock()}
  14. ${ce_getCustomerFacingCopyOnlyLock()}
  15. ${ce_getLocationVarietyWithoutFabricationLock()}

  ARTICLE HTML:
  ${html}`;

  return { success: true, prompt: prompt };
}

function applySimilarityFix(rawJson, sourceColumn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const targetCol = (sourceColumn === 'EU') ? 'GE' : 'CT';

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
    String(sh.getRange(targetCol + row).getValue() || '');

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
  // VALIDATE THE ENTIRE BATCH FIRST
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
        label + ' — NEW must be a string.'
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

  // Detect overlapping OLD targets.
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

  // ---------------------------------------------------------
  // ATOMIC SAFETY
  // If any fix fails, apply NOTHING.
  // ---------------------------------------------------------

  if (validationErrors.length > 0) {
    return {
      success: false,
      message:
        'Similarity fix batch rejected — no changes applied. ' +
        validationErrors.join(' | ')
    };
  }

  // ---------------------------------------------------------
  // APPLY ONLY AFTER EVERY FIX PASSES
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

  // W2B.1 final working HTML
  sh.getRange(row, 188).setValue(updatedHtml); // GF

  if (typeof logPipelineResume === 'function') {
    logPipelineResume(
      "W2B.1 — Similarity Fix Applied Atomically (GF)",
      ""
    );
  }

  return {
    success: true,
    message:
      fixes.length +
      '/' +
      fixes.length +
      ' similarity fix(es) applied atomically to GF. Applied: ' +
      applied.join(', '),
    appliedCount: fixes.length
  };
}

function passThroughSimilarityCheckToGF() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const html = String(sh.getRange(row, 187).getValue() || '').trim(); // GE = column 187
  if (!html) {
    return { success: false, message: "ERROR: Column GE (W2B.05 Fixed HTML) is empty for this row." };
  }

  sh.getRange(row, 188).setValue(html); // GF = column 188
  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B.1 — Similarity Check PASS, copied GE to GF", "");
  }

  return { success: true, message: "No genuine risk found. GE copied unchanged to column GF (W2B.1 Fixed HTML)." };
}