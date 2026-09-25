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
`You are fact-checking a rewritten article against its original source, to catch fabricated facts and unsupported embellishments introduced during an AI content rewrite pipeline.

GOVERNED PAGE CONTEXT:
MATERIAL: ${governedMaterial || "Not supplied"}
ARTICLE TYPE: ${articleType || "Not supplied"}
LOCALITY: ${governedLocality || "Not supplied"}

PAGE REWRITE BRIEF:
${rewriteBrief || "No active Rewrite Brief supplied."}

ORIGINAL SOURCE ARTICLE (ground truth — what actually happened):
${originalHtml}

REWRITTEN ARTICLE (column ${targetCol} — check this against the source above):
${generatedHtml}

TASK:
Compare the rewritten article against the original source. Identify every claim in the rewritten article that either:
1. FACTUAL CONTRADICTION — directly contradicts or invents something not present in the original (e.g. the rewrite says the floor moved/was unstable, but the original never mentions this).
2. UNSUPPORTED EMBELLISHMENT — adds dramatic framing, severity, or narrative colour with no supporting evidence in the original (e.g. "looked beyond saving", "felt embarrassed", "structural movement" when the original describes a routine, undramatic job).

Do NOT flag: natural paraphrasing, expanded technical explanation of a process that IS mentioned in the original (e.g. explaining why alkaline cleaner is used, when the original just says "we used alkaline cleaner"), or reasonable inference directly supported by stated facts (e.g. if the original says tiles were cracked and filled, describing the repair process in more technical detail is fine).

ADDITIONAL FACTUAL INTEGRITY CHECKS — HARD LOCK

In addition to checking factual claims against the supplied sources, inspect the article for the following internal contradictions and unsupported claims.

A. MATERIAL IDENTITY
Check that every material description matches the governed Material Entity.
Examples:
- Ceramic Tile and Porcelain Tile must not be described as natural stone or "stone".
- Natural stone terminology must not be transferred automatically to manufactured tile.
Flag any material-category contradiction.

B. LOCALITY INTEGRITY
For a locality-specific page, identify the governed locality from the supplied governance data.
Flag any unrelated town, county, city or region introduced into the article body unless it is explicitly supported as relevant context.
Raw search-query terms do not constitute evidence that another locality belongs in the article.

C. UNSUPPORTED BUSINESS OR AUTHORITY CLAIMS
Flag words or statements such as:
vetted, approved, accredited, certified, guaranteed, trusted, recommended, authorised, leading, award-winning
unless supplied evidence establishes the claim.

D. INVENTED PROJECT OR CUSTOMER NARRATIVE
Flag any statement implying a real assessment, project, customer, property or completed intervention unless the supplied source material establishes that event.
Examples:
"After assessing this installation..."
"The homeowner..."
"On this project..."
"We found..."

E. TREATMENT/MATERIAL CONTRADICTIONS
Check that treatments recommended in one part of the article do not contradict material guidance elsewhere.
Example: if the article correctly states that the ceramic tile itself normally needs no sealer, later wording must not imply routine resealing of the ceramic floor unless the wording clearly concerns a separate sealable component and is supported.

F. CLAIM-BEARING HEADINGS
Check H2/H3 headings for unsupported factual or quality claims as well as the body text.

Any one of these failures means the fact check must NOT return a clean PASS.

  MANDATORY MATERIAL-CATEGORY PRE-PASS CHECK

  Before returning PASS, determine the physical material category of the governed MATERIAL.

  Distinguish between:
  - natural stone;
  - fired clay / ceramic / porcelain tile;
  - other manufactured tile or flooring materials.

  The word "tile" describes an installation form and may legitimately apply to many materials, including natural stone.

  The word "stone", however, must not be used to describe the governed material itself unless that material is actually natural stone.

  Therefore:
  - Slate may legitimately be described as stone and as tile.
  - Marble, limestone, travertine and granite may legitimately be described as stone and as tile.
  - Ceramic, porcelain, terracotta, quarry tile, Victorian tile, Edwardian tile and encaustic tile must not be described as natural stone merely because they are installed as tiles.

  Before PASS, scan the complete rewritten article for any sentence where "stone" or "natural stone" refers back to the governed material.

  If the governed material is not natural stone and the article describes that material as stone, report:
  TYPE: FACTUAL CONTRADICTION

  Do not return PASS while such a contradiction remains.

  CLASSIFICATION LOCK:
  Every failure found under A-F above MUST be reported as:
  TYPE: FACTUAL CONTRADICTION

  Do not silently treat these integrity failures as stylistic, editorial, optional or outside the original-source comparison.

  The governed Material, Article Type, Locality and Rewrite Brief supplied above are authoritative governance evidence.

  MATERIAL IDENTITY RULE:
    Judge material terminology against the governed material category, not against a single banned word.

    Do not flag a term merely because it uses "tile" or "stone".
    Many governed materials may legitimately be described as tiles, including slate, terracotta, quarry tile, Victorian tile, Edwardian tile and encaustic tile.

    Flag only where the wording assigns the material to the wrong physical or material category.

    Examples:
    - Ceramic Tile described as natural stone → FACTUAL CONTRADICTION.
    - Porcelain Tile described as natural stone → FACTUAL CONTRADICTION.
    - Slate described as a tile installation → not automatically a contradiction.
    - Terracotta described as tile → not automatically a contradiction.
    - Quarry Tile described as tile → not automatically a contradiction.
    - Victorian, Edwardian or Encaustic flooring described as tile → not automatically a contradiction.

    The test is:
    Does the wording materially misclassify what the governed material actually is?

    If yes, report:
    TYPE: FACTUAL CONTRADICTION

    If no, do not flag it merely because of the words "stone" or "tile".
For each issue found, output:
ISSUE: [brief description]
TYPE: FACTUAL CONTRADICTION / UNSUPPORTED EMBELLISHMENT
QUOTE FROM REWRITE: "[exact fabricated/embellished text]"
SOURCE EVIDENCE: [what the original actually says on this point, or "not mentioned at all" if the original says nothing relevant]

If no issues are found, state: "No factual contradictions or unsupported embellishments found — rewrite is fully grounded in the source article."

End with:
OVERALL: PASS (no issues) or FAIL ([N] issue(s) found)`;

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