/* ============================================================
   ce_Stage2B_RewriteBriefCheck.gs
   STAGE 2B — REWRITE BRIEF COMPLIANCE CHECK (BODY LEVEL)
   Full-article check that the generated HTML body genuinely
   delivers on the Rewrite Brief: re-anchor entities present
   with real coverage, drift-prone language removed. Runs after
   W2B, before W3 humanisation.
============================================================ */

function buildRewriteBriefComplianceCheckPromptW2B() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();
  if (row < 2) return 'ERROR: Select a data row (row 2 or below).';

  // Prefer W2B.2 Fixed HTML (col GG / 189) — a previous fix already applied here
  var html = String(sheet.getRange(row, 189).getValue() || "").trim();
  if (!html) html = String(sheet.getRange(row, 188).getValue() || "").trim(); // fallback: GF (W2B.1 Fixed HTML, Case Studies)
  if (!html) html = String(sheet.getRange(row, 187).getValue() || "").trim(); // fallback: GE (W2B.05 Fixed HTML — used when W2B.1 Similarity was skipped, e.g. non-Case-Study article types)
  if (!html) html = String(sheet.getRange(row, 98).getValue() || "").trim(); // fallback: New HTML
  if (!html) return 'ERROR: No article HTML found in column GG, GF, GE, or column CT (New HTML). Run W2B.05 first.';

  const d = getActiveRowDataMap();
  const rewriteBrief = String(d["Page Rewrite Brief"] || "").trim();

  if (!rewriteBrief || rewriteBrief === "Not Applicable") {
    return 'ERROR: No active Rewrite Brief for this article — compliance check not applicable.';
  }

  // AUTO-SCOPE: if a previous FAIL result exists, delegate to the scoped
  // recheck instead of re-auditing the entire article from scratch.
  var headersForRecheck = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                               .map(function(h) { return String(h).trim(); });
  var recheckColIdx = headersForRecheck.indexOf('Rewrite Brief Compliance Check W2B');
  if (recheckColIdx > -1) {
    var previousSaved = String(sheet.getRange(row, recheckColIdx + 1).getValue() || "").trim();
    if (previousSaved.indexOf('FAIL') === 0) {
      var scopedPrompt = buildRewriteBriefComplianceRecheckPromptW2B_(previousSaved);
      if (scopedPrompt) return scopedPrompt;
      // scopedPrompt === null means no failing items were parseable — fall through to full audit
    }
  }

  const prompt = `
    STAGE 2B — REWRITE BRIEF COMPLIANCE CHECK (BODY LEVEL)
    ROLE: Senior UK SEO & Editorial Quality Auditor

    --- REWRITE BRIEF ---
    ${rewriteBrief}

    --- ARTICLE HTML ---
    ${html}

    --- CLASSIFICATION RULE (apply before doing any checking) ---
    The brief may name the same topic in two different roles, and these must not both become RE-ANCHOR requirements:
    - A topic named as part of the "observed query cluster" (i.e. what the page currently ranks for, per GSC) that is ALSO separately instructed to be removed, minimised, or kept out of the page's framing, is a DRIFT-CORRECTION signal — not a coverage requirement. Audit it ONLY as a DRIFT item (checking it is absent), never as a RE-ANCHOR item.
    - Only treat a topic as a RE-ANCHOR item if the brief asks for it to be genuinely present/covered, with no separate instruction elsewhere in the brief to remove or minimise that same topic.
    - If you are unsure whether a topic is a coverage requirement or a drift-correction signal, re-read the brief for any sentence containing "remove", "must not expand", "avoid", or "outside its scope" referring to that same topic — if found, it is DRIFT-ONLY.

    DRIFT ITEM SOURCE RULE:

    DRIFT items may be drawn from TWO governed sources only:

    1. SCOPE BOUNDARY
    Any named topic the SCOPE BOUNDARY explicitly says the article must not drift into, expand into, or become.

    2. QUERY PRECISION ACTION — EXPLICIT EXCLUSIONS
    Any specific term, phrase, material, service or locality that QUERY PRECISION ACTION explicitly says to:
    - exclude;
    - omit;
    - ignore;
    - treat as cross-location evidence only;
    - treat as unrelated;
    - or not work into the rewritten content.

    An explicit QUERY PRECISION exclusion is a governed DRIFT item even if it does not also appear in SCOPE BOUNDARY.

    Example:
    If QUERY PRECISION ACTION says:
    "exclude 'surrey' as cross-location evidence and exclude unrelated terms such as 'high-end construction tiles'"
    then both "surrey" and "high-end construction tiles" must be audited as DRIFT items and must not appear as article-body content unless the Rewrite Brief separately establishes a legitimate governed reason for them.

    Do not create DRIFT items from ordinary missing query terms that are merely unselected.
    Only explicit exclusion language creates a Query Precision DRIFT item.

    Do not create additional or fragmentary DRIFT items from the ESCALATION BOUNDARY clause — that clause describes when DIY work must stop and escalate to a professional, and is not a source of separate drift-topic names.

    MANDATORY QUERY PRECISION EXCLUSION CHECK — HARD LOCK

    Before returning the audit result, parse the QUERY PRECISION ACTION clause separately.

    For EVERY term or phrase explicitly described as:
    - excluded;
    - unrelated;
    - cross-location evidence only;
    - not to be worked into the content;
    - omitted;
    - or ignored;

    you MUST create a corresponding DRIFT ITEM in the audit output.

    Do not combine multiple exclusions into one generic DRIFT ITEM.

    Example:
    If QUERY PRECISION ACTION says:
    exclude "surrey" as cross-location evidence and exclude unrelated terms such as "professional tile supplies" and "high-end construction tiles"

    the audit MUST contain three separate DRIFT ITEM checks:
    - surrey
    - professional tile supplies
    - high-end construction tiles

    A clean CHECK 2B-BRIEF: PASS is prohibited unless every explicit QUERY PRECISION exclusion has been individually checked and has PRESENT: NO.

    Finding an excluded phrase anywhere in governed article-body prose means PRESENT: YES and the audit must FAIL.

    --- COMMERCIAL PRODUCT TABLE EXEMPTION ---
    The article HTML may contain one or more retained commercial product tables (recognisable by inline-styled headings like "Products often used during..." or "...are commonly paired with...", followed by a <table> of Amazon product links). These table headings and their surrounding product copy are NOT governed article content — they are preserved verbatim from the original page per a separate rule. Do NOT audit any product table heading or its product descriptions against RE-ANCHOR or DRIFT items. Only audit the surrounding prose paragraphs.

    --- TASK ---
    The Rewrite Brief above names specific entities/topics the article must re-anchor around with real coverage, and specific drift-prone language, framing, or topics that must be removed or minimised. Judge the full article body against both, applying the Classification Rule and the Commercial Product Table Exemption above.

    COVERAGE STANDARD: Coverage may be conceptual or mechanistic, not just verbatim. If the article
    explains the mechanism an entity name describes — even without using the governed term itself —
    treat it as genuinely covered and quote the sentence that demonstrates the mechanism. Only mark
    COVERED: NO where the underlying mechanism is genuinely absent, not merely unnamed.

    1. RE-ANCHOR CHECK: For each entity or topic the brief specifically names as something to re-anchor around (excluding any reclassified as DRIFT-ONLY per the Classification Rule), state whether the article gives it genuine substantive coverage (not just a passing mention) — YES or NO — quote the section or a short phrase where it appears, and state the exact section id attribute (e.g. section-3) of the <section> element containing that evidence. If covered across multiple sections, name the single section id with the strongest coverage.

    2. DRIFT CHECK: For each drift-prone term, framing, locality or topic the brief specifically says to remove, avoid or exclude — whether from SCOPE BOUNDARY or an explicit QUERY PRECISION ACTION exclusion, and including any reclassified from RE-ANCHOR per the Classification Rule — classify how it appears in the article outside any exempted commercial product table:

    - PASS — ABSENT: the topic does not appear.
    - PASS — CONTEXTUAL: the topic is mentioned briefly only to explain, qualify, warn, or support the governed subject and does not become its own workflow or intent.
    - PASS — ROUTED: the article deliberately sends the reader to another relevant article or hub for that topic, without developing the topic substantially on this page. A short routing sentence plus internal link is permitted.
    - FAIL — DEVELOPED: the article gives the drift topic substantive standalone coverage, procedural steps, decision logic, or enough detail that it becomes a competing intent on this page.

    Do NOT mark a routed internal link as drift merely because the linked topic is named. Judge the amount of on-page development, not the existence of the link itself.

    Quote the relevant phrase when present and state the exact section id attribute (e.g. section-2) of the <section> element containing it.

    OUTPUT FORMAT (exactly):
    RE-ANCHOR ITEM: [entity/topic text]
    COVERED: YES / NO
    EVIDENCE: [short quote or section reference, or "none"]
    SECTION: [section id, e.g. section-3, or "none" if COVERED is NO]

    (repeat for every re-anchor item)

    DRIFT ITEM: [drift term/framing text]
    STATUS: PASS — ABSENT / PASS — CONTEXTUAL / PASS — ROUTED / FAIL — DEVELOPED
    EVIDENCE: [short quote or routing/link evidence, or "none"]
    SECTION: [section id, e.g. section-2, or "none" if STATUS is PASS — ABSENT]

    (repeat for every drift item)

    Last line: "CHECK 2B-BRIEF: PASS" only if every re-anchor item is genuinely covered AND every drift item is classified as PASS — ABSENT, PASS — CONTEXTUAL, or PASS — ROUTED. Use "CHECK 2B-BRIEF: FAIL" only if one or more drift items are classified as FAIL — DEVELOPED, or a required re-anchor item is not covered.Last line: "CHECK 2B-BRIEF: PASS" only if every re-anchor item is genuinely covered AND no drift item is present. Otherwise "CHECK 2B-BRIEF: FAIL".
    `.trim();

  return prompt;
}

function saveRewriteBriefComplianceResultW2B(raw) {
  if (!raw || !raw.trim()) {
    return { success: false, message: 'No response provided.' };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row   = sheet.getActiveRange().getRow();
  if (row < 2) return { success: false, message: 'Select a data row first.' };

  var passed = /CHECK 2B-BRIEF:\s*PASS/i.test(raw);
  var failed = /CHECK 2B-BRIEF:\s*FAIL/i.test(raw);

  if (!passed && !failed) {
    return { success: false, message: 'Could not find CHECK 2B-BRIEF result line in response.' };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h) { return String(h).trim(); });
  var colIdx = headers.indexOf('Rewrite Brief Compliance Check W2B');
  if (colIdx === -1) {
    return { success: false, message: "Column 'Rewrite Brief Compliance Check W2B' not found in posts sheet." };
  }

  var reportText = raw.trim();

        if (failed) {

          var body = reportText
            .replace(/\n*CHECK 2B-BRIEF:\s*(PASS|FAIL)\s*$/i, '')
            .trim();

          var blocks = body.split(/\n(?=RE-ANCHOR ITEM:|DRIFT ITEM:)/);

          var failingBlocks = blocks.filter(function(block) {

            var reAnchorFail =
              /^RE-ANCHOR ITEM:/i.test(block.trim()) &&
              /COVERED:\s*NO/i.test(block);

            var driftFail =
              /^DRIFT ITEM:/i.test(block.trim()) &&
              (
                /STATUS:\s*FAIL\s*[—-]\s*DEVELOPED/i.test(block) ||
                /PRESENT:\s*YES/i.test(block)
              );

            return reAnchorFail || driftFail;
          });

          reportText =
            failingBlocks.join('\n\n') +
            '\n\nCHECK 2B-BRIEF: FAIL';

        } else {

          reportText =
            'No compliance failures found.' +
            '\n\nCHECK 2B-BRIEF: PASS';
        }

        var summary =
          (passed ? 'PASS' : 'FAIL') +
          ' — ' +
          reportText;
  var cell = sheet.getRange(row, colIdx + 1);
  cell.setNumberFormat('@');
  cell.setValue(summary);

  return {
    success: true,
    passed: passed,
    message: passed
      ? 'Article body complies with Rewrite Brief — proceed to W3 humanisation.'
      : 'Article body does not fully comply with Rewrite Brief — saved for review. Regenerate affected sections before proceeding.',
    rawResult: raw.trim()
  };
}

function applyRewriteBriefFixW2B(responseText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();

    if (row < 2) {
      return {
        success: false,
        message: 'Select a data row first.'
      };
    }

    // W2B.2 always works on the latest GG version.
    var html = String(
      sheet.getRange(row, 189).getValue() || ''
    );

    if (!html) {
      return {
        success: false,
        message: 'No article HTML found in GG.'
      };
    }

    // Remove accidental markdown fences.
    var cleanJson = String(responseText || '')
      .trim()
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/\s*```$/i, '')
      .trim();

    var fixes;

    try {
      fixes = JSON.parse(cleanJson);
    } catch (e) {
      return {
        success: false,
        message: 'Could not parse fix response as JSON: ' + e.message
      };
    }

    if (!Array.isArray(fixes) || fixes.length === 0) {
      return {
        success: false,
        message: 'No fixes found in JSON response.'
      };
    }

    var validationErrors = [];
    var seenOld = {};

    // ---------------------------------------------------------
    // VALIDATE THE ENTIRE BATCH BEFORE CHANGING ANYTHING
    // ---------------------------------------------------------

    fixes.forEach(function(fix, index) {

      var label =
        String(fix.fixLabel || ('Fix ' + (index + 1)));

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
          label + ' — NEW must be a string, including "" for deletion.'
        );
        return;
      }

      if (seenOld[fix.old]) {
        validationErrors.push(
          label + ' — duplicate OLD target in the same batch.'
        );
        return;
      }

      seenOld[fix.old] = true;

      var occurrences =
        html.split(fix.old).length - 1;

      if (occurrences === 0) {
        validationErrors.push(
          label + ' — OLD text was not found in GG.'
        );
        return;
      }

      if (occurrences > 1) {
        validationErrors.push(
          label +
          ' — OLD text occurs ' +
          occurrences +
          ' times in GG.'
        );
        return;
      }
    });

    // Detect overlapping OLD targets.
    for (var i = 0; i < fixes.length; i++) {
      for (var j = i + 1; j < fixes.length; j++) {

        var oldA = String(fixes[i].old || '');
        var oldB = String(fixes[j].old || '');

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

    // Atomic behaviour: if ONE fix fails validation,
    // NOTHING is applied.
    if (validationErrors.length > 0) {
      return {
        success: false,
        message:
          'Fix batch rejected — no changes applied. ' +
          validationErrors.join(' | ')
      };
    }

    // ---------------------------------------------------------
    // APPLY ONLY AFTER THE WHOLE BATCH PASSES VALIDATION
    // ---------------------------------------------------------

    var updatedHtml = html;
    var applied = [];

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

    // Save only after every replacement succeeded.
    sheet
      .getRange(row, 189)
      .setValue(updatedHtml);

    if (typeof logPipelineResume === 'function') {
      logPipelineResume(
        'W2B.2 — Exact Rewrite Brief Fix Applied (GG)',
        ''
      );
    }

    return {
      success: true,
      message:
        fixes.length +
        '/' +
        fixes.length +
        ' exact fix(es) applied atomically to GG. Applied: ' +
        applied.join(', '),
      appliedCount: fixes.length
    };

  } catch (e) {
    return {
      success: false,
      message: e.message
    };
  }
}

function buildLocationIntegrationPromptW2B() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();
  if (row < 2) return 'ERROR: Select a data row (row 2 or below).';

  var html = String(sheet.getRange(row, 151).getValue() || "").trim(); // EU = W2B Raw HTML
  if (!html) return 'ERROR: No W2B Raw HTML found in column EU — save W2B to the sheet first.';

  const d = getActiveRowDataMap();
  const location = String(d["Locality"] || d["Location"] || "").trim();
  if (!location) return 'ERROR: No Locality/Location set for this row — location integration is not applicable.';

  const prompt = `
STAGE 2B — LOCATION INTEGRATION (RUN BEFORE THE REWRITE BRIEF FIX PROMPT)
ROLE: Senior UK SEO content editor for a natural stone floor restoration website.

TASK: Strengthen this article's natural relevance to ${location} through genuine geographic audience context — not by repeating the location name and not by inventing local facts.

TECHNIQUE — GEOGRAPHIC HIERARCHY AND AUDIENCE FRAMING:
- Treat ${location} and any named towns or sub-regions as the ARTICLE'S AUDIENCE, not as evidence about a completed job, property, environmental condition or technical circumstance.
- Real, verifiable towns, settlements or recognised sub-regions within ${location} may be used where they improve natural audience framing.
- Do not add multiple place names simply because they are available.
- Suitable framing includes: "homeowners across [town], [town] and elsewhere in ${location}" or equivalent natural wording.
- Non-specific hypothetical framing is permitted only where it does not imply an observed property, project or local technical fact.

${ce_getLocationVarietyWithoutFabricationLock()}

INTEGRATION METHOD:
- Add the minimum geographic context necessary.
- Use 1 sentence by default.
- Use 2 where the article length or structure genuinely supports two distinct audience references.
- Use 3 only where each reference serves a different and useful contextual purpose.
- Never add a reference merely to reach a quota.
- Prefer adapting an existing sentence concerned with readership, applicability or guidance.
- Add a new standalone sentence only where modifying existing prose would sound unnatural or alter its meaning.
- Do not place location references in H2 headings.
- Do not cluster multiple location references together unless the wording naturally describes the article's audience.

RELEVANCE TEST:
Every location reference must contribute genuine audience or geographic orientation.
If removing the place name would leave exactly the same sentence with no meaningful loss of audience context, do not insert it there.

PRESERVATION:
Preserve all existing facts, figures, links, HTML structure and meaning exactly except for the minimum wording needed for location integration.
Do not rewrite surrounding prose for style.

ARTICLE HTML:
${html}

Return the complete corrected HTML only.
`.trim();

  return prompt;
}

function ce_stripHtmlCodeFences_(text) {
  return String(text || '')
    .trim()
    .replace(/^```html\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();
}


function saveLocationIntegrationResultW2B(responseHtml) {
  if (!responseHtml || !responseHtml.trim()) {
    return { success: false, message: 'No response provided.' };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row   = sheet.getActiveRange().getRow();
  if (row < 2) return { success: false, message: 'Select a data row first.' };

  var cell = sheet.getRange(row, 151); // EU — W2B Raw HTML, the column Fact Check reads by default
  cell.setNumberFormat('@');
  cell.setValue(ce_stripHtmlCodeFences_(responseHtml));

  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B — Location Integration Applied (EU)", "");
  }

  return {
    success: true,
    message: 'Location integration saved to column EU (W2B Raw HTML). Proceed to W2B.05 Fact Check next.'
  };
}

function buildRewriteBriefFixPromptW2B() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();
  if (row < 2) return 'ERROR: Select a data row (row 2 or below).';

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });
  const colIdx = headers.indexOf('Rewrite Brief Compliance Check W2B');
  if (colIdx === -1) return 'ERROR: Column "Rewrite Brief Compliance Check W2B" not found.';

  const savedResult = String(sheet.getRange(row, colIdx + 1).getValue() || '').trim();
  if (!savedResult) return 'ERROR: No saved compliance check result found. Run the compliance check first.';
  if (savedResult.indexOf('FAIL') !== 0) return 'ERROR: Saved result is not a FAIL — no fix needed.';

  var html = String(sheet.getRange(row, 189).getValue() || "").trim(); // GG = a previous fix already applied
  if (!html) html = String(sheet.getRange(row, 188).getValue() || "").trim(); // fallback: GF (W2B.1 Fixed HTML, Case Studies)
  if (!html) html = String(sheet.getRange(row, 187).getValue() || "").trim(); // fallback: GE (W2B.05 Fixed HTML — used when W2B.1 Similarity was skipped)
  if (!html) html = String(sheet.getRange(row, 98).getValue() || "").trim();
  if (!html) return 'ERROR: No article HTML found.';

  const d = getActiveRowDataMap();
  const rewriteBrief = String(d["Page Rewrite Brief"] || "").trim();

  // Extract failing section ids from the saved audit result
  var sectionIds = [];
  var itemsMissingSection = false;
  var itemBlocks = savedResult.split(/\n(?=RE-ANCHOR ITEM:|DRIFT ITEM:)/);
  var failingItems = [];
  itemBlocks.forEach(function(block) {
    var isReAnchorFail = /RE-ANCHOR ITEM:/.test(block) && /COVERED:\s*NO/i.test(block);
    var isDriftFail    = /DRIFT ITEM:/.test(block) && /STATUS:\s*FAIL\s*[—-]\s*DEVELOPED/i.test(block);
    if (!isReAnchorFail && !isDriftFail) return;
    var secMatch = block.match(/SECTION:\s*(section-[\w-]+)/i);
    if (secMatch) {
      var sid = secMatch[1].trim();
      if (sectionIds.indexOf(sid) === -1) sectionIds.push(sid);
    } else {
      itemsMissingSection = true;
    }
    failingItems.push(block.trim());
  });

  // Fallback: if any failing item had no section named (e.g. a RE-ANCHOR item
  // covered nowhere in the article), default it to the LAST section in the
  // article — scope-boundary and "stays within X" statements typically belong
  // there, and it guarantees the item is actually sent to the LLM for a fix
  // rather than silently dropped.
  if (itemsMissingSection) {
    var allSectionIds = [];
    var sidRe = /<section[^>]*\bid=["'](section-[\w-]+)["']/gi;
    var sidM;
    while ((sidM = sidRe.exec(html)) !== null) {
      if (allSectionIds.indexOf(sidM[1]) === -1) allSectionIds.push(sidM[1]);
    }
    if (allSectionIds.length > 0) {
      var fallbackSid = allSectionIds[allSectionIds.length - 1];
      if (sectionIds.indexOf(fallbackSid) === -1) sectionIds.push(fallbackSid);
    }
  }

  if (sectionIds.length === 0) {
    return 'ERROR: Could not identify specific failing section ids from the saved result. The audit result may predate the SECTION field — re-run the compliance check first.';
  }

  // Extract only the implicated <section id="..."> blocks from the full HTML
  var extractedSections = [];
  sectionIds.forEach(function(sid) {
    var re = new RegExp('<section[^>]*\\bid=["\']' + sid + '["\'][^>]*>[\\s\\S]*?<\\/section>', 'i');
    var m = html.match(re);
    if (m) extractedSections.push(m[0]);
  });

  if (extractedSections.length === 0) {
    return 'ERROR: Could not find the failing section(s) (' + sectionIds.join(', ') + ') in the article HTML by id.';
  }

  const prompt = `
  STAGE 2B — REWRITE BRIEF FIX (SECTION-SCOPED)
  ROLE: Senior UK SEO content editor for a natural stone floor restoration website.

  --- REWRITE BRIEF (must be satisfied) ---
  ${rewriteBrief}

  --- FAILING ITEMS FOR THE SECTIONS BELOW ONLY ---
  ${failingItems.join('\n\n')}

--- TASK ---
    You are not rewriting these sections generally.

    Treat each failed item above as a locked correction instruction.

    For each RE-ANCHOR ITEM marked "COVERED: NO":
    - Add only the minimum substantive coverage needed to satisfy that specific item.
    - Integrate it into the most relevant existing paragraph in the supplied section where possible.
    - Do not create a new standalone topic unless the failed item cannot be covered naturally within existing prose.

    For each DRIFT ITEM marked "STATUS: FAIL — DEVELOPED":
    - Treat the quoted EVIDENCE as content that MUST disappear from the corrected section.
    - If a sentence or paragraph primarily discusses the failed DRIFT topic, DELETE that entire sentence or paragraph. Do not rewrite, soften, summarise or paraphrase it.
    - A failed DRIFT topic must have ZERO substantive advice remaining after the correction.
    - Do not retain practical instructions, product recommendations, maintenance intervals, cleaning methods, aftercare advice or explanatory detail relating to that DRIFT topic.
    - Before returning the corrected section, search your own output for the failed DRIFT topic and its practical equivalents. If substantive coverage remains anywhere in the section, remove it before returning the section.
    - If the sentence also contains necessary in-scope information, rewrite only that sentence so the drift topic is completely removed.
    - Do NOT replace the removed wording with synonyms, paraphrases, examples, advice or another form of the same out-of-scope topic.
    - After editing, verify that the section no longer develops, explains, recommends or advises on the failed DRIFT topic anywhere.
    - Preserve an internal link only when it routes the reader away from the drift topic without giving that topic substantive coverage in this article.

    CRITICAL RULES:
    1. Do NOT return complete corrected sections.
    2. Return ONLY the exact text replacements required to resolve the failed items.
    3. Each "old" value must be copied VERBATIM from the supplied section HTML.
    4. Each "old" value must contain the smallest unique fragment needed for a safe replacement.
    5. Before returning a fix, verify that its "old" value appears exactly once in the supplied section HTML.
    6. The "new" value must contain only the minimum corrected replacement.
    7. If a DRIFT sentence or paragraph should be removed completely, return that exact sentence or paragraph as "old" and "" as "new".
    8. Do not alter any wording that is not required to resolve a failed RE-ANCHOR or DRIFT item.
    9. Do not create replacement wording that introduces another version of the same failed DRIFT topic.
    10. Preserve all existing links, HTML tags, figures, headings and factual claims unless the failed item specifically requires changing them.
    11. If a failed item cannot be corrected safely using an exact replacement from the supplied HTML, omit that fix rather than inventing one.
    12. Return ONLY a JSON array in this exact structure:

    [
      {
        "fixLabel": "short description",
        "old": "exact text copied from supplied HTML",
        "new": "minimum corrected replacement"
      }
    ]

    13. No markdown fences.
    14. No explanation before or after the JSON.
    15. Escape every double quote inside JSON string values as \\".
    16. ${ce_getConcreteLanguageStandard()}
    17. ${ce_getFixBatchIntegrityCheck()}
    18. ${ce_getContentPreservationRules()}
    19. ${ce_getSourceEvidencePriorityLock()}
    20. ${ce_getCustomerFacingCopyOnlyLock()}
    21. ${ce_getProjectGroundedMaintenanceLock()}
    22. ${ce_getLocationVarietyWithoutFabricationLock()}

    SECTION(S) TO FIX:
    ${extractedSections.join('\n\n')}

    Return the JSON array only.
    `.trim();

  return prompt;
}

function passThroughRewriteBriefCheckToGG() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  var html = String(sh.getRange(row, 188).getValue() || '').trim(); // GF = column 188 (W2B.1 Fixed HTML, Case Studies)
  var sourceLabel = 'GF';
  if (!html) {
    html = String(sh.getRange(row, 187).getValue() || '').trim(); // fallback: GE = column 187 (W2B.05 Fixed HTML — used when W2B.1 Similarity was skipped)
    sourceLabel = 'GE';
  }
  if (!html) {
    return { success: false, message: "ERROR: Column GF (W2B.1 Fixed HTML) and column GE (W2B.05 Fixed HTML) are both empty for this row." };
  }

  sh.getRange(row, 189).setValue(html); // GG = column 189
  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B.2 — Rewrite Brief Check PASS, copied " + sourceLabel + " to GG", "");
  }

  return { success: true, message: "No issues found. " + sourceLabel + " copied unchanged to column GG (W2B.2 Fixed HTML)." };
}

function promoteW2B2ToEU() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const html = String(sh.getRange(row, 189).getValue() || '').trim(); // GG = column 189
  if (!html) {
    return { success: false, message: "ERROR: Column GG (W2B.2 Fixed HTML) is empty for this row. Run W2B.2 first." };
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                     .map(function(h) { return String(h).trim(); });
  const colIdx = headers.indexOf('Rewrite Brief Compliance Check W2B');
  if (colIdx > -1) {
    const savedResult = String(sh.getRange(row, colIdx + 1).getValue() || '').trim();
    if (savedResult && savedResult.indexOf('PASS') !== 0) {
      return { success: false, message: "ERROR: Saved Rewrite Brief Compliance Check result is not a PASS. Resolve remaining failures before promoting to EU." };
    }
  }

  sh.getRange(row, 151).setValue(html); // EU = column 151
  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B.2 — PASS confirmed, promoted GG to EU", "");
  }

  return { success: true, message: "Confirmed PASS. Column GG copied to EU (W2B Raw HTML) — ready for the next stage." };
}