/**
 * ce_Stage2_3_CoherenceCheck.gs - W2B.3 WHOLE-ARTICLE COHERENCE & DEPTH CHECK
 * Runs after W2B.2's output (column GG) is ready. Reads the full, assembled
 * article and checks for cross-paragraph redundancy, thinness left behind by
 * fact-check removals, and links stranded in abrupt sentences by earlier
 * fixes. This is the only stage that reads the ENTIRE assembled article at
 * once specifically looking for repetition and depth gaps — none of
 * W2B.05/W2B.1/W2B.2 are chartered to catch this.
 *
 * Read pattern: prefers GI (a previous coherence fix already applied) and
 * falls back to GG (W2B.2 output) only on the first run for a row, so a
 * second round of fixes builds on the first instead of re-checking the
 * unfixed original.
 */



function saveCoherenceCheckResult(raw) {
  if (!raw || !raw.trim()) {
    return { success: false, message: 'No response provided.' };
  }

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();
  if (row < 2) return { success: false, message: 'Select a data row first.' };

  const passed = /OVERALL:\s*PASS/i.test(raw);
  const failed = /OVERALL:\s*FAIL/i.test(raw);

  if (!passed && !failed) {
    return { success: false, message: 'Could not find OVERALL result line in response.' };
  }

  const previousGH = String(sh.getRange(row, 190).getValue() || '').trim(); // GH
  const runMatch = previousGH.match(/^RUN\s+(\d+)/i);
  const runNumber = (runMatch ? parseInt(runMatch[1], 10) : 0) + 1;

  const summary = 'RUN ' + runNumber + ' — ' + (passed ? 'PASS' : 'FAIL') + ' — ' + raw.trim();
  const cell = sh.getRange(row, 190); // GH
  cell.setNumberFormat('@');
  cell.setValue(summary);

  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B.3 — Coherence Check Saved (GH)", "");
  }

  return {
    success: true,
    passed: passed,
    message: passed
      ? 'No coherence issues found — use the PASS button to copy through to GI.'
      : 'Coherence issues found — saved to column GH. Build the fix prompt next.',
    rawResult: raw.trim()
  };
}

function buildCoherenceCheckPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const headers =
    sh
      .getRange(
        1,
        1,
        1,
        sh.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });

  const colPostId =
    headers.indexOf('Post ID');

  const colTitle = 1; // column B

  if (colPostId === -1) {
    return {
      success: false,
      message:
        'ERROR: Post ID column not found.'
    };
  }

  const rowData =
    sh
      .getRange(
        row,
        1,
        1,
        sh.getLastColumn()
      )
      .getValues()[0];

  const postId =
    String(
      rowData[colPostId] || ''
    )
      .trim()
      .replace(/\.0$/, '');

  const title =
    String(
      rowData[colTitle] || ''
    ).trim();

  /*
   * W2B.3 operates on the current article.
   *
   * GI = W2B.3 working/final HTML.
   * GG = W2B.2 completed HTML.
   *
   * Prefer GI when it has already been initialised
   * by the automated runner.
   */
  var html =
    String(
      sh.getRange(row, 191).getValue() || ''
    ).trim();

  if (!html) {
    html =
      String(
        sh.getRange(row, 189).getValue() || ''
      ).trim();
  }

  if (!postId) {
    return {
      success: false,
      message:
        'ERROR: Post ID is empty for this row.'
    };
  }

  if (!html) {
    return {
      success: false,
      message:
        'ERROR: No article HTML found in column GI or GG. Run W2B.2 first.'
    };
  }

  /*
   * Original source HTML is used only to support
   * genuine THIN SECTION additions.
   */
  let originalHtml = '';

  try {

    const exportSheet =
      ss.getSheetByName('site-export');

    if (exportSheet) {

      const exportData =
        exportSheet
          .getDataRange()
          .getValues();

      const exportHeaders =
        exportData[0].map(function(h) {
          return String(h).trim();
        });

      const exportIdCol =
        exportHeaders.indexOf('ID');

      const exportHtmlCol =
        exportHeaders.indexOf(
          'Full Post HTML'
        );

      if (
        exportIdCol > -1 &&
        exportHtmlCol > -1
      ) {

        for (
          let i = 1;
          i < exportData.length;
          i++
        ) {

          const exportPostId =
            String(
              exportData[i][exportIdCol] ||
              ''
            )
              .trim()
              .replace(/\.0$/, '');

          if (
            exportPostId === postId
          ) {

            originalHtml =
              String(
                exportData[i][exportHtmlCol] ||
                ''
              ).trim();

            break;
          }
        }
      }
    }

  } catch (e) {

    /*
     * Non-fatal.
     * REDUNDANCY and STRANDED LINK checks
     * can still run without the source.
     */
    originalHtml = '';
  }

  const prompt = `
    W2B.3 — SINGLE-PASS WHOLE-ARTICLE COHERENCE AUDIT

    ROLE:
    Audit the complete ARTICLE HTML once.

    Your task is to identify ALL material coherence issues currently present in the article in this ONE audit.

    There will be NO iterative coherence checking after individual repairs.

    Therefore you MUST inspect the entire article now and return every qualifying issue you can identify.

    Do NOT stop after finding the first issue.

    Do NOT return one issue merely because it is the clearest issue.

    Do NOT assume another audit will discover remaining issues later.

    ==================================================
    ARTICLE HTML
    ==================================================

    ${html}

    ${
      originalHtml
        ? `
    ==================================================
    ORIGINAL SOURCE ARTICLE
    ==================================================

    Use this ONLY when evaluating a THIN SECTION.

    It may be used to identify genuine source-supported detail that was lost during previous editing.

    Do NOT use it to introduce unrelated material.

    ${originalHtml}
    `
        : ''
    }

    ==================================================
    ISSUES TO IDENTIFY
    ==================================================

    Identify ALL qualifying issues belonging to these three types only:

    1. REDUNDANCY

    Find separate paragraphs or sections that communicate materially the same underlying fact, explanation, diagnostic point or maintenance advice.

    Different wording does NOT make repeated meaning unique.

    For each REDUNDANCY issue:

    - identify the section that should change;
    - identify the fullest or best-placed explanation that should remain;
    - identify exactly what repeated meaning should be removed, shortened or merged;
    - do not remove unique information;
    - do not create a new explanation merely to replace a repeated one.

    2. THIN SECTION

    Identify a section that is materially under-developed compared with the role it is trying to perform.

    Do NOT flag a section merely because it is short.

    There must be a genuine explanatory gap.

    If ORIGINAL SOURCE ARTICLE is supplied:

    - additions must be supported by genuine information in that source;
    - do not invent facts, procedures, figures, examples or claims;
    - identify exactly what supported information may be restored.

    If no ORIGINAL SOURCE ARTICLE is supplied:

    - do NOT invent factual additions;
    - only flag a THIN SECTION where the missing explanatory requirement is evident from ARTICLE HTML itself.

    3. STRANDED LINK

    Identify an internal <a href="..."> link whose surrounding sentence is abrupt, generic or disconnected from the reason the reader would follow the link.

    For each STRANDED LINK:

    - identify the section containing the link;
    - preserve the exact href;
    - preserve the existing anchor text unless changing it is absolutely necessary to the stated repair;
    - repair only the contextual sentence around the link.

    ==================================================
    WHOLE-ARTICLE AUDIT — HARD RULES
    ==================================================

    1. Read the entire ARTICLE HTML before deciding the issue list.

    2. Return ALL qualifying coherence issues discovered in this one audit.

    3. Do NOT deliberately hold back an issue for a later run.

    4. Do NOT split one underlying problem into several artificial issues.

    5. Do NOT duplicate the same issue under different wording.

    6. Every issue must target ONE section_number that will be changed.

    7. A different section may be identified as locked_section_number where its fuller or correct treatment must remain unchanged.

    8. Do NOT invent section numbers.

    9. section_number must correspond to an actual governed article section.

    10. issue_type must be EXACTLY one of:

    REDUNDANCY
    THIN SECTION
    STRANDED LINK

    11. action must be EXACTLY one of:

    SHORTEN
    REMOVE
    MERGE
    EXPAND
    REPAIR_LINK

    12. Action compatibility:

    REDUNDANCY:
    SHORTEN
    REMOVE
    MERGE

    THIN SECTION:
    EXPAND

    STRANDED LINK:
    REPAIR_LINK

    13. "instruction" must describe the exact repair required.

    14. "preserve" must identify important existing content or purpose in the target section that must remain.

    15. Do NOT propose changes to headings.

    16. Do NOT propose changes to unrelated links.

    17. Do NOT propose stylistic polishing merely because wording could be improved.

    18. Do NOT conduct another fact check.

    19. Do NOT conduct another rewrite-brief compliance check.

    20. Do NOT create issues outside these three coherence categories.

    ==================================================
    REPAIR-SAFETY RULE
    ==================================================

    Each proposed repair must be designed so it can be applied without knowingly creating another coherence problem.

    Before including an issue, consider the proposed repair in the context of the WHOLE article.

    For REDUNDANCY:
    ensure the retained location still contains the necessary explanation.

    For REMOVE:
    ensure removing the targeted material does not remove information required by later text.

    For SHORTEN:
    ensure the unique purpose of the paragraph remains.

    For MERGE:
    ensure the repair does not create another duplicate explanation.

    For EXPAND:
    ensure the proposed addition does not duplicate information already adequately covered elsewhere.

    For REPAIR_LINK:
    ensure the repair keeps the existing link naturally connected to the surrounding subject.

    This is planning within the SAME audit.

    Do NOT perform a separate validation pass.

    ==================================================
    OUTPUT — JSON ONLY
    ==================================================

    Return exactly ONE JSON object.

    If NO coherence issues exist:

    {
      "status": "PASS",
      "issues": []
    }

    If one or more coherence issues exist:

    {
      "status": "FAIL",
      "issues": [
        {
          "section_number": 2,
          "issue_type": "REDUNDANCY",
          "action": "SHORTEN",
          "instruction": "Exact description of what must change.",
          "preserve": "Exact description of what must remain.",
          "locked_section_number": 1
        },
        {
          "section_number": 5,
          "issue_type": "THIN SECTION",
          "action": "EXPAND",
          "instruction": "Exact description of the supported information that may be added.",
          "preserve": "Existing purpose and information that must remain.",
          "locked_section_number": null
        }
      ]
    }

    ==================================================
    JSON HARD LOCK
    ==================================================

    - Return valid JSON only.
    - No markdown fences.
    - No commentary before the JSON.
    - No commentary after the JSON.
    - Top-level keys must be exactly:
      "status"
      "issues"
    - "status" must be exactly "PASS" or "FAIL".
    - "issues" must always be an array.
    - PASS requires an empty issues array.
    - FAIL requires at least one issue.
    - Every issue must contain exactly:
      "section_number"
      "issue_type"
      "action"
      "instruction"
      "preserve"
      "locked_section_number"
    - section_number must be an integer greater than 0.
    - locked_section_number must be an integer greater than 0 or null.
    - Do not return duplicate issues.
    - Do not return additional keys.

    FINAL SELF-CHECK BEFORE OUTPUT:

    Confirm internally that:

    - the whole article was reviewed;
    - every qualifying issue found is included;
    - no issue was intentionally left for a later audit;
    - every section_number exists;
    - every issue_type is permitted;
    - every action matches its issue_type;
    - no proposed repair knowingly duplicates or removes necessary information;
    - the final response is valid JSON.

    Return JSON only.
    `.trim();

  return {
    success: true,
    prompt: prompt,
    postId: postId,
    title: title,
    runNumber: 1
  };
}



function applyCoherenceFix(rawJson) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sh =
    ss.getSheetByName('posts');

  const row =
    sh.getActiveCell().getRow();

  if (
    !rawJson ||
    !String(rawJson).trim()
  ) {
    return {
      success: false,
      message:
        'No coherence repair response provided.'
    };
  }

  /*
   * =========================================================
   * STEP 1 — NORMALISE MODEL OUTPUT
   * =========================================================
   */

  let cleaned =
    String(rawJson)
      .trim()
      .replace(
        /^```json\s*/i,
        ''
      )
      .replace(
        /^```\s*/i,
        ''
      )
      .replace(
        /\s*```$/i,
        ''
      )
      .trim();

  /*
   * Recover the outer JSON object if the model
   * accidentally adds text around it.
   */
  const firstBrace =
    cleaned.indexOf('{');

  const lastBrace =
    cleaned.lastIndexOf('}');

  if (
    firstBrace > -1 &&
    lastBrace > firstBrace
  ) {
    cleaned =
      cleaned.substring(
        firstBrace,
        lastBrace + 1
      );
  }

  let parsed;

  try {

    parsed =
      JSON.parse(cleaned);

  } catch (e) {

    return {
      success: false,
      message:
        'Could not parse coherence repair JSON — ' +
        e.message
    };
  }

  /*
   * =========================================================
   * STEP 2 — VALIDATE TOP-LEVEL STRUCTURE
   * =========================================================
   */

  if (
    !parsed ||
    !Array.isArray(parsed.repairs)
  ) {
    return {
      success: false,
      message:
        'Coherence repair response must contain a "repairs" array.'
    };
  }

  if (
    parsed.repairs.length === 0
  ) {
    return {
      success: false,
      message:
        'Coherence repair response contained no repairs.'
    };
  }

  /*
   * =========================================================
   * STEP 3 — LOAD CURRENT ARTICLE
   * =========================================================
   */

  let html =
    String(
      sh.getRange(row, 191).getValue() || ''
    );

  if (!html) {
    html =
      String(
        sh.getRange(row, 189).getValue() || ''
      );
  }

  if (!html) {
    return {
      success: false,
      message:
        'No article HTML found in column GI or GG.'
    };
  }

  /*
   * Work against a separate candidate copy.
   *
   * Nothing is written to GI until EVERY repair
   * has been validated successfully.
   */
  let candidateHtml = html;

  const seenSections = {};

  const validatedRepairs = [];

  const brokenLinkPattern =
    /<a\s+href="\[https?:\/\/[^\]]+\]\(/i;

  /*
   * =========================================================
   * STEP 4 — VALIDATE EVERY SECTION REPAIR FIRST
   * =========================================================
   */

  for (
    let i = 0;
    i < parsed.repairs.length;
    i++
  ) {

    const repair =
      parsed.repairs[i] || {};

    const sectionNumber =
      Number(
        repair.section_number
      );

    const replacementHtml =
      typeof repair.replacement_html ===
      'string'
        ? repair.replacement_html.trim()
        : '';

    if (
      !Number.isInteger(
        sectionNumber
      ) ||
      sectionNumber < 1
    ) {
      return {
        success: false,
        message:
          'Invalid section_number in coherence repair.'
      };
    }

    if (
      seenSections[
        sectionNumber
      ]
    ) {
      return {
        success: false,
        message:
          'Section ' +
          sectionNumber +
          ' was returned more than once.'
      };
    }

    seenSections[
      sectionNumber
    ] = true;

    if (!replacementHtml) {
      return {
        success: false,
        message:
          'Section ' +
          sectionNumber +
          ' has no replacement_html.'
      };
    }

    /*
     * Reject accidental markdown-in-HTML links.
     */
    if (
      brokenLinkPattern.test(
        replacementHtml
      )
    ) {
      return {
        success: false,
        message:
          'Section ' +
          sectionNumber +
          ' contains malformed markdown-in-HTML link syntax.'
      };
    }

    /*
     * =====================================================
     * FIND THE EXISTING GOVERNED SECTION
     * =====================================================
     */

    const sectionRegex =
      new RegExp(
        '<section\\b[^>]*\\bid=["\']section-' +
        sectionNumber +
        '["\'][^>]*>[\\s\\S]*?<\\/section>',
        'i'
      );

    const existingMatch =
      html.match(sectionRegex);

    if (!existingMatch) {
      return {
        success: false,
        message:
          'Section ' +
          sectionNumber +
          ' does not exist in the current article.'
      };
    }

    const existingSection =
      existingMatch[0];

    /*
     * =====================================================
     * VERIFY REPLACEMENT TARGETS THE SAME SECTION
     * =====================================================
     */

    const replacementIdRegex =
      new RegExp(
        '^<section\\b[^>]*\\bid=["\']section-' +
        sectionNumber +
        '["\'][^>]*>',
        'i'
      );

    if (
      !replacementIdRegex.test(
        replacementHtml
      )
    ) {
      return {
        success: false,
        message:
          'Replacement for section ' +
          sectionNumber +
          ' does not preserve its section id.'
      };
    }

    if (
      !/<\/section>\s*$/i.test(
        replacementHtml
      )
    ) {
      return {
        success: false,
        message:
          'Replacement for section ' +
          sectionNumber +
          ' is not a complete section block.'
      };
    }

    /*
     * =====================================================
     * H2 HARD LOCK
     * =====================================================
     */

    const existingH2Match =
      existingSection.match(
        /<h2\b[^>]*>[\s\S]*?<\/h2>/i
      );

    const replacementH2Match =
      replacementHtml.match(
        /<h2\b[^>]*>[\s\S]*?<\/h2>/i
      );

    if (
      !existingH2Match ||
      !replacementH2Match
    ) {
      return {
        success: false,
        message:
          'Section ' +
          sectionNumber +
          ' is missing its governed H2.'
      };
    }

    if (
      existingH2Match[0] !==
      replacementH2Match[0]
    ) {
      return {
        success: false,
        message:
          'Section ' +
          sectionNumber +
          ' attempted to change its governed H2.'
      };
    }

    /*
     * Store the validated replacement.
     *
     * We deliberately do NOT apply anything yet.
     */
    validatedRepairs.push({
      sectionNumber:
        sectionNumber,

      existingSection:
        existingSection,

      replacementHtml:
        replacementHtml
    });
  }

  /*
   * =========================================================
   * STEP 5 — APPLY ALL VALIDATED REPAIRS ATOMICALLY
   * =========================================================
   */

  validatedRepairs.forEach(
    function(repair) {

      const currentSectionRegex =
        new RegExp(
          '<section\\b[^>]*\\bid=["\']section-' +
          repair.sectionNumber +
          '["\'][^>]*>[\\s\\S]*?<\\/section>',
          'i'
        );

      candidateHtml =
        candidateHtml.replace(
          currentSectionRegex,
          function() {
            return repair.replacementHtml;
          }
        );
    }
  );

  /*
   * =========================================================
   * STEP 6 — FINAL STRUCTURAL SANITY CHECK
   * =========================================================
   */

  for (
    let i = 0;
    i < validatedRepairs.length;
    i++
  ) {

    const sectionNumber =
      validatedRepairs[i]
        .sectionNumber;

    const finalSectionRegex =
      new RegExp(
        '<section\\b[^>]*\\bid=["\']section-' +
        sectionNumber +
        '["\'][^>]*>[\\s\\S]*?<\\/section>',
        'gi'
      );

    const matches =
      candidateHtml.match(
        finalSectionRegex
      ) || [];

    if (
      matches.length !== 1
    ) {
      return {
        success: false,
        message:
          'Section ' +
          sectionNumber +
          ' failed final structural validation.'
      };
    }
  }

  /*
   * =========================================================
   * STEP 7 — WRITE ONCE
   * =========================================================
   */

  sh
    .getRange(row, 191)
    .setValue(
      candidateHtml
    );

  if (
    typeof logPipelineResume ===
    'function'
  ) {
    logPipelineResume(
      'W2B.3 — Single-Pass Coherence Repairs Applied (GI)',
      ''
    );
  }

  return {
    success: true,
    message:
      validatedRepairs.length +
      ' section repair(s) applied atomically to GI.'
  };
}

function passThroughCoherenceCheckToGI() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  // Prefer GI (a previous fix already applied) — fall back to GG on first run
  var html = String(sh.getRange(row, 191).getValue() || '').trim(); // GI
  if (!html) html = String(sh.getRange(row, 189).getValue() || '').trim(); // GG

  if (!html) {
    return { success: false, message: "ERROR: No article HTML found in column GI or GG for this row." };
  }

  sh.getRange(row, 191).setValue(html); // GI
  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B.3 — Coherence Check PASS, confirmed in GI", "");
  }

  return { success: true, message: "No issues found. Confirmed in column GI (W2B.3 Final HTML)." };
}

function promoteCoherenceToEU() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const html = String(sh.getRange(row, 191).getValue() || '').trim(); // GI
  if (!html) {
    return { success: false, message: "ERROR: Column GI (W2B.3 Final HTML) is empty for this row. Run W2B.3 first." };
  }

  sh.getRange(row, 151).setValue(html); // EU
  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W2B.3 — Confirmed, promoted GI to EU", "");
  }

  return { success: true, message: "Column GI copied to EU (W2B Raw HTML) — this is now the source of truth for the article." };
}

function buildCoherenceFixPrompt(findingsText) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sh =
    ss.getSheetByName('posts');

  const row =
    sh.getActiveCell().getRow();

  /*
   * GI is the W2B.3 working copy.
   * Fall back to GG only if GI is empty.
   */
  var html =
    String(
      sh.getRange(row, 191).getValue() || ''
    ).trim();

  if (!html) {
    html =
      String(
        sh.getRange(row, 189).getValue() || ''
      ).trim();
  }

  if (!html) {
    return {
      success: false,
      message:
        'ERROR: No article HTML found in column GI or GG.'
    };
  }

  if (
    !findingsText ||
    !String(findingsText).trim()
  ) {
    return {
      success: false,
      message:
        'ERROR: No coherence audit JSON provided.'
    };
  }

  /*
   * Validate the audit JSON before giving it
   * to the repair model.
   */
  var auditJson;

  try {

    auditJson =
      JSON.parse(
        String(findingsText).trim()
      );

  } catch (e) {

    return {
      success: false,
      message:
        'ERROR: Coherence audit is not valid JSON — ' +
        e.message
    };
  }

  if (
    !auditJson ||
    String(
      auditJson.status || ''
    ).toUpperCase() !== 'FAIL' ||
    !Array.isArray(
      auditJson.issues
    ) ||
    auditJson.issues.length === 0
  ) {

    return {
      success: false,
      message:
        'ERROR: Coherence audit does not contain any repairable issues.'
    };
  }

  const prompt = `
    W2B.3 — SINGLE-PASS COHERENCE REPAIR

    ROLE:
    Apply ALL coherence repairs listed in the locked audit below.

    This is the ONLY repair pass.

    There will be NO repair loop and NO coherence recheck.

    You must therefore resolve every listed issue in this one response.

    ==================================================
    LOCKED COHERENCE AUDIT
    ==================================================

    ${JSON.stringify(auditJson, null, 2)}

    ==================================================
    ARTICLE HTML
    ==================================================

    ${html}

    ==================================================
    CORE EXECUTION RULE
    ==================================================

    The audit above is LOCKED.

    You are NOT diagnosing the article again.

    You are NOT allowed to:

    - add another issue;
    - omit a listed issue;
    - reinterpret the issue type;
    - change the requested action;
    - repair an unrelated weakness;
    - perform stylistic polishing outside the listed repairs.

    Apply only the repairs already authorised by the audit.

    ==================================================
    SECTION REPLACEMENT METHOD
    ==================================================

    For every affected section:

    1. Locate the complete existing governed section in ARTICLE HTML.

    2. A governed section has the form:

    <section id="section-N">
    ...
    </section>

    where N is the section_number.

    3. Return the COMPLETE corrected <section>...</section> block.

    4. Do NOT return old/new fragments.

    5. Do NOT return partial paragraphs.

    6. Do NOT return surrounding sections.

    7. Preserve the existing section id exactly.

    8. Preserve the existing H2 exactly.

    9. Preserve all unaffected content inside that section.

    10. Preserve all unaffected links, href values, images, figures, facts, locations and numbers.

    ==================================================
    MULTIPLE ISSUES IN THE SAME SECTION
    ==================================================

    More than one audit issue may target the same section_number.

    If that happens:

    - apply ALL authorised repairs for that section together;
    - return the section only ONCE;
    - combine the repairs into one final replacement_html;
    - do not return multiple competing versions of the same section.

    Therefore the number of returned repairs may be LOWER than the number of audit issues.

    There must be exactly ONE replacement per affected section.

    ==================================================
    LOCKED SECTIONS
    ==================================================

    If an issue contains a locked_section_number:

    - that locked section must remain completely unchanged;
    - do not return a replacement for the locked section unless that same section is independently targeted by another audit issue;
    - preserve the fuller or correct explanation identified there.

    ==================================================
    REDUNDANCY
    ==================================================

    For a REDUNDANCY issue:

    - apply only SHORTEN, REMOVE or MERGE as specified;
    - remove the repeated meaning identified by the audit;
    - preserve unique information;
    - preserve the intended role of the target section;
    - do not merely paraphrase the repetition;
    - do not create replacement text that repeats the same meaning in different words;
    - do not expand a redundancy repair.

    ==================================================
    THIN SECTION
    ==================================================

    For a THIN SECTION issue:

    - apply only the EXPAND instruction supplied by the audit;
    - use only information authorised by that instruction;
    - do not invent facts, examples, numbers, products, procedures or claims;
    - do not duplicate an explanation already adequately present elsewhere;
    - preserve the existing section purpose.

    ==================================================
    STRANDED LINK
    ==================================================

    For a STRANDED LINK issue:

    - repair only the context required by the audit;
    - preserve the exact href;
    - preserve the existing anchor text unless the audit explicitly requires otherwise;
    - do not move the link to another section;
    - do not convert HTML links to markdown.

    ==================================================
    HTML HARD LOCKS
    ==================================================

    For every replacement_html:

    - it must begin with the existing opening <section ...> tag;
    - it must end with </section>;
    - its section id must be exactly the original id;
    - its H2 must remain exactly unchanged;
    - do not add markdown fences;
    - do not add markdown links;
    - do not remove required internal links;
    - do not alter unrelated HTML structure;
    - do not include <html>, <body> or document wrappers.

    ==================================================
    OUTPUT FORMAT — JSON ONLY
    ==================================================

    Return exactly:

    {
      "repairs": [
        {
          "section_number": 2,
          "replacement_html": "<section id=\\"section-2\\">...</section>"
        },
        {
          "section_number": 5,
          "replacement_html": "<section id=\\"section-5\\">...</section>"
        }
      ]
    }

    ==================================================
    JSON HARD LOCK
    ==================================================

    - Return valid JSON only.
    - No markdown fences.
    - No explanation before the JSON.
    - No explanation after the JSON.
    - Top-level key must be exactly "repairs".
    - "repairs" must be an array.
    - Every repair object must contain exactly:
      "section_number"
      "replacement_html"
    - section_number must be an integer greater than 0.
    - Return each affected section_number only once.
    - Do not return an unaffected section.
    - Do not omit an affected section.
    - replacement_html must contain the COMPLETE corrected section.
    - replacement_html must use the same section id as section_number.

    ==================================================
    FINAL SELF-CHECK
    ==================================================

    Before returning:

    - confirm every audit issue has been applied;
    - confirm issues targeting the same section have been combined;
    - confirm every affected section appears exactly once;
    - confirm no unrelated section was changed;
    - confirm locked sections were respected;
    - confirm every H2 is unchanged;
    - confirm required links are preserved;
    - confirm the response is valid JSON.

    Return JSON only.
    `.trim();

  return {
    success: true,
    prompt: prompt
  };
}