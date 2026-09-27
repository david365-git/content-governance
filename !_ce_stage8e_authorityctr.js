/**
 * ce_Stage8E_AuthorityCTR.gs - W8E AUTHORITY & CTR ANALYSIS
 * Builds an audit prompt from the finished article's sheet columns and
 * saves the pasted LLM response back to the row.
 */

function buildW8EAnalysisPrompt() {

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const colHtml   = 'CT';
  const colH1     = 'CK';
  const colTitle  = 'CL';
  const colMeta   = 'CM';
  const colSchema = 'CN';
  const colPst    = 'CO';
  const colSerp   = 'FU';

  function getVal(col) {
    return sh.getRange(col + row).getValue();
  }

  const html      = getVal(colHtml);
  const h1        = getVal(colH1);
  const title     = getVal(colTitle);
  const meta      = getVal(colMeta);
  const schema    = getVal(colSchema);
  const pst       = getVal(colPst);
  const serpNotes = getVal(colSerp);

  if (!html) {
    return {
      success: false,
      message: "ERROR: New HTML (CT) is empty for this row."
    };
  }

  const prompt =
    `You are auditing a single finished article for Google authority potential,
    silo recovery, click-through rate and first-click satisfaction.

    The article is part of a post-migration content recovery process.

    Primary Search Term: ${pst}
    H1: ${h1}
    Meta Title: ${title}
    Meta Description: ${meta}

    Competitor SERP Notes:
    ${serpNotes || "No competitor SERP notes available for this row."}

    Schema (JSON-LD):
    ${schema}

    Article HTML:
    ${html}

    Return JSON ONLY.
    Do not use markdown fences.
    Do not include prose before or after the JSON.

    Use this exact top-level structure:

    {
      "overall_status": "PASS" | "PASS_WITH_FIXES" | "NEEDS_AUTHOR_INPUT",
      "auto_fixable_count": 0,
      "author_input_count": 0,
      "checks": [
        {
          "section": "GOOGLE AUTHORITY SIGNALS",
          "check": "Experience Proofing",
          "status": "PASS" | "PARTIAL" | "FAIL",
          "evidence": "specific evidence from the article, or 'not present'",
          "action_type": "NONE" | "AUTO-FIXABLE" | "NEEDS AUTHOR INPUT",
          "action": "specific action required, or empty string if none"
        }
      ]
    }

    Run these checks:

    GOOGLE AUTHORITY SIGNALS
    - Experience Proofing
    - Expertise Depth
    - Entity Alignment
    - Trust Signals

    SILO RECOVERY CHECK
    - Silo Identity Signal
    - Internal Link Architecture
    - PST / Scope Alignment

    CLICK-THROUGH RATE POTENTIAL
    - Title Tag
    - Meta Description
    - SERP Differentiation
    - Snippet Capture

    SEARCH INTENT & FIRST-CLICK SATISFACTION
    - Intent Match
    - Immediate Value
    - Actionable Takeaways

    RULES

    1. For every check, return one object in "checks".

    2. "status" must be exactly:
    PASS
    PARTIAL
    FAIL

    3. "action_type" must be exactly:
    NONE
    AUTO-FIXABLE
    NEEDS AUTHOR INPUT

    4. If status is PASS:
    - action_type must be NONE
    - action must be an empty string

    5. If status is PARTIAL or FAIL:
    - action_type must be AUTO-FIXABLE or NEEDS AUTHOR INPUT
    - action must state exactly what needs changing

    6. Use AUTO-FIXABLE only when the issue can be corrected safely from the existing article and supplied context without inventing facts.

    Examples:
    - improve Meta Title wording
    - improve Meta Description
    - tighten an existing explanation
    - add a self-contained snippet answer using facts already present
    - soften an absolute claim
    - improve an existing internal-link anchor

    7. Use NEEDS AUTHOR INPUT only when the fix depends on facts that are not present in the supplied material.

    Examples:
    - first-hand project detail
    - genuine experience claim
    - missing factual date
    - unverified technical claim
    - missing real-world observation

    8. Do not invent evidence, experience, project facts or competitor information.

    9. For Trust Signals:
    Schema datePublished/dateModified is sufficient freshness evidence.
    Do not request a visible update date in the article body.

    10. For SERP Differentiation:
    Use the supplied Competitor SERP Notes if available.
    If they are empty, say so explicitly.

    11. For each PARTIAL or FAIL action:
    state the exact field or article area affected.

    Examples:
    - Meta Title
    - Meta Description
    - first paragraph
    - H2 "Common Causes"
    - paragraph immediately below H2 "..."

    12. The "action" field must be concrete enough for an automated correction stage to understand.

    13. Determine overall_status as follows:

    PASS
    = every check is PASS.

    PASS_WITH_FIXES
    = one or more checks are PARTIAL or FAIL,
    but every required action is AUTO-FIXABLE.

    NEEDS_AUTHOR_INPUT
    = at least one required action is NEEDS AUTHOR INPUT.

    14. auto_fixable_count must equal the number of checks whose action_type is AUTO-FIXABLE.

    15. author_input_count must equal the number of checks whose action_type is NEEDS AUTHOR INPUT.

    16. Do not include duplicate fixes for the same underlying problem.

    17. Do not score, rank or give percentages.

    Return valid JSON only.`;

  return {
    success: true,
    prompt: prompt
  };
}

function buildW8EAutoFixPrompt() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const analysis = sh.getRange('FZ' + row).getValue();
  const html     = sh.getRange('CT' + row).getValue();

  if (!analysis) {
    return { success: false, message: "ERROR: No W8E analysis saved yet (column FZ is empty). Run W8E Analysis first." };
  }
  if (!html) {
    return { success: false, message: "ERROR: New HTML (CT) is empty for this row." };
  }

  const lines = String(analysis).split('\n');
  const autoItems = lines.filter(function(l) {
    return /^\s*\[AUTO-FIXABLE\]/i.test(l.trim());
  });

  if (autoItems.length === 0) {
    return { success: false, message: "No [AUTO-FIXABLE] items found in the saved analysis. Nothing to fix automatically." };
  }

  const prompt =
    `You are applying a fixed set of pre-approved text edits to an existing article. Do NOT return the whole article.

    For each fix below, return ONLY the exact original fragment being replaced (OLD) and its exact replacement (NEW). Each OLD fragment must be copied verbatim from the article HTML below — exact characters, exact tags — so it can be located by an automated find-and-replace. Each OLD fragment must be unique within the article (include enough surrounding text to make it unique if needed).

    Do not touch anything not listed in the fixes below.

    Return ONLY a JSON array, no markdown fences, no preamble, in this exact format:
    [
      {"fixLabel": "short label for this fix", "field": "html", "old": "exact original fragment", "new": "exact replacement fragment"}
    ]

    The "field" value must be one of: "html" (article body), "h1" (New H1), "metaTitle" (New Meta Title), "metaDescription" (New Meta Description), "schema" (Schema JSON-LD). Choose the field the fix actually targets — a title tag fix must use "metaTitle", a meta description fix must use "metaDescription", not "html". If a fix could apply to the article body, default to "html".

    CRITICAL JSON FORMATTING RULE: The old/new fragments contain HTML with double-quote attributes (e.g. style="..."). Every double quote character inside the old/new string values MUST be escaped as \\" so the result is valid JSON. Do not use curly/smart quotes (’ " ") anywhere — use straight quotes only, and escape them inside strings. Test that your output is valid JSON before returning it.

    FIXES TO APPLY:
    ${autoItems.join('\n')}

    ARTICLE HTML:
    ${html}`;

  return { success: true, prompt: prompt, count: autoItems.length };
}

function applyW8EAutoFix(rawJson) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  if (!rawJson || !rawJson.trim()) {
    return { success: false, message: "No response provided." };
  }

  let fixes;
  let cleaned = rawJson.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  try {
    fixes = JSON.parse(cleaned);
  } catch (e) {
    try {
      const repaired = cleaned.replace(/"(old|new|fixLabel|field)":\s*"([\s\S]*?)"(?=\s*[,}])/g, function(match, key, value) {
        const fixedValue = value.replace(/(?<!\\)"/g, '\\"');
        return '"' + key + '": "' + fixedValue + '"';
      });
      fixes = JSON.parse(repaired);
    } catch (e2) {
      return { success: false, message: "Could not parse response as JSON: " + e.message + " (repair attempt also failed: " + e2.message + ")" };
    }
  }

  if (!Array.isArray(fixes) || fixes.length === 0) {
    return { success: false, message: "No fixes found in parsed response." };
  }

  const FIELD_COLUMNS = {
    html: 'CT',
    h1: 'CK',
    metaTitle: 'CL',
    metaDescription: 'CM',
    schema: 'CN'
  };

  // Cache current values per column so multiple fixes to the same field apply in sequence.
  const cache = {};
  function getFieldValue(col) {
    if (!(col in cache)) {
      cache[col] = String(sh.getRange(col + row).getValue() || '');
    }
    return cache[col];
  }

  const applied = [];
  const failed = [];
  const touchedCols = {};

  fixes.forEach(function(fix) {
    const label = fix.fixLabel || '(unlabeled fix)';
    const fieldKey = FIELD_COLUMNS.hasOwnProperty(fix.field) ? fix.field : 'html';
    const col = FIELD_COLUMNS[fieldKey];
    const oldText = fix.old;
    const newText = fix.new;

    if (!oldText) { failed.push(label + ' — no OLD text provided'); return; }

    let current = getFieldValue(col);
    if (!current) {
      failed.push(label + ' — ' + fieldKey + ' (' + col + ') is empty for this row');
      return;
    }

    const occurrences = current.split(oldText).length - 1;
    if (occurrences === 0) {
      failed.push(label + ' — OLD fragment not found in ' + fieldKey + ' (' + col + ')');
      return;
    }
    if (occurrences > 1) {
      failed.push(label + ' — OLD fragment matches ' + occurrences + ' places in ' + fieldKey + ', skipped for safety');
      return;
    }

    cache[col] = current.replace(oldText, newText);
    touchedCols[col] = true;
    applied.push(label + ' (' + fieldKey + ')');
  });

  Object.keys(touchedCols).forEach(function(col) {
    sh.getRange(col + row).setValue(cache[col]);
  });

  if (applied.length > 0 && typeof logPipelineResume === 'function') {
    logPipelineResume("W8E — Auto-Fix Applied", "");
  }

  let message = applied.length + '/' + fixes.length + ' fix(es) applied.';
  if (applied.length > 0) message += '\nApplied: ' + applied.join(', ');
  if (failed.length > 0) message += '\nSkipped: ' + failed.join(' | ');

  return { success: applied.length > 0, message: message };
}

function saveW8EAnalysisResult(resultText) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  if (!resultText || !resultText.trim()) {
    return { success: false, message: "No result text provided." };
  }

  sh.getRange('FZ' + row).setValue(resultText.trim());

  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W8E — Authority & CTR Analysis", "");
  }

  return { success: true, message: "Saved to column FZ (W8E Authority & CTR Analysis)." };
}