/**
 * ce_Stage8E_AuthorityCTR.gs - W8E AUTHORITY & CTR ANALYSIS
 * Builds an audit prompt from the finished article's sheet columns and
 * saves the pasted LLM response back to the row.
 */

function buildW8EAnalysisPrompt() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const colHtml  = 'CT';
  const colH1    = 'CK';
  const colTitle = 'CL';
  const colMeta  = 'CM';
  const colSchema= 'CN';
  const colPst   = 'CO';
  const colSerp  = 'FU';

  function getVal(col) {
    return sh.getRange(col + row).getValue();
  }

  const html   = getVal(colHtml);
  const h1     = getVal(colH1);
  const title  = getVal(colTitle);
  const meta   = getVal(colMeta);
  const schema = getVal(colSchema);
  const pst    = getVal(colPst);
  const serpNotes = getVal(colSerp);

  if (!html) {
    return { success: false, message: "ERROR: New HTML (CT) is empty for this row." };
  }

  const prompt =
`You are auditing a single article for Google authority potential and click-through rate, as part of a post-migration content recovery project. Before May 2025 this article lived in a material-specific silo (e.g. /travertine/, /slate/); after a URL migration it was moved to a flat /home-garden/ structure, which caused a loss of topical authority. Your job is to check whether this specific article, as written, recovers that authority.

Primary Search Term: ${pst}
H1: ${h1}
Meta Title: ${title}
Meta Description: ${meta}

Competitor SERP Notes (from prior research):
${serpNotes || "No competitor SERP notes available for this row."}

Schema (JSON-LD):
${schema}

Article HTML:
${html}

Assess against these sections. For every point, give a PASS/FAIL/PARTIAL and one sentence of concrete evidence quoted or paraphrased from the actual article — not a generic statement. If you cannot find evidence, say "not present" rather than inferring.

1. GOOGLE AUTHORITY SIGNALS
- Experience Proofing: specific first-person or case-study detail?
- Expertise Depth: technical terms used correctly and explained?
- Entity Alignment: does the article anchor clearly to its core material entity throughout?
- Trust Signals: factual claims verifiable or hedged? For freshness, check the schema's dateModified/datePublished fields — this is sufficient evidence of freshness on its own. NEVER suggest adding a visible "last updated" or "reviewed and updated" date line into the article body — WordPress displays this automatically, and adding one in the HTML would create a duplicate, potentially conflicting date. Do not flag the absence of an on-page date as a Trust Signals issue at all.

2. SILO RECOVERY CHECK
- Silo Identity Signal: would a reader know which material silo this belongs to from the first 100 words?
- Internal Link Architecture: links to a silo hub or sibling articles (name them if present)?
- PST/Scope Alignment: does the body content match the stated Primary Search Term and scope?

3. CLICK-THROUGH RATE POTENTIAL
- Title Tag: specific, gives a concrete reason to click?
- Meta Description: states a concrete answer/value, not a vague teaser?
- SERP Differentiation: use the Competitor SERP Notes provided above to name one specific way this title+meta would stand out against the actual competing pages listed there. Only say "no real SERP data available" if the Competitor SERP Notes section above is genuinely empty — do not say this if notes are present.
- Snippet Capture: is there a clean, self-contained answer block likely to win a featured snippet?

4. SEARCH INTENT & FIRST-CLICK SATISFACTION
- Intent Match: informational or commercial, matches likely searcher intent?
- Immediate Value: core answer without excessive scrolling?
- Actionable Takeaways: clear and specific to this material/defect?

5. FINAL ACTION LIST
Every PARTIAL or FAIL verdict from sections 1-4 above MUST have a corresponding entry here — do not leave any PARTIAL or FAIL unaddressed. List the specific text edits, link additions, or heading changes needed, each tied to which section flagged it. For any item that requires adding new content (not just editing existing text), you must also state: (a) the exact location in the article where it should be inserted (e.g. "after the second paragraph", "immediately below the H2 'Common Causes'"), and (b) example draft text the author can adapt, written in the article's own voice and specific to the material and defect discussed — not a generic placeholder.

Tag every item as either AUTO-FIXABLE or NEEDS AUTHOR INPUT, using this test: if the fix is a self-contained text edit, rewording, or addition that does not depend on facts only the author knows (e.g. title tag wording, meta description, snippet block, keyphrase punctuation, softening an absolute claim), tag it AUTO-FIXABLE. If the fix requires a specific fact, date, or first-hand detail only the author can supply (e.g. a real first-person observation, a genuine review date, a specific project detail), tag it NEEDS AUTHOR INPUT — and in these cases the example draft text is a starting point for the author to correct, not text to be applied as-is.

Format each item as:
[AUTO-FIXABLE] or [NEEDS AUTHOR INPUT] — [Section name]: [description of the fix, location, and example text as applicable]

Rules: do not invent competitor data you don't have. Do not comment on sentence rhythm, em-dashes, or "AI-sounding" vocabulary. If a section has no issue, say "PASS, no action needed." Output plain text only — no markdown tables, no bold/italic asterisks, no headers with # symbols. Use simple line breaks and a colon after each label, e.g. "Experience Proofing: PARTIAL — evidence here."`;

  return { success: true, prompt: prompt };
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