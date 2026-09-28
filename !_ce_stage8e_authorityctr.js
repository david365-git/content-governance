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

  const html      = String(getVal(colHtml) || '').trim();
  const h1        = String(getVal(colH1) || '').trim();
  const title     = String(getVal(colTitle) || '').trim();
  const meta      = String(getVal(colMeta) || '').trim();
  const schema    = String(getVal(colSchema) || '').trim();
  const pst       = String(getVal(colPst) || '').trim();
  const serpNotes = String(getVal(colSerp) || '').trim();

  if (!html) {
    return {
      success: false,
      message: "ERROR: New HTML (CT) is empty for this row."
    };
  }


  const prompt =
    `You are performing the FINAL pre-publication audit of one finished article.

    Your job has TWO parts in ONE response:

    1. Audit the finished page for authority, silo recovery, CTR and first-click satisfaction.
    2. For every issue that can safely be corrected from the supplied material, return the exact deterministic replacement required.

    Do NOT create a second-stage repair plan.
    Do NOT ask another model to interpret your recommendations later.
    The fixes you return here will be applied directly by script.

    PRIMARY SEARCH TERM:
    ${pst}

    CURRENT H1:
    ${h1}

    CURRENT META TITLE:
    ${title}

    CURRENT META DESCRIPTION:
    ${meta}

    COMPETITOR SERP NOTES:
    ${serpNotes || "No competitor SERP notes available for this row."}

    CURRENT SCHEMA:
    ${schema}

    CURRENT ARTICLE HTML:
    ${html}


    ============================================================
    OUTPUT
    ============================================================

    Return VALID JSON ONLY.

    No markdown fences.
    No commentary before or after the JSON.

    Use exactly this top-level structure:

    {
      "overall_status": "PASS" | "PASS_WITH_FIXES" | "NEEDS_AUTHOR_INPUT",
      "auto_fixable_count": 0,
      "author_input_count": 0,
      "checks": [
        {
          "section": "GOOGLE AUTHORITY SIGNALS",
          "check": "Experience Proofing",
          "status": "PASS" | "PARTIAL" | "FAIL",
          "evidence": "specific evidence from the supplied page, or 'not present'",
          "action_type": "NONE" | "AUTO-FIXABLE" | "NEEDS AUTHOR INPUT",
          "action": "specific required action, or empty string"
        }
      ],
      "fixes": [
        {
          "fixLabel": "short description",
          "field": "html",
          "old": "exact current fragment",
          "new": "exact replacement fragment"
        }
      ]
    }


    ============================================================
    THE 14 REQUIRED CHECKS
    ============================================================

    GOOGLE AUTHORITY SIGNALS
    1. Experience Proofing
    2. Expertise Depth
    3. Entity Alignment
    4. Trust Signals

    SILO RECOVERY CHECK
    5. Silo Identity Signal
    6. Internal Link Architecture
    7. PST / Scope Alignment

    CLICK-THROUGH RATE POTENTIAL
    8. Title Tag
    9. Meta Description
    10. SERP Differentiation
    11. Snippet Capture

    SEARCH INTENT & FIRST-CLICK SATISFACTION
    12. Intent Match
    13. Immediate Value
    14. Actionable Takeaways


    ============================================================
    AUDIT RULES
    ============================================================

    1. Return exactly 14 check objects.

    2. Each governed check must appear exactly once.

    3. status must be exactly one of:
    PASS
    PARTIAL
    FAIL

    4. action_type must be exactly one of:
    NONE
    AUTO-FIXABLE
    NEEDS AUTHOR INPUT

    5. If status is PASS:
    - action_type must be NONE
    - action must be ""

    6. If status is PARTIAL or FAIL:
    - action_type must be AUTO-FIXABLE or NEEDS AUTHOR INPUT
    - action must state exactly what needs changing

    7. AUTO-FIXABLE means the correction can be made safely using ONLY information already present in the supplied article, schema, metadata or competitor notes.

    8. NEEDS AUTHOR INPUT means the correction requires a fact that is not safely established in the supplied material.

    Examples include:
    - first-hand experience
    - actual project history
    - exact experience duration
    - business/service scope that is contradictory or unverified
    - factual claims requiring confirmation

    9. Never invent first-hand experience, projects, business capabilities, service coverage, dates or outcomes.

    10. Trust Signals:
    datePublished/dateModified in Schema is sufficient freshness evidence.
    Never request a visible article-body update date.

    11. SERP Differentiation:
    use the supplied Competitor SERP Notes when available.
    Do not invent competitor evidence.

    12. Do not score or rank the article.

    13. Do not recommend stylistic rewrites merely because wording could be prettier.

    14. Only flag a problem when it materially affects authority, semantic alignment, search intent, CTR, factual reliability or user usefulness.



    ============================================================
    ATOMIC FIX RULES
    ============================================================

    For EVERY check marked AUTO-FIXABLE, normally return the corresponding safe replacement in the top-level "fixes" array.

    Each fix must use this exact form:

    {
      "fixLabel": "short description",
      "field": "html" | "h1" | "metaTitle" | "metaDescription" | "schema",
      "old": "exact existing fragment",
      "new": "exact replacement fragment"
    }

    The script will perform literal find-and-replace.

    Therefore:

    1. "old" MUST be copied verbatim from the supplied CURRENT field.

    2. Never paraphrase OLD.

    3. OLD must occur exactly once in that field.
    If necessary, include enough surrounding text or HTML to make it unique.

    4. NEW must make the smallest safe change necessary.

    5. Do not rewrite unrelated material.

    6. Do not return a fix for NEEDS AUTHOR INPUT.

    7. Do not use an automated fix to guess around an unresolved author-input issue.

    8. If an apparently AUTO-FIXABLE issue cannot actually be expressed as a safe exact replacement, classify it as NEEDS AUTHOR INPUT rather than returning an unsafe fix.

    9. For field "html":
    preserve existing HTML structure unless the identified issue genuinely requires a structural addition or replacement.

    10. For field "h1":
    OLD must equal the complete current H1.

    11. For field "metaTitle":
    OLD must equal the complete current Meta Title.
    NEW must be no more than 60 characters.

    12. For field "metaDescription":
    OLD must equal the COMPLETE current Meta Description field exactly as supplied.

    If the field contains:
    Yoast Keyphrase: ...

    preserve that line unchanged unless the issue explicitly concerns the keyphrase.

    The actual Meta Description prose after the fix must remain 140-160 characters and exactly two sentences.

    13. For field "schema":
    use the smallest unique exact JSON or JSON-LD fragment that can safely be replaced.
    Do not rewrite the whole schema unless there is no smaller safe replacement.

    14. Every object in "fixes" must correspond to an AUTO-FIXABLE check.

    15. Do not return duplicate fixes.

    16. If an AUTO-FIXABLE action genuinely requires two separate replacements, two fix objects are allowed for that check.

    17. The "fixes" array must be empty when there are no AUTO-FIXABLE issues.

    18. EXPERIENCE DURATION GOVERNANCE:

    The verified experience fact is:
    more than 30 years.

    Equivalent natural wording is allowed where it suits the surrounding sentence, including:
    - more than 30 years
    - over 30 years
    - 30 years and more
    - 30+ years
    - in excess of 30 years

    Treat these as semantically equivalent.

    Do not flag variation between these confirmed equivalents as a contradiction.

    If the page contains weaker or conflicting wording such as:
    - nearly 30 years
    - almost 30 years
    - around 30 years
    - approximately 30 years

    classify that wording as AUTO-FIXABLE and replace it with a natural equivalent of "more than 30 years" that fits the surrounding text.

    Do not require author input for this experience-duration fact.


    ============================================================
    OVERALL STATUS
    ============================================================

    Set overall_status to:

    PASS
    when every check passes.

    PASS_WITH_FIXES
    when one or more issues exist but every required action is safely AUTO-FIXABLE.

    NEEDS_AUTHOR_INPUT
    when at least one issue requires author input.

    auto_fixable_count must equal the number of CHECKS marked AUTO-FIXABLE.

    author_input_count must equal the number of CHECKS marked NEEDS AUTHOR INPUT.

    Return JSON only.`;

  return {
    success: true,
    prompt: prompt
  };
}

function buildW8EAutoFixPrompt() {

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const analysisRaw =
    String(
      sh.getRange('FZ' + row).getValue() || ''
    ).trim();

  const html =
    String(
      sh.getRange('CT' + row).getValue() || ''
    ).trim();

  const h1 =
    String(
      sh.getRange('CK' + row).getValue() || ''
    ).trim();

  const metaTitle =
    String(
      sh.getRange('CL' + row).getValue() || ''
    ).trim();

  const metaDescription =
    String(
      sh.getRange('CM' + row).getValue() || ''
    ).trim();

  const schema =
    String(
      sh.getRange('CN' + row).getValue() || ''
    ).trim();


  if (!analysisRaw) {
    return {
      success: false,
      message:
        "No W8E analysis saved yet in column FZ."
    };
  }


  let analysis;

  try {

    analysis =
      JSON.parse(
        analysisRaw
      );

  } catch (e) {

    return {
      success: false,
      message:
        "Saved W8E analysis is not valid JSON."
    };
  }


  if (
    !Array.isArray(
      analysis.checks
    )
  ) {

    return {
      success: false,
      message:
        "Saved W8E analysis contains no checks array."
    };
  }


  const autoItems =
    analysis.checks.filter(
      function(item) {

        return (
          item &&
          item.action_type ===
            "AUTO-FIXABLE"
        );
      }
    );


  if (
    autoItems.length === 0
  ) {

    return {
      success: false,
      message:
        "No AUTO-FIXABLE W8E items found."
    };
  }


  const fixesJson =
    JSON.stringify(
      autoItems.map(
        function(item) {

          return {
            section:
              item.section || "",

            check:
              item.check || "",

            action:
              item.action || ""
          };
        }
      ),
      null,
      2
    );


  const prompt =
    `You are applying a fixed set of pre-approved W8E corrections to an existing finished article.

    Do NOT reassess the article.
    Do NOT introduce additional improvements.
    Do NOT change anything that is not required by the supplied AUTO-FIXABLE actions.
    Do NOT act on anything marked NEEDS AUTHOR INPUT.

    AUTO-FIXABLE ACTIONS:
    ${fixesJson}

    CURRENT FIELDS

    H1:
    ${h1}

    META TITLE:
    ${metaTitle}

    META DESCRIPTION:
    ${metaDescription}

    SCHEMA:
    ${schema}

    ARTICLE HTML:
    ${html}

    TASK

    For each AUTO-FIXABLE action, produce the smallest safe replacement necessary.

    Return JSON ONLY.

    Use this exact structure:

    [
      {
        "fixLabel": "short description",
        "field": "html",
        "old": "exact existing fragment",
        "new": "exact replacement fragment"
      }
    ]

    FIELD must be exactly one of:

    html
    h1
    metaTitle
    metaDescription
    schema

    RULES

    1. OLD must be copied exactly from the supplied current field.

    2. OLD must uniquely identify the text being replaced.

    3. NEW must make only the requested correction.

    4. Do not change unrelated wording.

    5. Do not invent facts.

    6. Do not resolve any NEEDS AUTHOR INPUT issue.

    7. If an AUTO-FIXABLE action depends on unresolved author information, omit that fix rather than guessing.

    8. For HTML fixes:
    preserve all existing HTML structure unless the action specifically requires a structural change.

    9. For schema fixes:
    return only the exact schema fragment that needs changing, not the complete schema unless absolutely necessary.

    10. For Meta Title or Meta Description:
    OLD must be the complete current field value and NEW must be the complete replacement field value.

    11. Meta Title must remain no more than 60 characters.

    12. Meta Description must remain 140-160 characters.

    13. Preserve the existing Yoast Keyphrase line in the Meta Description field unless the fix explicitly concerns it.

    14. Return a valid JSON array only.

    15. No markdown fences.
    No explanation.
    No commentary.`;

  return {
    success: true,
    prompt: prompt,
    count: autoItems.length
  };
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

  if (!resultText || !String(resultText).trim()) {
    return {
      success: false,
      message: "No W8E result provided."
    };
  }

  let cleaned =
    String(resultText)
      .trim()
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/\s*```$/i, '')
      .trim();

  let parsed;

  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    return {
      success: false,
      message:
        "W8E result is not valid JSON: " +
        e.message
    };
  }


  /*
   * =========================================================
   * TOP-LEVEL VALIDATION
   * =========================================================
   */

  const allowedOverall = [
    "PASS",
    "PASS_WITH_FIXES",
    "NEEDS_AUTHOR_INPUT"
  ];

  if (
    allowedOverall.indexOf(
      parsed.overall_status
    ) === -1
  ) {
    return {
      success: false,
      message:
        "W8E overall_status is invalid."
    };
  }

  if (
    !Array.isArray(
      parsed.checks
    )
  ) {
    return {
      success: false,
      message:
        "W8E checks must be an array."
    };
  }

  if (
    !Array.isArray(
      parsed.fixes
    )
  ) {
    return {
      success: false,
      message:
        "W8E fixes must be an array."
    };
  }


  /*
   * =========================================================
   * EXPECT EXACTLY 14 GOVERNED CHECKS
   * =========================================================
   */

  const expectedChecks = [
    "Experience Proofing",
    "Expertise Depth",
    "Entity Alignment",
    "Trust Signals",
    "Silo Identity Signal",
    "Internal Link Architecture",
    "PST / Scope Alignment",
    "Title Tag",
    "Meta Description",
    "SERP Differentiation",
    "Snippet Capture",
    "Intent Match",
    "Immediate Value",
    "Actionable Takeaways"
  ];

  if (
    parsed.checks.length !==
    expectedChecks.length
  ) {
    return {
      success: false,
      message:
        "W8E returned " +
        parsed.checks.length +
        " checks; expected " +
        expectedChecks.length +
        "."
    };
  }


  /*
   * =========================================================
   * VALIDATE CHECKS
   * =========================================================
   */

  const allowedStatus = [
    "PASS",
    "PARTIAL",
    "FAIL"
  ];

  const allowedActionType = [
    "NONE",
    "AUTO-FIXABLE",
    "NEEDS AUTHOR INPUT"
  ];

  let autoCount = 0;
  let authorCount = 0;
  const seenChecks = {};
  const autoCheckNames = {};

  for (
    let i = 0;
    i < parsed.checks.length;
    i++
  ) {

    const item =
      parsed.checks[i] || {};

    const checkName =
      String(
        item.check || ""
      ).trim();

    if (
      expectedChecks.indexOf(
        checkName
      ) === -1
    ) {
      return {
        success: false,
        message:
          "Unexpected W8E check: " +
          checkName
      };
    }

    if (
      seenChecks[
        checkName
      ]
    ) {
      return {
        success: false,
        message:
          "Duplicate W8E check: " +
          checkName
      };
    }

    seenChecks[
      checkName
    ] = true;

    if (
      allowedStatus.indexOf(
        item.status
      ) === -1
    ) {
      return {
        success: false,
        message:
          "Invalid status for " +
          checkName +
          "."
      };
    }

    if (
      allowedActionType.indexOf(
        item.action_type
      ) === -1
    ) {
      return {
        success: false,
        message:
          "Invalid action_type for " +
          checkName +
          "."
      };
    }

    const evidence =
      String(
        item.evidence || ""
      ).trim();

    const action =
      String(
        item.action || ""
      ).trim();

    if (!evidence) {
      return {
        success: false,
        message:
          "Missing evidence for " +
          checkName +
          "."
      };
    }

    if (
      item.status === "PASS"
    ) {

      if (
        item.action_type !==
        "NONE"
      ) {
        return {
          success: false,
          message:
            "PASS check " +
            checkName +
            " must use action_type NONE."
        };
      }

      if (action) {
        return {
          success: false,
          message:
            "PASS check " +
            checkName +
            " must have an empty action."
        };
      }
    }

    if (
  item.status === "PARTIAL" ||
  item.status === "FAIL"
    ) {

      if (
        item.action_type ===
        "NONE"
      ) {
        return {
          success: false,
          message:
            checkName +
            " requires an action_type."
        };
      }

      if (!action) {
        return {
          success: false,
          message:
            checkName +
            " requires an action."
        };
      }


      /*
      * =========================================================
      * TRUST SIGNALS FALSE-POSITIVE GUARD
      *
      * If schema already contains datePublished or dateModified,
      * W8E must not ask the author to supply publication dates.
      * =========================================================
      */

      if (
        checkName ===
        "Trust Signals" &&
        item.action_type ===
        "NEEDS AUTHOR INPUT"
      ) {

        const currentSchema =
          String(
            sh.getRange(
              'CN' + row
            ).getValue() || ''
          );


        const hasSchemaDate =
          /"datePublished"\s*:/i.test(
            currentSchema
          ) ||
          /"dateModified"\s*:/i.test(
            currentSchema
          );


        const asksForDate =
          /publication date|published date|datePublished|modification date|modified date|dateModified/i.test(
            action
          );


        if (
          hasSchemaDate &&
          asksForDate
        ) {

          return {
            success: false,

            message:
              "Trust Signals incorrectly requested author-supplied publication dates even though schema already contains datePublished/dateModified."
          };
        }
      }
    }

    if (
      item.action_type ===
      "AUTO-FIXABLE"
    ) {
      autoCount++;

      autoCheckNames[
        checkName
      ] = true;
    }

    if (
      item.action_type ===
      "NEEDS AUTHOR INPUT"
    ) {
      authorCount++;
    }
  }


  /*
   * =========================================================
   * VERIFY COUNTS
   * =========================================================
   */

  if (
    Number(
      parsed.auto_fixable_count
    ) !==
    autoCount
  ) {
    return {
      success: false,
      message:
        "W8E auto_fixable_count does not match the checks."
    };
  }

  if (
    Number(
      parsed.author_input_count
    ) !==
    authorCount
  ) {
    return {
      success: false,
      message:
        "W8E author_input_count does not match the checks."
    };
  }


  /*
   * =========================================================
   * VERIFY OVERALL STATUS
   * =========================================================
   */

  let expectedOverall;

  if (
    authorCount > 0
  ) {
    expectedOverall =
      "NEEDS_AUTHOR_INPUT";

  } else if (
    autoCount > 0
  ) {
    expectedOverall =
      "PASS_WITH_FIXES";

  } else {
    expectedOverall =
      "PASS";
  }

  if (
    parsed.overall_status !==
    expectedOverall
  ) {
    return {
      success: false,
      message:
        "W8E overall_status should be " +
        expectedOverall +
        ", not " +
        parsed.overall_status +
        "."
    };
  }


  /*
   * =========================================================
   * CURRENT FIELD VALUES FOR ATOMIC FIX VALIDATION
   * =========================================================
   */

  const currentFields = {
    html:
      String(
        sh.getRange(
          'CT' + row
        ).getValue() || ''
      ),

    h1:
      String(
        sh.getRange(
          'CK' + row
        ).getValue() || ''
      ),

    metaTitle:
      String(
        sh.getRange(
          'CL' + row
        ).getValue() || ''
      ),

    metaDescription:
      String(
        sh.getRange(
          'CM' + row
        ).getValue() || ''
      ),

    schema:
      String(
        sh.getRange(
          'CN' + row
        ).getValue() || ''
      )
  };


  /*
   * =========================================================
   * VALIDATE ATOMIC FIX OBJECTS
   * =========================================================
   */

  const allowedFields = [
    "html",
    "h1",
    "metaTitle",
    "metaDescription",
    "schema"
  ];

  const seenFixes = {};

  for (
    let i = 0;
    i < parsed.fixes.length;
    i++
  ) {

    const fix =
      parsed.fixes[i] || {};

    const fixLabel =
      String(
        fix.fixLabel || ""
      ).trim();

    const field =
      String(
        fix.field || ""
      ).trim();

    const oldText =
      String(
        fix.old || ""
      );

    const newText =
      String(
        fix.new || ""
      );

    if (!fixLabel) {
      return {
        success: false,
        message:
          "W8E fix " +
          (i + 1) +
          " has no fixLabel."
      };
    }

    if (
      allowedFields.indexOf(
        field
      ) === -1
    ) {
      return {
        success: false,
        message:
          "W8E fix " +
          fixLabel +
          " has invalid field: " +
          field
      };
    }

    if (!oldText) {
      return {
        success: false,
        message:
          "W8E fix " +
          fixLabel +
          " has an empty OLD value."
      };
    }

    /*
 * Empty NEW is allowed only when deliberately
 * removing an HTML or schema fragment.
 */

    if (
      !newText &&
      field !== "html" &&
      field !== "schema"
    ) {
      return {
        success: false,
        message:
          "W8E fix " +
          fixLabel +
          " has an empty NEW value."
      };
    }

    if (
      oldText === newText
    ) {
      return {
        success: false,
        message:
          "W8E fix " +
          fixLabel +
          " does not change anything."
      };
    }


    /*
     * OLD must exist exactly once
     */

    const fieldValue =
      currentFields[
        field
      ];

    const firstIndex =
      fieldValue.indexOf(
        oldText
      );

    if (
      firstIndex === -1
    ) {
      return {
        success: false,
        message:
          "W8E fix " +
          fixLabel +
          " OLD text was not found in " +
          field +
          "."
      };
    }

    const secondIndex =
      fieldValue.indexOf(
        oldText,
        firstIndex +
        oldText.length
      );

    if (
      secondIndex !== -1
    ) {
      return {
        success: false,
        message:
          "W8E fix " +
          fixLabel +
          " OLD text is not unique in " +
          field +
          "."
      };
    }


    /*
     * COMPLETE-FIELD LOCKS
     */

    if (
      field === "h1" ||
      field === "metaTitle" ||
      field === "metaDescription"
    ) {

      if (
        oldText !==
        fieldValue
      ) {
        return {
          success: false,
          message:
            "W8E fix " +
            fixLabel +
            " must use the complete current " +
            field +
            " as OLD."
        };
      }
    }


    /*
     * TITLE LENGTH
     */

    if (
      field === "metaTitle" &&
      newText.length > 60
    ) {
      return {
        success: false,
        message:
          "W8E Meta Title fix exceeds 60 characters."
      };
    }


    /*
     * META DESCRIPTION VALIDATION
     */

    if (
      field === "metaDescription"
    ) {

      const oldYoastMatch =
        fieldValue.match(
          /\n\s*Yoast Keyphrase:\s*.+$/i
        );

      const newYoastMatch =
        newText.match(
          /\n\s*Yoast Keyphrase:\s*.+$/i
        );

      if (
        oldYoastMatch &&
        (
          !newYoastMatch ||
          newYoastMatch[0] !==
          oldYoastMatch[0]
        )
      ) {
        return {
          success: false,
          message:
            "W8E Meta Description fix did not preserve the Yoast Keyphrase line."
        };
      }


      const descriptionOnly =
        newText
          .replace(
            /\n\s*Yoast Keyphrase:\s*.+$/i,
            ''
          )
          .trim();


      if (
        descriptionOnly.length < 140 ||
        descriptionOnly.length > 160
      ) {
        return {
          success: false,
          message:
            "W8E Meta Description fix must be 140-160 characters."
        };
      }


      const sentenceCount =
        (
          descriptionOnly.match(
            /[.!?](?:\s|$)/g
          ) || []
        ).length;


      if (
        sentenceCount !== 2
      ) {
        return {
          success: false,
          message:
            "W8E Meta Description fix must contain exactly two sentences."
        };
      }
    }


    /*
     * DUPLICATE FIX PROTECTION
     */

    const fixKey =
      field +
      "||" +
      oldText;

    if (
      seenFixes[
        fixKey
      ]
    ) {
      return {
        success: false,
        message:
          "Duplicate W8E atomic fix detected: " +
          fixLabel
      };
    }

    seenFixes[
      fixKey
    ] = true;
  }


  /*
   * =========================================================
   * FIX COVERAGE CHECK
   *
   * If there are AUTO-FIXABLE checks there must be at least
   * one atomic fix.
   * =========================================================
   */

  if (
    autoCount > 0 &&
    parsed.fixes.length === 0
  ) {
    return {
      success: false,
      message:
        "W8E contains AUTO-FIXABLE checks but returned no atomic fixes."
    };
  }

  if (
    autoCount === 0 &&
    parsed.fixes.length > 0
  ) {
    return {
      success: false,
      message:
        "W8E returned atomic fixes when no checks are AUTO-FIXABLE."
    };
  }


  /*
   * =========================================================
   * SAVE NORMALISED JSON TO FZ
   * =========================================================
   */

  const normalised =
    JSON.stringify(
      parsed
    );

  sh.getRange(
    'FZ' + row
  ).setValue(
    normalised
  );


  return {
    success: true,

    message:
      "W8E JSON and atomic fixes validated and saved to column FZ.",

    overallStatus:
      parsed.overall_status,

    autoFixableCount:
      autoCount,

    authorInputCount:
      authorCount,

    fixCount:
      parsed.fixes.length,

    checks:
      parsed.checks,

    fixes:
      parsed.fixes
  };
}

function resolveW8EExperienceDuration() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sh =
    ss.getSheetByName('posts');

  const row =
    sh.getActiveCell().getRow();

  const confirmedText =
    'more than 30 years';


  /*
   * =========================================================
   * UPDATE FINAL HTML — CT
   * =========================================================
   */

  const htmlCell =
    sh.getRange(
      'CT' + row
    );

  let html =
    String(
      htmlCell.getValue() || ''
    );


  if (!html) {
    return {
      success: false,
      message:
        'Column CT is empty.'
    };
  }


  const conflictingPattern =
    /nearly 30 years/gi;


  const matches =
    html.match(
      conflictingPattern
    );


  if (
    matches &&
    matches.length
  ) {

    html =
      html.replace(
        conflictingPattern,
        confirmedText
      );

    htmlCell.setValue(
      html
    );
  }


  /*
   * =========================================================
   * UPDATE SAVED W8E JSON — FZ
   *
   * Remove ONLY the now-resolved experience-duration issue.
   * Do not invent or clear any genuine first-hand example issue.
   * =========================================================
   */

  const fzCell =
    sh.getRange(
      'FZ' + row
    );

  const raw =
    String(
      fzCell.getValue() || ''
    ).trim();


  let remainingAuthorItems = 0;


  if (raw) {

    try {

      const data =
        JSON.parse(
          raw
        );


      if (
        Array.isArray(
          data.checks
        )
      ) {

        data.checks.forEach(
          function(item) {

            if (
              !item ||
              item.check !==
              'Experience Proofing'
            ) {
              return;
            }


            /*
             * The experience duration is now confirmed.
             *
             * If the old action also mentioned adding a
             * first-hand example, retain only that unresolved
             * part rather than falsely marking the whole
             * check as PASS.
             */

            const oldAction =
              String(
                item.action || ''
              );


            const exampleRequired =
              /first-hand|first hand|repair example|project example|verifiable.*example/i.test(
                oldAction
              );


            if (
              exampleRequired
            ) {

              item.status =
                'PARTIAL';

              item.action_type =
                'NEEDS AUTHOR INPUT';

              item.action =
                'Add a genuine, verifiable first-hand ceramic tile repair or grout-restoration example only if one is available and appropriate to publish. Do not invent one.';

              item.evidence =
                'The conflicting experience-duration wording has been resolved to "more than 30 years", but no verified first-hand project example has been supplied.';

            } else {

              item.status =
                'PASS';

              item.action_type =
                'NONE';

              item.action =
                '';

              item.evidence =
                'The experience-duration wording is now consistent at "more than 30 years".';
            }
          }
        );


        /*
         * Recalculate counts from the actual checks.
         */

        const autoCount =
          data.checks.filter(
            function(item) {
              return (
                item &&
                item.action_type ===
                'AUTO-FIXABLE'
              );
            }
          ).length;


        const authorCount =
          data.checks.filter(
            function(item) {
              return (
                item &&
                item.action_type ===
                'NEEDS AUTHOR INPUT'
              );
            }
          ).length;


        data.auto_fixable_count =
          autoCount;

        data.author_input_count =
          authorCount;


        if (
          authorCount > 0
        ) {

          data.overall_status =
            'NEEDS_AUTHOR_INPUT';

        } else if (
          autoCount > 0
        ) {

          data.overall_status =
            'PASS_WITH_FIXES';

        } else {

          data.overall_status =
            'PASS';
        }


        remainingAuthorItems =
          authorCount;


        fzCell.setValue(
          JSON.stringify(
            data
          )
        );
      }

    } catch (e) {

      return {
        success: false,
        message:
          'Experience wording was updated, but FZ could not be updated: ' +
          e.message
      };
    }
  }


  return {

    success: true,

    replacements:
      matches
        ? matches.length
        : 0,

    remainingAuthorItems:
      remainingAuthorItems,

    message:
      '✔ Experience wording confirmed as "' +
      confirmedText +
      '". ' +
      (
        remainingAuthorItems
          ? remainingAuthorItems +
            ' author-input issue(s) remain.'
          : 'No author-input issues remain.'
      )
  };
}