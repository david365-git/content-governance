/**
 * ================================================================================
 * ce_Stage15B_H2Generation.gs - STAGE 1.5B — H2 GENERATION
 * ================================================================================
 * ce_Stage15B_H2Generation
 * Article-type-aware H2 creation
 *
 * Updated: April 2026
 * - Integrated Strategic Compass: Reads Column AD (Page Rewrite Brief).
 * - Thematic Governance: Forces H2s to align with specific rewrite goals.
 * ================================================================================
 */
/* ============================================================
   STAGE 1.5B — REQUIRED ENTITY COVERAGE VERIFICATION
   Builds a prompt asking the LLM to confirm every item in the
   article type's Required Entity Coverage list is addressed by
   at least one generated H2. Judgment-based check, not mechanical.
============================================================ */
/* ============================================================
   STAGE 1.5B — REWRITE BRIEF COMPLIANCE CHECK (H2-LEVEL)
   Builds a prompt asking the LLM to confirm the generated H2s
   reflect the Rewrite Brief's re-anchor entities and avoid its
   named drift-prone language. Early, cheap checkpoint before
   full body content is written at W2B.
============================================================ */
function buildRewriteBriefComplianceCheckPromptW15B() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();
  if (row < 2) return 'ERROR: Select a data row (row 2 or below).';
  const generatedH2s = String(sheet.getRange(row, 100).getValue() || "").trim();
  if (!generatedH2s) return 'ERROR: No H2 headings found in column CV (100). Run and save Stage 1.5B first.';
  const d = getActiveRowDataMap();
  const rewriteBrief = String(d["Page Rewrite Brief"] || "").trim();
  if (!rewriteBrief || rewriteBrief === "Not Applicable") {
    return 'CHECK 1.5B-BRIEF: PASS\nNo active Rewrite Brief — compliance check not applicable.';
  }
  var priorityEntities = [];
  try {
    var blueprintEntities = bc_getRecoveryBlueprintEntities();
    if (
      blueprintEntities &&
      blueprintEntities.priority &&
      blueprintEntities.priority.length > 0
    ) {
      priorityEntities = blueprintEntities.priority.filter(function(name) {
        var n = String(name || "").trim().toLowerCase();
        return n &&
          n.indexOf("none scored") === -1 &&
          n !== "none";
      });
    }
  } catch (e) {
    priorityEntities = [];
  }
  var priorityScopeNote =
    priorityEntities.length > 0
      ? "RE-ANCHOR SCOPE (Hard Lock): Only check RE-ANCHOR reflection for these Priority-tier entities:\n" +
        priorityEntities.join(", ") + "\n" +
        "Any other entity or topic named in the Rewrite Brief's RE-ANCHOR FOCUS section is Supporting/Peripheral tier " +
        "and is NOT required to be reflected at H2 level — it belongs in body content, verified separately. " +
        "Do not produce a RE-ANCHOR ITEM line for anything outside the Priority list above.\n"
      : "RE-ANCHOR SCOPE (Hard Lock): This row has no Priority-tier Recovery Blueprint entities scored 3+. " +
        "Do NOT produce any RE-ANCHOR ITEM lines — output only DRIFT ITEM lines below.\n";
  const prompt = `
STAGE 1.5B — REWRITE BRIEF COMPLIANCE CHECK (H2 LEVEL)
ROLE: SEO Content Structure Auditor
--- REWRITE BRIEF ---
${rewriteBrief}
--- GENERATED H2 HEADINGS ---
${generatedH2s}
--- ${priorityScopeNote}---
--- TASK ---
The Rewrite Brief above names specific entities/topics the article must re-anchor around, and specific drift-prone language or framing that must be removed or avoided.
Judge the H2 headings only — not full body content, which does not exist yet.
1. RE-ANCHOR CHECK:
For each Priority-tier entity listed in the RE-ANCHOR SCOPE above (if any), state whether the H2 set reflects it — directly named, or clearly implied by heading topic — YES or NO — and which H2 if applicable.
Do not check any entity outside this scope.
2. DRIFT CHECK:
For each drift-prone term, framing, or topic the brief specifically says to remove/avoid, state whether any H2 heading reintroduces it — YES (drift present, this is a problem) or NO (clean).
Judge conservatively:
only flag drift where a heading genuinely reintroduces that topic as its own subject, not where a word merely appears in passing within a different context.
3. ${ce_getSourceEvidencePriorityLock()}
OUTPUT FORMAT (exactly):
RE-ANCHOR ITEM: [entity/topic text]
REFLECTED: YES / NO
BY: [H2 heading text, or "none"]
(repeat for every Priority-tier re-anchor item — omit this section entirely if none exist)
DRIFT ITEM: [drift term/framing text]
REINTRODUCED: YES / NO
BY: [H2 heading text if reintroduced, or "none"]
(repeat for every drift item)
Last line:
"CHECK 1.5B-BRIEF: PASS"
only if every Priority-tier re-anchor item is reflected AND no drift item is reintroduced.
Otherwise:
"CHECK 1.5B-BRIEF: FAIL".
`.trim();
  return prompt;
}
function saveRewriteBriefComplianceResultW15B(raw) {
  if (!raw || !raw.trim()) {
    return {
      success: false,
      message: 'No response provided.'
    };
  }
  var ss =
    SpreadsheetApp.getActiveSpreadsheet();
  var sheet =
    ss.getSheetByName('posts');
  var row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return {
      success: false,
      message: 'Select a data row first.'
    };
  }
  var passed =
    /CHECK 1\.5B-BRIEF:\s*PASS/i.test(raw);
  var failed =
    /CHECK 1\.5B-BRIEF:\s*FAIL/i.test(raw);
  if (!passed && !failed) {
    return {
      success: false,
      message:
        'Could not find CHECK 1.5B-BRIEF result line in response.'
    };
  }
  var headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });
  var colIdx =
    headers.indexOf(
      'H2 - Rewrite Brief Compliance Check'
    );
  if (colIdx === -1) {
    return {
      success: false,
      message:
        "Column 'H2 - Rewrite Brief Compliance Check' not found in posts sheet."
    };
  }
  var summary =
    (passed ? 'PASS' : 'FAIL') +
    ' — ' +
    raw.trim();
  var cell =
    sheet.getRange(
      row,
      colIdx + 1
    );
  cell.setNumberFormat('@');
  cell.setValue(summary);
  return {
    success: true,
    passed: passed,
    message:
      passed
        ? 'H2s comply with Rewrite Brief — proceed to Stage 1.5C.'
        : 'H2s do not fully comply with Rewrite Brief — saved for review. Consider revising H2s before proceeding.',
    rawResult: raw.trim()
  };
}
/* ============================================================
   STAGE 1.5B — REWRITE BRIEF FIX PROMPT
============================================================ */
function buildRewriteBriefFixPromptW15B() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });
  const colIdx =
    headers.indexOf(
      'H2 - Rewrite Brief Compliance Check'
    );
  if (colIdx === -1) {
    return 'ERROR: Column "H2 - Rewrite Brief Compliance Check" not found.';
  }
  const savedResult =
    String(
      sheet
        .getRange(
          row,
          colIdx + 1
        )
        .getValue() || ''
    ).trim();
  if (!savedResult) {
    return 'ERROR: No saved compliance check result found. Run the compliance check first.';
  }
  if (savedResult.indexOf('FAIL') !== 0) {
    return 'ERROR: Saved result is not a FAIL — no fix needed.';
  }
  const generatedH2s =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ""
    ).trim();
  if (!generatedH2s) {
    return 'ERROR: No H2 headings found in column CV (100).';
  }
  const d =
    getActiveRowDataMap();
  const rewriteBrief =
    String(
      d["Page Rewrite Brief"] || ""
    ).trim();
  const articleType =
    d["Article Type"] ||
    "General";
  const material =
    d["Stone Type"] ||
    "UNKNOWN";
  const siblingHeadings =
    _getSiblingHeadingsForVariety(
      sheet,
      row,
      material
    );
  const minimumChangeStandard =
    bc_getMinimumChangeStandard_();
  const prompt = `
STAGE 1.5B — REWRITE BRIEF FIX
ROLE:
SEO Content Strategist
ARTICLE TYPE:
${articleType}
--- REWRITE BRIEF ---
${rewriteBrief}
--- PAGE SCOPE LOCK — HARD LOCK ---
ARTICLE TYPE: ${articleType}
MATERIAL: ${material}
Preserve the existing page subject and intent.
Do NOT broaden the page into a different service, diagnostic, maintenance, installation, restoration, product-ranking, or unrelated comparison topic.
A correction may resolve Rewrite Brief failures, but it must stay within the current ${articleType} subject and ${material} material scope.
Do NOT introduce a new competing material, service, problem, or decision framework merely to satisfy the validation result.
--- END PAGE SCOPE LOCK ---
--- CURRENT H2 HEADINGS ---
${generatedH2s}
--- END CURRENT H2 HEADINGS ---
--- COMPLIANCE CHECK FAILURE DETAIL ---
${savedResult}
--- END FAILURE DETAIL ---
${siblingHeadings
  ? '--- HEADINGS ALREADY USED ELSEWHERE IN THIS SILO (' +
    material +
    ') — HARD LOCK ---\n' +
    'Do NOT rewrite any H2 into a heading that duplicates the structural pattern, opening word, or phrasing shape of these existing headings from other ' +
    material +
    ' articles.\n' +
    siblingHeadings +
    '\n--- END EXISTING SILO HEADINGS ---\n'
  : ''}
${minimumChangeStandard}
--- TASK ---
The H2 set failed Rewrite Brief compliance only for the specific items identified in the failure detail above.
HARD FIX RULES:
1. Every H2 not directly implicated by a failed RE-ANCHOR or DRIFT item is LOCKED and must be returned verbatim.
2. Do not rewrite passing H2s for style, SEO improvement, variety, clarity or preference.
3. For each RE-ANCHOR ITEM marked "REFLECTED: NO", change only the H2 best suited to represent that missing governed topic.
4. Add the missing topic to an existing H2 only where it fits that section's governed role from Stage 1.5A.
5. Do not force a RE-ANCHOR topic into an unrelated section simply to make the check pass.
6. For each DRIFT ITEM marked "REINTRODUCED: YES", remove only the wording that creates the drift while preserving the section's intended topic.
7. Do not add, remove, merge, split or renumber sections.
8. Preserve the existing section-to-topic relationship.
9. Do not introduce any new topic, service, material, method, product, claim or decision framework that is not required by the Rewrite Brief.
10. If resolving a failure would require changing an unrelated passing H2 or changing the governed purpose of a section, do not invent a workaround.
CHANGE DISCIPLINE:
- Apply the MINIMUM CHANGE STANDARD above.
- Preserve the original opening phrase and structure unless the failed check specifically requires changing them.
- Passing headings must remain character-for-character unchanged.
- Every changed word must directly contribute to resolving a specific failed RE-ANCHOR or DRIFT item.
Continue to follow the symptom-first rule:
H2 headings must describe the homeowner-recognisable problem, condition or decision rather than lead with a technical mechanism, unless this Article Type explicitly permits otherwise.
OUTPUT FORMAT:
SECTION 1: Heading H2: [unchanged or corrected H2]
SECTION 2: Heading H2: [unchanged or corrected H2]
...
Return the COMPLETE H2 list, including headings that did not need changing.
Last line must be:
Stage 1.5B rewrite brief fix complete.
`.trim();
  return prompt;
}
/* ============================================================
   STAGE 1.5B — H2 HEADING GOVERNANCE CHECK (RULE 17)
============================================================ */
function buildH2GovernanceCheckPromptW15B() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const generatedH2s =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ""
    ).trim();
  if (!generatedH2s) {
    return 'ERROR: No H2 headings found in column CV (100). Run and save Stage 1.5B first.';
  }
  const d =
    getActiveRowDataMap();
  const material =
    d["Stone Type"] ||
    "UNKNOWN";
  const articleType =
    d["Article Type"] ||
    "General";
  const minimumChangeStandard =
    bc_getMinimumChangeStandard_();
  const prompt = `
STAGE 1.5B — H2 HEADING GOVERNANCE CHECK (RULE 17)
ROLE:
SEO Content Structure Auditor
MATERIAL:
${material}
ARTICLE TYPE:
${articleType}
--- GENERATED H2 HEADINGS ---
${generatedH2s}
--- END GENERATED H2 HEADINGS ---
${minimumChangeStandard}
--- TASK ---
Apply these four tests to every H2 heading above.
TEST 1 — SECTION MATCH:
The heading must accurately represent the section it belongs to.
FAIL only when the heading clearly:
- misrepresents the section;
- changes the section's purpose;
- is substantially narrower or broader than the section;
- or introduces a different topic.
Do NOT fail because another heading could sound slightly better.
TEST 2 — READER CLARITY:
The heading must be understandable to a homeowner.
FAIL only when the wording is genuinely confusing, highly technical, awkward enough to obscure the meaning, or uses terminology a homeowner is unlikely to understand.
Do NOT fail for:
- minor wording preferences;
- stylistic improvements;
- wording that is understandable but could later be humanised;
- generic suffixes unless they genuinely make the heading unclear.
TEST 3 — DUPLICATE INTENT:
The heading must have a distinguishable purpose from the other H2s.
FAIL only when another H2 would lead the reader to expect substantially the same section content.
Similar subject matter is acceptable where the section roles are different.
Do NOT fail merely because:
- several headings begin with When, Why, How or What;
- the material name appears several times;
- related sections naturally use related terminology.
TEST 4 — SCOPE AND SECTION ROLE:
The heading must remain within the existing article scope and preserve the section's structural role.
FAIL only when it clearly:
- introduces a new topic, service, method, material, outcome or decision;
- changes a diagnosis section into a remedy section;
- changes a process section into an outcome section;
- changes a maintenance section into a different subject;
- or otherwise changes the purpose established for that section.
TEST 5 — CLAIM SUPPORT:

The H2 must not make a factual, credential, quality, vetting or authority claim that the governed section brief does not substantiate.

Claim-bearing words include, but are not limited to:
vetted, approved, accredited, certified, guaranteed, trusted, recommended, authorised, award-winning, leading, best, expert-approved.

A heading may use one of these words ONLY when the supplied governance data or section brief explicitly establishes that claim.

If the section merely explains what a homeowner should look for in a contractor, a heading such as "Vetted Tiling Specialists..." is NOT supported.

When TEST 5 fails, make the smallest wording adjustment necessary while preserving the governed section role.

IMPORTANT — EARLY-STAGE STRUCTURAL CHECK:
This is an early structural validation stage, NOT a final editorial polish stage.
The purpose is to ensure that the H2s accurately represent the sections below them.
Humanisation and later editorial stages can improve phrasing.
Therefore, DO NOT FAIL for:
- wording that is acceptable but not perfect;
- minor stylistic preferences;
- repeated grammatical patterns;
- a heading simply because you can think of a better version;
- small differences in tone;
- phrases that remain understandable to a homeowner.
Only FAIL where a genuine structural, clarity, duplication or scope problem exists.
WHOLE-SET GOVERNANCE CHECK:
After checking each H2 individually, check the complete set.
Only FAIL the whole set for:
1. CLEAR DUPLICATE INTENT
Two or more headings substantially promise the same section.
2. GENUINELY CONFUSING SEQUENCE
The order is clearly contradictory or prevents the reader from understanding the progression.
Do NOT fail merely because another editorial order might also work.
3. OBVIOUS SUBJECT DRIFT
One or more headings move outside the page's intended material, problem, method or decision scope.
4. STRUCTURAL MISREPRESENTATION
A heading clearly fails to represent the section it belongs to.
Do NOT perform a phrase-shape variety test.
Do NOT fail repeated When, Why, How or What openings.
Do NOT perform final-copy humanisation here.
OUTPUT FORMAT:
H2: [heading text]
TEST 1 (Section Match): PASS / FAIL
TEST 2 (Reader Clarity): PASS / FAIL
TEST 3 (Duplicate Intent): PASS / FAIL
TEST 4 (Scope and Section Role): PASS / FAIL
TEST 5 (Claim Support): PASS / FAIL
REASON: [brief summary]
If ANY test is FAIL, also output:
PROBLEM:
[State the genuine structural problem.]
WHY THIS IS A PROBLEM:
[Explain briefly why it prevents the heading from doing its job.]
WHAT TO CHANGE:
[Give a correction instruction that follows the MINIMUM CHANGE STANDARD above.]
SUGGESTED ALTERNATIVES:
1. [First replacement H2]
2. [Second replacement H2]
3. [Third replacement H2]
Rules for suggested alternatives:
- Preserve the existing section purpose.
- Fix only the genuine failure.
- Do not introduce a new topic, service, material, process or outcome.
- Avoid duplicating another H2.
- Apply the MINIMUM CHANGE STANDARD above.
- Do not rewrite merely for stylistic improvement.
Repeat for every H2.
Last line:
CHECK 1.5B-H2GOV: PASS
only if there is no genuine structural failure.
Otherwise:
CHECK 1.5B-H2GOV: FAIL
`.trim();
  return prompt;
}
function buildH2TechnicalSafetyCheckPromptW15B() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const generatedH2s = String(
    sheet.getRange(row, 100).getValue() || ""
  ).trim();
  if (!generatedH2s) {
    return 'ERROR: No H2 headings found in column CV (100). Run and save Stage 1.5B first.';
  }
  const d = getActiveRowDataMap();
  const material =
    String(
      d["Stone Type"] || "UNKNOWN"
    ).trim();
  const articleType =
    String(
      d["Article Type"] || "General"
    ).trim();
  const rewriteBrief =
    String(
      d["Page Rewrite Brief"] || ""
    ).trim();
  const minimumChangeStandard =
    bc_getMinimumChangeStandard_();
  const techDNA =
    bc_getTechDNABlock(material);
  const technicalMatrix =
    bc_getTechnicalMatrixBlock(material);
  return `
STAGE 1.5B.2 — H2 TECHNICAL & SCOPE SAFETY CHECK
ROLE:
Senior Stone and Tile Technical Governance Auditor
MATERIAL:
${material}
ARTICLE TYPE:
${articleType}
PAGE REWRITE BRIEF:
${rewriteBrief || "No active Rewrite Brief supplied."}
--- CURRENT H2 HEADINGS ---
${generatedH2s}
--- END CURRENT H2 HEADINGS ---
--- MATERIAL GOVERNANCE ---
${techDNA}
${technicalMatrix}
--- END MATERIAL GOVERNANCE ---
${minimumChangeStandard}
TASK:
Check the H2 headings against the supplied material governance.
This is an early-stage technical safety check, not a final technical audit.
Only FAIL a heading where there is a clear and material technical or scope contradiction.
Do NOT invent technical rules that are not supported by the supplied governance.
Check specifically for these six failure modes:
1. PREMATURE SOLUTION JUMPING
FAIL only where a heading clearly pushes the reader into a specialist process, chemical, tool or treatment that is inappropriate for the Article Type or unsupported by the section role.
Do NOT fail merely because a heading mentions cleaning, sealing, repair, restoration or another process in a general way.
2. MATERIAL / METHOD INCOMPATIBILITY
FAIL where a heading clearly recommends, implies or normalises a process, chemical, coating, tool or treatment that conflicts with the supplied material governance.
Examples include:
- abrasive refinishing of glazed ceramic where prohibited;
- sealing a non-porous tile face where governance says it should not be sealed;
- acidic treatment where the material governance prohibits acid.
Only fail where the contradiction is clear.
3. SUBSTRATE VS SURFACE CONFUSION
FAIL where a heading clearly misrepresents a substrate, structural, moisture or installation problem as a simple surface-cleaning, polishing or sealing problem, or clearly does the reverse.
Do NOT fail where the heading is broad and does not make a false technical claim.
4. DIY VS PROFESSIONAL BOUNDARY BLURRING
FAIL only where a heading clearly presents a specialist, hazardous, machinery-dependent or technically controlled professional treatment as an ordinary homeowner DIY task.
Do NOT fail general homeowner-facing headings merely because they discuss the existence of a professional process.
5. UNREALISTIC OR ABSOLUTE OUTCOME CLAIMS
FAIL only where the heading clearly promises an outcome the supplied governance does not support, such as:
- guaranteed perfection;
- permanent removal of a permanent condition;
- invisible repair;
- complete reversal of irreversible damage;
- factory-new condition where not technically supportable.
Do NOT fail ordinary positive or outcome-oriented wording that does not make an absolute claim.
6. SEALER-AS-SOLUTION MISATTRIBUTION
FAIL only where the heading clearly presents sealing, coating or impregnation as the cure for an existing condition that the supplied material governance says requires a different treatment.
Do NOT fail a heading merely because it discusses sealing as part of maintenance, protection or decision-making.
IMPORTANT — EARLY-STAGE SAFETY GATE:
Only fail a heading where there is a genuine technical contradiction that could cause the article structure to point in the wrong direction.
Do NOT fail for:
- minor technical nuance;
- wording that could later be clarified;
- missing technical detail;
- homeowner-friendly simplification;
- a heading that is technically broad but not wrong;
- issues that can safely be corrected during later drafting, factual review or humanisation.
Do not introduce requirements from another material.
Do not rewrite the headings during this check.
OUTPUT FORMAT:
For each H2:
H2: [heading text]
TECHNICAL SAFETY: PASS / FAIL
REASON: [brief explanation]
If FAIL, also output:
PROBLEM:
[Plain-English explanation of the clear technical or scope contradiction.]
WHY THIS IS A PROBLEM:
[Explain the conflict with the supplied material governance.]
WHAT TO CHANGE:
[Give a correction instruction that follows the MINIMUM CHANGE STANDARD above. Do not write the replacement heading.]
Repeat for every H2.
Last line must be exactly:
CHECK 1.5B.2-TECH: PASS
or:
CHECK 1.5B.2-TECH: FAIL
`.trim();
}
/* ============================================================
   STAGE 1.5B.1 — H2 GOVERNANCE FIX (TARGETED, RULE 17)
============================================================ */
function buildH2GovernanceFixPromptW15B() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });
  const colIdx =
    headers.indexOf(
      'H2 Governance Check'
    );
  if (colIdx === -1) {
    return 'ERROR: Column "H2 Governance Check" not found.';
  }
  const savedResult =
    String(
      sheet
        .getRange(
          row,
          colIdx + 1
        )
        .getValue() || ''
    ).trim();
  if (!savedResult) {
    return 'ERROR: No saved governance check result found. Run the H2 governance check first.';
  }
  if (savedResult.indexOf('FAIL') !== 0) {
    return 'ERROR: Saved result is not a FAIL — no fix needed.';
  }
  const generatedH2s =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ""
    ).trim();
  if (!generatedH2s) {
    return 'ERROR: No H2 headings found in column CV (100).';
  }
  const d =
    getActiveRowDataMap();
  const material =
    d["Stone Type"] ||
    "UNKNOWN";
  const articleType =
    d["Article Type"] ||
    "General";
  const siblingHeadings =
    _getSiblingHeadingsForVariety(
      sheet,
      row,
      material
    );
  const minimumChangeStandard =
    bc_getMinimumChangeStandard_();
  const body =
    savedResult.replace(
      /^(PASS|FAIL)\s*—\s*/,
      ''
    );
  const blocks =
    body.split(
      /\n(?=H2:)/
    );
  const failingBlocks = [];
  blocks.forEach(function(block) {
    var isFail =
      /TEST\s*[1-4][^\n]*:\s*FAIL/i
        .test(block);
    var hasViolation =
      /PROBLEM:/i.test(block) ||
      /WHAT TO CHANGE:/i.test(block);
    if (isFail || hasViolation) {
      var h2Match =
        block.match(
          /H2:\s*(.+)/i
        );
      if (h2Match) {
        failingBlocks.push({
          heading:
            h2Match[1].trim(),
          detail:
            block.trim()
        });
      }
    }
  });
  if (failingBlocks.length === 0) {
    return 'ERROR: Could not identify specific failing H2s from the saved result. Check the saved governance result format.';
  }
  const failingList =
    failingBlocks
      .map(function(f) {
        return (
          '- "' +
          f.heading +
          '"\n  ' +
          f.detail
            .split('\n')
            .slice(1)
            .join('\n  ')
        );
      })
      .join('\n\n');
  const prompt = `
STAGE 1.5B.1 — H2 GOVERNANCE FIX (RULE 17 — TARGETED)
ROLE:
SEO Content Structure Auditor
MATERIAL:
${material}
ARTICLE TYPE:
${articleType}
--- CURRENT H2 HEADINGS ---
${generatedH2s}
--- END CURRENT H2 HEADINGS ---
--- FAILING HEADINGS ---
${failingList}
--- END FAILING HEADINGS ---
${siblingHeadings
  ? '--- HEADINGS ALREADY USED ELSEWHERE IN THIS SILO (' +
    material +
    ') — HARD LOCK ---\n' +
    'Do NOT rewrite any H2 into a heading that duplicates the structural pattern, opening word, or phrasing shape of these existing headings from other ' +
    material +
    ' articles.\n' +
    siblingHeadings +
    '\n--- END EXISTING SILO HEADINGS ---\n'
  : ''}
${minimumChangeStandard}
--- TASK ---
Rewrite ONLY the H2 headings listed above as failing.
Every other H2 heading in the current set must be copied back EXACTLY unchanged — same wording, same SECTION number, same order.
For each failing heading, apply Rule 17 corrections:
- If it failed Test 1 (Section Match):
  correct only the wording that misrepresents, narrows, broadens or changes the section's existing purpose.
- If it failed Test 2 (Reader Clarity):
  correct only the wording that is genuinely confusing, overly technical or unclear to a homeowner.
- If it failed Test 3 (Duplicate Intent):
  change only what is necessary to distinguish this section from the other H2 whose intended content substantially overlaps.
- If it failed Test 4 (Scope and Section Role):
  remove or correct only the wording that introduces a new topic or changes the governed purpose of the section.
For EACH failing H2:
1. Generate THREE possible corrected headings internally.
2. Test all three candidates against:
- Test 1 — Section Match
- Test 2 — Reader Clarity
- Test 3 — Duplicate Intent
- Test 4 — Scope and Section Role
- the whole-set governance rules
3. Reject any candidate that:
- introduces a new topic;
- narrows or broadens the section incorrectly;
- duplicates the intent of another H2;
- is genuinely confusing to a homeowner;
- changes the intended section role;
- moves outside the existing article scope.
4. Select the strongest remaining candidate.
5. Return ONLY that selected corrected heading in the final H2 list.
Apply the MINIMUM CHANGE STANDARD above.
Do not make any additional stylistic or structural change beyond what is required to make the failing H2 pass its identified Rule 17 test(s).
STRUCTURAL ROLE PRESERVATION — HARD LOCK:
When rewriting a failing H2, preserve the structural coverage role that heading is currently serving.
Do not change:
- a tools/materials heading into a process heading;
- a risk heading into a diagnosis heading;
- a diagnosis heading into a remedy heading;
- a process heading into an outcome heading;
- a decision heading into a generic informational heading;
- or any other section role merely to make the heading pass Rule 17.
The corrected H2 must continue to satisfy the same section function in the article structure.
If the failing H2 is currently carrying a required article-type or TSM coverage element, that coverage must remain present after correction.
Do not over-correct a heading merely because it contains a process-related word.
A process word does not automatically mean the heading contains two topics.
Do not expose the rejected candidate headings in the final output.
Do NOT rewrite, improve, or rephrase any heading not listed above as failing — even if you think it could be better.
Only the listed failing headings may change.
--- OUTPUT FORMAT ---
Output the COMPLETE H2 list, in the same SECTION N order as the current set, with only the failing headings corrected:
SECTION 1: Heading H2: [unchanged or corrected H2]
SECTION 2: Heading H2: [unchanged or corrected H2]
...
Last line must be:
Stage 1.5B complete. Waiting for Stage 1.5C.
`.trim();
  return prompt;
}
/* ============================================================
   STAGE 1.5B.2 — TECHNICAL & SCOPE SAFETY FIX
============================================================ */
function buildH2TechnicalSafetyFixPromptW15B(failureText) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  failureText =
    String(
      failureText || ''
    ).trim();
  if (!failureText) {
    return 'ERROR: No Technical & Scope Safety failure result was supplied.';
  }
  if (
    !/CHECK 1\.5B\.2-TECH:\s*FAIL/i
      .test(failureText)
  ) {
    return 'ERROR: Supplied Technical & Scope Safety result is not a FAIL.';
  }
  const generatedH2s =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ''
    ).trim();
  if (!generatedH2s) {
    return 'ERROR: No H2 headings found in column CV (100).';
  }
  const d =
    getActiveRowDataMap();
  const material =
    String(
      d['Stone Type'] ||
      'UNKNOWN'
    ).trim();
  const articleType =
    String(
      d['Article Type'] ||
      'General'
    ).trim();
  const rewriteBrief =
    String(
      d['Page Rewrite Brief'] ||
      ''
    ).trim();
  const siblingHeadings =
    _getSiblingHeadingsForVariety(
      sheet,
      row,
      material
    );
  const minimumChangeStandard =
    bc_getMinimumChangeStandard_();
  const techDNA =
    bc_getTechDNABlock(material);
  const technicalMatrix =
    bc_getTechnicalMatrixBlock(material);
  const prompt = `
STAGE 1.5B.2 — H2 TECHNICAL & SCOPE SAFETY FIX
ROLE:
Senior Stone and Tile Technical Governance Auditor
MATERIAL:
${material}
ARTICLE TYPE:
${articleType}
PAGE REWRITE BRIEF:
${rewriteBrief || 'No active Rewrite Brief supplied.'}
--- CURRENT H2 HEADINGS ---
${generatedH2s}
--- END CURRENT H2 HEADINGS ---
--- TECHNICAL & SCOPE SAFETY FAILURE RESULT ---
${failureText}
--- END FAILURE RESULT ---
--- MATERIAL GOVERNANCE ---
${techDNA}
${technicalMatrix}
--- END MATERIAL GOVERNANCE ---
${siblingHeadings
  ? '--- HEADINGS ALREADY USED ELSEWHERE IN THIS SILO (' +
    material +
    ') — HARD LOCK ---\n' +
    'Do NOT rewrite any H2 into a heading that duplicates the structural pattern, opening word, or phrasing shape of these existing headings from other ' +
    material +
    ' articles.\n' +
    siblingHeadings +
    '\n--- END EXISTING SILO HEADINGS ---\n'
  : ''}
${minimumChangeStandard}
--- TASK ---
Correct ONLY the H2 headings that are explicitly marked TECHNICAL SAFETY: FAIL in the failure result above.
HARD FIX RULES:
1. Every H2 marked TECHNICAL SAFETY: PASS is LOCKED and must be returned verbatim.
2. For each failing H2, use its PROBLEM, WHY THIS IS A PROBLEM and WHAT TO CHANGE instructions as the sole reason for modification.
3. Apply the MINIMUM CHANGE STANDARD above.
4. Remove or correct only the wording that creates the technical or scope contradiction.
5. Preserve the existing section role, subject and relationship to Stage 1.5A.
6. Do not add, remove, merge, split or renumber sections.
7. Do not convert:
- a diagnosis section into a remedy section;
- a risk section into a process section;
- a process section into an outcome section;
- a maintenance section into a restoration section;
- or any other section into a different governed role.
8. Do not introduce a new chemical, tool, treatment, coating, process, defect, outcome or technical claim merely to resolve the failure.
9. Do not invent technical facts that are not supported by the supplied material governance.
10. Do not broaden a homeowner-facing heading into specialist treatment language unless the failure result specifically requires that correction.
11. If the failure can be resolved by removing or replacing a short phrase, do not rewrite the whole H2.
12. If satisfying the failure would require changing the governed purpose of the section, do not invent a workaround.
13. Passing headings must remain character-for-character unchanged.
14. Every changed word must be directly traceable to a specific TECHNICAL SAFETY: FAIL item.
--- OUTPUT FORMAT ---
Return the COMPLETE H2 list in the existing SECTION order.
Use exactly:
SECTION 1: Heading H2: [unchanged or corrected H2]
SECTION 2: Heading H2: [unchanged or corrected H2]
...
Do not output explanations, rejected alternatives, analysis or commentary.
Last line must be exactly:
Stage 1.5B.2 technical safety fix complete.
`.trim();
  return prompt;
}
function saveH2GovernanceCheckResultW15B(raw) {
  if (!raw || !raw.trim()) {
    return {
      success: false,
      message: 'No response provided.'
    };
  }
  var ss =
    SpreadsheetApp.getActiveSpreadsheet();
  var sheet =
    ss.getSheetByName('posts');
  var row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return {
      success: false,
      message: 'Select a data row first.'
    };
  }
  var passed =
    /CHECK 1\.5B-H2GOV:\s*PASS/i
      .test(raw);
  var failed =
    /CHECK 1\.5B-H2GOV:\s*FAIL/i
      .test(raw);
  if (!passed && !failed) {
    return {
      success: false,
      message:
        'Could not find CHECK 1.5B-H2GOV result line in response.'
    };
  }
  var headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });
  var colIdx =
    headers.indexOf(
      'H2 Governance Check'
    );
  if (colIdx === -1) {
    return {
      success: false,
      message:
        "Column 'H2 Governance Check' not found in posts sheet."
    };
  }
  var summary =
    (passed ? 'PASS' : 'FAIL') +
    ' — ' +
    raw.trim();
  var cell =
    sheet.getRange(
      row,
      colIdx + 1
    );
  cell.setNumberFormat('@');
  cell.setValue(summary);
  return {
    success: true,
    passed: passed,
    message:
      passed
        ? 'H2s pass heading governance (Rule 17) — proceed to Stage 1.5C.'
        : 'One or more H2s fail heading governance — saved for review. Revise H2s before proceeding.',
    rawResult: raw.trim()
  };
}
function saveH2TechnicalSafetyCheckResultW15B(raw) {
  if (!raw || !raw.trim()) {
    return {
      success: false,
      message: 'No response provided.'
    };
  }
  var passed =
    /CHECK 1\.5B\.2-TECH:\s*PASS/i
      .test(raw);
  var failed =
    /CHECK 1\.5B\.2-TECH:\s*FAIL/i
      .test(raw);
  if (!passed && !failed) {
    return {
      success: false,
      message:
        'Could not find CHECK 1.5B.2-TECH result line in response.'
    };
  }
  return {
    success: true,
    passed: passed,
    message:
      passed
        ? 'H2s pass Technical & Scope Safety governance.'
        : 'One or more H2s fail Technical & Scope Safety governance.',
    rawResult: raw.trim()
  };
}
function buildEntityCoverageCheckPrompt() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const generatedH2s =
    String(
      sheet.getRange(row, 100).getValue() || ""
    ).trim();
  if (!generatedH2s) {
    return 'ERROR: No H2 headings found in column CV (100). Run and save Stage 1.5B first.';
  }
  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h || "").trim();
      });
  const values =
    sheet
      .getRange(
        row,
        1,
        1,
        headers.length
      )
      .getValues()[0];
  const d = {};
  headers.forEach(function(h, i) {
    d[h] = values[i];
  });
  const articleType =
    String(
      d["Article Type"] || ""
    ).trim();
  const material =
    String(
      d["Stone Type"] || ""
    ).trim();
  if (!articleType) {
    return 'ERROR: Article Type is empty on the selected posts row.';
  }
  const tierMap = {
    "Hub Page": "Tier 1",
    "Educational Guide": "Tier 1",
    "Method Guide": "Tier 2",
    "Service Page": "Tier 2",
    "Geo Service Page": "Tier 2",
    "Diagnostic Guide": "Tier 3",
    "Buyer Guide": "Tier 3",
    "FAQ Spoke": "Tier 3",
    "Case Study": "Tier 4"
  };
  const tierLabel =
    tierMap[articleType] ||
    "Tier 2";
  const NON_CONTENT_TSM_REQUIREMENTS = [
    'Internal Link Upward',
    'Internal Link',
    'Internal Link to Hub',
    'Internal Link to Tier 2',
    'Hub Link',
    'Lateral Link'
  ];
  let tsmRequirementNames = "";
  try {
    const tsmData =
      getTierStructuralRequirements(
        material,
        tierLabel,
        articleType
      );
    if (
      tsmData &&
      tsmData.length > 0
    ) {
      const contentTsmData =
        tsmData.filter(function(req) {
          return (
            NON_CONTENT_TSM_REQUIREMENTS
              .indexOf(req.name) === -1
          );
        });
      tsmRequirementNames =
        contentTsmData
          .map(function(req) {
            return '- ' + req.name;
          })
          .join('\n');
    }
  } catch (e) {
    tsmRequirementNames =
      '(Could not read TSM requirements: ' +
      e.message +
      ')';
  }
  let requiredEntityCoverage = "";
  try {
    const artSheet =
      ss.getSheets().find(
        function(s) {
          return (
            s.getSheetId() ==
            BC_SHEET_CONFIG.articleTypeControl
          );
        }
      );
    if (artSheet) {
      const artData =
        artSheet
          .getDataRange()
          .getValues();
      for (
        let i = 1;
        i < artData.length;
        i++
      ) {
        if (
          String(
            artData[i][0]
          ).trim() === articleType
        ) {
          requiredEntityCoverage =
            String(
              artData[i][2] || ""
            ).trim();
          break;
        }
      }
    }
  } catch (e) {
    return (
      'ERROR: Could not read Article Type Control Sheet — ' +
      e.message
    );
  }
  if (!requiredEntityCoverage) {
    return (
      'ERROR: No Required Entity Coverage found for article type: ' +
      articleType
    );
  }
  const prompt = `
STAGE 1.5B — ENTITY COVERAGE VERIFICATION
ROLE:
SEO Content Structure Auditor
ARTICLE TYPE:
${articleType}
MATERIAL:
${material}
--- REQUIRED ENTITY COVERAGE ---
${requiredEntityCoverage}
--- TSM REQUIREMENTS ---
${tsmRequirementNames || 'No TSM requirements found.'}
--- GENERATED H2 HEADINGS ---
${generatedH2s}
--- TASK ---
Two separate coverage lists are given above:
1. Required Entity Coverage
2. TSM Requirements
Check the generated H2 headings against BOTH lists independently.
An item satisfying one list does not automatically satisfy the other.
For each item in BOTH lists, state whether it is covered and by which H2(s).
OUTPUT FORMAT:
LIST: Required Entity Coverage
ITEM: [coverage item text]
COVERED: YES / NO
BY: [H2 heading text, or "none"]
Repeat for every Required Entity Coverage item.
Then:
LIST: TSM Requirements
ITEM: [TSM requirement name]
COVERED: YES / NO
BY: [H2 heading text, or "none"]
Repeat for every TSM Requirement item.
Last line:
CHECK 1.5B-COVERAGE: PASS
only if every item in BOTH lists is covered.
Otherwise:
CHECK 1.5B-COVERAGE: FAIL
`.trim();
  return prompt;
}
function saveEntityCoverageCheckResult(raw) {
  if (!raw || !raw.trim()) {
    return {
      success: false,
      message: 'No response provided.'
    };
  }
  var ss =
    SpreadsheetApp.getActiveSpreadsheet();
  var sheet =
    ss.getSheetByName('posts');
  var row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return {
      success: false,
      message: 'Select a data row first.'
    };
  }
  var passed =
    /CHECK 1\.5B-COVERAGE:\s*PASS/i
      .test(raw);
  var failed =
    /CHECK 1\.5B-COVERAGE:\s*FAIL/i
      .test(raw);
  if (!passed && !failed) {
    return {
      success: false,
      message:
        'Could not find CHECK 1.5B-COVERAGE result line in response.'
    };
  }
  var headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });
  var colIdx =
    headers.indexOf(
      'Entity Coverage Check'
    );
  if (colIdx === -1) {
    return {
      success: false,
      message:
        "Column 'Entity Coverage Check' not found in posts sheet."
    };
  }
  var summary =
    (passed ? 'PASS' : 'FAIL') +
    ' — ' +
    raw.trim();
  var cell =
    sheet.getRange(
      row,
      colIdx + 1
    );
  cell.setNumberFormat('@');
  cell.setValue(summary);
  return {
    success: true,
    passed: passed,
    message:
      passed
        ? 'All Required Entity Coverage items are addressed — saved to Entity Coverage Check column.'
        : 'One or more Required Entity Coverage items are missing — saved to Entity Coverage Check column.',
    rawResult: raw.trim()
  };
}
function _getSiblingHeadingsForVariety(
  sheet,
  currentRow,
  material
) {
  if (!material) {
    return '';
  }
  var lastRow =
    sheet.getLastRow();
  var lastCol =
    sheet.getLastColumn();
  var headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastCol
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });
  var stoneTypeCol =
    headers.indexOf(
      'Stone Type'
    );
  var headingsCol =
    headers.indexOf(
      'Current Headings'
    );
  var articleTypeCol =
    headers.indexOf(
      'Article Type'
    );
  if (
    stoneTypeCol === -1 ||
    headingsCol === -1
  ) {
    return '';
  }
  if (lastRow < 2) {
    return '';
  }
  var data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastCol
      )
      .getValues();
  var siblingBlocks = [];
  for (
    var i = 0;
    i < data.length;
    i++
  ) {
    var actualRow =
      i + 2;
    if (
      actualRow === currentRow
    ) {
      continue;
    }
    var rowMaterial =
      String(
        data[i][stoneTypeCol] || ''
      ).trim();
    if (
      rowMaterial !== material
    ) {
      continue;
    }
    var headingsText =
      String(
        data[i][headingsCol] || ''
      ).trim();
    if (!headingsText) {
      continue;
    }
    var rowArticleType =
      articleTypeCol > -1
        ? String(
            data[i][articleTypeCol] || ''
          ).trim()
        : '';
    siblingBlocks.push(
      (
        rowArticleType
          ? '[' +
            rowArticleType +
            '] '
          : ''
      ) +
      headingsText
    );
  }
  return siblingBlocks.join(
    '\n---\n'
  );
}
function buildStage15BPrompt() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  extractCurrentHeadingsToPosts();
  const d =
    getActiveRowDataMap();
  const material =
    d["Stone Type"] ||
    "UNKNOWN";
  const articleType =
    d["Article Type"] ||
    "General";
  const siblingHeadings =
    _getSiblingHeadingsForVariety(
      sheet,
      row,
      material
    );
  const primaryEntity =
    d["Primary Entity"] ||
    material;
  const location =
    String(
      d["Locality"] ||
      d["Location"] ||
      ""
    ).trim();
  const rewriteBrief =
    String(
      d["Page Rewrite Brief"] ||
      "Not Applicable"
    ).trim();
  const sfWhatTheyFound =
    String(
      d["ac_entry_condition"] ||
      ""
    ).trim();
  const sfWhyTheyCalled =
    String(
      d["ac_homeowner_perception"] ||
      ""
    ).trim();
  const sfFloorBehaviour =
    String(
      d["ac_material_behaviour"] ||
      ""
    ).trim();
  const sfConstraint =
    String(
      d["ac_what_forced_careful_decisions"] ||
      d["ac_constraint"] ||
      ""
    ).trim();
  const sfApproach =
    String(
      d["ac_process_emphasis"] ||
      ""
    ).trim();
  const sfOutcome =
    String(
      d["ac_result_type"] ||
      ""
    ).trim();
  const sfStoryDirection =
    String(
      d["ac_narrative_archetype"] ||
      ""
    ).trim();
  const hasStoryFramework =
    sfWhatTheyFound ||
    sfStoryDirection;
  const articleAngle =
    String(
      d["Article Angle"] ||
      ""
    ).trim();
  const isCaseStudy =
    articleType ===
    "Case Study";
  const ARTICLE_ANGLE_DEFINITIONS = {
    "Diagnostic-Cause-Led": {
      overview:
        "Open by identifying the technical cause of the problem before describing the fix.",
      mapping:
        "SECTION 1 (Project Context): frame the entry point as noticing the SYMPTOM, then naming the underlying CAUSE. SECTION 2 (Problem & Intervention): explain the cause in more depth, then the correction. SECTION 3 (Measurable Outcome): the result once the cause was addressed. SECTION 4 (Maintenance Handover & Escalation Link): frame around preventing this same CAUSE from recurring."
    },
    "Chronological/Step-by-Step": {
      overview:
        "Follow the actual sequence of the job as it happened, stage by stage, without reordering for diagnostic effect.",
      mapping:
        "SECTION 1 (Project Context): arrival and initial assessment. SECTION 2 (Problem & Intervention): the job as a sequence of steps/stages in the order they happened. SECTION 3 (Measurable Outcome): the result at the end of the sequence. SECTION 4 (Maintenance Handover & Escalation Link): frame as the final practical stage in the sequence."
    },
    "Contrast/Before-After-Led": {
      overview:
        "Structure the narrative around explicit before/after comparison as the organising device.",
      mapping:
        "SECTION 1 (Project Context): introduce the before condition clearly. SECTION 2 (Problem & Intervention): the intervention, still referencing what it changed from. SECTION 3 (Measurable Outcome): explicit before-vs-after framing. SECTION 4 (Maintenance Handover & Escalation Link): frame as what keeps the after state from sliding back to before."
    },
    "Assessment-Pathway/Self-Test-Led": {
      overview:
        "Structure the narrative around helping the reader assess their own floor's condition, ending with self-assessment language.",
      mapping:
        "SECTION 1 (Project Context): frame the entry point in terms a reader can self-recognise. SECTION 2 (Problem & Intervention): present the professional decision rationale. SECTION 3 (Measurable Outcome): frame the outcome partly as a capability statement. SECTION 4 (Maintenance Handover & Escalation Link): close with self-check maintenance language."
    },
    "Finish-Recovery/Shine-Led": {
      overview:
        "Centre the narrative on the visual/finish transformation as the core throughline, with technical process subordinate to the finish story.",
      mapping:
        "SECTION 1 (Project Context): frame the entry point around finish or appearance loss. SECTION 2 (Problem & Intervention): explain why cleaning alone cannot fix a finish problem, then the mechanical correction. SECTION 3 (Measurable Outcome): emphasise the finish outcome. SECTION 4 (Maintenance Handover & Escalation Link): frame around protecting the restored finish."
    }
  };
  const articleAngleData =
    ARTICLE_ANGLE_DEFINITIONS[
      articleAngle
    ];
  const articleAngleBlock =
    (
      isCaseStudy &&
      articleAngle &&
      articleAngleData
    )
      ? `
--- ARTICLE ANGLE — HARD LOCK ---
This Case Study has been assigned the "${articleAngle}" structural shape.
This governs HOW each of the fixed TSM sections is framed and worded.
Story Framework tells you WHAT happened.
Article Angle tells you HOW to frame and sequence that content.
SHAPE OVERVIEW:
${articleAngleData.overview}
MANDATORY SECTION-BY-SECTION MAPPING:
${articleAngleData.mapping}
This mapping is a Hard Lock.
Every H2 must reflect its section's assigned framing above while preserving the Story Framework content.
CRITICAL — SECTION 4 SCOPE:
This project's restoration is COMPLETE.
Section 4 is a maintenance handover to the homeowner who just paid for this restoration.
It must contain genuine, specific advice for keeping THIS floor in good condition going forward.
--- END ARTICLE ANGLE ---
`
      : '';
  const structure15A =
    String(
      sheet
        .getRange(row, 99)
        .getValue() || ""
    ).trim();
  if (!structure15A) {
    return 'ERROR: No Stage 1.5A output found. Run W1.5A first.';
  }
  const recoveryBlueprint =
    String(
      d["Recovery Blueprint"] ||
      ""
    ).trim();
  let requiredEntityCoverage = "";
  let onPageExpectations = "";
  let reasoning = "";
  try {
    const artSheet =
      ss.getSheets().find(
        function(s) {
          return (
            s.getSheetId() ==
            BC_SHEET_CONFIG.articleTypeControl
          );
        }
      );
    if (artSheet) {
      const artData =
        artSheet
          .getDataRange()
          .getValues();
      for (
        let i = 1;
        i < artData.length;
        i++
      ) {
        if (
          String(
            artData[i][0]
          ).trim() === articleType
        ) {
          requiredEntityCoverage =
            String(
              artData[i][2] ||
              ""
            ).trim();
          onPageExpectations =
            String(
              artData[i][3] ||
              ""
            ).trim();
          reasoning =
            String(
              artData[i][8] ||
              ""
            ).trim();
          break;
        }
      }
    }
  } catch (e) {
    requiredEntityCoverage =
      "(Article Type Control Sheet read error: " +
      e.message +
      ")";
    onPageExpectations =
      "(Article Type Control Sheet read error: " +
      e.message +
      ")";
  }
  const isGeoServicePage =
    articleType ===
    "Geo Service Page";
  const h2RuleBlock =
    isGeoServicePage
      ? `
--- ARTICLE TYPE H2 RULE: GEO SERVICE PAGE ---
This page is a Geo Service Page.
The symptom-first rule does NOT apply.
H2s for a Geo Service Page must:
1. Confirm the service exists and is delivered locally.
2. Describe what the professional process involves at principle level.
3. Set honest outcome expectations.
4. Move the reader toward making contact.
On-Page Content Expectations:
${onPageExpectations}
${reasoning
  ? 'Content Strategy: ' +
    reasoning
  : ''}
${location
  ? 'Target location: ' +
    location
  : ''}
`
      : `
--- ARTICLE TYPE H2 RULE: ${articleType.toUpperCase()} ---
On-Page Content Expectations:
${onPageExpectations}
${reasoning
  ? 'Content Strategy: ' +
    reasoning
  : ''}
Apply the symptom-first rule below.
H2s must describe the homeowner's problem, not the technical mechanism.
`;
  const workingTitle =
    buildWorkingTitleForH2Stage_(d);
  const prompt = `
STAGE 1.5B — H2 GENERATION
ROLE:
SEO Content Strategist
MATERIAL:
${material}
ARTICLE TYPE:
${articleType}
PRIMARY ENTITY:
${primaryEntity}
WORKING TITLE:
${workingTitle}
${location
  ? 'LOCATION: ' +
    location
  : ''}
--- HEADING QUALITY TARGET — HARD LOCK ---
Every H2 must satisfy all four of these:
1. SEARCH INTENT & UX
The H2 sequence must guide the reader logically from problem to solution.
Do not create overlapping or redundant headings covering the same ground twice.
2. E-E-A-T SIGNAL
Headings should demonstrate hands-on expertise through natural specificity without becoming a summary.
3. SCANNABILITY — HARD LOCK
H2s must be short and punchy.
A heading should be understandable in one glance.
Rich technical vocabulary belongs in H3s or body text, not crammed into the H2 itself.
4. SINGLE TOPIC PER HEADING
Each H2 must cover exactly one main idea.
Do not combine two different section purposes into one heading.
${hasStoryFramework
  ? `
--- STORY FRAMEWORK — PRIMARY GOVERNING INSTRUCTION ---
WHAT THE FLOOR PRESENTED WITH:
${sfWhatTheyFound}
WHY THE HOMEOWNER CALLED:
${sfWhyTheyCalled}
WHAT THE FLOOR WAS DOING:
${sfFloorBehaviour}
WHAT FORCED CAREFUL DECISIONS:
${sfConstraint}
WHAT GOVERNED THE APPROACH:
${sfApproach}
WHAT CHANGED FOR THE HOMEOWNER:
${sfOutcome}
STORY DIRECTION:
${sfStoryDirection}
H2 STORY FRAMEWORK RULES:
- Section 1 H2 MUST open from the entry point represented by:
  "${sfWhatTheyFound}"
- The narrative must move through:
  "${sfConstraint}"
  as a turning point.
- The process section H2 must reflect:
  "${sfApproach}"
- The outcome section H2 must reflect:
  "${sfOutcome}"
- The overall story direction is:
  "${sfStoryDirection}"
- Do NOT default to generic dark, dirty or dull framing unless that matches the entry point above.
--- END STORY FRAMEWORK ---
`
  : ''}
${articleAngleBlock}
--- PRIMARY THEMATIC GOVERNANCE ---
REWRITE BRIEF:
${rewriteBrief}
If the Rewrite Brief above is NOT "Not Applicable", adapt the tone, vocabulary and specific angle of every H2 to align with it.
The H2s must still respect the article-type rules and governed section roles.
--- REQUIRED ENTITY COVERAGE — HARD LOCK ---
This article type MUST cover the following elements somewhere across its H2 sections:
${requiredEntityCoverage ||
  'Not specified for this article type.'}
Every item listed above must be represented by at least one H2 heading or clearly covered within a governed section.
You do not need one H2 per item.
Do not omit any required coverage item.
--- SYSTEM INSTRUCTION ---
Stage 1.5A has created a section structure below.
Your ONLY job is to create an H2 heading for each section.
The H2 rule that applies depends on the article type — read the ARTICLE TYPE H2 RULE block carefully before generating any headings.
${h2RuleBlock}
--- SYMPTOM-FIRST RULE ---
Applies to all article types EXCEPT Geo Service Page.
SYMPTOM-FIRST vs MECHANISM-FIRST:
WRONG:
"Geological Formation & Internal Structure"
CORRECT:
"Why some marks will not come out no matter what you use"
The H2 must describe the HOMEOWNER'S PROBLEM, not the technical mechanism.
CRITICAL RULE:
Never copy the TSM Requirement name as the H2.
ENTITY NAMES FORBIDDEN IN H2S:
- Calcite Crystal Structure
- Impregnating Sealer
- Porosity
- Micro-Scratching
- Acid Etching
Use plain language instead.
--- RECOVERY BLUEPRINT ---
${recoveryBlueprint || 'No Recovery Blueprint available.'}
--- SECTION STRUCTURE FROM STAGE 1.5A ---
${structure15A}
--- HUB-INTRO HANDLING RULE ---
If the section structure includes:
HUB-INTRO: id="hub-intro" | H2="None" | CONTENT="One orienting sentence and a quick-links navigation list only"
Do NOT generate an H2 for HUB-INTRO.
Carry HUB-INTRO forward unchanged as a structural block.
Generate H2 headings for SECTION 1 onward only.
${siblingHeadings
  ? '--- HEADINGS ALREADY USED ELSEWHERE IN THIS SILO (' +
    material +
    ') — HARD LOCK ---\n' +
    'These headings already exist on other published ' +
    material +
    ' articles. Do NOT repeat the same structural pattern, opening word, or phrasing shape as these — genuine variety across the silo matters for avoiding a detectable content footprint.\n' +
    siblingHeadings +
    '\n--- END EXISTING SILO HEADINGS ---\n'
  : ''}
--- TASK ---
For each numbered section in the structure above, create an H2 heading.
Apply the ARTICLE TYPE H2 RULE for this article type.
If HUB-INTRO is present, preserve it unchanged and do not generate an H2 for it.
OUTPUT FORMAT:
SECTION 1: Heading H2: [Your H2]
SECTION 2: Heading H2: [Your H2]
...
--- OUTPUT INSTRUCTION ---
If HUB-INTRO is present in the section structure, preserve it unchanged.
Output ONLY HUB-INTRO if present, followed by the section numbers with H2s.
If no HUB-INTRO is present, start with SECTION 1.
Last line must be:
Stage 1.5B complete. Waiting for Stage 1.5C.
`.trim();
  return prompt;
}
function buildWorkingTitleForH2Stage_(d) {
  var primaryEntity =
    String(
      d["Primary Entity"] ||
      ""
    ).trim();
  var primaryTerm =
    String(
      d["Primary Search Term"] ||
      ""
    ).trim();
  var material =
    String(
      d["Stone Type"] ||
      ""
    ).trim();
  if (primaryTerm) {
    return (
      primaryTerm
        .charAt(0)
        .toUpperCase() +
      primaryTerm.slice(1)
    );
  }
  if (primaryEntity) {
    return primaryEntity;
  }
  return (
    material ||
    "this article"
  );
}
/* ============================================================
   STAGE 1.5B — ENTITY COVERAGE FIX
============================================================ */
function buildEntityCoverageFixPromptW15B() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });
  const colIdx =
    headers.indexOf(
      'Entity Coverage Check'
    );
  if (colIdx === -1) {
    return 'ERROR: Column "Entity Coverage Check" not found.';
  }
  const savedResult =
    String(
      sheet
        .getRange(
          row,
          colIdx + 1
        )
        .getValue() || ''
    ).trim();
  if (!savedResult) {
    return 'ERROR: No saved coverage check result found. Run the coverage check first.';
  }
  if (
    savedResult.indexOf(
      'FAIL'
    ) !== 0
  ) {
    return 'ERROR: Saved result is not a FAIL — no fix needed.';
  }
  const generatedH2s =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ""
    ).trim();
  if (!generatedH2s) {
    return 'ERROR: No H2 headings found in column CV (100).';
  }
  const d =
    getActiveRowDataMap();
  const articleType =
    d["Article Type"] ||
    "General";
  const material =
    d["Stone Type"] ||
    "UNKNOWN";
  const siblingHeadings =
    _getSiblingHeadingsForVariety(
      sheet,
      row,
      material
    );
  const minimumChangeStandard =
    bc_getMinimumChangeStandard_();
  const prompt = `
STAGE 1.5B — ENTITY COVERAGE FIX
ROLE:
SEO Content Strategist
ARTICLE TYPE:
${articleType}
MATERIAL:
${material}
--- COVERAGE CHECK FAILURE DETAIL ---
${savedResult}
--- END COVERAGE CHECK FAILURE DETAIL ---
--- CURRENT H2 HEADINGS ---
${generatedH2s}
--- END CURRENT H2 HEADINGS ---
${siblingHeadings
  ? '--- HEADINGS ALREADY USED ELSEWHERE IN THIS SILO (' +
    material +
    ') — HARD LOCK ---\n' +
    'Do NOT rewrite any H2 into a heading that duplicates the structural pattern, opening word, or phrasing shape of these existing headings from other ' +
    material +
    ' articles.\n' +
    siblingHeadings +
    '\n--- END EXISTING SILO HEADINGS ---\n'
  : ''}
${minimumChangeStandard}
--- TASK ---
The H2 set failed Entity Coverage only for the specific items marked COVERED: NO above.
Apply the MINIMUM CHANGE STANDARD above to resolve only the COVERED: NO items.
HARD FIX RULES:
1. Every H2 associated only with COVERED: YES items is LOCKED and must be returned verbatim.
2. Change only an H2 whose wording genuinely needs adjustment to represent a COVERED: NO item.
3. Do not rewrite passing headings for style, variety, SEO improvement or better wording.
4. Do not change a heading merely because another version sounds stronger.
5. Do not add, remove, merge, split or renumber sections.
6. Preserve the existing section-to-topic relationship from Stage 1.5A.
7. A COVERED: NO item must be resolved by an actual H2 wording change where the current headings do not clearly represent it.
8. Do not merely explain that an existing vague H2 could cover the missing item.
The returned H2 itself must make the required coverage clear.
9. If one existing H2 can safely represent the missing item without changing its governed section role, apply the MINIMUM CHANGE STANDARD above to that H2 only.
10. If satisfying a missing item would require changing the governed purpose of a section, inventing a new section, or changing a COVERED: YES heading, do NOT force the change.
NO NEW FACTS — HARD LOCK:
Do NOT invent a specific technical detail, product, method name, quantity or outcome that is not already supported by:
- the existing heading;
- the Stage 1.5A section role;
- or the supplied governance context.
The article body has not yet been written.
A corrected heading must clarify the governed TOPIC, not assert an unverified fact about what happened.
When uncertain, choose wording that satisfies the missing coverage item while preserving the existing section role and applying the MINIMUM CHANGE STANDARD above.
Continue to follow the symptom-first rule:
H2 headings should describe the homeowner-recognisable problem, condition or decision rather than lead with a technical mechanism, unless the Article Type rules explicitly permit otherwise.
PAGE SCOPE LOCK — HARD LOCK:
ARTICLE TYPE:
${articleType}
MATERIAL:
${material}
Preserve the existing page subject and intent.
Do NOT broaden the page into a different:
- service;
- diagnostic;
- maintenance;
- installation;
- restoration;
- product-ranking;
- or unrelated comparison topic.
A correction may clarify missing coverage, but it must stay within the current ${articleType} subject and ${material} material scope.
Do NOT introduce a new competing material, service, problem, or decision framework merely to satisfy a validation item.
OUTPUT FORMAT:
SECTION 1: Heading H2: [unchanged or corrected H2]
SECTION 2: Heading H2: [unchanged or corrected H2]
...
Return the COMPLETE H2 list, including headings that did not need changing.
Last line must be:
Stage 1.5B coverage fix complete.
`.trim();
  return prompt;
}
/* ============================================================
   STAGE 1.5B — ENTITY COVERAGE REPAIR RECOMMENDATIONS
============================================================ */
function buildEntityCoverageRepairRecommendationsW15B() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });
  const colIdx =
    headers.indexOf(
      'Entity Coverage Check'
    );
  if (colIdx === -1) {
    return 'ERROR: Column "Entity Coverage Check" not found.';
  }
  const savedResult =
    String(
      sheet
        .getRange(
          row,
          colIdx + 1
        )
        .getValue() || ''
    ).trim();
  if (!savedResult) {
    return 'ERROR: No saved Entity Coverage result found.';
  }
  if (
    savedResult.indexOf(
      'FAIL'
    ) !== 0
  ) {
    return 'ERROR: Entity Coverage result is not a FAIL.';
  }
  const generatedH2s =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ''
    ).trim();
  if (!generatedH2s) {
    return 'ERROR: No H2 headings found in column CV (100).';
  }
  const d =
    getActiveRowDataMap();
  const articleType =
    String(
      d["Article Type"] ||
      "General"
    ).trim();
  const material =
    String(
      d["Stone Type"] ||
      "UNKNOWN"
    ).trim();
  const rewriteBrief =
    String(
      d["Page Rewrite Brief"] ||
      ""
    ).trim();
  const prompt = `
STAGE 1.5B — ENTITY COVERAGE REPAIR RECOMMENDATIONS
ROLE:
Senior SEO Content Structure Auditor
ARTICLE TYPE:
${articleType}
MATERIAL:
${material}
PAGE REWRITE BRIEF:
${rewriteBrief || "No active Rewrite Brief supplied."}
--- CURRENT H2 HEADINGS ---
${generatedH2s}
--- END CURRENT H2 HEADINGS ---
--- ENTITY COVERAGE FAILURE ---
${savedResult}
--- END ENTITY COVERAGE FAILURE ---
TASK:
The manually edited H2 set has failed Entity Coverage.
Identify ONLY the specific items marked:
COVERED: NO
For each missing item:
1. Identify the EXISTING SECTION that is the best fit to carry that missing coverage.
2. Prefer repairing one existing heading rather than adding a new section.
3. Preserve every H2 that already satisfies another required coverage item.
4. Do not damage or remove any item already marked COVERED: YES.
5. Preserve the existing article type, page subject, material scope and section role.
6. Do not introduce a new service, material, technical claim, process, problem or intent.
7. Do not rewrite the complete H2 set.
8. Generate THREE candidate replacement H2s for the selected section.
9. Internally test all three candidates against:
- the missing Entity Coverage requirement;
- Rule 17;
- the current H2 set for duplicate intent;
- the Page Rewrite Brief;
- the existing structural role of the selected section.
10. Return only candidates that fix the missing coverage without creating another obvious governance failure.
NO NEW FACTS — HARD LOCK:
The article body has not been written yet.
Do not invent:
- technical methods;
- products;
- equipment;
- guarantees;
- locations not already present;
- service claims;
- outcomes;
- qualifications;
- or other facts not already supported by the existing headings or governance context.
OUTPUT FORMAT:
MISSING REQUIREMENT:
[exact missing coverage item]
BEST SECTION TO REPAIR:
SECTION [N]: Heading H2: [current heading]
WHY THIS SECTION:
[brief plain-English explanation]
WHAT NEEDS TO CHANGE:
[clear instruction]
SUGGESTED ALTERNATIVES:
1. [replacement H2]
2. [replacement H2]
3. [replacement H2]
Repeat this block for every COVERED: NO item.
Do not output a complete replacement H2 set.
Last line must be exactly:
ENTITY COVERAGE REPAIR RECOMMENDATIONS COMPLETE
`.trim();
  return prompt;
}