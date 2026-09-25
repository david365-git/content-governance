/**
 * ================================================================================
 * ce_Stage15C_Enrichment.gs - STAGE 1.5C — SECTION ENRICHMENT
 * ================================================================================
 *
 * Add images, visual patterns, entities, and content briefs
 *
 * Reads:  Col 99  (CU) — W1.5A Audit Check
 *         Col 100 (CV) — W1.5B H2s
 * Saves:  Col 101 (CW) — W1.5C Enriched Plan
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * ================================================================================
 */
function buildStage15CPrompt() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const d =
    getActiveRowDataMap();
  const material =
    d["Stone Type"] ||
    "UNKNOWN";
  const articleType =
    d["Article Type"] ||
    "General";
  const primaryEntity =
    d["Primary Entity"] ||
    material;
  const feedsHub =
    String(
      d["Feeds Hub"] ||
      ""
    ).trim();
  // --------------------------------------------------
  // FRESH UPSTREAM SOURCE — W1.5A
  // --------------------------------------------------
  const structure15A =
    String(
      sheet
        .getRange(row, 99)
        .getValue() || ""
    ).trim();
  if (!structure15A) {
    return 'ERROR: No Stage 1.5A output found in CU. Run W1.5A first.';
  }
  // --------------------------------------------------
  // FRESH UPSTREAM SOURCE — W1.5B
  // --------------------------------------------------
  const h2s15B =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ""
    ).trim();
  if (!h2s15B) {
    return 'ERROR: No Stage 1.5B output found in CV. Run W1.5B first.';
  }
  if (
    h2s15B.indexOf(
      'FINAL FAILURE REPORT'
    ) !== -1
  ) {
    return (
      'ERROR: W1.5C blocked — Column CV contains a FINAL FAILURE REPORT. ' +
      'Resolve the W1.5B heading failure before running W1.5C.'
    );
  }
  // --------------------------------------------------
  // ORIGINAL HTML
  // --------------------------------------------------
  const postId =
    String(
      d["Post ID"] ||
      ""
    ).trim();
  let originalHtml = "";
  if (postId) {
    try {
      const exportSheet =
        ss.getSheetByName(
          "site-export"
        );
      if (exportSheet) {
        const exportData =
          exportSheet
            .getDataRange()
            .getValues();
        if (exportData.length > 0) {
          const exportHeaders =
            exportData[0];
          const idCol =
            exportHeaders.indexOf(
              "ID"
            );
          const htmlCol =
            exportHeaders.indexOf(
              "Full Post HTML"
            );
          if (
            idCol > -1 &&
            htmlCol > -1
          ) {
            for (
              let i = 1;
              i < exportData.length;
              i++
            ) {
              if (
                String(
                  exportData[i][idCol]
                ).trim() === postId
              ) {
                originalHtml =
                  String(
                    exportData[i][htmlCol] ||
                    ""
                  ).trim();
                break;
              }
            }
          }
        }
      }
    } catch (e) {
      // Original HTML is enrichment context only.
      // Do not stop W1.5C solely because site-export could not be read.
      originalHtml = "";
    }
  }
  // --------------------------------------------------
  // GOVERNANCE CONTEXT
  // --------------------------------------------------
  const recoveryBlueprint =
    String(
      d["Recovery Blueprint"] ||
      ""
    ).trim();
  const entityGapAnalysis =
    String(
      d["W0 Entity Gap Analysis"] ||
      ""
    ).trim();
  // --------------------------------------------------
  // EXISTING IMAGE INVENTORY
  // --------------------------------------------------
  let imageCount = 0;
  let imageList = "";
  if (originalHtml) {
    const cleanHtml =
      originalHtml.replace(
        /""/g,
        '"'
      );
    const imgMatches =
      cleanHtml.match(
        /<img[^>]+>/gi
      );
    if (imgMatches) {
      const SKIP_FILENAMES =
        /David_Allen\.jpg/i;
      const SKIP_DOMAINS =
        /m\.media-amazon\.com|gravatar\.com|logo|icon|avatar|sprite|pixel|tracking/i;
      const SKIP_THUMBS =
        /\-\d+x\d+(@\d+x)?\.(jpg|jpeg|png|webp|gif|avif)$/i;
      let filteredCount = 0;
      imgMatches.forEach(
        function(tag) {
          const srcMatch =
            tag.match(
              /src=["']([^"']+)["']/i
            );
          if (!srcMatch) {
            return;
          }
          const fullSrc =
            srcMatch[1];
          const parts =
            fullSrc.split('/');
          const filename =
            parts[
              parts.length - 1
            ];
          if (
            SKIP_FILENAMES.test(
              filename
            )
          ) {
            return;
          }
          if (
            SKIP_DOMAINS.test(
              fullSrc
            )
          ) {
            return;
          }
          if (
            SKIP_THUMBS.test(
              filename
            )
          ) {
            return;
          }
          filteredCount++;
          imageList +=
            "IMAGE " +
            filteredCount +
            ": " +
            filename +
            "\n";
        }
      );
      imageCount =
        filteredCount;
    }
  }
  // --------------------------------------------------
  // EXISTING VIDEO INVENTORY
  // --------------------------------------------------
  let videoCount = 0;
  let videoList = "";
  if (originalHtml) {
    const iframeRe =
      /<iframe[^>]+src=["']([^"']*(?:youtube|vimeo|wistia|loom|embed)[^"']*)["'][^>]*>/gi;
    let vm;
    let vidIdx = 0;
    while (
      (vm = iframeRe.exec(
        originalHtml
      )) !== null
    ) {
      videoCount++;
      vidIdx++;
      videoList +=
        "VIDEO " +
        vidIdx +
        ": " +
        vm[1] +
        "\n";
    }
  }
  let prompt = `
STAGE 1.5C — SECTION ENRICHMENT
ROLE:
SEO Content Strategist
MATERIAL:
${material}
ARTICLE TYPE:
${articleType}
PRIMARY ENTITY:
${primaryEntity}
--- SYSTEM INSTRUCTION ---
You have the governed section structure from Stage 1.5A and the validated H2 set from Stage 1.5B.
Your job is to enrich that EXISTING structure.
You may add:
1. Image assignments
2. Video assignments
3. Visual patterns
4. Entities to include
5. Section content briefs
6. Required internal-link placement
You must NOT change:
- section numbers;
- section order;
- H2 wording;
- TSM Requirement ownership;
- Word Budget ownership;
- HUB-INTRO status;
- or the governed purpose of any section.
Stage 1.5C enriches the approved structure. It does not redesign it.
--- IMAGE PRESERVATION — HARD LOCK ---
ORIGINAL IMAGE COUNT:
${imageCount}
${imageList || "No eligible site images found in original HTML."}
If ORIGINAL IMAGE COUNT is ${imageCount}, you must assign all ${imageCount} eligible images somewhere in the enriched plan.
Do not silently discard an eligible original image.
David_Allen.jpg is excluded because it is reserved for the author bio.
Amazon product images from m.media-amazon.com are excluded from section image assignments.
IMAGE ASSIGNMENT RULE:
Assign each eligible image to the section where its existing context is most relevant.
Multiple images may be assigned to one section if the section can support them.
The Images field must use the EXACT filename shown above.
Examples:
Images: marble-crystals.webp
Images: marble-before.webp, marble-after.webp
Do not invent filenames.
ESCALATION-BOUNDARY PRODUCT IMAGE RULE — HARD LOCK:
If an image depicts a professional-only product or material named inside an escalation boundary, assign it to the professional escalation or handoff section.
Do not place a professional-only product image beside DIY avoidance advice in a way that visually suggests homeowner use.
IMAGE CONTEXT RULE — HARD LOCK:
Every image used in the eventual article must have substantive explanatory text before it.
No two images should be planned consecutively without substantive text between them.
If a section's word budget cannot support the assigned images, move the lowest-priority image to another genuinely relevant section.
--- VIDEO PRESERVATION — HARD LOCK ---
ORIGINAL VIDEO COUNT:
${videoCount}
${videoList || "No video embeds found in original HTML."}
${videoCount > 0
  ? `
The original HTML contains ${videoCount} video embed(s).
ALL must be preserved.
Video 1 must be assigned to the HEADER using:
Header Video: [exact src URL]
Video 2 onwards may be assigned to the most contextually appropriate numbered section using:
Videos: [exact src URL]
Use the EXACT URL shown above.
Do not modify, shorten or replace it.
`
  : `
There are no original videos to preserve.
`}
--- VISUAL PATTERN SELECTION ---
Use visual patterns strategically.
Available patterns:
1. Symptom-Anchor
2. Diagnostic-Sequence
3. Comparison-Paragraphs
4. Mechanism-Blockquote
5. H3-Subheadings
6. None
Do not assign a pattern simply to create variety.
SYMPTOM-ANCHOR:
Use where the section covers homeowner-recognisable conditions or variants.
DIAGNOSTIC-SEQUENCE:
Use only when the content genuinely contains at least four discrete diagnostic or procedural items that a reader needs to scan.
COMPARISON-PARAGRAPHS:
Use when the section genuinely contrasts two conditions, choices or states.
MECHANISM-BLOCKQUOTE:
Use where one critical explanatory principle merits being isolated.
H3-SUBHEADINGS:
Before assigning H3-Subheadings, count the entities planned for that section.
If fewer than 3 distinct entities are being covered, H3-Subheadings is NOT eligible.
Do not use H3s merely to make a short section look structured.
Use a maximum of approximately 3-4 visual patterns across the entire article.
Sections below roughly 200 words rarely need a visual pattern.
`;
  prompt += `
--- ENTITY SELECTION ---
Use entities from the Recovery Blueprint below.
Select only entities genuinely relevant to the section topic.
Prefer 1-3 entities per section.
Do not overload a section merely because more entities are available.
RECOVERY BLUEPRINT:
${recoveryBlueprint || "No Recovery Blueprint available."}
--- W0 ENTITY GAP ANALYSIS — PRIORITY CONTEXT ---
The following entities were flagged at W0 as missing from the page's current search footprint.
Where relevant to the governed section topic, prefer these gap entities over generic Recovery Blueprint entities.
Do not force a gap entity into an unrelated section.
${entityGapAnalysis || "No W0 Entity Gap Analysis available for this row."}
--- INTERNAL LINK TO HUB — HARD LOCK ---
Feeds Hub URL:
${feedsHub || "none available — omit internal link"}
If a section's TSM Requirement is:
Internal Link to Hub
then that section's Content Brief must explicitly direct the reader toward the hub.
Its:
Internal link:
field must contain the exact Feeds Hub URL above, unless no hub URL exists.
Do not populate Internal link for unrelated sections.
--- SECTION LOAD BUDGET — HARD LOCK ---
Before assigning enrichment elements to a section, calculate the approximate load:
- Image assigned: 40 words
- Visual pattern assigned other than None: 30 words
- Each entity assigned: 15 words
- Internal link assigned: 25 words
Subtract this load from the section's maximum Word Budget.
If fewer than 60 words remain, the section is overloaded.
Reduce load in this order:
1. Remove the visual pattern first.
2. Reduce entities beyond the single most important one.
3. Move an image to another genuinely relevant section if necessary.
Never assign:
image + visual pattern + 3 entities + internal link
to a section whose maximum word budget is below 270 words.
--- CONTENT BRIEF RULES ---
Write 1-2 sentences describing what the section will cover.
The brief must:
- remain specific to this article;
- preserve the governed section role;
- avoid introducing a new topic;
- focus on the practical reader outcome;
- avoid making unsupported factual claims;
- avoid expanding beyond the supplied section structure.
ENTITY NAME PHRASING RULE — HARD LOCK:
If a Content Brief references a governed entity from the Recovery Blueprint or Entities to include field, do not write that entity in Title Case as though it were a branded or named process.
Translate it into natural descriptive language.
WRONG:
"Detail the Two-Stage Filling Approach."
CORRECT:
"Detail the two-step filling process."
WRONG:
"Explain the Tailored Maintenance Handover."
CORRECT:
"Explain the maintenance routine given to the homeowner, tailored to this floor."
The Content Brief tells the later writer what concept to cover.
It must not accidentally encourage governance entity labels to appear verbatim in body copy.
--- AUTHORITY DISTRIBUTION RULE — HARD LOCK ---
For Case Study structures using the standard four-section authority sequence:
Section 1:
entry condition and problem identification only.
Do not front-load:
- technical mechanism;
- process rationale;
- specialist judgement;
- or final outcome.
Section 2:
technical mechanism and cause explanation.
Section 3:
professional judgement and process rationale.
Section 4:
outcome, transformation, maintenance handover or the governed final-section role.
Where the Stage 1.5A structure differs from this standard pattern, follow Stage 1.5A rather than forcing the four-section example.
--- FRESH GOVERNED SECTION STRUCTURE FROM W1.5A ---
${structure15A}
--- FRESH VALIDATED H2 SET FROM W1.5B ---
${h2s15B}
--- STRUCTURAL PRESERVATION — HARD LOCK ---
The Stage 1.5A and Stage 1.5B content above is authoritative.
For every numbered section:
1. Preserve the SECTION number exactly.
2. Preserve the H2 wording exactly.
3. Preserve the TSM Requirement assigned by Stage 1.5A.
4. Preserve the Word Budget assigned by Stage 1.5A.
5. Do not move one section's TSM Requirement to another section.
6. Do not add a new numbered section.
7. Do not remove a numbered section.
8. Do not merge or split sections.
9. Do not rewrite an H2 for clarity, style, SEO or preference.
10. Do not repair Stage 1.5A or Stage 1.5B here.
Stage 1.5C adds enrichment only.
--- HUB-INTRO HANDLING RULE ---
If Stage 1.5A or Stage 1.5B includes:
HUB-INTRO: id="hub-intro" | H2="None" | CONTENT="One orienting sentence and a quick-links navigation list only"
preserve HUB-INTRO as a special structural block.
Do not convert HUB-INTRO into a numbered section.
Output HUB-INTRO exactly in this format:
HUB-INTRO:
Content Brief: One orienting sentence and a quick-links navigation list only
Visual Pattern: None
List eligible: NO — navigation only
Images: None
Videos: None
Entities to include: None
Internal link: None
--- LIST ELIGIBILITY — RULE 16 GATE ---
For each numbered section decide whether a list is genuinely appropriate.
Mark:
List eligible: YES — procedural/diagnostic only
only when ALL THREE conditions are true:
1. The items are genuinely discrete.
2. There will be at least four items.
3. The reader benefits from scanning and identifying them rather than reading connected explanation.
If any condition fails:
List eligible: NO — prose required
A section with fewer than four discrete items must be marked NO.
METHOD GUIDE NAMED PRODUCT EXCEPTION — HARD LOCK:
If the Article Type is Method Guide and the original HTML contains named proprietary products inside an existing list or step sequence, preserve those named product steps in the generation brief rather than generalising them away.
Where applicable mark:
List eligible: YES — named product steps preserved from original HTML
Do not invent a product name that is not present in the source HTML.
--- OUTPUT FORMAT — HARD LOCK ---
If HUB-INTRO exists, output it first in the exact special format above.
Then output EVERY numbered section.
For each numbered section use exactly this structure:
SECTION [N]: Heading H2: [copy the Stage 1.5B H2 verbatim]
TSM Requirement: [copy the Stage 1.5A TSM Requirement verbatim]
Word Budget: [copy the Stage 1.5A Word Budget verbatim]
Content Brief: [1-2 sentences]
Visual Pattern: [Symptom-Anchor / Diagnostic-Sequence / Comparison-Paragraphs / Mechanism-Blockquote / H3-Subheadings / None]
List eligible: [YES — procedural/diagnostic only / YES — named product steps preserved from original HTML / NO — prose required]
Images: [exact eligible filenames separated by commas, or None]
Videos: [exact iframe src URL(s), or None]
Entities to include: [0-3 entities from Recovery Blueprint that directly support THIS SECTION'S governed TSM Requirement, Primary Entity and Rewrite Brief scope, or "None"]

ENTITY ROLE GATE — HARD LOCK:
A Recovery Blueprint entity is NOT automatically required body coverage merely because it is technically valid for the material.

Before assigning any entity:
1. Confirm it directly serves the section's stated TSM Requirement.
2. Confirm it supports the page's Primary Entity and Article Type.
3. Confirm it does not introduce an intent prohibited or minimised by the Rewrite Brief SCOPE BOUNDARY.
4. Confirm it does not turn a principle-level or decision-support section into a standalone cleaning, repair, maintenance, sealing, polishing or other method workflow.
5. If the entity belongs mainly to a secondary or adjacent intent, leave it out of the section and use "None" unless the section explicitly owns that supporting intent.

The Rewrite Brief scope overrides Recovery Blueprint completeness.
Do not force technical entities into a section merely to achieve entity coverage.
Internal link: [exact Feeds Hub URL only where governed, otherwise None]
Do not omit any required field.
Do not add extra fields.
Do not include commentary between sections.
--- COMPLETENESS CHECK BEFORE OUTPUT ---
Before finalising, internally verify:
1. Every Stage 1.5A numbered section is present exactly once.
2. No extra numbered section exists.
3. Every H2 matches Stage 1.5B exactly.
4. Every section contains all required output fields.
5. Every eligible original image is assigned exactly once unless the same source image is genuinely reused in the original content.
6. Every original video is preserved.
7. HUB-INTRO is preserved if present.
8. Internal Link to Hub is assigned only to the governed section.
9. No section has been overloaded beyond its available word budget.
10. No enrichment choice changes the governed purpose of a section.
Last line must be exactly:
Stage 1.5C complete. Waiting for Stage 1.5D.
No preamble.
No explanation.
`;
  return prompt.trim();
}
/* ============================================================
   STAGE 1.5C — TARGETED REPAIR PROMPT
   Used only when saveStage15CEnrichedPlan() rejects the
   generated plan.
   The repair prompt works from:
   - the CURRENT upstream CU/CV;
   - the rejected W1.5C output;
   - the exact validation failure.
   It must not redesign passing sections.
============================================================ */
function buildStage15CRepairPrompt(
  rejectedPlan,
  validationFailure
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return 'ERROR: Select a data row (row 2 or below).';
  }
  const structure15A =
    String(
      sheet
        .getRange(row, 99)
        .getValue() || ''
    ).trim();
  const h2s15B =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ''
    ).trim();
  if (!structure15A) {
    return 'ERROR: Column CU is empty — W1.5A must complete before W1.5C repair.';
  }
  if (!h2s15B) {
    return 'ERROR: Column CV is empty — W1.5B must complete before W1.5C repair.';
  }
  rejectedPlan =
    String(
      rejectedPlan || ''
    ).trim();
  validationFailure =
    String(
      validationFailure || ''
    ).trim();
  if (!rejectedPlan) {
    return 'ERROR: No rejected W1.5C plan supplied.';
  }
  if (!validationFailure) {
    return 'ERROR: No W1.5C validation failure supplied.';
  }
  const minimumChangeStandard =
    bc_getMinimumChangeStandard_();
  const prompt = `
STAGE 1.5C — ENRICHED PLAN REPAIR
ROLE:
SEO Content Structure Repair Auditor
--- AUTHORITATIVE W1.5A STRUCTURE ---
${structure15A}
--- END W1.5A STRUCTURE ---
--- AUTHORITATIVE W1.5B H2 SET ---
${h2s15B}
--- END W1.5B H2 SET ---
--- REJECTED W1.5C PLAN ---
${rejectedPlan}
--- END REJECTED PLAN ---
--- VALIDATION FAILURE ---
${validationFailure}
--- END VALIDATION FAILURE ---
${minimumChangeStandard}
--- TASK ---
Repair ONLY the defect identified in the validation failure.
HARD RULES:
1. Preserve every unaffected section exactly where possible.
2. Preserve all Stage 1.5A section numbers and order.
3. Preserve every Stage 1.5B H2 character-for-character.
4. Preserve every correct TSM Requirement.
5. Preserve every correct Word Budget.
6. Do not rewrite Content Briefs merely for style or preference.
7. Do not redistribute images, videos, entities, patterns or links unless the validation failure specifically requires it.
8. Do not introduce new sections.
9. Do not delete governed sections.
10. Do not merge or split sections.
11. Do not invent images, filenames, video URLs, entities, links, products, technical claims or outcomes.
12. If one missing field or malformed line can be fixed directly, do only that.
13. If one section is responsible for the failure, do not alter unrelated sections.
14. Apply the MINIMUM CHANGE STANDARD above.
--- OUTPUT FORMAT ---
Return the COMPLETE corrected W1.5C enriched plan.
Preserve the same output structure used by Stage 1.5C.
Do not explain the repair.
Last line must be exactly:
Stage 1.5C complete. Waiting for Stage 1.5D.
`.trim();
  return prompt;
}
/* ============================================================
   STAGE 1.5C — SHARED STRUCTURE HELPERS
   These helpers give the W1.5C save function an authoritative
   map of the CURRENT CU/CV structure before anything is allowed
   to overwrite CW.
============================================================ */
function ce_getStage15CGovernedRequirements_() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();
  const sheet =
    ss.getSheetByName('posts');
  const row =
    sheet.getActiveRange().getRow();
  if (row < 2) {
    return {
      success: false,
      message: 'Select a data row first.'
    };
  }
  const structure15A =
    String(
      sheet
        .getRange(row, 99)
        .getValue() || ''
    ).trim();
  const h2s15B =
    String(
      sheet
        .getRange(row, 100)
        .getValue() || ''
    ).trim();
  if (!structure15A) {
    return {
      success: false,
      message:
        'Column CU is empty — W1.5A must complete before W1.5C.'
    };
  }
  if (!h2s15B) {
    return {
      success: false,
      message:
        'Column CV is empty — W1.5B must complete before W1.5C.'
    };
  }
  const expectedSections = [];
  const expectedH2s = {};
  const structureRegex =
    /^SECTION\s+(\d+):/gmi;
  let structureMatch;
  while (
    (
      structureMatch =
        structureRegex.exec(
          structure15A
        )
    ) !== null
  ) {
    const number =
      Number(
        structureMatch[1]
      );
    if (
      expectedSections.indexOf(
        number
      ) === -1
    ) {
      expectedSections.push(
        number
      );
    }
  }
  expectedSections.sort(
    function(a, b) {
      return a - b;
    }
  );
  if (
    expectedSections.length === 0
  ) {
    return {
      success: false,
      message:
        'Could not determine governed section numbers from W1.5A in CU.'
    };
  }
  const h2Regex =
    /^SECTION\s+(\d+):\s*(?:Heading H2:\s*)?(.+)$/gmi;
  let h2Match;
  while (
    (
      h2Match =
        h2Regex.exec(
          h2s15B
        )
    ) !== null
  ) {
    const number =
      Number(
        h2Match[1]
      );
    const heading =
      String(
        h2Match[2] || ''
      ).trim();
    if (heading) {
      expectedH2s[number] =
        heading;
    }
  }
  const missingH2s = [];
  expectedSections.forEach(
    function(number) {
      if (
        !expectedH2s[number]
      ) {
        missingH2s.push(
          number
        );
      }
    }
  );
  if (
    missingH2s.length > 0
  ) {
    return {
      success: false,
      message:
        'Current W1.5B output is incomplete — missing H2 for SECTION ' +
        missingH2s.join(', SECTION ') +
        '.'
    };
  }
  const expectsHubIntro =
    /^HUB-INTRO:/mi.test(
      structure15A
    ) ||
    /^HUB-INTRO:/mi.test(
      h2s15B
    );
  return {
    success: true,
    row: row,
    structure15A:
      structure15A,
    h2s15B:
      h2s15B,
    expectedSections:
      expectedSections,
    expectedH2s:
      expectedH2s,
    expectsHubIntro:
      expectsHubIntro
  };
}
/* ============================================================
   Remove accidental markdown code fences from W1.5C output.
   This does not alter the plan itself.
============================================================ */
function ce_stripStage15CCodeFences_(text) {
  return String(
    text || ''
  )
    .trim()
    .replace(
      /^```(?:text|plaintext|markdown)?\s*/i,
      ''
    )
    .replace(
      /\s*```$/i,
      ''
    )
    .trim();
}