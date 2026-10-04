/**
 * ================================================================================
 * ce_Stage15A_StructureAudit.gs - STAGE 1.5A — STRUCTURE AUDIT & MAPPING
 * ================================================================================
 * ce_Stage15A_StructureAudit
 * Audit checks and TSM-to-section mapping
 *
 * Updated: April 2026
 * - Integrated Strategy Filter: Reads Column AG (Secondary Intent Decisions).
 * - Implements "N" Exclusion: Automatically removes prohibited sections.
 * - Word Re-allocation: Moves "N" section budgets into Section 1.
 * ================================================================================
 */

function buildStage15APrompt() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();
  if (row < 2) return 'ERROR: Select a data row (row 2 or below).';

  const d             = getActiveRowDataMap();
  const material      = d["Stone Type"]      || "UNKNOWN";
  const articleType   = d["Article Type"]    || "General";
  const primaryEntity = d["Primary Entity"] || material;
  const commFlag      = d["COMM Flag"]      || "No";
  const geoctaFlag    = d["GEOCTA Flag"]    || "No";

  // New Variable: Read the Strategy Decisions (Column AG)
  const secondaryIntents = String(d["Secondary Intent Decisions"] || "");

// Existing article headings/structure from Column CT
const existingStructure = String(d["Current Headings"] || "").trim();

  // Auto-detect tier from Article Type
  const tierMap = {
    "Hub Page":          "Tier 1",
    "Educational Guide": "Tier 1",
    "Method Guide":      "Tier 2",
    "Service Page":      "Tier 2",
    "Geo Service Page":  "Tier 2",
    "Diagnostic Guide":  "Tier 3",
    "Buyer Guide":       "Tier 3",
    "FAQ Spoke":         "Tier 3",
    "Case Study":        "Tier 4"
  };

  const tierLabel = tierMap[articleType] || "Tier 2";

  // Get TSM requirements for this tier
  let tsmData = getTierStructuralRequirements(material, tierLabel, articleType);
  if (!tsmData || tsmData.length === 0) {
    return 'ERROR: No TSM requirements found for ' + tierLabel +
            '. Check "Tier Structural Coverage Matrix" sheet.';
  }

   // ==========================================
  // FILTER EXCLUDED SECTIONS AND THEIR BUDGET
  // ==========================================

  tsmData = tsmData.filter(function(req) {

    // If Strategy marks this section "N", remove both
    // the section and its associated word budget.
    if (
      secondaryIntents.includes(
        "N — " + req.name
      )
    ) {
      return false;
    }

    return true;
  });

  // ==========================================

  // Read On-Page Content Expectations from Article Type Control Sheet
  let onPageExpectations = "";
  let reasoning          = "";

  try {
    const artSheet = ss.getSheets().find(
      function(s) { return s.getSheetId() == BC_SHEET_CONFIG.articleTypeControl; }
    );
    if (artSheet) {
      const artData = artSheet.getDataRange().getValues();
      for (let i = 1; i < artData.length; i++) {
        if (String(artData[i][0]).trim() === articleType) {
          onPageExpectations = String(artData[i][3] || "").trim();
          reasoning          = String(artData[i][8] || "").trim();
          break;
        }
      }
    }
  } catch (e) {
    onPageExpectations = "(Article Type Control Sheet read error: " + e.message + ")";
  }

  if (!onPageExpectations) {
    onPageExpectations = "(No on-page expectations defined for this article type.)";
  }

  const tierWordCount =
    tsmData.reduce(function(total, req) {
      return total + Number(req.wordCount || 0);
    }, 0);

  // Build TSM block using the NEW FILTERED DATA
  let tsmBlock = '';
  tsmData.forEach(function(req, index) {
    const minWords = req.wordCount;
    const maxWords = Math.round(minWords * 1.2);
    // Note: We use index+1 to keep the numbering clean for the LLM
    tsmBlock += 'Order ' + (index + 1) + ': ' + req.name +
                ' (' + minWords + '–' + maxWords + ' words)\n';
  });

  const prompt = `
STAGE 1.5A — STRUCTURE AUDIT & MAPPING

MATERIAL: ${material}
ARTICLE TYPE: ${articleType}
TIER: ${tierLabel}
PRIMARY ENTITY: ${primaryEntity}
MINIMUM WORD BUDGET: ${tierWordCount} words
COMM FLAG: ${commFlag}
GEOCTA FLAG: ${geoctaFlag}

--- FAILURE RULE ---
If any check fails, you MUST output a one-sentence reason immediately after the FAIL result.
Format: AUDIT CHECK [N]: FAIL — [one sentence reason]
A FAIL result with no reason is not permitted.

--- AUDIT CHECKS ---
Run the following 11 checks on the content structure:

AUDIT CHECK 1 — ARTICLE TYPE CLASSIFICATION
Verify the article type matches the structural expectations below.
AUDIT CHECK 1: PASS / FAIL

AUDIT CHECK 2 — TIER WORD BUDGET RANGE
Minimum word budget for ${tierLabel} is ${tierWordCount} words. 
AUDIT CHECK 2: PASS / FAIL

AUDIT CHECK 3 — TSM REQUIREMENT COVERAGE
Count the number of TSM structural requirements listed below.

For Case Study, Method Guide, Diagnostic Guide and Educational Guide articles, these are COVERAGE REQUIREMENTS, not a required number of article sections or H2 headings.
A single TSM requirement may be covered across multiple chronological article sections where this improves the clarity of the project story.
Existing useful H2/H3 sections must not be removed, merged or suppressed merely to make the article section count equal the TSM requirement count.

For Case Study, Method Guide, Diagnostic Guide and Educational Guide articles, do NOT count the header block or HUB-INTRO as TSM sections.

If the project problem or challenge is only described in the opening introduction and there is no clear H2 introducing that problem, create a dedicated Problem/Challenge H2 before the intervention sequence.

Do not add this H2 if an existing H2 already clearly introduces the project problem.

For all other article types, retain the existing TSM section-count interpretation unless another article-type rule explicitly states otherwise.

AUDIT CHECK 3: ${tsmData.length} TSM requirements must be covered

AUDIT CHECK 4 — WORD BUDGET DISTRIBUTION
Each section has a governed word range defined below.
AUDIT CHECK 4: PASS / FAIL

AUDIT CHECK 5 — COMM FLAG VALIDATION
COMM: ${commFlag}
AUDIT CHECK 5: PASS / FAIL / NOT APPLICABLE

AUDIT CHECK 6 — GEOCTA FLAG VALIDATION
GEOCTA: ${geoctaFlag}
AUDIT CHECK 6: PASS / FAIL / NOT APPLICABLE

AUDIT CHECK 7 — PRIMARY ENTITY ALIGNMENT
AUDIT CHECK 7: PASS / FAIL

AUDIT CHECK 8 — ARTICLE TYPE EXPECTATIONS
Validate only against the block below.
AUDIT CHECK 8: PASS / FAIL

AUDIT CHECK 9 — TSM ORDER COVERAGE

For Case Study, Method Guide, Diagnostic Guide and Educational Guide articles:
Verify every TSM Order from 1 to ${tsmData.length} is covered at least once across the mapped sections.
A TSM Order MAY appear on more than one section where separate reader-facing sections genuinely contribute to the same coverage role.
Do not treat repeated use of an Order as duplication when those sections cover distinct parts of that requirement.

For all other article types:
Verify every TSM Order from 1 to ${tsmData.length} appears exactly once across the mapped sections — no Order missing and no Order duplicated.

Sections do NOT need to appear in numerical TSM Order sequence where a different narrative sequence better serves the reader.

AUDIT CHECK 9: PASS / FAIL

AUDIT CHECK 10 — SECTION STRUCTURE VALIDATION

For Case Study, Method Guide, Diagnostic Guide and Educational Guide articles, do NOT require the number of article sections to equal the number of TSM requirements.

PASS if:
- every required TSM coverage area is represented somewhere in the article structure;
- the section order tells the project story clearly and chronologically;
- useful existing H2/H3 sections are retained where they help the story;
- additional sections are permitted where they represent genuine project stages such as problem, assessment, test clean, cleaning, rinsing, drying, repair, sealing, maintenance or outcome.

FAIL only if the proposed structure omits a required TSM coverage area, contains unnecessary or duplicate sections, or leaves the project story confused or out of sequence.

For all other article types, retain the existing rule that section count must match the TSM requirement count unless another article-type rule explicitly states otherwise.

Do NOT count the header block. Do NOT count HUB-INTRO if present.
AUDIT CHECK 10: PASS / FAIL

AUDIT CHECK 11 — WORD BUDGET TOTAL
Total must fall within: ${tierWordCount}–${Math.round(tierWordCount * 1.2)} words.
AUDIT CHECK 11: PASS / FAIL

--- ARTICLE TYPE EXPECTATIONS ---
Article Type: ${articleType}
On-Page Content Expectations: ${onPageExpectations}
${reasoning ? 'Content Strategy: ' + reasoning : ''}

--- TSM REQUIREMENTS (${material} ${tierLabel}) ---
${tsmBlock}

--- EXISTING ARTICLE STRUCTURE FROM COLUMN CT ---
${existingStructure || "(No existing structure found in Column CT.)"}

For Case Study, Method Guide, Diagnostic Guide and Educational Guide articles:
Use the existing structure above as the starting point.
Preserve useful existing H2/H3 headings where they represent genuine reader-facing stages or topics.
Only merge, remove, rename or reorder them where necessary to satisfy governance, correct a confused structure, or improve the logical sequence.
Do not collapse useful existing sections merely to match the number of TSM requirements.

--- TSM-TO-SECTION MAPPING ---
${articleType === "Hub Page" ? "HUB-INTRO HARD LOCK: This is a Hub Page. You MUST output HUB-INTRO as the FIRST line of your section mapping — before SECTION 1. A missing HUB-INTRO is a mapping failure.\n" : ""}

${["Case Study", "Method Guide", "Diagnostic Guide", "Educational Guide"].indexOf(articleType) !== -1 ? `
FLEXIBLE COVERAGE STRUCTURE RULE:
Treat each TSM Order below as a COVERAGE ROLE, not as a fixed one-section-per-order requirement.

Every TSM Order must be covered somewhere in the final article structure, but a single TSM Order may be represented across multiple sections where that better reflects the real project chronology.

Do NOT remove, merge or suppress useful existing H2/H3 sections merely to make the number of article sections equal the number of TSM Orders.

First inspect the existing article structure.

Preserve useful existing H2/H3 sections where they represent genuine project stages.

If the original article is haphazard, reorganise only as much as necessary to create a clear chronological project story.

For Case Study articles only: if the project problem or challenge is only described in the opening introduction and there is no clear H2 introducing it, create a dedicated Problem/Challenge section before the intervention sequence.

Do not add that section if an existing H2 already clearly introduces the problem.

A TSM role such as Problem & Intervention may legitimately span several sections, for example:
- homeowner problem or challenge;
- assessment or test clean;
- main cleaning or restoration work;
- rinsing or extraction;
- drying;
- repair;
- sealing or protection.

Only include stages that actually occurred in the project.

SECTION ORDER:
Choose the section sequence that best tells the real project story clearly and chronologically.

Every required TSM coverage role must be represented, but TSM Orders do NOT need a one-to-one relationship with article sections.
` : `
Map every TSM Order below to a section. Every TSM Order must appear exactly once across the mapped sections — none may be omitted or duplicated.

SECTION ORDER (Content-Driven — Soft Lock): You do not need to map TSM Orders sequentially (1, 2, 3...). Choose the narrative sequence that best serves the reader's understanding for THIS specific article — for example, leading with the most urgent symptom, or grouping related causes together, even if that means Order 3 appears before Order 1. Every TSM Order must still be covered exactly once; only the SEQUENCE in which they appear as sections is flexible.
`}

For Hub Page articles, output HUB-INTRO as a separate structural block before SECTION 1.
HUB-INTRO is mandatory, is not a TSM section, and does not count toward the ${tsmData.length} required TSM sections.
Use this exact format:
HUB-INTRO: id="hub-intro" | H2="None" | CONTENT="One orienting sentence and a quick-links navigation list only"

Then output the section mapping using this format:
SECTION 1: Order [N] — [Requirement Name] — [min]–[max] words

For Case Study, Method Guide, Diagnostic Guide and Educational Guide articles, the same Order [N] may appear on more than one SECTION where several genuine reader-facing sections collectively satisfy that TSM coverage role.

Do not duplicate an Order merely to increase section count. Repeat an Order only where the article genuinely needs separate sections for distinct parts of that coverage area.
...

--- OUTPUT INSTRUCTION ---
Output ONLY the 11 audit results followed by the section mapping.
For Hub Page articles, include the HUB-INTRO mapping line before SECTION 1.
Last line: "Stage 1.5A complete. Waiting for Stage 1.5B."
`.trim();

  return prompt;
}