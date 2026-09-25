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
  // NEW LOGIC: FILTER & RE-ALLOCATE BUDGET
  // ==========================================
  let spareWords = 0;
  tsmData = tsmData.filter(function(req) {
    // If the Strategy Column marks this section name with an "N", remove it
    if (secondaryIntents.includes("N — " + req.name)) {
      spareWords += (req.wordCount || 0); 
      return false; // Remove from list
    }
    return true; // Keep in list
  });

  // Re-allocate deleted section words to the Geological/Primary section (Order 1)
  if (tsmData.length > 0 && spareWords > 0) {
    tsmData[0].wordCount = (tsmData[0].wordCount || 0) + spareWords;
  }
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

  const tierWordCount = getTierWordCount(tierLabel, articleType);

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

AUDIT CHECK 3 — TSM REQUIREMENT COUNT
Count the number of mapped TSM sections listed below.
Do NOT count the header block. Do NOT count HUB-INTRO if present.
AUDIT CHECK 3: ${tsmData.length} sections required

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

AUDIT CHECK 9 — TSM ORDER COMPLETENESS
Verify every TSM Order from 1 to ${tsmData.length} appears exactly once across the mapped sections — no Order missing, no Order duplicated. Sections do NOT need to appear in sequential Order — a content-driven sequence is permitted (see SECTION ORDER rule below).
AUDIT CHECK 9: PASS / FAIL

AUDIT CHECK 10 — SECTION COUNT MATCH
Expected: ${tsmData.length} mapped sections
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

--- TSM-TO-SECTION MAPPING ---
${articleType === "Hub Page" ? "HUB-INTRO HARD LOCK: This is a Hub Page. You MUST output HUB-INTRO as the FIRST line of your section mapping — before SECTION 1. A missing HUB-INTRO is a mapping failure.\n" : ""}Map every TSM Order below to a section. Every TSM Order must appear exactly once across the mapped sections — none may be omitted or duplicated.

SECTION ORDER (Content-Driven — Soft Lock): You do not need to map TSM Orders sequentially (1, 2, 3...). Choose the narrative sequence that best serves the reader's understanding for THIS specific article — for example, leading with the most urgent symptom, or grouping related causes together, even if that means Order 3 appears before Order 1. Every TSM Order must still be covered exactly once; only the SEQUENCE in which they appear as sections is flexible.

For Hub Page articles, output HUB-INTRO as a separate structural block before SECTION 1.
HUB-INTRO is mandatory, is not a TSM section, and does not count toward the ${tsmData.length} required TSM sections.
Use this exact format:
HUB-INTRO: id="hub-intro" | H2="None" | CONTENT="One orienting sentence and a quick-links navigation list only"

Then output the TSM section mapping using this format:
SECTION 1: Order [N] — [Requirement Name] — [min]–[max] words
...

--- OUTPUT INSTRUCTION ---
Output ONLY the 11 audit results followed by the section mapping.
For Hub Page articles, include the HUB-INTRO mapping line before SECTION 1.
Last line: "Stage 1.5A complete. Waiting for Stage 1.5B."
`.trim();

  return prompt;
}