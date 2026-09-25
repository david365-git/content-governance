/**
 * ============================================================
 * GPT STRUCTURE-LOCKED GOVERNANCE ENGINE — 3 SECTION OUTPUT
 * File: Analysis-Script-ChatGPT.gs
 * FULL 12 COLUMN ENFORCEMENT
 * ============================================================
 */

/* ============================================================
   UI LOADER
============================================================ */

function openActiveRowTDFExportSidebarGPT() {
  const html = HtmlService
    .createHtmlOutputFromFile('ActiveRowTDFSidebar-ChatGPT')
    .setTitle('Active Row TDF Export — GPT Structure Locked')
    .setWidth(700);

  SpreadsheetApp.getUi().showSidebar(html);
}


/* ============================================================
   TIER DETECTION HELPERS
============================================================ */

function detectTierFromPageTypeBlockGPT(pageTypeBlock) {
  if (pageTypeBlock.includes("Tier 1")) return 1;
  if (pageTypeBlock.includes("Tier 2")) return 2;
  if (pageTypeBlock.includes("Tier 3")) return 3;
  if (pageTypeBlock.includes("Tier 4")) return 4;
  return 0;
}

function buildTierConstraintsGPT(tier) {

  if (tier === 4) {
    return `
TIER 4 — SUPPORTING SPOKE
- Preserve 95%+ original narrative.
- Structural alignment only.
- Upward internal linking only.
`;
  }

  if (tier === 3) {
    return `
TIER 3 — DIAGNOSTIC
- Problem-solver framing.
- Controlled refinement permitted.
`;
  }

  if (tier === 2) {
    return `
TIER 2 — SERVICE
- Action-led framing.
- Moderate structural refinement allowed.
`;
  }

  if (tier === 1) {
    return `
TIER 1 — HUB ROOT
- Authority positioning required.
- Structural expansion permitted.
- Internal links must flow outward.
`;
  }

  return `NO TIER DETECTED.`;
}


/* ============================================================
   ARTICLE TYPE CONTROL BUILDER
============================================================ */

function buildArticleTypeControlBlockGPT(ss) {

  const sheet = ss.getSheetByName("Article Type Control Sheet");
  if (!sheet) return "ARTICLE TYPE CONTROL SHEET NOT FOUND.";

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIndex = name => headers.indexOf(name);

  const typeCol = colIndex("Canonical Article Type Label");
  const primaryCol = colIndex("Primary Entity (Sheet 702468283)");
  const coverageCol = colIndex("Required Entity Coverage");
  const pageExpectCol = colIndex("On-Page Content Expectations");
  const schemaExpectCol = colIndex("Schema / Entity Expectations");

  let block = `
ARTICLE TYPE CONTROL — STRUCTURAL HARD LOCK

Article Type classification MUST be determined exclusively by dominant structural organisation pattern.
Commercial tone, geo modifiers, CTAs, or keyword prominence must NOT override structural classification.

Select ONE Canonical Article Type from the following definitions:
`;

  for (let i = 1; i < data.length; i++) {

    const row = data[i];
    if (!row[typeCol]) continue;

    block += `
---------------------------------------------------------------------

${i}. ${row[typeCol]}
Primary Entity Focus: ${row[primaryCol]}
Required Entity Coverage: ${row[coverageCol]}
On-Page Content Expectations: ${row[pageExpectCol]}
Schema / Entity Expectations: ${row[schemaExpectCol]}
`;
  }

  block += `
---------------------------------------------------------------------

After selecting the Article Type:

- Fully execute Required Entity Coverage.
- Fully satisfy On-Page Content Expectations.
- Maintain Primary Entity Focus throughout.
- Do NOT introduce structural elements from other Article Types.
- Do NOT compress required structural sections.
`;

  return block.trim();
}


/* ============================================================
   MAIN PROMPT BUILDER
============================================================ */

function getActiveRowTDFDataGPT() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const activeRow = sheet.getActiveRange().getRow();

  if (activeRow < 2)
    return { part1: "ERROR: Select a valid row.", part2: "", part3: "" };

  const headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(activeRow,1,1,headerRow.length).getValues()[0];

  const colIndex = name => headerRow.indexOf(name);
  const stoneType = rowValues[colIndex("Stone Type")] || "UNKNOWN";


  /* ============================================================
     TIER DETECTION BLOCK
  ============================================================ */

  const tierSheet = ss.getSheetByName('tier levels sheet');
  let pageTypeBlock = "";

  if (tierSheet) {
    const tierHeaders = tierSheet.getRange(1,1,1,tierSheet.getLastColumn()).getValues()[0];
    const stoneColIndex = tierHeaders.indexOf(stoneType);

    if (stoneColIndex !== -1) {
      const labels = tierSheet.getRange(2,1,4,1).getValues();
      const values = tierSheet.getRange(2,stoneColIndex+1,4,1).getValues();
      pageTypeBlock = labels.map((row,i)=>
        `${row[0]}: ${values[i][0]}`
      ).join("\n");
    }
  }

  const tier = detectTierFromPageTypeBlockGPT(pageTypeBlock);
  const tierRules = buildTierConstraintsGPT(tier);


  /* ============================================================
     SECTION 1 — GOVERNANCE (START)
  ============================================================ */

  const part1 = `
SECTION 1 — GOVERNANCE ENGINE (STRUCTURE LOCKED)

You are generating a STRICT tab-delimited database row.
You are NOT writing advisory content.

OUTPUT CONTRACT — NON-NEGOTIABLE

Generate EXACTLY 12 columns.
The row MUST contain EXACTLY 11 tab characters.
If the tab count is incorrect, DO NOT output.

Wrap output in triple backticks.
No commentary before or after.

// COLUMN ORDER (MANDATORY)

1. Post ID
   - Must be a unique identifier (numeric or alphanumeric).
   - No empty values allowed.
   - Should not contain any special characters or HTML.

2. New H1
   - The H1 should be a clear, descriptive headline that summarizes the content.
   - Limit to 70 characters.
   - No HTML tags other than <h1>.
   - No line breaks allowed.

3. New Meta Title (≤60 chars)
   - Limit to 60 characters (including spaces).
   - Only descriptive text, no HTML or schema code.
   - Must be concise and accurate.
   - No line breaks or empty values.

4. New Meta Description (≤155 chars)
   - Limit to 155 characters (including spaces).
   - Provide a concise summary of the content.
   - No HTML or schema code.
   - Must not be empty.
   - No line breaks.

5. Schema (complete <script type='application/ld+json'> block, single line)
   - Must be a complete <script type='application/ld+json'> block, formatted as a single line.
   - Only valid JSON-LD code, no Meta Title or Description.
   - No line breaks, spaces, or extra content outside the <script> tag.

6. Primary Search Term
   - Must be a single keyword or keyphrase relevant to the page content.
   - No empty values allowed.
   - No line breaks.   
   
7. Content Reinforcement (pipe-separated H2 directives)
   - Contains content reinforcement directives (H2 headings) relevant to the page.
   - Use pipe-separated values (|), no HTML or schema code.
   - No empty values or line breaks.

8. Technical Directives (pipe-separated HTML blocks)
   - Contains technical directives (steps involved in restoration) in a clear, structured format.
   - Must be pipe-separated, no line breaks.
   - Should not contain HTML tags unless part of the directive.

9. Strategic Reasoning (minimum 3 pipe-separated segments)
    - Contains at least 3 segments explaining the strategic reasoning behind the content.
    - Must be pipe-separated, no empty values or line breaks.

10. Topical Intent
    - Contains relevant content categories or keywords.
    - Must be pipe-separated if multiple segments are required.
    - No line breaks or empty values.

11. Matrix Role (must reference Tier)
    - Should clearly reference the page type (e.g., Case Study, Service Page).
    - Must include the Tier (Tier 1, Tier 2, etc.).
    - No empty values allowed.

12. Key Decisions Explained (pipe-separated)
    - Explains key decisions regarding the content of the page.
    - At least one pipe-separated segment.
    - No empty values or line breaks.
    
// VALIDATION RULES

- No column may be empty.
- No line breaks inside cells.
- Replace internal line breaks with pipe characters ( | ).

END OF SECTION 1.
WAIT FOR THE NEXT SECTION BEFORE PROCEEDING.
DO NOT GENERATE OUTPUT.
`.trim();


  /* ============================================================
     SECTION 2 — ESTATE CONTEXT (START)
  ============================================================ */

  const articleTypeBlock = buildArticleTypeControlBlockGPT(ss);

  const part2 = `
SECTION 2 — ESTATE CONTEXT

TIER CLASSIFICATION DATA:
${pageTypeBlock}

Tier classification governs:
- Structural scope
- Link direction
- Expansion allowance
- Reinforcement depth

------------------------------------------------------------

${articleTypeBlock}

END OF SECTION 2.
WAIT FOR THE NEXT SECTION BEFORE PROCEEDING.
DO NOT GENERATE OUTPUT.
`.trim();


  /* ============================================================
     SECTION 3 — PAGE DATA (START)
  ============================================================ */

  const headerLine = headerRow.join("\t");
  const dataLine = rowValues.join("\t");

  const part3 = `
SECTION 3 — PAGE DATA

${headerLine}
${dataLine}

You now have:
- Full structural contract
- Tier constraints
- Article Type structural definitions
- Active row dataset

END OF SECTION 3.
YOU NOW HAVE THE FULL PROMPT.
PROCEED WITH FINAL TDF OUTPUT.
`.trim();


  return {
    part1: part1,
    part2: part2,
    part3: part3
  };
}