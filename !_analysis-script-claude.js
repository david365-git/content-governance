/**
 * ============================================================
 * ACTIVE ROW TDF EXPORT SIDEBAR
 * FULL GOVERNANCE + HUB OVERRIDE + KEY DECISIONS COLUMN
 * ============================================================
 */

function openActiveRowTDFExportSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('ActiveRowTDFSidebar-Claude')
    .setTitle('Active Row TDF Export')
    .setWidth(700);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getActiveRowTDFData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const activeRange = sheet.getActiveRange();
  const activeRow = activeRange.getRow();

  if (activeRow < 2) return "ERROR: Please select a valid data row (not header).";

  function normalizeString(str) {
    return str.toString().replace(/\u00A0/g,' ').replace(/\s+/g,' ').trim().toLowerCase();
  }

  function clean(v) {
    return String(v ?? "")
      .replace(/\t/g," ")
      .replace(/\r?\n/g," ")
      .trim();
  }

  // ------------------------------------------------------------
  // HEADER + STONE TYPE
  // ------------------------------------------------------------
  const headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const stoneTypeColIdx = headerRow.indexOf("Stone Type");
  if (stoneTypeColIdx === -1) return "ERROR: 'Stone Type' column not found.";

  const activeStoneType = sheet.getRange(activeRow, stoneTypeColIdx+1).getValue();
  const normalizedStone = normalizeString(activeStoneType);

  // ------------------------------------------------------------
  // HUB OVERRIDE CHECK (Column A contains "H")
  // ------------------------------------------------------------
  const hubIndicator = sheet.getRange(activeRow,1).getValue();
  const isHubOverride = normalizeString(hubIndicator).includes("h");

  // ------------------------------------------------------------
  // TECHNICAL MATRIX
  // ------------------------------------------------------------
  const techSheet = ss.getSheetByName('technical');
  let technicalMatrix = `NO TECHNICAL MATRIX FOUND for "${activeStoneType}"`;

  if (techSheet) {
    const techData = techSheet.getRange("B2:C14").getValues();
    const match = techData.find(r => normalizeString(r[0]) === normalizedStone);
    if (match && match[1]) technicalMatrix = match[1];
  }

  // ------------------------------------------------------------
  // TIER CLASSIFICATION (UNCHANGED LOGIC)
  // ------------------------------------------------------------
  const tierSheet = ss.getSheetByName('tier levels sheet');
  let pageTypeBlock = `NO PAGE TYPE DATA FOUND for "${activeStoneType}"`;

  if (tierSheet) {
    const tierHeaders = tierSheet.getRange(1,1,1,tierSheet.getLastColumn()).getValues()[0];
    const stoneColIndex = tierHeaders.findIndex(h => normalizeString(h) === normalizedStone);

    if (stoneColIndex !== -1) {
      const labels = tierSheet.getRange(2,1,4,1).getValues();
      const values = tierSheet.getRange(2,stoneColIndex+1,4,1).getValues();

      pageTypeBlock = labels.map((row,i)=>
        `${row[0] || 'Unknown Tier'}: ${(values[i] && values[i][0]) || 'N/A'}`
      ).join("\n");
    }
  }

  // ------------------------------------------------------------
  // HUB INVENTORY
  // ------------------------------------------------------------
  const allData = sheet.getDataRange().getValues();

  const colIndex = name => headerRow.indexOf(name);

  const titleIdx = colIndex("Title");
  const urlIdx = colIndex("URL");
  const dateIdx = colIndex("Published Date");
  const metaTitleIdx = colIndex("Current Meta Title");
  const metaDescIdx = colIndex("Current Meta Description");
  const headingsIdx = colIndex("Current Headings");

  let hubRows = [];

  for (let i=1; i<allData.length; i++) {

    if (i+1 === activeRow) continue;

    const rowStone = normalizeString(allData[i][stoneTypeColIdx]);

    if (rowStone === normalizedStone) {

      hubRows.push([
        clean(allData[i][titleIdx]),
        clean(allData[i][urlIdx]),
        clean(allData[i][dateIdx]),
        clean(allData[i][metaTitleIdx]),
        clean(allData[i][metaDescIdx]),
        clean(allData[i][headingsIdx])
      ].join(" | "));
    }
  }

  const hubInventoryBlock = `
HUB ARTICLE INVENTORY:
These are all hub articles eligible for internal linking. Use full URLs only. Do not use Page IDs.

Title | URL | Published Date | Current Meta Title | Current Meta Description | Current Headings
${hubRows.join("\n")}
`.trim();

  // ------------------------------------------------------------
  // GOVERNANCE PROMPT
  // ------------------------------------------------------------
  const instructionBlock = `
Act as a Senior SEO Strategist specializing in high-value stone restoration.

STRATEGIC CONTEXT FOR ${activeStoneType.toString().toUpperCase()}:
${technicalMatrix}

CLASSIFICATION RULES:
${pageTypeBlock}

${isHubOverride ? `
HUB OVERRIDE RULE:
- Column A contains "H".
- This page MUST be treated as the primary Hub regardless of Tier classification.
- Do not downgrade or reclassify.
- Internal links must flow outward to supporting pages.
- Strategic Reasoning must acknowledge this override based on Base GSC intent signal.
` : ``}

PAGE TITLE (H1) ALIGNMENT:
- Retain exactly if aligned.
- Rewrite only if misaligned.
- Output into Column 3 (New H1).
- Do not optimise unnecessarily.
- Explain decision in Strategic Reasoning.

META TITLE ALIGNMENT:
- Retain exactly if aligned.
- Rewrite only if misaligned or over 60 characters.
- Explain decision in Strategic Reasoning.

META DESCRIPTION ALIGNMENT:
- Retain exactly if aligned.
- Rewrite only if misaligned or over 155 characters.
- Explain decision in Strategic Reasoning.

INTERNAL LINK GOVERNANCE:
- Select only from HUB ARTICLE INVENTORY.
- Use full URL exactly as listed.
- Do not invent URLs.
- Do not suggest the active page.
- Prefer complementary intent or higher-tier authority pages.

Col 9: Content Reinforcement — must contain H2 heading suggestions formatted as: 
REPLACE: [existing heading text] → [new heading as ready-to-paste HTML h2 tag] or 
NEW SECTION: [new heading as ready-to-paste HTML h2 tag followed by a 1-2 sentence 
content brief in plain text] or PROMOTE: [existing H3 text] → [new heading as 
ready-to-paste HTML h2 tag]. Each entry must be justified by GSC query data or page 
role. Separate entries with pipes. Include only as many as the page structure 
genuinely requires.

TECHNICAL DIRECTIVES FORMAT REQUIREMENT:
- Technical Directives occupies Column 10.
- Each directive must be delivered as ready-to-paste HTML code, not as an instruction.
- Do NOT write directives as implementation instructions (e.g. "Add hub navigation block near top of page").
- DO write each directive as the actual HTML block to be implemented, preceded by a single line comment explaining its purpose.
- Separate each HTML block with a pipe character ( | ) between them.
- Example structure:
  <!-- Hub navigation block — place after opening paragraph --> <div style='...'>[html content]</div> | 
  <!-- Replace tumbled-specific opening — swap for this block --> <p>[html content]</p>
- If HTML contains quotes, use single quotes for inline styles to avoid breaking the TDF cell.
- Each HTML block must be self-contained and pasteable directly into WordPress or a page editor without further interpretation.
- MOBILE CONSTRAINT RULE: Do NOT use border-left on any block element. Do NOT use any styling that fixes or constrains width. Use border-top instead of border-left for any decorative line. All blocks must be full-width and flow naturally within a single-column mobile layout.

KEY DECISIONS EXPLAINED REQUIREMENT:
- This is Column 14. It must always be populated.
- Compose a concise structured explanation covering:
  H1 decision, Meta Title decision, Meta Description decision,
  Schema selection, Internal Link logic, and any Hub Override.
- Replace any internal line breaks with pipe characters ( | ).
- If this column is empty in your output, your output is invalid.

STRATEGIC REASONING FORMAT REQUIREMENT:
- Strategic Reasoning occupies Column 11.
- Do NOT use line breaks inside this column.
- Replace all internal line breaks with a pipe character ( | ) instead.
- Each pipe-separated segment should represent a distinct strategic point.
- Minimum 3 pipe-separated segments required.
- Example structure:
  GSC data confirms broad ranking intent across 200+ query variants | 
  URL migration in May 2025 is the primary cause of impression decline | 
  Hub Override active — internal links flow outward to supporting pages | 
  No cannibalisation detected — supporting page must be role-constrained
- Do NOT write Strategic Reasoning as continuous prose without pipes.
- Do NOT output Strategic Reasoning outside Column 11.
SCHEMA MANDATE:
- Commercial intent -> Service or LocalBusiness.
- Educational intent -> HowTo or Article.
SCHEMA FORMAT REQUIREMENT:
The schema value in Column 6 must be output as a complete, pasteable HTML block including the opening and closing script tags.
Format exactly as:
<script type='application/ld+json'>{json content}</script>
- Use single quotes on the type attribute to avoid breaking the TDF cell.
- The JSON content must be minimised with no internal line breaks.
- No explanatory text before or after the script block.
- The entire value must sit in a single cell without breaking the row.

HTML COLUMN REQUIREMENT:
Generate a final column titled 'New HTML' (Column 15).

This column must contain the complete revised WordPress page body content as a single minimised HTML string with no line breaks.

SCOPE RULES:
- Include all body content: opening paragraphs, all h2 and h3 headings, all body copy, all images with their original src/srcset/sizes attributes preserved, all lists, all iframes, all existing affiliate product blocks, and all Technical Directive blocks integrated in their correct positions.
- Do NOT include the h1 tag — this is in Column 3.
- Do NOT include the meta title — this is in Column 4.
- Do NOT include the meta description — this is in Column 5.
- Do NOT include any schema JSON-LD script tags — this is in Column 6.

TECHNICAL DIRECTIVE INTEGRATION:
- Column 10 may contain zero, one, or multiple Technical Directive HTML blocks.
- If Column 10 is empty, output the existing page content revised only for heading structure and copy changes from Column 9.
- If Column 10 contains directives, each directive includes a single line comment specifying its insertion position. Read that comment and insert the HTML block at the described position within the content flow.
- Do not output Technical Directives as a separate block at the end. Every directive must be merged into the content at the position described by its comment.
- If a directive comment says to replace existing content, remove the original content and substitute the directive HTML in its place.
- If a directive comment says to add new content, insert the directive HTML at the described position without removing anything.

FORMAT RULES:
- Output as a single unbroken minimised HTML string with no internal line breaks.
- Use single quotes for all inline style attributes to avoid breaking the TDF cell.
- No doctype, no html tag, no head tag, no body tag, no script tags.
- No explanatory text before or after the HTML string.
- The entire column value must be pasteable as a single cell into a spreadsheet without breaking the row.

TASK:
Generate a single-line tab-delimited (TDF) output with exactly 14 columns.
Do not output any explanatory text outside the TDF row.
Before generating output, internally count your tab characters.
A valid row must contain exactly 13 tab characters (14 columns).
If your count is fewer than 13, you have collapsed columns — do not output.
Generate a single-line tab-delimited (TDF) output with exactly 14 columns.
BEFORE OUTPUTTING: Verify all 14 columns contain content:
Col 1: Post ID — must be a number
Col 2: URL — must be a full URL
Col 3: New H1 — must be a title string
Col 4: New Meta Title — must be under 60 characters
Col 5: New Meta Description — must be under 155 characters
Col 6: Schema — must be valid JSON-LD
Col 7: Internal Link Target — must be a full URL from HUB ARTICLE INVENTORY
Col 8: Primary Search Term — must be a single query string
Col 9: Content Reinforcement — must contain suggestions separated by pipes
Col 10: Technical Directives — must contain implementation instructions separated by pipes
Col 11: Strategic Reasoning — must be a prose paragraph
Col 12: Topical Intent — must be a short label
Col 13: Matrix Role — must reference Tier classification
Col 14: Key Decisions Explained — must cover H1, Meta Title, Meta Description, Schema, Internal Link, and Hub Override decisions
Col 15: New HTML - This column must contain the complete revised WordPress page body content as a single minimised HTML string with no line breaks.
If any column is empty, populate it before outputting.
OUTPUT FORMAT — CRITICAL:
Output the TDF row inside a code block. Output must be a single unbroken line of tab-separated values. Exactly 14 columns separated by tab characters. No line breaks anywhere in the output. No line breaks inside cells — replace any internal line breaks with a pipe character ( | ) instead. No explanatory text before or after the TDF row. No headers. No wrappers. The entire output must be pasteable as a single row into a spreadsheet.
Column Order:
1. Post ID
2. URL
3. New H1
4. New Meta Title
5. New Meta Description
6. Schema (JSON-LD)
7. Internal Link Target
8. Primary Search Term
9. Content Reinforcement (h2)
10. Technical Directives (h3)
11. Strategic Reasoning
12. Topical Intent
13. Page Type
14. Key Decisions Explained
`.trim();

  // ------------------------------------------------------------
  // ACTIVE ROW EXPORT (ALL COLUMNS INCLUDING GSC)
  // ------------------------------------------------------------
  const rowValues = sheet.getRange(activeRow,1,1,headerRow.length).getValues()[0];
  const headerLine = headerRow.join("\t");
  const dataLine = rowValues.map(clean).join("\t");

  return instructionBlock + "\n\n" + headerLine + "\n" + dataLine + "\n\n" + hubInventoryBlock;
}


/**
 * ============================================================
 * CORE FIELD EXPORT — ACTIVE ROW ONLY
 * ADDITIVE FUNCTION — DOES NOT MODIFY EXISTING LOGIC
 * ============================================================
 */

function getCoreFieldExportFromActiveRow() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const activeRow = sheet.getActiveRange().getRow();

  if (activeRow < 2) return "ERROR: Please select a valid data row (not header).";

  const headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(activeRow,1,1,headerRow.length).getValues()[0];

  function clean(v) {
    return String(v ?? "")
      .replace(/\r?\n/g," ")
      .trim();
  }

  const colIndex = name => headerRow.indexOf(name);

  return `
Post ID:
${clean(rowValues[colIndex("Post ID")])}

New H1:
${clean(rowValues[colIndex("New H1")])}

New HTML:
${clean(rowValues[colIndex("New HTML")])}

New Meta Title:
${clean(rowValues[colIndex("New Meta Title")])}

New Meta Description:
${clean(rowValues[colIndex("New Meta Description")])}

New Schema:
${clean(rowValues[colIndex("Schema (JSON-LD)")])}
`.trim();
}
/**
 * ============================================================
 * UPDATE SITE-EXPORT ROW FROM ACTIVE POSTS ROW
 * ============================================================
 */

function updateSiteExportRow() {
  const status = document.getElementById('status');
  const messageBox = document.getElementById('updateMessage');

  status.innerText = "Updating site-export row...";
  messageBox.innerText = "";

  google.script.run
    .withSuccessHandler(function(response) {
      status.innerText = "Complete.";
      messageBox.style.color = "#28a745";
      messageBox.innerText = response;
    })
    .withFailureHandler(function(err) {
      status.innerText = "Failed.";
      messageBox.style.color = "#dc3545";
      messageBox.innerText = "Update failed: " + err.message;
    })
    .updateSiteExportRowFromActiveRow();
}
function updateSiteExportRowFromActiveRow() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName('posts');
  const exportSheet = ss.getSheetByName('site-export');

  if (!exportSheet) {
    return "ERROR: 'site-export' sheet not found.";
  }

  const activeRow = postsSheet.getActiveRange().getRow();
  if (activeRow < 2) {
    return "ERROR: Please select a valid row in 'posts'.";
  }

  const postHeaders = postsSheet.getRange(1,1,1,postsSheet.getLastColumn()).getValues()[0];
  const exportHeaders = exportSheet.getRange(1,1,1,exportSheet.getLastColumn()).getValues()[0];
  const postData = postsSheet.getRange(activeRow,1,1,postHeaders.length).getValues()[0];

  const colIndex = (headers, name) =>
    headers.findIndex(h => String(h).trim().toLowerCase() === name.trim().toLowerCase());

  const postIdIndex = colIndex(postHeaders, "Post ID");
  const exportIdIndex = colIndex(exportHeaders, "ID");

  if (postIdIndex === -1) return "ERROR: 'Post ID' column not found in posts.";
  if (exportIdIndex === -1) return "ERROR: 'ID' column not found in site-export.";

  const postId = Number(postData[postIdIndex]);
  const exportData = exportSheet.getDataRange().getValues();

  for (let i = 1; i < exportData.length; i++) {

    const exportPostId = Number(exportData[i][exportIdIndex]);

    if (exportPostId === postId) {

      // 🔹 Explicit Field Mapping
      const fieldMap = {
        "New H1": "Title",
        "New Meta Title": "Meta Title",
        "New Meta Description": "Meta Description",
        "Schema (JSON-LD)": "WPCode Header Schema",
        "New HTML": "Full Post HTML"
      };

      Object.keys(fieldMap).forEach(postField => {

        const exportField = fieldMap[postField];

        const postCol = colIndex(postHeaders, postField);
        const exportCol = colIndex(exportHeaders, exportField);

        if (postCol !== -1 && exportCol !== -1) {
          exportSheet.getRange(i+1, exportCol+1).setValue(postData[postCol]);
        }

      });

      // 🔹 Update Last Updated Date
      const lastUpdatedCol = colIndex(exportHeaders, "Last Updated Date");
      if (lastUpdatedCol !== -1) {
        exportSheet.getRange(i+1, lastUpdatedCol+1).setValue(new Date());
      }
      // 🔹 Update POSTS sheet tracking columns (dd/mm/yyyy)
      const postsHeaders = postHeaders;

      const siteExportUpdateCol = colIndex(postsHeaders, "site-export update Date-Time");
      const headingsUpdateCol = colIndex(postsHeaders, "Current Headings update Date-Time");

      const now = new Date();
      const formatted = Utilities.formatDate(
        now,
         Session.getScriptTimeZone(),
           "dd/MM/yyyy HH:mm"
          );

      if (siteExportUpdateCol !== -1) {
        postsSheet.getRange(activeRow, siteExportUpdateCol + 1)
          .setValue(formatted)
          .setNumberFormat("dd/mm/yyyy hh:mm");
      }

      if (headingsUpdateCol !== -1) {
        postsSheet.getRange(activeRow, headingsUpdateCol + 1)
          .setValue(formatted)
          .setNumberFormat("dd/mm/yyyy hh:mm");
      }

runHeadingAuditFromActiveRow();
      return "Row " + (i+1) + " updated successfully in site-export and Current Headings row updated.";
    }
  }

  return "No matching ID found in site-export.";
}

/**
 * ============================================================
 * UPDATE WORDPRESS UPDATE DATE (POSTS SHEET)
 * Triggered when Post Fields copy button is clicked
 * ============================================================
 */
function updateWordPressUpdateDateFromActiveRow() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName('posts');

  const activeRow = postsSheet.getActiveRange().getRow();
  if (activeRow < 2) return;

  const headers = postsSheet.getRange(1,1,1,postsSheet.getLastColumn()).getValues()[0];

  const colIndex = name =>
    headers.findIndex(h => String(h).trim().toLowerCase() === name.trim().toLowerCase());

  const wpUpdateCol = colIndex("Wordpress Update Date - Time");

  if (wpUpdateCol !== -1) {

    const now = new Date();

    postsSheet.getRange(activeRow, wpUpdateCol + 1)
      .setValue(now)
      .setNumberFormat("dd/mm/yyyy hh:mm");

  }

}