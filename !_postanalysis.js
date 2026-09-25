/**
 * Opens the Post Analysis sidebar.
 */
function showPostAnalysisExportSidebar() {
  const html = HtmlService.createTemplateFromFile('post-analysis-prompt')
    .evaluate()
    .setTitle('Post Analysis Export')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * REFACTORED ENGINE - Minimal Service Calls
 * Loads everything into memory once to eliminate the "Loading..." hang.
 */
function getPostAnalysisTSV() {
  const ss = SpreadsheetApp.getActive();
  
  // 1. BULK DATA LOAD (Do this once at the start)
  const sheets = {
    posts: ss.getSheetByName('posts'),
    export: ss.getSheetByName('site-export'),
    tech: ss.getSheetByName('technical'),
    pers: ss.getSheetByName('personas'),
    prompts: ss.getSheetByName('prompts')
  };

  if (!sheets.posts || !sheets.export) return "Error: Essential sheets missing.";

  // Get all data arrays upfront
  const postsValues = sheets.posts.getDataRange().getValues();
  const exportValues = sheets.export.getRange(1, 1, 275, 15).getValues(); // Hard limit for speed
  const promptValues = sheets.prompts.getRange("C1:C30").getValues();
  const techValues = sheets.tech.getDataRange().getValues();
  const persValues = sheets.pers.getDataRange().getValues();

  // 2. EXTRACT PROMPTS (From memory, not the sheet)
  const mainPrompt = promptValues[1][0];      // C2
  const secondaryPrompt = promptValues[22][0]; // C23
  const authorBoxPrompt = promptValues[29][0]; // C30

  // 3. SCAN POSTS FOR MARKERS
  const postsHeader = postsValues[0].map(h => String(h).trim().toLowerCase());
  const urlIdx = postsHeader.indexOf('url') !== -1 ? postsHeader.indexOf('url') : postsHeader.indexOf('canonical url');
  const titleIdx = postsHeader.indexOf('title');

  let targetUrl = "";
  let hubInfo = "No HUB ('h') found.";
  let foundW = false;
  let foundH = false;

  for (let i = 1; i < postsValues.length; i++) {
    if (foundW && foundH) break;
    
    // Safety check: only run the expensive "Hidden" check if a marker exists
    const marker = String(postsValues[i][0]).toLowerCase().trim();
    if (marker === 'w' || marker === 'h') {
      if (sheets.posts.isRowHiddenByFilter(i + 1)) continue;

      if (marker === 'w') {
        targetUrl = pa_fastNorm_(postsValues[i][urlIdx]);
        foundW = true;
      }
      if (marker === 'h') {
        hubInfo = `HUB TITLE: ${postsValues[i][titleIdx]}\nHUB URL: ${postsValues[i][urlIdx]}`;
        foundH = true;
      }
    }
  }

  if (!targetUrl) return "Error: 'w' marker not found in visible rows.";

  // 4. FIND MATCH IN SITE-EXPORT (In-memory search)
  const exportHeader = exportValues[0].map(h => String(h).trim().toLowerCase());
  const canonIdx = exportHeader.indexOf('canonical url');
  const stoneIdx = exportHeader.indexOf('stone type');
  
  const matchRow = exportValues.find(r => pa_fastNorm_(r[canonIdx]) === targetUrl);
  if (!matchRow) return "Error: URL not found in site-export.";

  const stoneType = String(matchRow[stoneIdx] || "").trim().toLowerCase();

  // 5. TECHNICAL & PERSONA LOOKUPS (In-memory search)
  const stoneTechnical = techValues.find(r => String(r[1]).toLowerCase().trim() === stoneType)?.[2] || "N/A";
  const stonePersona = persValues.find(r => String(r[1]).toLowerCase().trim() === stoneType)?.[2] || "N/A";

  // 6. ASSEMBLE OUTPUT
  return [
    "--- PROMPT INSTRUCTIONS (C2) ---", mainPrompt, "",
    "--- SITE EXPORT DATA ---", exportHeader.join('\t'), matchRow.map(v => pa_sanitiseTSV_(v)).join('\t'), "",
    "--- TECHNICAL DATA ---", stoneTechnical, "",
    "--- STONE PERSONA ---", stonePersona, "",
    "--- ADDITIONAL INSTRUCTIONS (C23) ---", secondaryPrompt, "",
    "--- EXAMPLE AUTHOR BOX (C30) ---", authorBoxPrompt, "",
    "--- PARENT HUB REFERENCE ('h') ---", hubInfo
  ].join("\n");
}

/** Helper Functions */
function pa_fastNorm_(url) {
  return String(url || "").trim().toLowerCase().replace(/\/$/, "");
}

function pa_sanitiseTSV_(v) {
  let str = (v instanceof Date) ? v.toLocaleDateString('en-GB') : String(v ?? '');
  return str.replace(/\t/g, ' ').replace(/\r\n|\r|\n/g, ' ').trim();
}