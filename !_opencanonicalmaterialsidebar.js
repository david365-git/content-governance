/*******************************************************
 * Abbey — Canonical Material Row Hider (Sidebar)
 * Option 2: Dropdown populated in HTML (no server call)
 *******************************************************/

const TARGET_GID = 835712004;


function openCanonicalMaterialSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('CanonicalMaterialSidebar')
    .setTitle('Filter by Canonical Material');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Sidebar calls this when user clicks Run */
function runCanonicalMaterialHide_(selectedCanonicalName) {
  const chosen = String(selectedCanonicalName || '').trim();
  if (!chosen) throw new Error('No Canonical Material Name selected.');

  const ss = SpreadsheetApp.getActive();
  const sheet = getSheetById_(ss, TARGET_GID);
  if (!sheet) throw new Error('Target sheet not found (gid=' + TARGET_GID + ').');

  // Remove native Google filter (standard filter)
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  // Clear Column A from row 2 down
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  }

  // Force UI refresh
  SpreadsheetApp.flush();

  // Dynamic header scan (row 1): find "Stone Type" and "Entity Role"
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error('Target sheet has no columns.');

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const stoneTypeCol = findHeaderIndex_(headers, 'Stone Type');   // 1-based
  const entityRoleCol = findHeaderIndex_(headers, 'Entity Role'); // 1-based (found as requested)

  if (!stoneTypeCol) throw new Error('Header "Stone Type" not found on row 1 of target sheet.');
  if (!entityRoleCol) throw new Error('Header "Entity Role" not found on row 1 of target sheet.');

  // Unhide all data rows first (predictable repeated runs)
  if (lastRow >= 2) sheet.showRows(2, lastRow - 1);

  if (lastRow < 2) {
    return { ok: true, message: 'No data rows to process.' };
  }

  // Read all data in one call
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const chosenKey = chosen.toLowerCase();

  // Hide contiguous blocks for speed
  let hideStart = null;
  let hideCount = 0;

  for (let i = 0; i < data.length; i++) {
    const rowIndex = i + 2; // actual sheet row
    const stoneVal = String(data[i][stoneTypeCol - 1] ?? '').trim();

    const matches = stoneTypeMatches_(stoneVal, chosenKey);

    if (!matches) {
      if (hideStart === null) {
        hideStart = rowIndex;
        hideCount = 1;
      } else {
        hideCount++;
      }
    } else if (hideStart !== null) {
      sheet.hideRows(hideStart, hideCount);
      hideStart = null;
      hideCount = 0;
    }
  }

  if (hideStart !== null) sheet.hideRows(hideStart, hideCount);

  SpreadsheetApp.flush();

  return { ok: true, message: 'Done. Kept rows where Stone Type matches: ' + chosen };
}

/* -------------------------
   Helpers
------------------------- */

function getSheetById_(ss, gid) {
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (sh.getSheetId() === gid) return sh;
  }
  return null;
}

function findHeaderIndex_(headersRow, headerText) {
  const target = String(headerText).trim().toLowerCase();
  for (let i = 0; i < headersRow.length; i++) {
    if (String(headersRow[i]).trim().toLowerCase() === target) return i + 1; // 1-based
  }
  return 0;
}

function stoneTypeMatches_(stoneVal, chosenKey) {
  if (!stoneVal) return false;

  // Split on commas, slashes, pipes (multi-value patterns)
  const parts = String(stoneVal)
    .split(/[,/|]/)
    .map(s => s.trim().toLowerCase())
    .filter(Boolean);

  if (parts.length === 0) return false;
  return parts.includes(chosenKey);
}
