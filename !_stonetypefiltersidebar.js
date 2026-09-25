/**
 * 1. CONFIGURATION
 * We will use the Sheet ID 835712004 for the 'posts' tab.
 */
const TARGET_SHEET_ID = 835712004;

/**
 * Launcher for the Sidebar.
 */
function openStoneTypeSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("StoneTypeFilterSidebar")
    .setTitle("Stone Filter & Governance")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Fetches unique stone types for the dropdown.
 */
function getStoneTypeOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() == TARGET_SHEET_ID);
  if (!sheet) return ["Error: Sheet ID not found"];
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => String(h).trim());
  const stoneColIdx = headers.indexOf("Stone Type");
  
  if (stoneColIdx === -1) return ["Check Header 'Stone Type'"];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ["(No Data)"];

  const vals = sheet.getRange(2, stoneColIdx + 1, lastRow - 1, 1).getDisplayValues().flat();
  const uniqueTypes = [...new Set(vals.filter(v => v.trim() !== ""))];
  
  return ["(Show all)", ...uniqueTypes].sort();
}

/**
 * Main Filter and Hub Marker Logic.
 */
function applyStoneTypeSelection(stoneType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() == TARGET_SHEET_ID);
  const ui = SpreadsheetApp.getUi();
  
  if (!sheet) {
    ui.alert("Could not find sheet with ID: " + TARGET_SHEET_ID);
    return;
  }

  // 1. DYNAMIC HEADER SCAN (Find Stone Type and Entity Role)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => String(h).trim());
  
  const stoneColIdx = headers.indexOf("Stone Type");
  const roleColIdx = headers.indexOf("Entity Role");

  if (stoneColIdx === -1 || roleColIdx === -1) {
    ui.alert("Missing Headers! Looking for 'Stone Type' and 'Entity Role'. Found: " + headers.join(" | "));
    return;
  }

  // 2. RESET VIEW
  // Remove manual filters and show all rows to start fresh
  if (sheet.getFilter()) { sheet.getFilter().remove(); }
  sheet.showRows(1, sheet.getMaxRows());
  
  // Clear Column A (Hub Markers) below header
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).clearContent();
  SpreadsheetApp.flush();

  if (!stoneType || stoneType === "(Show all)") return;

  // 3. APPLY FILTER LOGIC
  const lastRow = sheet.getLastRow();
  // We use DisplayValues to match exactly what you see on the screen
  const stoneData = sheet.getRange(1, stoneColIdx + 1, lastRow, 1).getDisplayValues().flat();
  const roleData = sheet.getRange(1, roleColIdx + 1, lastRow, 1).getDisplayValues().flat();
  
  let matchCount = 0;
  
  // Loop starts at 1 to skip the header row
  for (let i = 1; i < stoneData.length; i++) {
    const currentRowValue = stoneData[i].trim();
    const rowIndex = i + 1;

    if (currentRowValue === stoneType.trim()) {
      matchCount++;
      // Mark as 'H' if the Entity Role is 'Hub'
      if (roleData[i].toLowerCase().includes("hub")) {
        sheet.getRange(rowIndex, 1).setValue("H");
      }
    } else {
      // Hide if it's not a match
      sheet.hideRows(rowIndex);
    }
  }

  SpreadsheetApp.flush();
  ui.alert("SUCCESS\nFound " + matchCount + " rows for " + stoneType);
}

