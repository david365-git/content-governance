/**
 * STONE FILTER LOGIC - ISOLATED FILE
 */

// Unique variable name to avoid collision with other .gs files
var STONE_FILTER_GID = 835712004;

/**
 * 1. SIDEBAR LAUNCHER
 */
function openStoneTypeSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("StoneTypeFilterSidebar")
    .setTitle("Stone Filter & Hub Marker")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 2. DROPDOWN POPULATOR
 */
function getStoneTypeOptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets().find(function(s) { return s.getSheetId() == STONE_FILTER_GID; });
  if (!sheet) return ["Error: Sheet ID not found"];
  
  // Get headers from Row 1
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });
  var stoneColIdx = headers.indexOf("Stone Type");
  
  if (stoneColIdx === -1) return ["Check Header 'Stone Type'"];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return ["(No Data)"];

  // Use DisplayValues to handle exactly what's visible in the cells
  var vals = sheet.getRange(2, stoneColIdx + 1, lastRow - 1, 1).getDisplayValues().flat();
  var uniqueTypes = [...new Set(vals.filter(function(v) { return v.trim() !== ""; }))];
  
  return ["(Show all)", ...uniqueTypes].sort();
}

/**
 * 3. MAIN FILTER & HUB MARKER LOGIC
 */
function applyStoneTypeSelection(stoneType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets().find(function(s) { return s.getSheetId() == STONE_FILTER_GID; });
  var ui = SpreadsheetApp.getUi();
  
  if (!sheet) {
    ui.alert("Could not find sheet with ID: " + STONE_FILTER_GID);
    return;
  }

  // Find Column Indices from Row 1
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });
  
  var stoneColIdx = headers.indexOf("Stone Type");
  var roleColIdx = headers.indexOf("Entity Role");

  if (stoneColIdx === -1 || roleColIdx === -1) {
    ui.alert("Missing Headers! Please ensure Row 1 contains 'Stone Type' and 'Entity Role'.");
    return;
  }

  // RESET: Unhide everything and clear old 'H' markers in Col A
  if (sheet.getFilter()) { sheet.getFilter().remove(); }
  sheet.showRows(1, sheet.getMaxRows());
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).clearContent();
  }
  SpreadsheetApp.flush();

  // If "Show all" is selected, we stop here after the reset
  if (!stoneType || stoneType === "(Show all)") return;

  // FETCH DATA
  var lastRow = sheet.getLastRow();
  // We fetch the data from the columns found above
  var stoneData = sheet.getRange(1, stoneColIdx + 1, lastRow, 1).getDisplayValues().flat();
  var roleData = sheet.getRange(1, roleColIdx + 1, lastRow, 1).getDisplayValues().flat();
  
  var matchCount = 0;
  
  // Loop starts at index 1 (Row 2) to skip headers
  for (var i = 1; i < stoneData.length; i++) {
    var currentRowStoneValue = stoneData[i].trim();
    var rowIndex = i + 1;

    if (currentRowStoneValue === stoneType.trim()) {
      matchCount++;
      // Check for Hub status in the Entity Role column
      if (roleData[i].toLowerCase().indexOf("hub") !== -1) {
        sheet.getRange(rowIndex, 1).setValue("H");
      }
    } else {
      // Row does not match selection: Hide it
      sheet.hideRows(rowIndex);
    }
  }

  SpreadsheetApp.flush();
  ui.alert("SUCCESS\nFiltered " + matchCount + " rows for " + stoneType);
}

/**
 * 4. RESET VIEW FUNCTION
 */
function showAllStoneRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets().find(function(s) { return s.getSheetId() == STONE_FILTER_GID; });
  if (sheet) {
    sheet.showRows(1, sheet.getMaxRows());
    if (sheet.getMaxRows() > 1) {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).clearContent();
    }
    SpreadsheetApp.flush();
  }
}