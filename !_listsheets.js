/**
 * Opens the Navigator Sidebar.
 */
function openNavigator() {
  const html = HtmlService.createTemplateFromFile('Navigator')
    .evaluate()
    .setTitle('Sheet Navigator')
    .setWidth(150);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns an array of all sheet names in the spreadsheet.
 */
function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(sheet => sheet.getName());
}

/**
 * Switches the active sheet to the one selected in the sidebar.
 */
function setActiveSheetByName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.activate();
  }
}
