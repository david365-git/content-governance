/**
 * MASTER GOVERNANCE CONTROLLER
 * Function name synchronized with your Executions Log
 */
function openCanonicalMaterialSidebar(selectedMaterial) {
  // SAFETY CHECK: Prevent the "null: text" error
  if (!selectedMaterial) {
    throw new Error("No material was selected from the dropdown.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  
  if (!sheet) throw new Error("Sheet 'posts' not found.");

  // 1. UPDATE AP1 (The Step we are verifying)
  sheet.getRange("AO1").setValue("Active Governance:");
  sheet.getRange("AP1").setValue(selectedMaterial);
  
  // 2. Clear Column A
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  }
  
  // 3. Force update to the screen
  SpreadsheetApp.flush();

  // 4. Header Scan
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const stoneTypeIdx = headers.indexOf("Stone Type");

  if (stoneTypeIdx === -1) throw new Error("Column 'Stone Type' not found.");

  // 5. Native Filter Logic
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  
  const range = sheet.getDataRange();
  const filter = range.createFilter();
  
  // Apply criteria: This is where the 'null' error usually happens
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(selectedMaterial) 
    .build();

  filter.setColumnFilterCriteria(stoneTypeIdx + 1, criteria);

  return "Success: " + selectedMaterial + " is now active.";
}