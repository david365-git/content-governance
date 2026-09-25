function generateSheetIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const targetSheetName = "sheet_guid";
  let indexSheet = ss.getSheetByName(targetSheetName);

  // Create the sheet if it doesn't exist, otherwise clear it
  if (!indexSheet) {
    indexSheet = ss.insertSheet(targetSheetName);
  } else {
    indexSheet.clear();
  }

  // Set up the headers
  const data = [["Sheet Name", "GID", "Link"]];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const gid = sheet.getSheetId();
    const url = `${ss.getUrl()}#gid=${gid}`;
    
    // Create a clickable formula for Column C
    const linkFormula = `=HYPERLINK("${url}", "Open ${name}")`;
    
    data.push([name, gid, linkFormula]);
  });

  // Write all data to the sheet at once for better performance
  indexSheet.getRange(1, 1, data.length, 3).setValues(data);
  
  // Basic formatting: Bold the header row
  indexSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
  indexSheet.activate();
}