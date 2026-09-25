function buildSheetsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const OUTPUT_SHEET_NAME = "Sheets List";
  const MAX_COLUMNS = 20; // A–T
  
  const output = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!output) {
    throw new Error('Sheet "Sheets List" does not exist.');
  }
  
  // Clear everything but retain sheet ID
  output.clear();
  
  let outputRow = 1;
  
  sheets.forEach(sheet => {
    if (sheet.getName() === OUTPUT_SHEET_NAME) return;
    
    const sheetName = sheet.getName();
    const tabColor = sheet.getTabColor();
    
    // Write sheet name
    output.getRange(outputRow, 1).setValue(sheetName);
    
    // Get first row values (A–T only)
    const lastCol = Math.min(sheet.getLastColumn(), MAX_COLUMNS);
    if (lastCol > 0) {
      const headerValues = sheet
        .getRange(1, 1, 1, lastCol)
        .getValues();
      
      output
        .getRange(outputRow, 2, 1, lastCol)
        .setValues(headerValues);
    }
    
    if (tabColor) {
      const rowRange = output.getRange(outputRow, 1, 1, MAX_COLUMNS + 1);
      rowRange.setBackground(tabColor);
      
      const rgb = hexToRgb(tabColor);
      const brightness = (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
      
      if (brightness < 128) {
        rowRange.setFontColor("#ffffff");
      } else {
        rowRange.setFontColor("#000000");
      }
    }
    
    outputRow++;
  });
  
  output.autoResizeColumns(1, MAX_COLUMNS + 1);
}


// Helper: Convert hex to RGB
function hexToRgb(hex) {
  hex = hex.replace("#", "");
  return {
    r: parseInt(hex.substring(0, 2), 16),
    g: parseInt(hex.substring(2, 4), 16),
    b: parseInt(hex.substring(4, 6), 16)
  };
}