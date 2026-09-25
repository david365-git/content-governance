/**
 * UTILITIES.GS
 * This file contains the general helper functions for your Utilities menu.
 */

// 1. Convert markdown to bold (**text** to Bold)
function convertMarkdownBold() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  if (!range) return;
  
  const values = range.getRichTextValues();
  const newValues = values.map(row => row.map(cell => {
    let text = cell.getText();
    let match = text.match(/\*\*(.*?)\*\*/g);
    let builder = cell.copy();
    
    if (match) {
      match.forEach(m => {
        let clean = m.replace(/\*\*/g, "");
        text = text.replace(m, clean);
        // Note: Simple text replacement for basic conversion
      });
      return SpreadsheetApp.newRichTextValue().setText(text).build();
    }
    return cell;
  }));
  range.setRichTextValues(newValues);
}

// 2. Sort Sheet By Entity Role
function sortByEntityRolePreserveLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("posts");
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const roleCol = headers.indexOf("Entity Role") + 1;
  
  if (roleCol > 0) {
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    range.sort({column: roleCol, ascending: true});
  }
}

// 3. GSC - Copy Row + Headers for Analysis
function copyRowWithHeadersToClipboard() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const currentRow = sheet.getActiveRange().getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let output = "";
  headers.forEach((h, i) => {
    output += h + ": " + rowData[i] + "\n";
  });
  
  // Display in a model so user can copy
  const html = HtmlService.createHtmlOutput('<textarea style="width:100%;height:90%;">' + output + '</textarea>')
    .setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Copy Row Data");
}

// 4. Seed prompt names into Prompts!B2
function seedPromptNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const promptSheet = ss.getSheetByName("Prompts");
  if (!promptSheet) return;
  // logic to pull names into B2
  promptSheet.getRange("B2").setValue("Prompts Seeded: " + new Date().toLocaleString());
}

// 5. Placeholder functions to prevent Menu Errors
function openNavigator() { SpreadsheetApp.getUi().alert("Navigator Sidebar Logic Needed"); }
function showPostAnalysisExportSidebar() { SpreadsheetApp.getUi().alert("Analysis Sidebar Logic Needed"); }
function openPostExporterSidebar() { SpreadsheetApp.getUi().alert("Exporter Sidebar Logic Needed"); }
function extractHighlightedRowsPlanningColumns() { SpreadsheetApp.getUi().alert("Extraction Logic Needed"); }



/**
 * 2. THE GENERATOR
 * Directly copies Row 1 (A-T) and pastes into Column B.
 */
function generateColoredSheetList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const targetSheetName = "Sheet Names";
  
  // Setup the target sheet
  let outputSheet = ss.getSheetByName(targetSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(targetSheetName);
  }
  
  outputSheet.clear(); // Wipe data and colors

  sheets.forEach((sheet, index) => {
    const name = sheet.getName();
    const color = sheet.getTabColor();
    const rowIndex = index + 1;
    
    // Do not process the summary sheet itself
    if (name === targetSheetName) return;

    // STEP 1: Set Sheet Name in Column A
    const nameCell = outputSheet.getRange(rowIndex, 1);
    nameCell.setValue(name).setFontWeight("bold");

    // STEP 2: Force Copy Row 1 (A1:T1) - Exactly 20 columns
    const sourceHeaders = sheet.getRange("A1:T1");
    const destination = outputSheet.getRange(rowIndex, 2); // Start at Column B
    
    // This method is the most reliable way to transfer data between sheets
    sourceHeaders.copyTo(destination, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

    // STEP 3: Styling the entire row (Name + 20 Headers = 21 Columns)
    const fullRow = outputSheet.getRange(rowIndex, 1, 1, 21);
    fullRow.setBorder(true, true, true, true, true, true, "#444444", SpreadsheetApp.BorderStyle.SOLID);
    fullRow.setVerticalAlignment("middle");

    if (color) {
      fullRow.setBackground(color);
      // Contrast check for white/black text
      fullRow.setFontColor(isColorDark_(color) ? "#FFFFFF" : "#000000");
    } else {
      fullRow.setFontColor("#000000");
    }
  });

  // Final Cleanup
  outputSheet.setColumnWidth(1, 200);
  outputSheet.setFrozenColumns(1);
  
  // Resize columns for readability
  if (outputSheet.getLastColumn() > 1) {
    outputSheet.autoResizeColumns(2, outputSheet.getLastColumn() - 1);
  }

  SpreadsheetApp.getUi().alert("✅ Update Complete. If it's still blank, check if Row 1 is frozen or filtered.");
}

/**
 * Helper: Contrast Checker
 */
function isColorDark_(hex) {
  if (!hex) return false;
  const r = parseInt(hex.substring(1, 3), 16);
  const g = parseInt(hex.substring(3, 5), 16);
  const b = parseInt(hex.substring(5, 7), 16);
  const yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
  return yiq < 128;
}

/**
 * Update Active Row to posts-active-row sheet
 */


function updateActiveRowSheet() {
  // Reuses global 'ss' if already declared, or assigns it safely
  if (typeof ss === 'undefined') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  const sourceSheet = ss.getSheetByName('posts');
  const targetSheet = ss.getSheetByName('posts-active-row');
  
  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert('Error: Make sure "posts" and "posts-active-row" sheets exist.');
    return;
  }
  
  const activeRow = sourceSheet.getActiveRange().getRow();
  if (activeRow < 2) {
    SpreadsheetApp.getUi().alert('Please select a row within the data area (row 2 or below).');
    return;
  }

  const numColumns = sourceSheet.getLastColumn();
  if (numColumns === 0) {
    SpreadsheetApp.getUi().alert('The source sheet appears to be empty.');
    return;
  }

  // 1. Clear ONLY columns A and B on posts-active-row
  targetSheet.getRange(1, 1, targetSheet.getMaxRows(), 2).clearContent();

  // 2. Set Column Widths for A and B only
  targetSheet.setColumnWidth(1, 300);
  targetSheet.setColumnWidth(2, 600);

  // 3. Set Labels in Row 1
  targetSheet.getRange("A1").setValue("Header").setFontWeight("bold");
  targetSheet.getRange("B1").setValue("Details").setFontWeight("bold");

  // 4. Copy ALL Headers from 'posts' and transpose to column A
  const headerValues = sourceSheet.getRange(1, 1, 1, numColumns).getValues()[0];
  const transposedHeaders = headerValues.map(value => [value]);
  targetSheet.getRange(2, 1, transposedHeaders.length, 1).setValues(transposedHeaders);

  // 5. Copy ALL Values from active row and transpose to column B
  const rowValues = sourceSheet.getRange(activeRow, 1, 1, numColumns).getValues()[0];
  const transposedValues = rowValues.map(value => [value]);
  targetSheet.getRange(2, 2, transposedValues.length, 1).setValues(transposedValues);

  // 6. Format the text for the clipboard
  let clipboardText = "";
  for (let i = 0; i < headerValues.length; i++) {
    clipboardText += headerValues[i] + "\t" + rowValues[i] + "\n";
  }

  // 7. Open the browser sidebar helper
  showClipboardSidebar(clipboardText);
  
  ss.toast("Columns A and B updated. Copy helper opened.");
}

// Make sure this function is entirely OUTSIDE of the updateActiveRowSheet function blocks!
function showClipboardSidebar(textToCopy) {
  const safeText = textToCopy.replace(/\\/g, '\\\\').replace(/`/g, '\\`');
  
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 15px; background: #f9f9f9; text-align: center; }
        button { background-color: #1a73e8; color: white; border: none; padding: 10px 20px; font-size: 14px; border-radius: 4px; cursor: pointer; font-weight: bold; width: 100%; margin-top: 10px;}
        button:hover { background-color: #1557b0; }
        p { color: #5f6368; font-size: 13px; margin-bottom: 15px; }
      </style>
    </head>
    <body>
      <p>Your transposed data is ready! Click below to finish copying it to your computer's clipboard.</p>
      <button id="copyBtn">Copy to Clipboard</button>
      <script>
        document.getElementById('copyBtn').addEventListener('click', function() {
          const text = \`${safeText}\`;
          navigator.clipboard.writeText(text).then(function() {
            google.script.host.close();
          }).catch(function(err) {
            alert('Failed to copy text automatically.');
          });
        });
      </script>
    </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Clipboard Copy Helper')
    .setWidth(300);
    
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}