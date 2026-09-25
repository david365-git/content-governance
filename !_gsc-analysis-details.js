/**
 * EXTRACTION ENGINE: GSC Focused
 * 1. Pulls instruction from 'prompts'!C3
 * 2. Extracts only specific GSC columns from the active row
 * 3. Formats for PC clipboard
 */
function copyRowWithHeadersToClipboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const selectedRowIndex = activeRange.getRow();
  
  // 1. Get the Prompt from 'prompts'!C3
  const promptsSheet = ss.getSheetByName('prompts');
  if (!promptsSheet) {
    SpreadsheetApp.getUi().alert("Error: 'prompts' sheet not found.");
    return;
  }
  const mainPrompt = promptsSheet.getRange(3, 3).getValue();

  // 2. Safety check for row selection
  if (selectedRowIndex === 1) {
    SpreadsheetApp.getUi().alert("Please select a data row (Row 2 or below).");
    return;
  }

  // 3. Define the TARGET columns
  const targetHeaders = ["Title", "URL", "Clicks", "Impressions", "CTR", "Position", "Queries"];
  
  // 4. Map the headers to their column indices
  const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = targetHeaders.map(target => {
    return {
      name: target,
      index: sheetHeaders.findIndex(h => String(h).trim().toLowerCase() === target.toLowerCase())
    };
  });

  // 5. Extract data for the selected row
  const fullRowData = sheet.getRange(selectedRowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  
  const finalHeaders = [];
  const finalData = [];

  colMap.forEach(col => {
    if (col.index !== -1) {
      finalHeaders.push(col.name);
      finalData.push(fullRowData[col.index]);
    }
  });

  // 6. Format Output
  let finalOutput = "--- GSC ANALYSIS PROMPT ---\n" + mainPrompt + "\n\n";
  finalOutput += "--- FILTERED GSC DATA ---\n";
  finalOutput += finalHeaders.join('\t') + "\n" + finalData.join('\t');

  // 7. Open Desktop-Optimized Sidebar
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: sans-serif; padding: 15px; background-color: #fff;">' +
    '<p style="font-size: 14px; font-weight: bold; color: #1a73e8; margin-bottom: 10px;">Filtered GSC Package (Row ' + selectedRowIndex + ')</p>' +
    '<textarea id="box" readonly style="width: 100%; height: 350px; font-family: Consolas, monospace; font-size: 12px; padding: 10px; border: 1px solid #ccc; border-radius: 4px; background: #f9f9f9; resize: none;">' + finalOutput + '</textarea>' +
    '<button id="copyBtn" onclick="doCopy()" style="width: 100%; padding: 12px; margin-top: 12px; background: #1a73e8; color: white; border: none; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 14px;">Copy Filtered Data</button>' +
    '</div>' +
    '<script>' +
    '  function doCopy() {' +
    '    var box = document.getElementById("box");' +
    '    box.select();' +
    '    document.execCommand("copy");' +
    '    var btn = document.getElementById("copyBtn");' +
    '    btn.innerText = "✓ Copied!";' +
    '    btn.style.background = "#0d652d";' +
    '    setTimeout(function() { google.script.host.close(); }, 1000);' +
    '  }' +
    '</script>'
  )
  .setTitle('GSC Focused Export')
  .setWidth(450)
  .setHeight(520);

  SpreadsheetApp.getUi().showSidebar(html);
}