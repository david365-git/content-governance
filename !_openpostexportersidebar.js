/**
 * Opens the Post Exporter Sidebar.
 */
function openPostExporterSidebar() {
  const html = HtmlService.createTemplateFromFile('PostExporterSidebar')
    .evaluate()
    .setTitle('Posts Filter Export')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Fetches headers from row 1 of the 'posts' sheet.
 */
function getPostsHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  if (!sheet) return [];
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.map(h => h.toString().trim()).filter(h => h !== "");
}

/**
 * Fetches Templates from 'prompts' sheet.
 * Processes the 'post-columns' cell as a vertical list.
 */
function getSavedPrompts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const promptSheet = ss.getSheetByName('prompts');
  if (!promptSheet) return [];

  const data = promptSheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  
  const nameIdx = headers.indexOf('prompt name');
  const colsIdx = headers.indexOf('post-columns');

  if (nameIdx === -1 || colsIdx === -1) return [];

  const prompts = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameIdx]).trim();
    const colList = String(data[i][colsIdx]).trim();
    if (name && colList) {
      prompts.push({ name: name, columns: colList });
    }
  }
  return prompts;
}

/**
 * Processes extraction based on the rows currently SELECTED by your mouse.
 */
function processExtraction(selectedHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('posts');
  if (!sourceSheet) return "Error: 'posts' sheet not found.";

  const selection = sourceSheet.getActiveRangeList();
  if (!selection) return "Error: No rows selected.";

  const data = sourceSheet.getDataRange().getValues();
  const rawHeaders = data[0].map(h => h.toString().trim());
  const colIndices = selectedHeaders.map(h => rawHeaders.indexOf(h));
  
  const outputData = [selectedHeaders];
  const selectedRowNumbers = [];
  const ranges = selection.getRanges();
  
  ranges.forEach(range => {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    for (let r = 0; r < numRows; r++) {
      selectedRowNumbers.push(startRow + r);
    }
  });

  const uniqueRows = [...new Set(selectedRowNumbers)].sort((a, b) => a - b);

  uniqueRows.forEach(rowNum => {
    if (rowNum === 1) return; 
    const rowIndex = rowNum - 1;
    if (data[rowIndex]) {
      const rowValues = colIndices.map(idx => idx !== -1 ? data[rowIndex][idx] : "");
      outputData.push(rowValues);
    }
  });

  let exportSheet = ss.getSheetByName('Filtered Export');
  if (!exportSheet) {
    exportSheet = ss.insertSheet('Filtered Export');
  } else {
    exportSheet.clear();
  }

  if (outputData.length > 1) {
    exportSheet.getRange(1, 1, outputData.length, selectedHeaders.length).setValues(outputData);
    exportSheet.getRange(1, 1, 1, selectedHeaders.length)
               .setFontWeight("bold").setBackground("#d9ead3")
               .setBorder(true, true, true, true, true, true);
    exportSheet.autoResizeColumns(1, selectedHeaders.length);
    sourceSheet.getRange("A1").activate();
    ss.setActiveSheet(exportSheet);
    return "Success! " + (outputData.length - 1) + " rows extracted.";
  }
  return "No rows selected.";
}