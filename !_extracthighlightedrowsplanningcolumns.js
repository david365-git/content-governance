/**
 * Extract highlighted (selected) rows and ONLY the planning columns you asked for,
 * with headers as the first row, into a new sheet.
 *
 * Columns extracted (by header text in row 1):
 * Title, Stone Type, Article Type, Primary Entity, Entity Role,
 * Observed Query Cluster (GSC), Primary Query Cluster Owned,
 * Clicks, Impressions, Position
 *
 * Usage:
 * 1) Highlight any cells across the rows you want to extract (can be any columns).
 * 2) Menu: Extract → Extract highlighted rows (planning columns)
 */


function extractHighlightedRowsPlanningColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();


  const headerRow = 1;
  const lastCol = sheet.getLastColumn();
  const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  const requiredHeaders = [
    'Title',
    'Stone Type',
    'Article Type',
    'Primary Entity',
    'Entity Role',
    'Observed Query Cluster (GSC)',
    'Primary Query Cluster Owned\n→ What this page is allowed to own',
    'Clicks',
    'Impressions',
    'Position'
  ];

  // Build header -> column index map (case-sensitive match by default)
  const headerToCol = new Map();
  headerValues.forEach((h, idx) => headerToCol.set(String(h).trim(), idx + 1));

  // Resolve required column indices (and fail fast if any are missing)
  const colIndexes = [];
  for (const h of requiredHeaders) {
    const col = headerToCol.get(h);
    if (!col) {
      SpreadsheetApp.getUi().alert(
        'Missing required header in row 1:\n\n' +
          h +
          '\n\nFix the header text (must match exactly), then run again.'
      );
      return;
    }
    colIndexes.push(col);
  }

  // Collect unique row numbers from the selection(s), skipping header row
  const selectedRows = new Set();
  rangeList.getRanges().forEach(r => {
    const start = r.getRow();
    const end = start + r.getNumRows() - 1;
    for (let row = start; row <= end; row++) {
      if (row !== headerRow) selectedRows.add(row);
    }
  });

  if (selectedRows.size === 0) {
    SpreadsheetApp.getUi().alert('Your selection only includes the header row. Please highlight some data rows.');
    return;
  }

  // Sort rows top-to-bottom
  const rows = Array.from(selectedRows).sort((a, b) => a - b);

  // Build output: first row = headers (same order as requiredHeaders)
  const output = [requiredHeaders.slice()];

  // Fetch each row’s required columns
  // (Efficient enough for typical selections; keeps logic simple and reliable)
  rows.forEach(rowNum => {
    const rowData = colIndexes.map(colNum => sheet.getRange(rowNum, colNum).getValue());
    output.push(rowData);
  });

  // Write to (or create) output sheet
  const outName = 'Extracted (Highlighted)';
  let outSheet = ss.getSheetByName(outName);
  if (!outSheet) outSheet = ss.insertSheet(outName);

  outSheet.clearContents();
  outSheet.getRange(1, 1, output.length, requiredHeaders.length).setValues(output);

  // Simple formatting
  outSheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight('bold');
  outSheet.setFrozenRows(1);
  outSheet.autoResizeColumns(1, requiredHeaders.length);

  outSheet.activate();
}
