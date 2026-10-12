/**
 * FI — Featured Image sidebar: read the selected post's SEO Meta Title.
 * Read-only; does not modify any sheet cell.
 */
function ce_fiGetSeoMetaTitle() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if (!sheet) throw new Error('posts sheet not found.');
  var row = sheet.getActiveCell().getRow();
  if (row < 2) return '';
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var column = headers.findIndex(function (h) {
    return String(h || '').trim() === 'New Meta Title';
  }) + 1;
  if (!column) throw new Error('New Meta Title column not found.');
  return String(sheet.getRange(row, column).getDisplayValue() || '').trim();
}
