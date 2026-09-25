/**
 * Restores specific posts columns from site-export for the active row,
 * matching on Post ID (posts col D) against site-export's "ID" column.
 */
function restoreActiveRowFromSiteExport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName("posts");
  var exportSheet = ss.getSheetByName("site-export");

  if (!postsSheet) return "ERROR: 'posts' sheet not found.";
  if (!exportSheet) return "ERROR: 'site-export' sheet not found.";

  var activeRow = postsSheet.getActiveRange().getRow();
  if (activeRow < 2) return "ERROR: Select a valid data row in 'posts' first.";

  var postId = String(postsSheet.getRange(activeRow, 4).getValue() || "").trim(); // col D
  if (!postId) return "ERROR: Post ID (column D) is empty for this row.";

  // Find site-export's "ID" column dynamically by header
  var exportHeaders = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn())
    .getValues()[0]
    .map(function(h) { return String(h).trim(); });
  var idColIdx = exportHeaders.indexOf("ID");
  if (idColIdx === -1) return "ERROR: 'ID' column not found in site-export headers.";

  // Find the matching row in site-export
  var exportData = exportSheet.getDataRange().getValues();
  var matchRow = null;
  for (var i = 1; i < exportData.length; i++) {
    if (String(exportData[i][idColIdx]).trim() === postId) {
      matchRow = exportData[i];
      break;
    }
  }
  if (!matchRow) return "ERROR: No matching Post ID '" + postId + "' found in site-export.";

  // Column letter -> index (1-based) helper
  function colLetterToIndex(letter) {
    var col = 0;
    for (var j = 0; j < letter.length; j++) {
      col = col * 26 + (letter.toUpperCase().charCodeAt(j) - 64);
    }
    return col;
  }

  // Mapping: posts column letter -> site-export column letter
  var mapping = [
    { postsCol: "CK", exportCol: "B" },
    { postsCol: "CL", exportCol: "J" },
    { postsCol: "CM", exportCol: "K" },
    { postsCol: "CN", exportCol: "M" },
    { postsCol: "CZ", exportCol: "O" }
  ];

  var written = [];
  mapping.forEach(function(m) {
    var exportColIdx = colLetterToIndex(m.exportCol) - 1; // zero-based for matchRow array
    var value = matchRow[exportColIdx];
    var postsColIdx = colLetterToIndex(m.postsCol); // 1-based for getRange
    var cell = postsSheet.getRange(activeRow, postsColIdx);
    try {
      cell.setPlainTextValue(String(value));
    } catch (e) {
      cell.setValue(value);
    }
    written.push(m.postsCol + " ← " + m.exportCol);
  });

  return "Restored row " + activeRow + " (Post ID " + postId + "): " + written.join(", ");
}