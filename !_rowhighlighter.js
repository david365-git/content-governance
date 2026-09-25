/**
 * ULTRA-FAST, HIGH VISIBILITY & F.LUX OPTIMIZED
 * Clears other 'w's, highlights active row background, and sets text color.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const value = e.value;
  const activeRow = range.getRow();
  
  // 1. Guard Clause
  if (range.getColumn() !== 1 || value !== 'w' || activeRow === 1) return;

  // 2. Visual Settings
  const highlightBg = "#00ffff"; // Cyan (high contrast for f.lux)
  const textColor = "#85200c";   // Your requested Deep Red
  const resetColor = null;       // Transparent/Default

  // 3. Find and process all 'w' markers
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const textFinder = sheet.getRange(1, 1, lastRow, 1).createTextFinder('w');
  const occurrences = textFinder.findAll();

  occurrences.forEach(cell => {
    const rowNum = cell.getRow();
    const rowRange = sheet.getRange(rowNum, 1, 1, lastCol);
    
    if (rowNum === activeRow) {
      // Apply High-Visibility styles to the new selection
      rowRange.setBackground(highlightBg);
      rowRange.setFontColor(textColor);
      rowRange.setFontWeight("normal");
    } else {
      // Clear and Reset only visible rows
      if (!sheet.isRowHiddenByFilter(rowNum)) {
        cell.clearContent();
        rowRange.setBackground(resetColor);
        rowRange.setFontColor(resetColor); // Resets to default
        rowRange.setFontWeight("normal");
      }
    }
  });
}