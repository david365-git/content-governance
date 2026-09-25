/**
 * 🧬 ABBEY STRATEGY: SELECTION FORMATTER (V5)
 * 
 * CHANGE LOG:
 * Date: 2026-03-02
 * Reason: Added support for pipe "|" delimiter to convert into line breaks.
 * Silent execution retained.
 */

function formatAbbeySelection() {
  const range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  const values = range.getValues(); 
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  
  const output = [];

  for (let i = 0; i < numRows; i++) {
    let rowData = [];
    
    for (let j = 0; j < numCols; j++) {
      let cellText = String(values[i][j]);

      if (!cellText || cellText === "null" || cellText === "undefined" || cellText === "") {
        rowData.push("");
        continue;
      }

      // --- APPLY STRATEGIC FORMATTING ---
      
      let formatted = cellText

        // 0. Convert pipe delimiter to real line breaks
        // Convert pipe delimiter into clean section spacing
        .replace(/\s*\|\s*/g, "\n\n")

        // 1. Force breaks after Heading tags
        .replace(/<\/h[1-6]>/gi, "$&\n\n")
        
        // 2. Technical Directives: Match keywords even if touching <p>
        .replace(/(<p>)?(REPLACE|REMOVE|ADD|UPDATE)\s/gi, "\n\n$1$2 ")
        
        // 3. Warning Box handling
        .replace(/(<p>)?(\[Warning Box\]:|WARNING BOX:)/gi, "\n\n$1$2")

        // 4. Strategic Reasoning formatting
        .replace(/(POST-URL-MIGRATION ENTITY DRIFT DIAGNOSIS:)/gi, "$1\n")
        .replace(/(MONITOR DRIFT CORRECTION:|CANONICAL AUTHORITY RESTORATION:|CANNIBALISATION MANAGEMENT:)/gi, "\n\n$1\n")
        .replace(/(Analysis reveals)/gi, "\n\n$1")
        .replace(/(This "Entity Drift")/gi, "\n\n$1")
        .replace(/(By reclaiming)/gi, "\n\n$1");

      // Clean up excessive line breaks
      formatted = formatted.replace(/\n{4,}/g, "\n\n").trim();

      rowData.push(formatted);
    }
    output.push(rowData);
  }

  // Write back to selected range only
  range.setValues(output);
  range.setWrap(true);
  range.setVerticalAlignment("top");

  SpreadsheetApp.flush();
}