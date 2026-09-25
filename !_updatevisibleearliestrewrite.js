/**
 * Updates the 'Earliest Rewrite Date' column for VISIBLE rows only.
 * Respects active filters to avoid overwriting non-target stone types.
 */
function updateVisibleRewriteDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 1. Map Columns Dynamically using your specific header name
  const driftIdx     = headers.indexOf("Drift Status");
  const govStatusIdx = headers.indexOf("Entity Governance Status");
  const govDateIdx   = headers.indexOf("Entity Governance Date");
  const rewriteIdx   = headers.indexOf("Earliest Rewrite Date");

  if (rewriteIdx === -1) {
    SpreadsheetApp.getUi().alert("Error: Column 'Earliest Rewrite Date' not found. Please check Row 1.");
    return;
  }

  let updateCount = 0;

  // 2. Loop through data rows
  for (let i = 1; i < data.length; i++) {
    const rowNumber = i + 1;

    // 3. Only process if the row is visible (not filtered out)
    if (!sheet.isRowHiddenByFilter(rowNumber)) {
      
      const driftStatus = data[i][driftIdx];
      const govStatus   = data[i][govStatusIdx];
      const govDate     = data[i][govDateIdx];

      // Use the internal calculation logic
      const result = calculateEarliestRewrite(driftStatus, govStatus, govDate);
      
      // Update only the target cell
      sheet.getRange(rowNumber, rewriteIdx + 1).setValue(result);
      updateCount++;
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Updated " + updateCount + " visible rows.", "Governance Complete");
}

/**
 * Logic Helper: Matches your formula logic for date offsets.
 */
function calculateEarliestRewrite(driftStatus, govStatus, lastUpdated) {
  const ds = String(driftStatus || "").trim();
  const gs = String(govStatus || "").trim();
  
  if (ds === "") return "WAITING: Drift Status";
  if (gs.toLowerCase() !== "clear") return "PROHIBITED: Entity Status";
  
  let luDate = new Date(lastUpdated);
  if (isNaN(luDate.getTime())) return "WAITING: Last Updated";

  let daysToAdd = 0;
  if (ds === "Significant Drift") daysToAdd = 21;
  else if (ds === "Minor Drift") daysToAdd = 30;
  else if (ds === "Aligned") daysToAdd = 90;
  else return "WAITING: Valid Drift Enum";

  luDate.setDate(luDate.getDate() + daysToAdd);
  // Returns date in DD/MM/YYYY format
  return Utilities.formatDate(luDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
}