/**
 * Automatically handles row selection.
 * Note: Since we want the Dashboard to stay open, we don't 're-launch' 
 * the sidebar here, as that would reset your current tab.
 */
function onSelectionChange(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() === "posts") {
    const activeRow = e.range.getRow();
    // This is a hook for future auto-refresh logic if needed
    if (activeRow > 1) {
      console.log("Active row changed to: " + activeRow);
    }
  }
}

/**
 * Returns an object containing the prompt and title directly to the Dashboard.
 */
function launchHumanActionGenerator() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== "posts") {
    throw new Error("Please switch to the 'posts' sheet.");
  }

  const activeRow = sheet.getActiveCell().getRow();
  if (activeRow < 2) {
    throw new Error("Select a row with data (Row 2 or below).");
  }

  return fetchAndCombineData(activeRow);
}

/**
 * Logic to pull data and return it as a structured Object.
 */
function fetchAndCombineData(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName("posts");
  const promptsSheet = ss.getSheetByName("prompts");
  
  if (!postsSheet) throw new Error("'posts' sheet not found.");

  // Fetch headers to find dynamic column indices
  const headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];
  
  const idxPostTitle = headers.indexOf("Title") + 1;
  const idxStrategic = headers.indexOf("Strategic Action Required") + 1;
  const idxStatus = headers.indexOf("Rewrite Status") + 1;
  const idxBrief = 23; // Fixed Column 23

  const data = {
    title: idxPostTitle > 0 ? postsSheet.getRange(row, idxPostTitle).getValue() : "N/A",
    strategic: idxStrategic > 0 ? postsSheet.getRange(row, idxStrategic).getValue() : "N/A",
    status: idxStatus > 0 ? postsSheet.getRange(row, idxStatus).getValue() : "N/A",
    brief: postsSheet.getRange(row, idxBrief).getValue() || "N/A"
  };

  let promptBase = "";
  if (promptsSheet) {
    promptBase = promptsSheet.getRange("D61").getValue();
  }

  // Build the big text block for the LLM
  const fullPrompt = `PROMPT:\n${promptBase || "[Empty Prompt]"}\n\n` +
          `--- ROW DATA (Row ${row}) ---\n` +
          `Title: ${data.title}\n` +
          `Strategic Action Required: ${data.strategic}\n` +
          `Rewrite Status: ${data.status}\n` +
          `Page Rewrite Brief: ${data.brief}`;

  // NEW: Return as an object so the HTML can put "title" in the preview 
  // and "fullPrompt" in the copy-paste box.
  return {
    title: data.title,
    fullPrompt: fullPrompt
  };
}