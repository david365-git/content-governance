/**
 * ce_Stage8E_Suggestions.gs - W8E SUGGESTIONS REVIEW
 * Loads and saves the W8E analysis text (column FZ) for annotation.
 */

function getW8ESuggestionsForActiveRow() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const text = sh.getRange('FZ' + row).getValue();

  if (!text) {
    return { success: false, message: "No W8E analysis saved yet for this row (column FZ is empty). Run W8E first." };
  }

  return { success: true, text: text };
}

function saveW8ESuggestionsForActiveRow(text) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  if (typeof text !== 'string') {
    return { success: false, message: "No text provided." };
  }

  sh.getRange('FZ' + row).setValue(text);

  if (typeof logPipelineResume === 'function') {
    logPipelineResume("W8E — Suggestions Annotated", "");
  }

  return { success: true, message: "Saved to column FZ." };
}