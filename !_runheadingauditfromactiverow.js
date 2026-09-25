/**
 * ==========================================================
 * HEADING AUDIT — SITE EXPORT LOOKUP (STRICT SINGLE ROW)
 * ==========================================================
 * - Runs from custom menu
 * - Requires exactly one row selected in "posts"
 * - Matches Post ID to site-export ID
 * - Extracts H2–H4 from Full Post HTML
 * - Writes result into Current Headings column
 * ==========================================================
 */


function runHeadingAuditFromActiveRow() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName("posts");
  const exportSheet = ss.getSheetByName("site-export");
  const ui = SpreadsheetApp.getUi();

  if (!postsSheet || !exportSheet) {
    ui.alert("Required sheets not found (posts / site-export).");
    return;
  }

  const range = postsSheet.getActiveRange();

  if (!range) {
    ui.alert("No active selection.");
    return;
  }

  // STRICT: Only allow exactly one row selected
  if (range.getNumRows() !== 1) {
    ui.alert("Please select exactly ONE row before running the Heading Audit.");
    return;
  }

  const activeRow = range.getRow();

  if (activeRow < 2) {
    ui.alert("Please select a valid data row (not header).");
    return;
  }

  const postsData = postsSheet.getDataRange().getValues();
  const postsHeaders = postsData[0];

  const postIdCol = postsHeaders.indexOf("Post ID");
  const currentHeadingsCol = postsHeaders.indexOf("Current Headings");

  if (postIdCol === -1 || currentHeadingsCol === -1) {
    ui.alert("Required columns not found in posts sheet (Post ID / Current Headings).");
    return;
  }

  const postId = postsData[activeRow - 1][postIdCol];

  if (!postId) {
    ui.alert("No Post ID found in selected row.");
    return;
  }

  const exportData = exportSheet.getDataRange().getValues();
  const exportHeaders = exportData[0];

  const exportIdCol = exportHeaders.indexOf("ID");
  const htmlCol = exportHeaders.indexOf("Full Post HTML");

  if (exportIdCol === -1 || htmlCol === -1) {
    ui.alert("Required columns not found in site-export sheet (ID / Full Post HTML).");
    return;
  }

  let html = null;

  for (let i = 1; i < exportData.length; i++) {
    if (exportData[i][exportIdCol] == postId) {
      html = exportData[i][htmlCol];
      break;
    }
  }

  if (!html || typeof html !== "string") {
    postsSheet
      .getRange(activeRow, currentHeadingsCol + 1)
      .setValue("No HTML Found");
    ui.alert("No HTML found for Post ID: " + postId);
    return;
  }

  // Extract H2–H4
  const regex = /<(h[2-4])[^>]*>(.*?)<\/h[2-4]>/gsi;
  let matches;
  let results = [];

  while ((matches = regex.exec(html)) !== null) {
    let cleanContent = matches[2]
      .replace(/<\/?[^>]+(>|$)/g, "")
      .trim();

    results.push("<" + matches[1] + "> " + cleanContent);
  }

  const finalOutput = results.length > 0
    ? results.join("\n")
    : "No Headings Found";

  postsSheet
    .getRange(activeRow, currentHeadingsCol + 1)
    .setValue(finalOutput);

  // ui.alert("Heading Audit Complete for Post ID: " + postId);
}