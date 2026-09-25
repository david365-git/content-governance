/**
 * Opens the "Inform-rewrites" sidebar.
 * This function is called by the 'Open Rewrite Export' menu item.
 */
function showInformRewritesSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Inform-rewrites')
    .setTitle('Export Row Data')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * High-speed data extraction for the Inform-rewrites sidebar.
 * It looks for rows marked with 'w' in Column A and extracts specific SEO governance columns.
 */
function getTabDelimitedData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return "No data found.";

  const headers = data[0];
  
  // Specific governance and rewrite columns requested for the export
  const targetHeaders = [
    "Title", "URL", "Article Type", "Primary Entity", "Entity Role", "Entity Type",
    "Supporting Entities (Core)", "Peripheral Entities (Link Out)", "Feeds Hub",
    "Governed Page Entity Rules (Entity Rules Document)", "Entity Governance Status",
    "Entity Governance Date", "Publish Justification (Demand vs Boundary) → Why it is prioritised now",
    "Rewrite Status", "Rewrite Role Lock", "Rewrite Governance Summary",
    "Cannibalisation Guardrail", "Rewrite Danger Signals", "Safe Handoff Pages",
    "GAS Handoff Document → Operational handoff / execution artefact",
    "Observed Query Cluster (GSC)", "GSC Intent Evidence (Context Only)",
    "Primary Query Cluster Owned → What this page is allowed to own"
  ];

  // Map header names to column indices for speed
  const colIndices = targetHeaders.map(h => headers.indexOf(h));
  
  // Start the output with the header row
  const output = [targetHeaders.join('\t')];

  // Filter rows where Column A (index 0) is 'w'
  const rows = data.slice(1)
    .filter(row => String(row[0]).toLowerCase() === 'w')
    .map(row => colIndices.map(idx => (idx !== -1 ? row[idx] : "")).join('\t'));

  if (rows.length === 0) return "No rows marked with 'w' found in Column A.";

  return output.concat(rows).join('\n');
}