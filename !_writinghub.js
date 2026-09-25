function showWritingHubSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('for-writing-hub-page')
    .setTitle('Writing Hub Data Export').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getWritingHubData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const targetHeaders = ["Title", "URL", "Stone Type", "Article Type", "Primary Entity", "Entity Role", "Entity Type", "Supporting Entities (Core)", "Peripheral Entities (Link Out)", "Feeds Hub", "Rewrite Status", "Governed Page Entity Rules (Entity Rules Document)"];
  const colIndices = targetHeaders.map(h => data[0].indexOf(h));
  const outputRows = [targetHeaders.join('\t')];
  
  data.slice(1).forEach(row => {
    if (String(row[0]).toLowerCase() === 'w') {
      outputRows.push(colIndices.map(idx => (idx !== -1 ? row[idx] : "")).join('\t'));
    }
  });
  return outputRows.join('\n');
}