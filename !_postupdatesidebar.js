function openPostUpdateSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('PostUpdateSidebar_html')
    .setTitle('Post Update Block')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getPostUpdateBlockFromActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRange = sheet.getActiveRange();

  if (!activeRange) {
    throw new Error('Please select a row first.');
  }

  const row = activeRange.getRow();
  const lastColumn = sheet.getLastColumn();

  if (row < 2) {
    throw new Error('Please select a data row, not the header row.');
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const values = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

  function getValueByHeader(headerName) {
    const index = headers.findIndex(h => h.trim() === headerName);
    return index === -1 ? '' : values[index];
  }

  function cleanValue(value) {
    return String(value == null ? '' : value).trim();
  }

  function stripKeyphraseFromSchema(value) {
    return String(value == null ? '' : value)
      .replace(/\s*Yoast Keyphrase:\s*.*$/gim, '')
      .trim();
  }

  const postId = cleanValue(getValueByHeader('Post ID'));
  const newH1 = cleanValue(getValueByHeader('New H1'));
  const newHtml = cleanValue(getValueByHeader('New HTML'));
  const newMetaTitle = cleanValue(getValueByHeader('New Meta Title'));

  let newMetaDescription = cleanValue(getValueByHeader('New Meta Description'));
  let yoastKeyphrase = cleanValue(getValueByHeader('Yoast Keyphrase'));

  const embeddedKeyphraseMatch = newMetaDescription.match(/Yoast Keyphrase:\s*(.*)$/i);
  if (embeddedKeyphraseMatch) {
    if (!yoastKeyphrase) {
      yoastKeyphrase = cleanValue(embeddedKeyphraseMatch[1]);
    }
    newMetaDescription = cleanValue(
      newMetaDescription.replace(/\s*Yoast Keyphrase:\s*.*$/i, '')
    );
  }

  const newSchema = stripKeyphraseFromSchema(getValueByHeader('Schema (JSON-LD)'));

  const block =
    'Post ID: ' + postId + '\n' +
    'New H1: ' + newH1 + '\n' +
    'New HTML: ' + newHtml + '\n' +
    'New Meta Title: ' + newMetaTitle + '\n' +
    'New Meta Description: ' + newMetaDescription + '\n' +
    'Yoast Keyphrase: ' + yoastKeyphrase + '\n' +
    'New Schema: ' + newSchema;

  return block;
}