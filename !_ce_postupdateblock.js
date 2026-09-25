/**
 * ce_PostUpdateBlock.gs - WIDE MODELESS DIALOG FIELD EDITOR
 * Loads/saves individual fields (H1, Meta Title, Meta Description, Yoast,
 * Schema, New HTML) for the active row, opened alongside other sidebars.
 */

function openPostUpdateBlockDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ce_Dialog_PostUpdateBlock')
    .setTitle('Post Update Block');
  SpreadsheetApp.getUi().showSidebar(html);
}

function _puColMap() {
  return {
    h1:             'CK',
    metaTitle:      'CL',
    metaDescription:'CM',
    schema:         'CN',
    html:           'CT'
  };
}

function getPostUpdateField(field) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();
  const map = _puColMap();

  if (field === 'yoast') {
    const metaVal = String(sh.getRange('CM' + row).getValue() || '');
    const match = metaVal.match(/Yoast Keyphrase:\s*(.*)$/i);
    return { success: true, value: match ? match[1].trim() : '' };
  }

  if (field === 'metaDescription') {
    let metaVal = String(sh.getRange('CM' + row).getValue() || '');
    metaVal = metaVal.replace(/\s*Yoast Keyphrase:\s*.*$/i, '').trim();
    return { success: true, value: metaVal };
  }

  const col = map[field];
  if (!col) return { success: false, message: 'Unknown field: ' + field };

  const value = sh.getRange(col + row).getValue();
  return { success: true, value: String(value == null ? '' : value) };
}

function savePostUpdateField(field, value) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  if (field === 'yoast') {
    let metaVal = String(sh.getRange('CM' + row).getValue() || '');
    metaVal = metaVal.replace(/\s*Yoast Keyphrase:\s*.*$/i, '').trim();
    sh.getRange('CM' + row).setValue(metaVal + '\nYoast Keyphrase: ' + value.trim());
    return { success: true, message: 'Yoast Keyphrase saved into Meta Description (CM).' };
  }

  if (field === 'metaDescription') {
    const existing = String(sh.getRange('CM' + row).getValue() || '');
    const match = existing.match(/(\s*Yoast Keyphrase:\s*.*)$/i);
    const keyphraseSuffix = match ? match[1] : '';
    sh.getRange('CM' + row).setValue(value.trim() + keyphraseSuffix);
    return { success: true, message: 'Meta Description saved (CM), Yoast line preserved.' };
  }

  const map = _puColMap();
  const col = map[field];
  if (!col) return { success: false, message: 'Unknown field: ' + field };

  sh.getRange(col + row).setValue(value);

  if (typeof logPipelineResume === 'function') {
    logPipelineResume('Post Update Block — ' + field + ' saved', '');
  }

  return { success: true, message: field + ' saved to column ' + col + '.' };
}