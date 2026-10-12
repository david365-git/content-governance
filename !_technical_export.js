/**
 * TECHNICAL MATRIX — SINGLE GEMINI AUDIT EXPORT
 * Read-only. On the active row of "technical", combines original C
 * and six section columns N:S into one labelled UTF-8 text file.
 */
function technicalExportOpen() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== 'technical') {
    ui.alert('Open the technical tab and select a material row first.');
    return;
  }

  const active = sheet.getActiveRange();
  if (!active || active.getRow() < 2) {
    ui.alert('Select a material row below row 1.');
    return;
  }

  const row = active.getRow();
  const material = String(sheet.getRange(row, 2).getDisplayValue() || '').trim();
  if (!material) {
    ui.alert('Column B has no material name on this row.');
    return;
  }

  const original = String(sheet.getRange(row, 3).getValue() || '');
  const headings = sheet.getRange(1, 14, 1, 6).getDisplayValues()[0];
  const sections = sheet.getRange(row, 14, 1, 6).getValues()[0];

  const parts = [
    'TECHNICAL MATRIX — INDEPENDENT AUDIT PACK',
    'Material: ' + material,
    'Source row: ' + row,
    'Contents: Column C original reference and Columns N–S structured sections.',
    'Note: Original reference may contain unverified claims. Audit all content independently.',
    '',
    '================================================================',
    'COLUMN C — ORIGINAL TECHNICAL SPECIFICATION',
    '================================================================',
    original || '[EMPTY]'
  ];

  for (let index = 0; index < 6; index++) {
    const column = String.fromCharCode(78 + index); // N through S
    parts.push(
      '',
      '================================================================',
      'COLUMN ' + column + ' — ' + (headings[index] || 'SECTION ' + (index + 1)),
      '================================================================',
      String(sections[index] || '').trim() || '[EMPTY — NOT POPULATED]'
    );
  }

  const file = {
    name: material.replace(/[^a-zA-Z0-9_-]/g, '_') + '_Technical_Audit_C_N-S.txt',
    content: parts.join('\n')
  };

  // Escape JSON for safe inclusion in the HTML script element.
  const payload = JSON.stringify(file)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  const html = `<!doctype html>
<html><head><base target="_top"><style>
body{font-family:Arial,sans-serif;padding:18px;color:#202124}
button{display:block;width:100%;padding:12px;border:0;border-radius:5px;background:#176b45;color:white;cursor:pointer}
</style></head>
<body>
<h3>Gemini Technical Audit Export</h3>
<p id="material"></p>
<p>One text file: original specification (C) and all six labelled sections (N–S).</p>
<button id="download">Download audit text file</button>
<script>
const file = ${payload};
document.getElementById('material').textContent = 'Material: ' + ${JSON.stringify(material)};
document.getElementById('download').addEventListener('click', function() {
  const blob = new Blob([file.content], {type:'text/plain;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = file.name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(function(){URL.revokeObjectURL(url);}, 2000);
});
</script></body></html>`;

  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(300),
    'Technical Audit Export — ' + material
  );
}
