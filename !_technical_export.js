/**
 * TECHNICAL MATRIX — THREE TEXT-FILE EXPORTS
 * Read-only: active material row on the "technical" tab.
 * Exports C (original), L (master) and M (AI knowledge).
 */
function technicalExportOpen() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== 'technical') {
    ui.alert('Open the technical tab and select the material row first.');
    return;
  }

  const range = sheet.getActiveRange();
  if (!range || range.getRow() < 2) {
    ui.alert('Select a material row below row 1.');
    return;
  }

  const row = range.getRow();
  const material = String(sheet.getRange(row, 2).getDisplayValue() || '').trim();
  if (!material) {
    ui.alert('Column B has no material name on the selected row.');
    return;
  }

  const safeName = material.replace(/[^a-zA-Z0-9_-]/g, '_');
  const columns = [
    {column: 3, suffix: 'C_Original', label: 'Download C — Original'},
    {column: 12, suffix: 'L_Master', label: 'Download L — Master'},
    {column: 13, suffix: 'M_AI', label: 'Download M — AI knowledge'}
  ];

  const files = columns.map(function(item) {
    return {
      name: safeName + '_' + item.suffix + '.txt',
      label: item.label,
      content: String(sheet.getRange(row, item.column).getValue() || '')
    };
  });

  const payload = JSON.stringify({material: material, files: files})
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  const html = `<!doctype html>
<html>
<head>
<base target="_top">
<style>
body{font-family:Arial,sans-serif;padding:18px;color:#202124}
button{display:block;width:100%;margin:10px 0;padding:12px;border:0;border-radius:5px;background:#176b45;color:white;cursor:pointer;text-align:left}
button:disabled{background:#bbb;cursor:not-allowed}
small{display:block;color:#666;margin-top:4px}
</style>
</head>
<body>
<h3>Technical data export</h3>
<p id="material"></p>
<p>Download these three text files and upload them to Gemini.</p>
<div id="downloads"></div>
<script>
const data = ${payload};
document.getElementById('material').textContent = 'Material: ' + data.material;
data.files.forEach(function(file) {
  const button = document.createElement('button');
  button.textContent = file.label;
  button.disabled = !file.content;
  if (!file.content) button.textContent += ' (empty)';
  button.addEventListener('click', function() {
    const blob = new Blob([file.content], {type:'text/plain;charset=utf-8'});
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = file.name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function() { URL.revokeObjectURL(url); }, 2000);
  });
  document.getElementById('downloads').appendChild(button);
});
</script>
</body>
</html>`;

  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(440).setHeight(370),
    'Technical Export — ' + material
  );
}
