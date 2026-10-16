/**
 * TECHNICAL MATRIX — SINGLE GEMINI AUDIT EXPORT
 * Read-only. On the active row of "technical", combines original C
 * concise technical knowledge M and six section columns N:S into one labelled UTF-8 text file.
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
  const concise = String(sheet.getRange(row, 13).getValue() || '');
  const headings = sheet.getRange(1, 14, 1, 6).getDisplayValues()[0];
  const sections = sheet.getRange(row, 14, 1, 6).getValues()[0];

  const auditPrompt = [
    'GEMINI — INDEPENDENT TECHNICAL MATRIX AUDIT',
    '',
    'Act as an independent senior UK stone, tile and hard-floor cleaning/restoration specialist and technical auditor for the material named below.',
    'AUDIT ONLY. Do not rewrite the supplied specifications.',
    '',
    'The export contains Column B (material name), Column C (original technical reference), Column M (AI Technical Knowledge — Concise) and six structured technical sections in Columns N–S.',
    'Do not assume any column offset. Judge each column by its actual displayed label.',
    '',
    'Check:',
    '1. Each of N–S matches its actual heading and preserves the important material-specific information from C.',
    '2. M is a concise, accurate representation of the same technical knowledge.',
    '3. C, M and N–S contain no material contradictions or unsupported technical claims.',
    '4. Material behaviour, diagnosis, chemistry, preparation, repairs, refinishing, sealing, moisture and aftercare advice are technically accurate and appropriately qualified.',
    '5. Genuine safety restrictions and practical safeguards are present, including appropriate dust control during cutting or mortar removal.',
    '6. State whether previously reported Critical or Major issues remain, where identifiable from the supplied text.',
    '',
    'Audit rules:',
    '- Column C is the primary reference but may itself contain incorrect advice; flag any actual technical errors in it.',
    '- Do not require generic asbestos or lead discussions as routine tile-floor issues.',
    '- Do not demand arbitrary manufacturer-specific details or fail for harmless repetition, style or wording differences.',
    '- Identify only well-grounded technical issues that would matter to the safety, accuracy or quality of AI-generated articles.',
    '- Do not invent facts, projects, processes or material characteristics.',
    '',
    'For each issue, return: Column | Severity (Critical/Major/Minor) | Exact Existing Text | Recommended Replacement.',
    'For every Critical or Major finding, provide exact replacement text and its correct column.',
    'Conclude with precisely one verdict:',
    'PASS — technically suitable and correctly allocated, with no remaining Critical or Major issues.',
    'REVISE — significant corrections are required.',
    'FAIL — serious inaccurate or unsafe technical guidance remains.',
    'On PASS, explicitly confirm Columns M–S are fit for use as this material reference for the AI article-generation workflow.',
    '',
    '========== BEGIN REFERENCE MATERIAL TO AUDIT =========='
  ].join('\\n');

  const parts = [
    auditPrompt,
    '',
    'TECHNICAL MATRIX — INDEPENDENT AUDIT PACK',
    'Material: ' + material,
    'Source row: ' + row,
    'Contents: Column C original reference, Column M concise knowledge and Columns N–S structured sections.',
    'Note: Original reference may contain unverified claims. Audit all content independently.',
    '',
    '================================================================',
    'COLUMN C — ORIGINAL TECHNICAL SPECIFICATION',
    '================================================================',
    original || '[EMPTY]',
    '',
    '================================================================',
    'COLUMN M — AI TECHNICAL KNOWLEDGE — CONCISE',
    '================================================================',
    concise || '[EMPTY]'
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
    name: material.replace(/[^a-zA-Z0-9_-]/g, '_') + '_Technical_Audit_C_M-S.txt',
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
<p>One text file: audit prompt, original specification (C), concise knowledge (M) and all six sections (N–S).</p>
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
