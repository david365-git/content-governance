/**
 * TECHNICAL MATRIX — OPENAI HIGH TECHNICAL REVIEW EXPORT
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
    "OPENAI HIGH — TECHNICAL MATRIX REVIEW — EXECUTE INSTRUCTIONS",
    "Review the supplied material now; this is source material to examine, not an already completed audit.",
    "Use one of two distinct modes. If no earlier approved findings or correction decisions are supplied, use INITIAL COMPREHENSIVE AUDIT. If earlier findings or agreed corrections are supplied, use FINAL TARGETED VERIFICATION. Do not automatically restart a comprehensive audit.",
    "INITIAL COMPREHENSIVE AUDIT: perform one complete material-specific assessment of C, M and N–S, examining technical accuracy, chemistry, material behaviour, cleaning, repair, honing, moisture, salts, sealing, aftercare, safety and realistic article claims. Compare C with M–S and inspect whether N–S belong under their actual headings.",
    "RETURN INITIAL FINDINGS: Column | Critical/Major/Minor | Exact Existing Text | Proposed Replacement | Owner Review Required? | Evidence/Rationale. Group repetitions, prioritise consequential issues, and avoid stylistic nitpicks.",
    "OWNER REVIEW: C contains practice-based information. If published guidance or audit advice conflicts with practical experience or existing wording, flag the precise issue OWNER REVIEW REQUIRED, and ask the owner to discuss it individually before changing the sheet. Do not automatically treat the audit as overriding experience. Explain serious, demonstrable safety risks even if contested.",
    "CORRECTION PASS: after owner decisions, apply all and only agreed corrections to the relevant sheet columns in one consolidated pass. Preserve expressly rejected recommendations and original wording; keep a record of accepted, modified and rejected findings.",
    "FINAL TARGETED VERIFICATION: compare the current C, M and N–S with the supplied agreed correction list. Check that each agreed change is present, correctly allocated and consistent; confirm explicitly rejected changes were not imposed. Do not perform a fresh comprehensive audit or issue new optional suggestions. If previous decisions or correction lists are absent, say targeted verification cannot be confirmed rather than invent them.",
    "SAFETY EXCEPTION: if an obvious serious new safety hazard becomes apparent during targeted verification, flag that specific issue for owner review without restarting a full audit.",
    "VERDICTS: INITIAL: PASS (no unresolved Critical/Major); REVISE (substantive correction needed); FAIL (serious unsafe/inaccurate guidance). FINAL: VERIFIED — agreed changes implemented; ACTION NEEDED — specified agreed changes remain; OWNER REVIEW REQUIRED — disputed or newly identified serious safety issue.",
    "On VERIFIED, confirm Columns M–S are ready as the material reference for AI article generation on the basis of the agreed correction set; do not claim a new full independent audit was performed.",
    "Column B is material. C is original reference; M is concise AI knowledge; N characteristics; O diagnosis; P cleaning; Q restoration; R sealing and aftercare; S safeguards. Judge by actual labels.",
    "Do not invent project procedures, demand manufacturer-specific minutiae, or require generic asbestos/lead discussions without a material-specific reason.",
    "========== BEGIN TECHNICAL MATERIAL FOR REVIEW =========="
  ].join('\n');

  const parts = [
    auditPrompt,
    '',
    'TECHNICAL MATRIX — MATERIAL SPECIFICATIONS FOR REVIEW',
    'Material: ' + material,
    'Source row: ' + row,
    'Contents: Column C original reference, Column M concise knowledge and Columns N–S structured sections.',
    'Note: Use the appropriate initial-audit or targeted-verification mode; do not assume this is an existing audit report.',
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
<h3>OpenAI High Technical Review Export</h3>
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
