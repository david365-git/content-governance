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
    "INDEPENDENT TECHNICAL MATRIX AUDIT — EXECUTE THESE INSTRUCTIONS",
    "",
    "Act as an independent senior UK stone, tile and hard-floor cleaning/restoration specialist and technical auditor for the material named below.",
    "Perform the audit now. The material below is the technical specification to audit, NOT an existing audit report.",
    "AUDIT ONLY. Do not rewrite the complete specifications.",
    "",
    "AUDIT MODE — INITIAL OR SUPPLEMENTARY",
    "This prompt applies equally to initial audits and supplementary audits following corrections.",
    "Always perform a complete fresh independent technical review, including previously unreported issues. Do not presume earlier corrections or approvals are correct.",
    "If an earlier audit report is supplied, check whether its Critical and Major issues have been resolved, and identify any newly discovered significant issues.",
    "If no earlier audit report is supplied, continue normally without requesting one or commenting on its absence.",
    "Do not describe the supplied technical material as merely an audit pack or as a missing audit response. Your task is to carry out the audit.",
    "",
    "MATERIAL AND COLUMNS",
    "B: Material name.",
    "C: Original technical specification (primary reference, but not unquestionable authority).",
    "M: AI Technical Knowledge — Concise.",
    "N: Material Characteristics & Behaviour.",
    "O: Problems, Defects & Diagnosis.",
    "P: Cleaning & Surface Preparation.",
    "Q: Restoration, Repairs & Interventions.",
    "R: Protection, Sealing & Aftercare.",
    "S: Technical Restrictions & Safeguards.",
    "Judge columns by their actual displayed headings; never assume column offsets.",
    "",
    "AUDIT REQUIREMENTS",
    "1. Check every technical statement for material-specific accuracy.",
    "2. Check N–S against their actual headings and the important material-specific information in C.",
    "3. Verify M is a concise, accurate account of the same technical knowledge.",
    "4. Identify contradictions within and between C, M and N–S.",
    "5. Check behaviour, diagnosis, chemistry, surface preparation, mechanical treatment, repairs, sealing, moisture, salts and aftercare.",
    "6. Check practical safety measures, including appropriate dust controls for cutting, grinding and grout or mortar removal.",
    "7. Identify unsupported guarantees, unjustified generalisations and invented mandatory treatment steps.",
    "8. Check whether article-generation rules could produce inaccurate claims, invented processes or unrealistic outcomes.",
    "9. Where earlier findings are actually provided, identify which Critical and Major issues remain unresolved.",
    "",
    "AUDIT RULES",
    "- Flag technical errors in C as well as M–S; C may contain incorrect claims.",
    "- Identify technically meaningful issues, not harmless repetition, style, or wording preferences.",
    "- Do not demand arbitrary manufacturer-specific detail.",
    "- Do not require generic asbestos or lead discussions as routine tile-floor issues.",
    "- Do not invent facts, projects, treatments, processes or material properties.",
    "- Apply the same technical standard to initial and supplementary audits.",
    "- Report any outstanding Critical or Major issue, even if it was missed previously.",
    "",
    "EXPERIENCE-BASED CONFLICTS — OWNER REVIEW REQUIRED",
    "The current technical reference includes knowledge developed through practical restoration experience. Do not automatically treat newer guidance, common industry practice or published sources as overriding that experience.",
    "When a recommendation conflicts with the current reference or a documented practical method, clearly flag it as OWNER REVIEW REQUIRED. Quote the specific existing wording, explain the evidence and technical reason for the proposed change, identify the precise point of disagreement, and state whether the issue is a demonstrated safety error, a contextual difference, or uncertainty.",
    "Keep the proposed replacement as a recommendation only; do not presume owner agreement or instruct automatic application of a disputed change.",
    "After the audit is returned, each OWNER REVIEW REQUIRED finding must be discussed individually with the owner before changing the affected spreadsheet cells. Ask for their reasoning, floor conditions and practical experience, then reconcile the wording together.",
    "Do not use owner review to disregard a substantiated serious safety risk: clearly identify any immediate hazard and retain its Critical or Major severity while seeking resolution.",
    "",
    "FINDINGS FORMAT",
    "For each finding provide: Column | Severity | Exact Existing Text | Recommended Replacement | Review Status (Standard / OWNER REVIEW REQUIRED) | Reason and Evidence.",
    "Critical: unsafe advice or fundamental technical error with potentially serious consequences.",
    "Major: significant technical inaccuracy, contradiction, omission or unsupported claim that could mislead AI-generated articles.",
    "Minor: useful improvement that does not prevent technical approval.",
    "For every Critical or Major issue provide exact replacement text suitable for direct insertion into Google Sheets, naming every affected column.",
    "Group identical issues across columns instead of needlessly repeating them.",
    "",
    "FINAL CONSISTENCY CHECK",
    "Assess C, M, N, O, P, Q, R and S individually for technical accuracy, correct allocation, completeness and consistency.",
    "State whether earlier significant findings are resolved only when the earlier audit is supplied.",
    "",
    "FINAL VERDICT — conclude with precisely one",
    "PASS — technically suitable and correctly allocated, with no remaining Critical or Major issues.",
    "REVISE — significant corrections are required.",
    "FAIL — serious inaccurate or unsafe technical guidance remains.",
    "On PASS, explicitly confirm Columns M–S are fit for use as this material reference for the AI article-generation workflow.",
    "",
    "========== BEGIN TECHNICAL MATERIAL FOR AUDIT =========="
  ].join('\n');

  const parts = [
    auditPrompt,
    '',
    'TECHNICAL MATRIX — MATERIAL SPECIFICATIONS FOR AUDIT',
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
