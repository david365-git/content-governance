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
    "MODE SELECTION — do not infer a prior audit or agreed correction list from the conversation, labels such as DRAFT, or assumptions. FINAL TARGETED VERIFICATION is permitted ONLY when an explicit, itemised owner-approved correction register is included in the material supplied to this audit window. Otherwise perform the INITIAL COMPREHENSIVE AUDIT once, with a complete actionable findings register. If the user explicitly requested verification but the approved register is missing, report VERIFICATION NOT POSSIBLE and explain the missing input; do not claim to have verified past agreements.",
    "INITIAL COMPREHENSIVE AUDIT: perform one complete material-specific assessment of C, M and N–S, examining technical accuracy, chemistry, material behaviour, cleaning, repair, honing, moisture, salts, sealing, aftercare, safety and realistic article claims. Compare C with M–S and inspect whether N–S belong under their actual headings.",
    "MANDATORY COMPLETE FINDINGS REGISTER (INITIAL): for EVERY significant issue, including those requiring owner decisions, give: numbered issue ID; affected columns; severity; verbatim existing text from each affected column (or clearly distinguish equivalent wording); technically precise concern; complete ready-to-insert replacement for each affected column; reason/evidence; and status STANDARD or OWNER REVIEW REQUIRED. Never merely list subjects such as 'chemical stripping' or 'sealing'. Do not summarise, omit, or defer findings to a subsequent exchange. Group identical text across columns when appropriate. Finish the full register in this same report.",
    "STRICT INDEPENDENT-AUDITOR BOUNDARY: You are an external report-producing auditor, not the owner’s collaborator, assistant, editor or correction operator. This audit window has NO role in resolving findings with the owner, applying changes to Google Sheets, deciding acceptance, or planning next steps. Do not say we, let us, we can, shall we, next we should, I can help, or propose discussing issues one by one. Output only the completed standalone audit findings and verdict, then STOP.",
    "OUTPUT TERMINATION RULE: End immediately after the final verdict and, where relevant, a neutral statement that findings are referred to the separate owner-review conversation. No questions, invitations, offers, suggested next actions, conversational wrap-up or requests for documents. Owner decisions and amendments are handled exclusively in a different ChatGPT conversation.",
    "OWNER REVIEW: C contains practice-based information. If published guidance or audit advice conflicts with practical experience or existing wording, flag the precise issue OWNER REVIEW REQUIRED in the completed audit report, quoting the current wording, proposed replacement, evidence and precise point of conflict. Do not ask the owner questions in this audit window; the main ChatGPT review conversation will discuss each flagged issue with the owner before any spreadsheet change. Do not automatically treat the audit as overriding experience. Explain serious, demonstrable safety risks even if contested.",
    "AUDIT WINDOW — REPORT ONLY: Complete the entire audit or verification report in one response. NEVER ask clarifying questions, request owner decisions, offer interactive choices, or wait for follow-up. State uncertainties in findings. All decisions happen in the main ChatGPT conversation. The report must be self-contained and detailed enough for that conversation to review each unresolved issue without returning to the audit window. Do not turn the audit report into a short summary.",
    "CORRECTION-PASS CONTEXT ONLY: a separate owner-review conversation handles approvals and spreadsheet corrections. The auditor must NEVER claim to apply spreadsheet edits, offer to implement corrections or participate in owner review. Its only deliverable is the standalone findings report.",
    "FINAL TARGETED VERIFICATION: check EACH item of the explicitly supplied approved correction register against current C, M and N–S. For each item report its issue ID, affected columns, exact current offending text if still present, exact approved replacement, and VERIFIED or ACTION NEEDED. List each owner-review item in the supplied register with its full original wording, proposed replacement, evidence, and reason it needs a decision; never replace issue detail with topic names. Do not conduct another comprehensive audit or invent unresolved findings not recorded in the supplied list. If the register is absent, state VERIFICATION NOT POSSIBLE rather than pretending to know earlier decisions.",
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



/**
 * CROSS-MATERIAL TECHNICAL AUDIT — ACTIVE COLUMN
 * Read-only. Export all fourteen materials for the selected technical column
 * with full C/M reference texts and a complete independent-auditor prompt.
 * Supported active columns: C and M:S.
 */
function technicalExportCrossMaterialActiveColumn() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== 'technical') {
    ui.alert('Open the technical sheet and select a cell in C or M–S first.');
    return;
  }
  const active = sheet.getActiveRange();
  if (!active) {
    ui.alert('Select a technical-column cell first.');
    return;
  }
  const col = active.getColumn();
  if (col !== 3 && (col < 13 || col > 19)) {
    ui.alert('Select any cell in technical column C or M–S. The selected column will be audited across all materials.');
    return;
  }

  const letter = String.fromCharCode(64 + col);
  const sections = {
    3: 'Original technical specification',
    13: 'AI technical knowledge reference',
    14: 'Material characteristics and behaviour',
    15: 'Problems, defects and diagnosis',
    16: 'Cleaning and surface preparation',
    17: 'Restoration, repairs and interventions',
    18: 'Protection, sealing and aftercare',
    19: 'Technical restrictions and safeguards'
  };
  const title = sections[col];
  const first = 2;
  const last = 15;
  const count = last - first + 1;
  const materialNames = sheet.getRange(first, 2, count, 1).getDisplayValues();
  const selected = sheet.getRange(first, col, count, 1).getValues();
  const refsC = sheet.getRange(first, 3, count, 1).getValues();
  const refsM = sheet.getRange(first, 13, count, 1).getValues();

  const prompt = [
    'INDEPENDENT CROSS-MATERIAL TECHNICAL CONSISTENCY AUDIT — EXECUTE THESE INSTRUCTIONS',
    'Selected column: ' + letter + ' — ' + title,
    'Source: technical!' + letter + first + ':' + letter + last + ' of Entity-Led Content Governance Map.',
    '',
    'ROLE',
    'Act as an independent, technically rigorous senior UK professional in natural-stone, historic-clay, terrazzo, ceramic-tile, porcelain-tile and grout restoration. Audit the COMPLETE selected column across ALL FOURTEEN material rows, comparing equivalent technical claims between every material. This is an independent report, not a request to edit or simplify the matrix.',
    '',
    'SCOPE: NOT A SINGLE-ISSUE AUDIT',
    'Inspect all relevant claims about stone or ceramic composition, porosity and surface behaviour; chemical compatibility; cleaning, rinsing and extraction; rotary and steam methods; mechanical honing, polishing and abrasion; grout and joints; filling, cracks, lippage and structural defects; moisture, salts, drying and DPMs; impregnating, wax and topical finishes; maintenance; realistic outcomes, safeguards and operational scope. The DPM/topical-sealer comparison is an EXAMPLE ONLY, not the main issue or sole indicator. Do not focus on it disproportionately.',
    'All fourteen material rows MUST be examined: Ceramic Tile, Grout, Limestone, Marble, Porcelain Tile, Quarry Tile, Sandstone, Slate, Stone, Terracotta, Terrazzo, Tile, Travertine and Victorian Tile.',
    '',
    'SOURCE READING RULE',
    'The material entries labelled SELECTED COLUMN are the principal audit targets. Complete material-specific Column C and Column M extracts are included for context and checking whether apparent contradictions arise from condensed safeguards or alternative details. Do NOT silently substitute C or M for the selected text. Always reference the exact column and row for every quoted passage. Do not infer that absence of a statement proves a conflict; identify actual omissions only where technically material.',
    '',
    'TECHNICAL ACCURACY RULES — MANDATORY',
    '1. Do NOT make different materials uniform when their chemistry, firing, porosity, manufacture or construction justify different methods. Explicitly preserve valid material-specific exceptions.',
    '2. Two rows repeating the same inaccurate statement does not make it correct. Compare against technically supportable material behaviour and evidence; indicate uncertainty where evidence is insufficient.',
    '3. Distinguish true contradictions, acceptable material-specific differences, omitted necessary conditions, unsupported universal claims and conflicts between the selected column and its own C/M material references.',
    '4. Retain technical detail. Do not rewrite whole specifications, truncate nuance, replace specialist practice with generic safety prose, or simplify accurate distinctions just for shorter articles.',
    '5. Do not confuse open surface pits trapping slurry mechanically with stone or grout absorbing liquid. Do not confuse removable contamination with etching, damage, existing coatings or wear.',
    '6. Do not assume hidden bedding condition, substrate type, active rising damp or the existence of a DPM without evidence. Understand different floor eras and materials.',
    '7. Operational policy: topical coatings are considered only when a DPM is confirmed AND floor conditions and the system are suitable. A DPM alone never proves suitability. Suitable vapour-permeable impregnating sealers are preferred for appropriate moisture-active older floors lacking confirmed DPM; do not assert every floor must be sealed. This is one policy among many to check, not the main subject.',
    '8. Controlled professional steam treatment may be suitable for assessed historic floors; routine domestic steam mopping is a different process. Professional hot-water extraction and controlled wetting have material-specific limitations.',
    '9. Suitable light wax treatment can produce exceptional patina in sympathetic, low-traffic historic hallways. Reclaimed original historic clay tiles are preferred when obtainable and compatible; modern replacements may be porcelain and are not equivalent.',
    '10. Flexible cement grout is not actually able to accommodate substantial floor movement. Cosmetic repairs do not correct unstable substrate or bedding; substantial subfloor repairs are not part of the restoration service.',
    '11. Do NOT introduce slip-resistance assessments, advice, claims or terminology into findings or proposed cell wording. This business does not perform such assessments.',
    '12. Every treatment and proposed result depends on assessing that particular floor and agreeing realistic limitations and outcomes with the client.',
    '13. Where outside references and experienced practice differ, identify the exact disagreement as OWNER REVIEW REQUIRED; do not silently override the latter. Cite independently verified sources only if actually available; never invent citations.',
    '',
    'OUTPUT REQUIREMENTS — ONE COMPLETE STANDALONE REPORT',
    'A. Confirm that every one of the 14 selected-column entries was inspected, explicitly naming each. Report missing/empty source cells.',
    'B. Number findings AC-01, AC-02 etc. Group by substantive technical topic, not by material or mere keyword matching. Cover ALL identified significant issues; do not stop after a few examples.',
    'C. For EVERY finding give: severity (Critical/Major/Minor), classification (CONTRADICTION / JUSTIFIED MATERIAL DIFFERENCE / MISSING QUALIFICATION / UNSUPPORTED GENERALISATION / CROSS-REFERENCE CONFLICT), exact quoted existing text with material and cell references, precise technical explanation and real-world consequences, an owner-review flag if practice or policy is disputed, and confidence (High/Medium/Low).',
    'D. When a correction is justified, give exact ready-to-insert replacement wording specific to EACH affected selected-column cell. Do not invent quotations or propose rewriting an entire cell; preserve correct surrounding passages.',
    'E. If interpreting a selected-column statement requires checking another sheet column, use the included C/M material references but distinguish the additional-column evidence clearly. Where C/M is insufficient, mark REQUIRES ADDITIONAL COLUMN CHECK rather than guessing.',
    'F. Conclude with a priority findings register plus a separate list of legitimate differences that MUST remain. Provide a PASS / REVISE / FAIL verdict with reasons.',
    'G. This is REPORT ONLY. Do not claim to alter the spreadsheet. Do not ask the owner questions, offer next steps, or conduct a second audit. Output a complete report and STOP.',
    '',
    '========== START OF EXACT SHEET SOURCE =========='
  ].join('\n');

  const parts = [
    prompt,
    '',
    'SELECTED AUDIT COLUMN: ' + letter + ' — ' + title,
    'MATERIAL ROW RANGE: technical!2:15',
    'This export is read-only and uses the sheet values at the time of download.'
  ];

  for (let index = 0; index < count; index++) {
    const row = first + index;
    const material = String(materialNames[index][0] || '').trim() || '[UNNAMED MATERIAL]';
    const body = String(selected[index][0] || '').trim();
    const original = String(refsC[index][0] || '').trim();
    const aiReference = String(refsM[index][0] || '').trim();
    parts.push(
      '',
      '================================================================',
      'MATERIAL: ' + material + ' | ROW ' + row + ' | SELECTED CELL ' + letter + row,
      '================================================================',
      '### AUDIT TARGET: ' + material + ' — technical!' + letter + row,
      body || '[EMPTY CELL]',
      '',
      '### MATERIAL REFERENCE: ' + material + ' — technical!C' + row,
      col === 3 ? '[SAME AS AUDIT TARGET ABOVE]' : (original || '[EMPTY REFERENCE]'),
      '',
      '### AI REFERENCE: ' + material + ' — technical!M' + row,
      col === 13 ? '[SAME AS AUDIT TARGET ABOVE]' : (aiReference || '[EMPTY REFERENCE]')
    );
  }
  parts.push('', '========== END OF EXACT SHEET SOURCE ==========');

  const file = {
    name: 'Cross_Material_Audit_Column_' + letter + '_Rows_2-15.txt',
    content: parts.join('\n')
  };
  const payload = JSON.stringify(file)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');

  const html = '<!doctype html><html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;padding:18px;color:#202124}' +
    'button{width:100%;padding:12px;border:0;border-radius:5px;background:#176b45;color:#fff;cursor:pointer}' +
    '</style></head><body>' +
    '<h3>Cross-Material Technical Audit Export</h3>' +
    '<p>Selected: Column ' + letter + ' — ' + title + '</p>' +
    '<p>Downloads all 14 materials with exact row/cell references, full C/M material references and the independent-audit prompt.</p>' +
    '<button id="download">Download audit .txt</button>' +
    '<script>const file=' + payload + ';' +
    'document.getElementById("download").onclick=function(){' +
    'const blob=new Blob([file.content],{type:"text/plain;charset=utf-8"});' +
    'const url=URL.createObjectURL(blob);const link=document.createElement("a");' +
    'link.href=url;link.download=file.name;document.body.appendChild(link);link.click();link.remove();' +
    'setTimeout(function(){URL.revokeObjectURL(url);},2000);};' +
    '</script></body></html>';

  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(480).setHeight(310),
    'Cross-Material Audit — Column ' + letter
  );
}
