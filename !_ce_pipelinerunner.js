/* ============================================================
   ce_PipelineRunner
   Permanent working sheet for the pipeline.
   Cleared and repopulated for each active row run.

   Sheet name: "Pipeline Runner"
   Triggered via custom menu: Pipeline → [stage actions]

   Layout:
     Rows 1-5    Header — post identity data
     Rows 6-10   W-1 Authority Brief
     Rows 11-13  W0B Location Context (Geo pages only)
     Rows 14-19  W0 Stage 0 Forensic
     Rows 20-28  W1A-W1D Data and Rules
     Rows 29-31  W1.5A Audit
     Rows 32-35  W1.5B Section Plan
     Rows 36-38  W2A Pre-Write
     Rows 39-44  W2B Execute
     Rows 45-48  W4B Image Alt and Caption
     Rows 49-53  W4 HTML Audit + Push (renumbered — W3 removed)
     Rows 54-57  W4 Final HTML Audit
     Rows 58-63  W5 Governance
     Rows 64-68  W5C Schema Builder
     Row  69     Push All to Posts Sheet
============================================================ */

/* ── Colours ── */
var PR_COLOURS = {
  header:    '#1a1a2e',
  headerFg:  '#ffffff',
  section:   '#1565c0',
  sectionFg: '#ffffff',
  geo:       '#00695c',
  geoFg:     '#ffffff',
  prompt:    '#e3f2fd',
  paste:     '#fff8e1',
  result:    '#e8f5e9',
  fail:      '#ffebee',
  pending:   '#f5f5f5',
  label:     '#263238',
  labelFg:   '#ffffff',
  separator: '#eceff1'
};

/* ── Column constants ── */
var PR_COL_LABEL   = 1; // A
var PR_COL_STATUS  = 2; // B
var PR_COL_NOTE    = 3; // C
var PR_COL_CONTENT = 2; // B — visible content column

/* ============================================================
   MENU BUILDER
   Called from onOpen() in the main script.
   Adds "Pipeline" menu to the spreadsheet toolbar.
============================================================ */
function buildPipelineMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('▶ Pipeline')
    .addItem('① Initialise Runner for Active Row', 'initPipelineRunner')
    .addSeparator()
    .addItem('W-1  Run Automated Check', 'prRunW1Check')
    .addSeparator()
    .addItem('W0B  Save Location Context', 'prSaveLocationContext')
    .addSeparator()
    .addItem('W0   Generate Forensic Prompt', 'prGenerateW0')
    .addItem('W0   Push 5 Fields to Posts', 'prPushW0Fields')
    .addSeparator()
    .addItem('W1A  Generate Prompt', 'prGenerateW1A')
    .addItem('W1B  Generate Prompt', 'prGenerateW1B')
    .addItem('W1C  Generate Prompt', 'prGenerateW1C')
    .addItem('W1D  Generate Prompt', 'prGenerateW1D')
    .addSeparator()
    .addItem('W1.5A  Generate Audit Prompt', 'prGenerateW15A')
    .addItem('W1.5B  Generate Section Plan Prompt', 'prGenerateW15B')
    .addItem('W1.5B  Save Section Plan to Posts', 'prSaveSectionPlan')
    .addSeparator()
    .addItem('W2A  Generate Pre-Write Prompt', 'prGenerateW2A')
    .addItem('W2B  Generate Execute Prompt', 'prGenerateW2B')
    .addItem('W2B  Save HTML to Posts', 'prSaveHTML')
    .addSeparator()
    .addItem('W2C  Generate Image Suggestion Prompt', 'prGenerateW2C')
    .addItem('W2C  Save Suggestions to Posts', 'prSaveW2C')
    .addSeparator()
    .addItem('W4B  Load Figure Inventory', 'prLoadFigures')
    .addItem('W4B  Generate Alt/Caption Prompt', 'prGenerateW4B')
    .addSeparator()
    .addSeparator()
    .addItem('W4   Run Audit', 'prRunAudit')
    .addItem('W4   Push HTML to Posts', 'prPushHTML')
    .addSeparator()
    .addItem('W5   Generate Governance Prompt', 'prGenerateW5')
    .addItem('W5   Push Governance Fields', 'prPushGovernance')
    .addSeparator()
    .addItem('W5C  Build Schema', 'prBuildSchema')
    .addItem('W5B  Run Schema Audit', 'prRunSchemaAudit')
    .addSeparator()
    .addItem('⬆ Push ALL Results to Posts Sheet', 'prPushAll')
    .addSeparator()
    .addItem('W7   Generate Video Prompt', 'openW7Sidebar')
    .addSeparator()
    .addItem('📋 Open Prompt Copy / Paste Sidebar', 'openPipelineRunnerSidebar')
    .addToUi();
}

/* ============================================================
   HELPER — Get or create Pipeline Runner sheet
============================================================ */
function getPRSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Pipeline Runner') || ss.insertSheet('Pipeline Runner');
}

function resetPRSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName('Pipeline Runner');
  if (existing) {
    ss.deleteSheet(existing);
  }
  var newSh = ss.insertSheet('Pipeline Runner');
  ss.moveActiveSheet(ss.getSheets().length);
  // Set column widths — col A label, col B status, col C-Z content (no merging needed)
  newSh.setColumnWidth(1, 130);  // A — label
  newSh.setColumnWidth(2, 900);  // B — content (single wide column — no merge needed)
  // Hide cols C onwards — not needed
  newSh.hideColumns(3, 24);
  return newSh;
}

/* ============================================================
   HELPER — Write a section header row
============================================================ */
function prWriteSectionHeader(sh, row, label, colour, fgColour) {
  colour  = colour  || PR_COLOURS.section;
  fgColour = fgColour || PR_COLOURS.sectionFg;
  sh.getRange(row, 1)
    .setValue(label)
    .setBackground(colour)
    .setFontColor(fgColour)
    .setFontFamily('JetBrains Mono')
    .setFontSize(9)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sh.getRange(row, 2)
    .setValue('')
    .setBackground(colour);
  sh.setRowHeight(row, 22);
}

/* ============================================================
   HELPER — Write a label + content row
============================================================ */
function prWriteRow(sh, row, label, content, bgColour, isWrap) {
  bgColour = bgColour || PR_COLOURS.pending;
  isWrap   = isWrap !== false;
  sh.getRange(row, PR_COL_LABEL)
    .setValue(label || '')
    .setBackground(PR_COLOURS.label)
    .setFontColor(PR_COLOURS.labelFg)
    .setFontFamily('JetBrains Mono')
    .setFontSize(8)
    .setFontWeight('bold')
    .setWrap(true);
  sh.getRange(row, 2)
    .setValue(content || '')
    .setBackground(bgColour)
    .setFontFamily('JetBrains Mono')
    .setFontSize(8)
    .setWrap(true)
    .setVerticalAlignment('top');
  if (isWrap) {
    sh.setRowHeight(row, 21);
  }
}

/* ============================================================
   HELPER — Write status to col B
============================================================ */
function prSetStatus(sh, row, status) {
  var colours = {
    'PENDING': ['#f5f5f5', '#666'],
    'RUNNING': ['#fff3e0', '#e65100'],
    'PASS':    ['#e8f5e9', '#1b5e20'],
    'FAIL':    ['#ffebee', '#b71c1c'],
    'DONE':    ['#e3f2fd', '#0d47a1']
  };
  var c = colours[status] || colours['PENDING'];
  sh.getRange(row, PR_COL_STATUS)
    .setValue(status)
    .setBackground(c[0])
    .setFontColor(c[1])
    .setFontFamily('JetBrains Mono')
    .setFontSize(7)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

/* ============================================================
   HELPER — Write prompt content (blue background)
============================================================ */
function prWritePrompt(sh, row, label, promptText) {
  prWriteRow(sh, row, label, promptText, PR_COLOURS.prompt, true);
  // Set taller row for prompts
  sh.setRowHeight(row, 80);
}

/* ============================================================
   HELPER — Write paste area (yellow background)
============================================================ */
function prWritePasteArea(sh, row, label, instruction) {
  prWriteRow(sh, row, label, instruction || '← PASTE CHATGPT OUTPUT HERE', PR_COLOURS.paste, true);
  sh.setRowHeight(row, 80);
}

/* ============================================================
   INIT — Initialise Pipeline Runner for active posts row
   Clears the sheet and writes the full structure.
   Called from menu: Pipeline → Initialise Runner
============================================================ */
/* ============================================================
   W-1 READINESS CHECK — inlined from v76_W1Check.gs
============================================================ */

/* ============================================================
   v76_W1Check.gs
   W-1 AUTOMATED READINESS CHECK
   Runs five checks against entity analysis columns and GSC data.
   Returns a single traffic-light result with one clear action.

   Called by initPipelineRunner() — replaces raw brief display.
   Also callable from menu: Pipeline → W-1 Run Readiness Check

   OUTPUT to Pipeline Runner rows 7-10:
     Row 7:  Overall status — GO / REVIEW / BLOCKED
     Row 8:  Check results — one line per check
     Row 9:  Required action — what to do next
     Row 10: LLM prompt (if BLOCKED or REVIEW requires one)
============================================================ */


/* ============================================================
   PROMPT BUILDERS — called when checks fail
============================================================ */

function buildEntityAnalysisPrompt(d) {
  var p = '';
  p += 'ENTITY ANALYSIS REQUIRED\n\n';
  p += 'The following page has not been through entity analysis.\n';
  p += 'Run the Master Governance Controller prompt against this page to generate the 24 governance columns.\n\n';
  p += 'Page Title: ' + (d['Title'] || 'Unknown') + '\n';
  p += 'URL: ' + (d['URL'] || 'Unknown') + '\n';
  p += 'Stone Type: ' + (d['Stone Type'] || 'Unknown') + '\n\n';
  p += 'After running the entity analysis, paste the TSV output into the posts sheet starting at the Article Type column.';
  return p;
}

function buildQueryClusterPrompt(d) {
  var p = '';
  p += 'QUERY CLUSTER IDENTIFICATION REQUIRED\n\n';
  p += 'The Primary Query Cluster Owned column is empty for this page.\n';
  p += 'Based on the page details below, identify the primary query cluster this page should own.\n\n';
  p += 'Page Title: ' + (d['Title'] || 'Unknown') + '\n';
  p += 'URL: ' + (d['URL'] || 'Unknown') + '\n';
  p += 'Article Type: ' + (d['Article Type'] || 'Unknown') + '\n';
  p += 'Stone Type: ' + (d['Stone Type'] || 'Unknown') + '\n';
  p += 'Primary Entity: ' + (d['Primary Entity'] || 'Unknown') + '\n';
  p += 'Observed Query Cluster: ' + (d['Observed Query Cluster'] || 'None') + '\n\n';
  p += 'OUTPUT: State the single primary query cluster this page should own.\n';
  p += 'Format: Primary Query Cluster Owned: [your answer]\n';
  p += 'Then paste this value into col V (Primary Query Cluster Owned) in the posts sheet.';
  return p;
}

function buildRewriteResolutionPrompt(d) {
  var p = '';
  p += 'GOVERNANCE VIOLATION — REWRITE REQUIRED\n\n';
  p += 'This page has been flagged as Needs Rewrite. The governance issue must be resolved before content generation.\n\n';
  p += 'Page Title: ' + (d['Title'] || 'Unknown') + '\n';
  p += 'Article Type: ' + (d['Article Type'] || 'Unknown') + '\n';
  p += 'Rewrite Governance Summary: ' + (d['Rewrite Governance Summary'] || 'Not specified') + '\n';
  p += 'Cannibalisation Guardrail: ' + (d['Cannibalisation Guardrail'] || 'Not specified') + '\n';
  p += 'Page Rewrite Brief: ' + (d['Page Rewrite Brief'] || 'Not specified') + '\n\n';
  p += 'TASK: Review the governance issue described above.\n';
  p += 'Confirm whether:\n';
  p += '1. The article type classification is correct\n';
  p += '2. The content scope is within the allowed boundaries\n';
  p += '3. The rewrite brief is actionable\n\n';
  p += 'OUTPUT: Confirm the rewrite brief is understood and state the specific governance change required.\n';
  p += 'Then update the Rewrite Status column in the posts sheet to "Governance Tightening Recommended" if resolved,\n';
  p += 'or update the Page Rewrite Brief with a more specific instruction.';
  return p;
}

/* ============================================================
   WRAPPERS — called by initPipelineRunner
============================================================ */

/* ============================================================
   v76_W1Check.gs
   W-1 AUTOMATED READINESS CHECK
   Runs five checks against entity analysis columns and GSC data.
   Returns a single traffic-light result with one clear action.

   Called by initPipelineRunner() — replaces raw brief display.
   Also callable from menu: Pipeline → W-1 Run Readiness Check

   OUTPUT to Pipeline Runner rows 7-10:
     Row 7:  Overall status — GO / REVIEW / BLOCKED
     Row 8:  Check results — one line per check
     Row 9:  Required action — what to do next
     Row 10: LLM prompt (if BLOCKED or REVIEW requires one)
============================================================ */

/* Local colour constants — mirrors PR_COLOURS in v76_PipelineRunner.gs */
var W1_COLOURS = {
  label:     '#263238',
  labelFg:   '#ffffff',
  result:    '#e8f5e9',
  fail:      '#ffebee',
  paste:     '#fff8e1',
  pending:   '#f5f5f5',
  prompt:    '#e3f2fd',
  separator: '#eceff1'
};

function runW1ReadinessCheck(d) {
  // d is passed in from initPipelineRunner — avoids re-reading wrong active sheet
  if (!d) d = getActiveRowDataMap();
  var sh = getPRSheet();

  var rows   = [];  // Each entry: { status, label, detail, prompt }
  var blocked = false;
  var review  = false;

  /* ── CHECK 1 — Entity analysis present ── */
  var hasEntity = String(d['Article Type'] || '').trim() !== '' &&
                  String(d['Primary Entity'] || '').trim() !== '';
  if (!hasEntity) {
    blocked = true;
    rows.push({
      status: 'BLOCKED',
      label:  'CHECK 1 — Entity Analysis Missing',
      detail: 'Article Type or Primary Entity is empty. The governance prompt has not been run for this page.',
      prompt: _w1Prompt_EntityAnalysis(d)
    });
  } else {
    rows.push({
      status: 'OK',
      label:  'CHECK 1 — Entity Analysis',
      detail: d['Article Type'] + ' | ' + d['Primary Entity'],
      prompt: ''
    });
  }

  /* ── CHECK 2 — Query cluster present ── */
  var primaryCluster  = String(d['Primary Query Cluster Owned'] || '').trim();
  var observedCluster = String(d['Observed Query Cluster'] || '').trim();
  if (!primaryCluster) {
    if (observedCluster) {
      review = true;
      rows.push({
        status: 'REVIEW',
        label:  'CHECK 2 — Primary Query Cluster Missing',
        detail: 'Observed cluster exists but Primary Query Cluster Owned is empty.',
        prompt: _w1Prompt_QueryCluster(d)
      });
    } else {
      blocked = true;
      rows.push({
        status: 'BLOCKED',
        label:  'CHECK 2 — No Query Cluster',
        detail: 'Both Primary Query Cluster Owned and Observed Query Cluster are empty.',
        prompt: _w1Prompt_QueryCluster(d)
      });
    }
  } else {
    rows.push({
      status: 'OK',
      label:  'CHECK 2 — Query Cluster',
      detail: primaryCluster,
      prompt: ''
    });
  }

  /* ── CHECK 3 — Problem Angle selected ── */
  var problemAngle = String(d['Problem Angle'] || '').trim();
  if (!problemAngle) {
    review = true;
    rows.push({
      status: 'REVIEW',
      label:  'CHECK 3 — Problem Angle Not Set',
      detail: 'No angle selected. The article has no governing opening sentence instruction.',
      prompt: _w1Prompt_ProblemAngle(d)
    });
  } else {
    rows.push({
      status: 'OK',
      label:  'CHECK 3 — Problem Angle',
      detail: problemAngle.split('—')[0].trim(),
      prompt: ''
    });
  }

  /* ── CHECK 4 — GSC signal ── */
  var bImpr  = parseFloat(String(d['Benmrk Impr']    || '0').replace(/,/g,'')) || 0;
  var mImpr  = parseFloat(String(d['Montr Impr']     || '0').replace(/,/g,'')) || 0;
  var mClicks = parseFloat(String(d['Montr Clicks']  || '0').replace(/,/g,'')) || 0;
  var mPos   = parseFloat(String(d['Montr Position'] || '0').replace(/,/g,'')) || 0;
  var delta  = mImpr - bImpr;
  var pct    = bImpr > 0 ? Math.round((delta / bImpr) * 100) : 0;

  if (bImpr === 0 && mImpr === 0) {
    rows.push({
      status: 'OK',
      label:  'CHECK 4 — GSC Signal',
      detail: 'No GSC data — new or unindexed page. Proceed.',
      prompt: ''
    });
  } else if (delta < -10 && pct < -20) {
    review = true;
    rows.push({
      status: 'REVIEW',
      label:  'CHECK 4 — Authority Loss Detected',
      detail: 'Impressions: ' + bImpr + ' → ' + mImpr + ' (' + pct + '%). Recovery focus required.',
      prompt: _w1Prompt_GSCRecovery(d, bImpr, mImpr, pct)
    });
  } else if (mImpr > 0 && mClicks === 0) {
    review = true;
    rows.push({
      status: 'REVIEW',
      label:  'CHECK 4 — CTR Issue',
      detail: mImpr + ' impressions, 0 clicks' + (mPos ? ' at position ' + mPos.toFixed(1) : '') + '. Title and meta need strengthening.',
      prompt: _w1Prompt_CTR(d, mImpr, mPos)
    });
  } else {
    rows.push({
      status: 'OK',
      label:  'CHECK 4 — GSC Signal',
      detail: mImpr + ' impressions | ' + mClicks + ' clicks' + (mPos ? ' | pos ' + mPos.toFixed(1) : ''),
      prompt: ''
    });
  }

  /* ── CHECK 5 — Rewrite status ── */
  var rewriteStatus = String(d['Rewrite Status'] || '').trim();
  var rewriteBrief  = String(d['Page Rewrite Brief'] || '').trim();
  if (rewriteStatus === 'Needs Rewrite') {
    blocked = true;
    rows.push({
      status: 'BLOCKED',
      label:  'CHECK 5 — Governance Violation',
      detail: 'Rewrite Status = Needs Rewrite. ' + (rewriteBrief || 'See Rewrite Governance Summary column.'),
      prompt: _w1Prompt_RewriteViolation(d)
    });
  } else if (rewriteStatus === 'Governance Tightening Recommended') {
    review = true;
    rows.push({
      status: 'REVIEW',
      label:  'CHECK 5 — Governance Tightening',
      detail: rewriteBrief || 'See Page Rewrite Brief column.',
      prompt: _w1Prompt_GovernanceTightening(d)
    });
  } else if (!rewriteStatus) {
    review = true;
    rows.push({
      status: 'REVIEW',
      label:  'CHECK 5 — Rewrite Status Empty',
      detail: 'Rewrite Status not set — entity analysis may not be complete.',
      prompt: _w1Prompt_EntityAnalysis(d)
    });
  } else {
    rows.push({
      status: 'OK',
      label:  'CHECK 5 — Rewrite Status',
      detail: rewriteStatus,
      prompt: ''
    });
  }

  /* ── WRITE RESULTS TO PIPELINE RUNNER ── */
  var overallBg, overallFg, overallMsg, overallAction;
  if (blocked) {
    overallBg     = '#ffebee'; overallFg = '#b71c1c';
    overallMsg    = '🔴 BLOCKED — fix the issue(s) below before proceeding';
    overallAction = 'Copy the prompt from the BLOCKED row below → paste into ChatGPT → paste the result into the posts sheet → re-run Initialise Runner';
  } else if (review) {
    overallBg     = '#fff8e1'; overallFg = '#e65100';
    overallMsg    = '🟡 REVIEW — warnings present. Copy any PROMPT rows below to ChatGPT, paste results back, then use menu ▶ Pipeline → W0 Generate Forensic Prompt';
    overallAction = 'Optional: copy prompts from REVIEW rows to ChatGPT and update the posts sheet. Then NEXT STEP: Run Pipeline → W0 Generate Forensic Prompt.';
  } else {
    overallBg     = '#e8f5e9'; overallFg = '#1b5e20';
    overallMsg    = '🟢 GO — all checks passed';
    overallAction = 'NEXT STEP: Run Pipeline → W0 Generate Forensic Prompt';
  }

  // Row 7 — overall status
  sh.getRange(7, 2)
    .setValue(overallMsg)
    .setBackground(overallBg).setFontColor(overallFg)
    .setFontFamily('JetBrains Mono').setFontSize(10).setFontWeight('bold')
    .setVerticalAlignment('middle');
  sh.setRowHeight(7, 26);

  // Row 8 — action instruction
  sh.getRange(8, 1).setValue('ACTION')
    .setBackground('#263238').setFontColor('#fff')
    .setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
  sh.getRange(8, 2)
    .setValue(overallAction)
    .setBackground(overallBg).setFontColor(overallFg)
    .setFontFamily('JetBrains Mono').setFontSize(9).setFontWeight('bold').setWrap(true);
  sh.setRowHeight(8, 26);

  // Rows 9 onwards — one row per check
  var currentRow = 9;
  rows.forEach(function(r) {
    var bg = r.status === 'OK' ? '#e8f5e9' : (r.status === 'BLOCKED' ? '#ffebee' : '#fff8e1');
    var fg = r.status === 'OK' ? '#1b5e20' : (r.status === 'BLOCKED' ? '#b71c1c' : '#e65100');

    // Check label row
    sh.getRange(currentRow, 1).setValue(r.status)
      .setBackground(bg).setFontColor(fg)
      .setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold')
      .setHorizontalAlignment('center');
    sh.getRange(currentRow, 2)
      .setValue(r.label + ': ' + r.detail)
      .setBackground(bg).setFontColor(fg)
      .setFontFamily('JetBrains Mono').setFontSize(8).setWrap(true);
    sh.setRowHeight(currentRow, 20);
    currentRow++;

    // Prompt row — only if not OK
    if (r.prompt) {
      sh.getRange(currentRow, 1).setValue('PROMPT')
        .setBackground('#c62828').setFontColor('#fff')
        .setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
      sh.getRange(currentRow, 2)
        .setValue(r.prompt)
        .setBackground('#e3f2fd')
        .setFontFamily('JetBrains Mono').setFontSize(8).setWrap(true).setVerticalAlignment('top');
      sh.setRowHeight(currentRow, 100);
      currentRow++;
    }
  });

  // Separator after checks
  sh.getRange(currentRow, 2).setBackground('#eceff1').setValue('');
  sh.setRowHeight(currentRow, 6);

  return { blocked: blocked, review: review, nextRow: currentRow + 1 };
}

/* ── PROMPT BUILDERS ── */

function _w1Prompt_EntityAnalysis(d) {
  return 'Copy this prompt into ChatGPT:\n\n' +
    'I am running entity analysis on a web page. ' +
    'Page Title: ' + (d['Title'] || 'Unknown') + '. ' +
    'URL: ' + (d['URL'] || 'Unknown') + '. ' +
    'Stone Type: ' + (d['Stone Type'] || 'Unknown') + '. ' +
    'Run the Master Governance Controller prompt against this page and return the 24-column TSV row. ' +
    'I will paste the TSV result into the posts sheet Article Type column for this row, then re-run Initialise Runner.';
}

function _w1Prompt_QueryCluster(d) {
  return 'Copy this prompt into ChatGPT:\n\n' +
    'Identify the primary query cluster for this page. ' +
    'Page Title: ' + (d['Title'] || 'Unknown') + '. ' +
    'URL: ' + (d['URL'] || 'Unknown') + '. ' +
    'Article Type: ' + (d['Article Type'] || 'Unknown') + '. ' +
    'Stone Type: ' + (d['Stone Type'] || 'Unknown') + '. ' +
    'Primary Entity: ' + (d['Primary Entity'] || 'Unknown') + '. ' +
    'Observed Query Cluster: ' + (d['Observed Query Cluster'] || 'None') + '.\n\n' +
    'Return exactly one line:\n' +
    'Primary Query Cluster Owned: [3-6 words describing the query this page should rank for]\n\n' +
    'After pasting into ChatGPT — paste the result into the Primary Query Cluster Owned column in the posts sheet, then re-run Initialise Runner.';
}

function _w1Prompt_ProblemAngle(d) {
  return 'Copy this prompt into ChatGPT:\n\n' +
    'Select the Problem Angle for this article. ' +
    'Page Title: ' + (d['Title'] || 'Unknown') + '. ' +
    'Article Type: ' + (d['Article Type'] || 'Unknown') + '. ' +
    'Primary Entity: ' + (d['Primary Entity'] || 'Unknown') + '. ' +
    'Primary Query Cluster: ' + (d['Primary Query Cluster Owned'] || 'Unknown') + '.\n\n' +
    'Choose exactly one angle from this list:\n' +
    'WHY — explains why the problem occurs at material level\n' +
    'ID — helps reader identify which specific condition they have\n' +
    'FIX — explains how the problem is professionally corrected\n' +
    'PREVENT — explains how to prevent the problem recurring\n' +
    'CHOOSE — helps reader evaluate and select a specialist\n' +
    'LOCAL — offers professional correction in a specific location\n' +
    'PROVE — documents a real project as proof of outcome\n' +
    'DECIDE — helps reader decide if professional intervention is worth it\n\n' +
    'Return exactly one line:\n' +
    'Problem Angle: [CODE] — [one sentence describing this article\'s specific approach]\n\n' +
    'After pasting into ChatGPT — paste the result into the Problem Angle column in the posts sheet, then re-run Initialise Runner.';
}

function _w1Prompt_GSCRecovery(d, bImpr, mImpr, pct) {
  return 'Copy this prompt into ChatGPT:\n\n' +
    'Generate a recovery brief for this page which has lost search visibility. ' +
    'Page Title: ' + (d['Title'] || 'Unknown') + '. ' +
    'URL: ' + (d['URL'] || 'Unknown') + '. ' +
    'Article Type: ' + (d['Article Type'] || 'Unknown') + '. ' +
    'Primary Query Cluster: ' + (d['Primary Query Cluster Owned'] || 'Unknown') + '. ' +
    'Benchmark impressions before May 2025 migration: ' + bImpr + '. ' +
    'Monitor impressions after migration: ' + mImpr + ' (' + pct + '% change). ' +
    'Benchmark queries: ' + (d['Benmrk Queries'] || 'None recorded') + '. ' +
    'Monitor queries: ' + (d['Post ID'] ? bc_getQueriesFromSheet(String(d['Post ID']), 'GSC Monitor') || 'None recorded' : d['Montr Queries'] || 'None recorded') + '.\n\n' +
    'Return exactly one line:\n' +
    'Page Rewrite Brief: [max 60 words — which query clusters to recover, what entity coverage to strengthen, what the opening section must address]\n\n' +
    'After pasting into ChatGPT — paste the result into the Page Rewrite Brief column in the posts sheet, then re-run Initialise Runner.';
}

function _w1Prompt_CTR(d, mImpr, mPos) {
  return 'Copy this prompt into ChatGPT:\n\n' +
    'Generate an improved H1 and Meta Description for this page. It is ranking but getting zero clicks. ' +
    'Page Title: ' + (d['Title'] || 'Unknown') + '. ' +
    'URL: ' + (d['URL'] || 'Unknown') + '. ' +
    'Stone Type: ' + (d['Stone Type'] || 'Unknown') + '. ' +
    'Primary Query Cluster: ' + (d['Primary Query Cluster Owned'] || 'Unknown') + '. ' +
    'Current H1: ' + (d['New H1'] || 'Not set') + '. ' +
    'Current Meta Description: ' + (d['New Meta Description'] || 'Not set') + '. ' +
    'Impressions: ' + mImpr + ' | Clicks: 0 | Position: ' + (mPos ? mPos.toFixed(1) : 'unknown') + '.\n\n' +
    'Return exactly two lines:\n' +
    'New H1: [max 60 chars — opens with the reader\'s problem, no trailing full stop]\n' +
    'New Meta Description: [max 155 chars — addresses the reader\'s problem, ends with a call to action]\n\n' +
    'After pasting into ChatGPT — paste the results into New H1 and New Meta Description columns in the posts sheet, then re-run Initialise Runner.';
}

function _w1Prompt_RewriteViolation(d) {
  return 'Copy this prompt into ChatGPT:\n\n' +
    'The governance analysis for this page returned Rewrite Status: Needs Rewrite. ' +
    'Page Title: ' + (d['Title'] || 'Unknown') + '. ' +
    'URL: ' + (d['URL'] || 'Unknown') + '. ' +
    'Rewrite Governance Summary: ' + (d['Rewrite Governance Summary'] || 'Not specified') + '. ' +
    'Page Rewrite Brief: ' + (d['Page Rewrite Brief'] || 'Not specified') + '.\n\n' +
    'Run the governance prompt again for this page using the current page content and GSC data. ' +
    'Return the full updated 24-column TSV row.\n\n' +
    'After pasting into ChatGPT — paste the TSV result into the posts sheet Article Type column for this row, then re-run Initialise Runner.';
}

function _w1Prompt_GovernanceTightening(d) {
  var rewriteBrief = String(d['Page Rewrite Brief'] || '').trim();
  var govSummary   = String(d['Rewrite Governance Summary'] || '').trim();

  // If brief exists and is substantial — confirm and proceed, no action needed
  if (rewriteBrief && rewriteBrief.length >= 20) {
    return 'GOVERNANCE NOTE — no action required.\n\n' +
      'Rewrite brief confirmed:\n' +
      rewriteBrief + '\n\n' +
      'NEXT STEP: Use menu \u25b6 Pipeline \u2192 W0 Generate Forensic Prompt';
  }

  // Brief is missing or too vague — generate a prompt to fix it
  return 'TASK: Generate a specific rewrite brief for this page.\n\n' +
    'Page Title: ' + (d['Title'] || 'Unknown') + '\n' +
    'Article Type: ' + (d['Article Type'] || 'Unknown') + '\n' +
    'Primary Entity: ' + (d['Primary Entity'] || 'Unknown') + '\n' +
    'Primary Query Cluster: ' + (d['Primary Query Cluster Owned'] || 'Unknown') + '\n' +
    'Governance Summary: ' + (govSummary || 'Not specified') + '\n\n' +
    'Write a rewrite brief of maximum 60 words that specifies:\n' +
    '1. What the opening section must address\n' +
    '2. What entity coverage to strengthen\n' +
    '3. What scope boundary to observe\n\n' +
    'OUTPUT FORMAT (one line only):\n' +
    'Page Rewrite Brief: [your brief here]\n\n' +
    'Paste this into the Page Rewrite Brief column in the posts sheet for this row.\n' +
    'Then re-run Pipeline → Initialise Runner.';
}

function runW1AutoCheck(d) {
  return runW1ReadinessCheck(d);
}

function writeW1CheckToRunner(sh, result) {
  return 11;
}

function prRunW1Check() {
  // Read from posts sheet explicitly
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange() ? posts.getActiveRange().getRow() : 2;
  var headers = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0]
                     .map(function(h) { return String(h).trim(); });
  var vals    = posts.getRange(row, 1, 1, headers.length).getValues()[0];
  var d = {};
  headers.forEach(function(h, i) { d[h] = vals[i]; });
  runW1ReadinessCheck(d);
}


function initPipelineRunner() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var posts   = ss.getSheetByName('posts');
  var activeRow = posts.getActiveRange().getRow();

  if (activeRow < 2) {
    SpreadsheetApp.getUi().alert('Select a data row in the posts sheet first, then run Initialise Runner.');
    return;
  }

  // Read data explicitly from posts sheet BEFORE switching to Pipeline Runner
  var headers = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0]
                     .map(function(h) { return String(h).trim(); });
  var vals    = posts.getRange(activeRow, 1, 1, headers.length).getValues()[0];
  var d = {};
  headers.forEach(function(h, i) { d[h] = vals[i]; });

  var sh = resetPRSheet();  // Delete and recreate — guaranteed clean sheet

  // Set column widths
  sh.setColumnWidth(PR_COL_LABEL,   120); // A — label
  sh.setColumnWidth(PR_COL_STATUS,   55); // B — status
  sh.setColumnWidth(PR_COL_NOTE,    100); // C — note
  // D onwards — merge into content area, set wide
  for (var c = PR_COL_CONTENT; c <= 26; c++) {
    sh.setColumnWidth(c, 80);
  }

  // ── ROW 1 — Main header ──
  sh.getRange(1, 2)
    .setValue('ABBEY FLOOR CARE — CONTENT ENGINE v76 — PIPELINE RUNNER')
    .setBackground(PR_COLOURS.header)
    .setFontColor(PR_COLOURS.headerFg)
    .setFontFamily('JetBrains Mono')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 28);

  // ── ROW 2 — Post identity ──
  sh.getRange(2, PR_COL_LABEL).setValue('POST ID').setBackground('#37474f').setFontColor('#fff').setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
  sh.getRange(2, PR_COL_STATUS).setValue(d['Post ID'] || '').setBackground('#eceff1').setFontFamily('JetBrains Mono').setFontSize(8);
  sh.getRange(2, PR_COL_NOTE).setValue('ARTICLE TYPE').setBackground('#37474f').setFontColor('#fff').setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
  sh.getRange(2, 2).setValue(d['Article Type'] || '').setBackground('#eceff1').setFontFamily('JetBrains Mono').setFontSize(8);
  sh.getRange(2, 9).setValue('STONE TYPE').setBackground('#37474f').setFontColor('#fff').setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
  sh.getRange(2, 2).setValue('STONE TYPE: ' + (d['Stone Type'] || '')).setBackground('#eceff1').setFontFamily('JetBrains Mono').setFontSize(8);
  sh.setRowHeight(2, 22);

  // ── ROW 3 — Title ──
  sh.getRange(3, PR_COL_LABEL).setValue('TITLE').setBackground('#37474f').setFontColor('#fff').setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
  sh.getRange(3, 2).setValue(d['Title'] || '').setBackground('#eceff1').setFontFamily('JetBrains Mono').setFontSize(8).setWrap(false);
  sh.setRowHeight(3, 22);

  // ── ROW 4 — URL ──
  sh.getRange(4, PR_COL_LABEL).setValue('URL').setBackground('#37474f').setFontColor('#fff').setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
  sh.getRange(4, 2).setValue(d['URL'] || '').setBackground('#eceff1').setFontFamily('JetBrains Mono').setFontSize(8).setWrap(false);
  sh.setRowHeight(4, 22);

  // ── ROW 5 — Query cluster and problem angle ──
  sh.getRange(5, PR_COL_LABEL).setValue('PRIMARY CLUSTER').setBackground('#37474f').setFontColor('#fff').setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
  sh.getRange(5, 2).setValue(d['Primary Query Cluster Owned'] || '').setBackground('#eceff1').setFontFamily('JetBrains Mono').setFontSize(8);
  sh.getRange(5, 14).setValue('PROBLEM ANGLE').setBackground('#37474f').setFontColor('#fff').setFontFamily('JetBrains Mono').setFontSize(8).setFontWeight('bold');
  sh.getRange(5, 2).setValue('PROBLEM ANGLE: ' + (d['Problem Angle'] || 'NOT SET — run W-1')).setBackground('#eceff1').setFontFamily('JetBrains Mono').setFontSize(8);
  sh.setRowHeight(5, 22);

  // ── ROW 6 — Separator ──
  sh.getRange(6, 2).setValue('').setBackground(PR_COLOURS.separator);
  sh.setRowHeight(6, 6);

  // ── W-1 AUTOMATED CHECK (rows 7-13) ──
  // Run all five checks automatically — no manual reading required
  var w1Result = runW1AutoCheck(d);
  var nextRow  = writeW1CheckToRunner(sh, w1Result);
  // Also run problem register check and append
  _prCheckProblemRegisterToRunner(sh, nextRow, d);

  // ── W0B LOCATION CONTEXT — dynamic row after W-1 checks ──
  var w0bStart = nextRow;
  var isGeo = String(d['Article Type'] || '').trim() === 'Geo Service Page' ||
              String(d['Article Type'] || '').trim() === 'Case Study';
  var geoBg = isGeo ? '#00695c' : '#78909c';
  sh.getRange(w0bStart, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(w0bStart, 6);
  var w0bHdr = w0bStart + 1;
  prWriteSectionHeader(sh, w0bHdr, 'W0B — LOCATION CONTEXT' + (isGeo ? '  |  Menu: Pipeline → W0B Save Location Context' : '  |  NOT APPLICABLE — not a Geo Service Page'), geoBg);
  prWriteRow(sh, w0bHdr+1, 'LOCALITY', d['Locality'] || (isGeo ? 'ENTER IN POSTS SHEET col DK' : 'n/a'), isGeo ? PR_COLOURS.paste : '#eceff1');
  sh.setRowHeight(w0bHdr+1, 22);
  prWriteRow(sh, w0bHdr+2, 'PARENT AREA', d['Parent Area'] || (isGeo ? 'ENTER IN POSTS SHEET col DL' : 'n/a'), isGeo ? PR_COLOURS.paste : '#eceff1');
  sh.setRowHeight(w0bHdr+2, 22);
  prWriteRow(sh, w0bHdr+3, 'LOCATION CONTEXT', d['Location Context'] || (isGeo ? 'NOT BUILT — use menu: Pipeline → W0B' : 'n/a'), isGeo ? PR_COLOURS.result : '#eceff1', true);
  sh.setRowHeight(w0bHdr+3, 60);
  prSetStatus(sh, w0bHdr, isGeo ? (d['Location Context'] ? 'DONE' : 'PENDING') : 'DONE');
  var w0bEnd = w0bHdr + 4;
  sh.getRange(w0bEnd, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(w0bEnd, 6);

  // ── W0 FORENSIC (rows 17-23) ──
  prWriteSectionHeader(sh, 17, 'W0 — STAGE 0 FORENSIC ANALYSIS  |  Menu: Pipeline → W0 Generate Forensic Prompt');
  prWritePrompt(sh, 18, 'PROMPT', 'Use menu: Pipeline → W0 Generate Forensic Prompt — then copy to ChatGPT');
  prWritePasteArea(sh, 19, 'PASTE OUTPUT', '← Paste ChatGPT Stage 0 output here');
  prWriteRow(sh, 20, 'PRIMARY SEARCH TERM', d['Primary Search Term'] || 'Not yet pushed', PR_COLOURS.result);
  sh.setRowHeight(20, 22);
  prWriteRow(sh, 21, 'STRATEGIC REASONING', d['Strategic Reasoning'] || 'Not yet pushed', PR_COLOURS.result, true);
  sh.setRowHeight(21, 40);
  prWriteRow(sh, 22, 'TOPICAL INTENT', d['Topical Intent'] || 'Not yet pushed', PR_COLOURS.result);
  sh.setRowHeight(22, 22);
  prWriteRow(sh, 23, 'MATRIX ROLE / KEY DECISIONS', (d['Matrix Role'] || '') + ' | ' + (d['Key Decisions Explained'] || ''), PR_COLOURS.result, true);
  sh.setRowHeight(23, 40);
  prSetStatus(sh, 17, d['Primary Search Term'] ? 'DONE' : 'PENDING');

  // ── ROW 24 — Separator ──
  sh.getRange(24, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(24, 6);

  // ── W1A-W1D DATA AND RULES (rows 25-35) ──
  var stages1 = [
    { row: 25, label: 'W1A — DATA LOADING',     key: 'buildStage1APrompt' },
    { row: 28, label: 'W1B — STRUCTURAL RULES',  key: 'buildStage1BPrompt' },
    { row: 31, label: 'W1C — PARAGRAPH RULES',   key: 'buildStage1CPrompt' },
    { row: 34, label: 'W1D — OUTPUT QUALITY',    key: 'buildStage1DPrompt' }
  ];
  stages1.forEach(function(s) {
    prWriteSectionHeader(sh, s.row, s.label + '  |  Menu: Pipeline → ' + s.label.split('—')[0].trim() + ' Generate Prompt');
    prWritePrompt(sh, s.row + 1, 'PROMPT', 'Use menu to generate — copy to ChatGPT — respond "Stage received"');
    prWriteRow(sh, s.row + 2, 'STATUS', 'Paste LLM acknowledgement here or leave blank — copy prompt and continue', PR_COLOURS.paste);
    sh.setRowHeight(s.row + 2, 22);
    prSetStatus(sh, s.row, 'PENDING');
  });

  // ── ROW 37 — Separator ──
  sh.getRange(37, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(37, 6);

  // ── W1.5A AUDIT (rows 38-41) ──
  prWriteSectionHeader(sh, 38, 'W1.5A — AUDIT  |  Menu: Pipeline → W1.5A Generate Audit Prompt');
  prWritePrompt(sh, 39, 'PROMPT', 'Use menu to generate — copy to ChatGPT');
  prWritePasteArea(sh, 40, 'PASTE AUDIT OUTPUT', '← Paste ChatGPT audit output here');
  prWriteRow(sh, 41, 'AUDIT RESULT', 'Paste result summary here for reference', PR_COLOURS.pending);
  sh.setRowHeight(41, 22);
  prSetStatus(sh, 38, 'PENDING');

  // ── ROW 42 — Separator ──
  sh.getRange(42, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(42, 6);

  // ── W1.5B SECTION PLAN (rows 43-47) ──
  prWriteSectionHeader(sh, 43, 'W1.5B — SECTION PLAN  |  Menu: Pipeline → W1.5B Generate Section Plan Prompt');
  prWritePrompt(sh, 44, 'PROMPT', 'Use menu to generate — copy to ChatGPT');
  prWritePasteArea(sh, 45, 'PASTE SECTION PLAN', '← Paste approved section plan here');
  sh.setRowHeight(45, 100);
  var storedPlan = String(d['Approved Section Plan'] || '').trim();
  prWriteRow(sh, 46, 'STORED PLAN', storedPlan || 'Not yet saved', storedPlan ? PR_COLOURS.result : PR_COLOURS.pending, true);
  sh.setRowHeight(46, 60);
  prWriteRow(sh, 47, 'ACTION', 'After pasting plan above — use menu: Pipeline → W1.5B Save Section Plan to Posts', PR_COLOURS.paste);
  sh.setRowHeight(47, 22);
  prSetStatus(sh, 43, storedPlan ? 'DONE' : 'PENDING');

  // ── ROW 48 — Separator ──
  sh.getRange(48, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(48, 6);

  // ── W2A PRE-WRITE (rows 49-51) ──
  prWriteSectionHeader(sh, 49, 'W2A — PRE-WRITE  |  Menu: Pipeline → W2A Generate Pre-Write Prompt');
  prWritePrompt(sh, 50, 'PROMPT', 'Use menu to generate — copy to ChatGPT');
  prWriteRow(sh, 51, 'LLM RESPONSE', 'Paste pre-write acknowledgement here for reference', PR_COLOURS.paste);
  sh.setRowHeight(51, 22);
  prSetStatus(sh, 49, 'PENDING');

  // ── ROW 52 — Separator ──
  sh.getRange(52, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(52, 6);

  // ── W2B EXECUTE (rows 53-60) ──
  prWriteSectionHeader(sh, 53, 'W2B — EXECUTE  |  Menu: Pipeline → W2B Generate Execute Prompt');
  prWritePrompt(sh, 54, 'PART 1', 'Use menu to generate Part 1 — copy to ChatGPT');
  prWritePrompt(sh, 55, 'PART 2', 'Use menu to generate Part 2 — copy to ChatGPT');
  prWritePrompt(sh, 56, 'PART 3', 'Use menu to generate Part 3 — copy to ChatGPT');
  prWritePasteArea(sh, 57, 'PASTE HTML OUTPUT', '← Paste ChatGPT HTML output here');
  sh.setRowHeight(57, 200);
  var storedHTML = String(d['New HTML'] || '').trim();
  prWriteRow(sh, 58, 'STORED HTML', storedHTML ? storedHTML.substring(0, 200) + '...' : 'Not yet saved', storedHTML ? PR_COLOURS.result : PR_COLOURS.pending, true);
  sh.setRowHeight(58, 40);
  prWriteRow(sh, 59, 'ACTION', 'After pasting HTML above — use menu: Pipeline → W2B Save HTML to Posts', PR_COLOURS.paste);
  sh.setRowHeight(59, 22);
  prWriteRow(sh, 60, 'WORD COUNT', 'Will be calculated when HTML is saved', PR_COLOURS.pending);
  sh.setRowHeight(60, 22);
  prSetStatus(sh, 53, storedHTML ? 'DONE' : 'PENDING');

  // ── ROW 61 — Separator ──
  sh.getRange(61, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(61, 6);

  // ── W4B IMAGE ALT & CAPTION (rows 62-66) ──
  prWriteSectionHeader(sh, 62, 'W4B — IMAGE ALT & CAPTION  |  Menu: Pipeline → W4B Load Figure Inventory');
  prWriteRow(sh, 63, 'FIGURE INVENTORY', 'Use menu: Pipeline → W4B Load Figure Inventory', PR_COLOURS.pending, true);
  sh.setRowHeight(63, 60);
  prWritePrompt(sh, 64, 'ALT/CAPTION PROMPT', 'Use menu: Pipeline → W4B Generate Alt/Caption Prompt');
  prWritePasteArea(sh, 65, 'PASTE LLM OUTPUT', '← Paste alt/caption output here');
  prWriteRow(sh, 66, 'ACTION', 'After pasting — use menu to push alt/caption data to posts sheet', PR_COLOURS.paste);
  sh.setRowHeight(66, 22);
  prSetStatus(sh, 62, 'PENDING');

  // ── ROW 67 — Separator ──
  sh.getRange(67, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(67, 6);

  // ── W4 HTML PASTE + AUDIT (rows 68-73) ──
  prWriteSectionHeader(sh, 68, 'W4 — HTML AUDIT + PUSH  |  Menu: Pipeline → W4 Run Audit, then W4 Push HTML to Posts');
  prWritePasteArea(sh, 69, 'HTML FOR AUDIT', '← Paste Stage 2B HTML output here');
  sh.setRowHeight(69, 120);
  prWriteRow(sh, 70, 'AUDIT RESULTS', 'Will appear here after audit', PR_COLOURS.pending, true);
  sh.setRowHeight(70, 120);
  prWriteRow(sh, 71, 'PASS COUNT', 'Pending', PR_COLOURS.pending);
  sh.setRowHeight(71, 22);
  prWriteRow(sh, 72, 'CORRECTION PROMPT', 'Will appear here if audit fails', PR_COLOURS.pending, true);
  sh.setRowHeight(72, 22);
  prWriteRow(sh, 73, 'ACTION', 'Use menu: Pipeline → W4 Push HTML to Posts after audit passes', PR_COLOURS.paste);
  sh.setRowHeight(73, 22);
  prSetStatus(sh, 68, 'PENDING');

  // ── ROW 74 — Separator ──
  sh.getRange(74, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(74, 6);

  // ── W4 FINAL HTML AUDIT (rows 75-79) ──
  prWriteSectionHeader(sh, 75, 'W4 — FINAL HTML AUDIT  |  Menu: Pipeline → W4 Run Audit');
  prWriteRow(sh, 70, 'AUDIT RESULTS', 'Use menu: Pipeline → W4 Run Audit', PR_COLOURS.pending, true);
  sh.setRowHeight(70, 80);
  prWriteRow(sh, 71, 'PASS COUNT', '', PR_COLOURS.pending);
  sh.setRowHeight(77, 22);
  prWriteRow(sh, 72, 'CORRECTION PROMPT', 'If audit fails — correction prompt will appear here', PR_COLOURS.pending, true);
  sh.setRowHeight(72, 60);
  prWriteRow(sh, 79, 'ACTION', 'If audit passes — use menu: Pipeline → W4 Push HTML to Posts', PR_COLOURS.paste);
  sh.setRowHeight(79, 22);
  prSetStatus(sh, 68, 'PENDING');

  // ── ROW 80 — Separator ──
  sh.getRange(80, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(80, 6);

  // ── W5 GOVERNANCE (rows 81-87) ──
  prWriteSectionHeader(sh, 81, 'W5 — GOVERNANCE  |  Menu: Pipeline → W5 Generate Governance Prompt');
  prWritePrompt(sh, 82, 'PROMPT', 'Use menu to generate — copy to ChatGPT');
  prWritePasteArea(sh, 83, 'PASTE LLM OUTPUT', '← Paste H1 + Meta Title + Meta Description output here');
  sh.setRowHeight(83, 80);
  prWriteRow(sh, 84, 'NEW H1', d['New H1'] || 'Not yet pushed', PR_COLOURS.result);
  sh.setRowHeight(84, 22);
  prWriteRow(sh, 85, 'NEW META TITLE', d['New Meta Title'] || 'Not yet pushed', PR_COLOURS.result);
  sh.setRowHeight(85, 22);
  prWriteRow(sh, 86, 'NEW META DESC', d['New Meta Description'] || 'Not yet pushed', PR_COLOURS.result, true);
  sh.setRowHeight(86, 40);
  prWriteRow(sh, 87, 'ACTION', 'After pasting — use menu: Pipeline → W5 Push Governance Fields', PR_COLOURS.paste);
  sh.setRowHeight(87, 22);
  prSetStatus(sh, 81, d['New H1'] ? 'DONE' : 'PENDING');

  // ── ROW 88 — Separator ──
  sh.getRange(88, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(88, 6);

  // ── W5C SCHEMA (rows 89-94) ──
  prWriteSectionHeader(sh, 89, 'W5C — SCHEMA BUILDER  |  Menu: Pipeline → W5C Build Schema');
  prWriteRow(sh, 90, 'PRE-FLIGHT', 'Use menu: Pipeline → W5C Build Schema — pre-flight runs automatically', PR_COLOURS.pending, true);
  sh.setRowHeight(90, 22);
  var storedSchema = String(d['Schema (JSON-LD)'] || '').trim();
  prWriteRow(sh, 91, 'SCHEMA OUTPUT', storedSchema || 'Not yet built', storedSchema ? PR_COLOURS.result : PR_COLOURS.pending, true);
  sh.setRowHeight(91, 80);
  prWriteRow(sh, 92, 'W5B AUDIT RESULT', 'Use menu: Pipeline → W5B Run Schema Audit', PR_COLOURS.pending, true);
  sh.setRowHeight(92, 40);
  prWriteRow(sh, 93, 'CORRECTION PROMPT', 'If schema audit fails — correction prompt appears here', PR_COLOURS.pending, true);
  sh.setRowHeight(93, 40);
  prWriteRow(sh, 94, 'ACTION', 'Schema is pushed automatically when built — W5B audit runs automatically after', PR_COLOURS.paste);
  sh.setRowHeight(94, 22);
  prSetStatus(sh, 89, storedSchema ? 'DONE' : 'PENDING');

  // ── ROW 95 — Separator ──
  sh.getRange(95, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(95, 6);

  // ── W2C IMAGE SUGGESTION (rows 97-99) ──
  sh.getRange(97, 2).setBackground(PR_COLOURS.separator);
  sh.setRowHeight(97, 6);
  prWriteSectionHeader(sh, 98, 'W2C — IMAGE SUGGESTION  |  Menu: Pipeline → W2C Generate Image Suggestion Prompt');
  prWritePrompt(sh, 99, 'PROMPT', 'Use menu to generate — copy to ChatGPT');
  prWritePasteArea(sh, 100, 'PASTE LLM OUTPUT', '← Paste JSON suggestions here, then use menu: Pipeline → W2C Save Suggestions to Posts');
  sh.setRowHeight(100, 100);
  prSetStatus(sh, 98, 'PENDING');

  // ── PUSH ALL (row 96) ──
  sh.getRange(96, 2)
    .setValue('⬆  PIPELINE COMPLETE — USE MENU: Pipeline → Push ALL Results to Posts Sheet  ⬆')
    .setBackground('#1b5e20')
    .setFontColor('#ffffff')
    .setFontFamily('JetBrains Mono')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(96, 30);

  // Freeze header rows
  sh.setFrozenRows(5);

  // Navigate to Pipeline Runner sheet
  ss.setActiveSheet(sh);
  sh.getRange(7, 1).activate();

  SpreadsheetApp.getUi().alert(
    'Pipeline Runner initialised for:\n' +
    (d['Title'] || 'Unknown') + '\n\n' +
    'Use the Pipeline menu to run each stage.\n' +
    'Paste ChatGPT outputs directly into the yellow cells.'
  );
}

/* ============================================================
   W-1 — BUILD AUTHORITY BRIEF
   Reads entity analysis columns, builds brief, writes to runner
   and pushes to posts sheet col DN.
============================================================ */
function prBuildAuthorityBrief() {
  var sh  = getPRSheet();
  var res = buildAuthorityBrief();
  if (!res.success) {
    prWriteRow(sh, 8, 'BRIEF', '❌ ' + res.message, PR_COLOURS.fail, true);
    prSetStatus(sh, 7, 'FAIL');
    return;
  }
  prWriteRow(sh, 8, 'BRIEF', res.brief, PR_COLOURS.result, true);
  sh.setRowHeight(8, 150);
  // Push to sheet
  var pushRes = pushAuthorityBriefToSheet(res.brief);
  if (pushRes.success) {
    prSetStatus(sh, 7, 'DONE');
  } else {
    prWriteRow(sh, 10, 'REGISTER CHECK', '⚠ Built but not saved: ' + pushRes.message, PR_COLOURS.fail);
    prSetStatus(sh, 7, 'PASS');
  }
}

/* ============================================================
   W-1 — CHECK PROBLEM REGISTER
   Checks estate coverage and writes result to runner row 10.
============================================================ */
function prCheckProblemRegister() {
  var sh = getPRSheet();
  var d  = getActiveRowDataMap();
  var problem   = String(d['Primary Query Cluster Owned'] || d['Observed Query Cluster'] || '').trim();
  var activeUrl = String(d['URL'] || '').trim();

  if (!problem) {
    prWriteRow(sh, 10, 'REGISTER CHECK', '❌ No Primary Query Cluster found — run entity analysis first', PR_COLOURS.fail);
    return;
  }

  var res = checkProblemRegister(problem, String(d['Stone Type'] || ''), activeUrl);

  if (!res.success) {
    prWriteRow(sh, 10, 'REGISTER CHECK', '❌ ' + res.message, PR_COLOURS.fail);
    return;
  }

  if (!res.found) {
    prWriteRow(sh, 10, 'REGISTER CHECK', '✔ NEW ANGLE — no other articles address this problem', PR_COLOURS.result);
  } else {
    var matchText = res.count + ' article(s) address this problem:\n';
    res.matches.forEach(function(m) {
      matchText += '  [' + (m.angleCode || 'no angle') + '] ' + m.articleType + ' — ' + m.title + '\n';
    });
    matchText += '\nSelect a different angle to differentiate this article.';
    prWriteRow(sh, 10, 'REGISTER CHECK', matchText, res.count > 2 ? PR_COLOURS.fail : PR_COLOURS.paste, true);
    sh.setRowHeight(10, 20 + (res.count * 18));
  }
}

/* ============================================================
   INTERNAL HELPER — Write problem register result to runner
   Called during init to append register check after W-1 checks.
============================================================ */
function _prCheckProblemRegisterToRunner(sh, startRow, d) {
  var problem   = String(d['Primary Query Cluster Owned'] || d['Observed Query Cluster'] || '').trim();
  var activeUrl = String(d['URL'] || '').trim();

  if (!problem) return;

  var res = checkProblemRegister(problem, String(d['Stone Type'] || ''), activeUrl);
  if (!res.success) return;

  var bg      = res.found ? (res.count > 2 ? '#ffebee' : '#fff8e1') : '#e8f5e9';
  var fg      = res.found ? (res.count > 2 ? '#b71c1c' : '#e65100') : '#1b5e20';
  var emoji   = res.found ? (res.count > 2 ? '❌' : '⚠') : '✔';
  var content = emoji + ' PROBLEM REGISTER: ' + res.message;

  if (res.found && res.matches) {
    res.matches.forEach(function(m) {
      content += '\n  [' + (m.angleCode || 'no angle') + '] ' +
                 (m.articleType || '') + ' - ' + (m.title || m.url || '');
    });
    if (!d['Problem Angle']) {
      content += '\n\nACTION: Enter the angle code in col DO of the posts sheet to differentiate this article.';
    }
  }

  sh.getRange(startRow, 2)
    .setValue(content)
    .setBackground(bg)
    .setFontColor(fg)
    .setFontFamily('JetBrains Mono')
    .setFontSize(8)
    .setWrap(true)
    .setVerticalAlignment('top');
  sh.setRowHeight(startRow, res.found ? 20 + (res.count * 18) : 22);
}

/* ============================================================
   W0B — SAVE LOCATION CONTEXT
   Reads locality/parent area from runner rows 13-14,
   generates location context prompt, displays in row 15.
============================================================ */
function prSaveLocationContext() {
  var sh = getPRSheet();
  var res = buildLocationContextPrompt();
  if (!res.success) {
    prWriteRow(sh, 15, 'LOCATION CONTEXT', '❌ ' + res.message, PR_COLOURS.fail);
    prSetStatus(sh, 12, 'FAIL');
    return;
  }
  // Display the prompt for copying
  prWriteRow(sh, 15, 'LOCATION CONTEXT PROMPT', res.prompt + '\n\n← Copy this prompt → paste into ChatGPT → paste result into posts sheet col DM', PR_COLOURS.prompt, true);
  sh.setRowHeight(15, 120);
  prSetStatus(sh, 12, 'RUNNING');
  SpreadsheetApp.getUi().alert('Location Context prompt generated in row 15.\nCopy the prompt, paste into ChatGPT, then paste the result into col DM of the posts sheet and re-initialise the runner.');
}

/* ============================================================
   W0 — GENERATE FORENSIC PROMPT
============================================================ */
function prGenerateW0() {
  var sh  = getPRSheet();
  prSetStatus(sh, 17, 'RUNNING');
  var status = checkGSCDataStatus();
  if (!status.success) {
    var bypassRes = populateStage0FromGovernance();
    if (bypassRes.success) {
      prWriteRow(sh, 18, 'PROMPT', '✔ GOVERNANCE BYPASS — no GSC monitor data. Stage 0 fields populated directly from governance pipeline.\n\nFields written: ' + bypassRes.written.join(', ') + '\n\n' + (bypassRes.notFound && bypassRes.notFound.length > 0 ? 'Columns not found: ' + bypassRes.notFound.join(', ') : ''), PR_COLOURS.result, true);
      sh.setRowHeight(18, 80);
      prSetStatus(sh, 17, 'DONE');
    } else {
      prWriteRow(sh, 18, 'PROMPT', '❌ BYPASS FAILED — ' + bypassRes.message, PR_COLOURS.fail, true);
      prSetStatus(sh, 17, 'FAIL');
    }
    return;
  }
  var prompt = generateForensicPrompt();
  prWriteRow(sh, 18, 'PROMPT', prompt, PR_COLOURS.prompt, true);
  sh.setRowHeight(18, 200);
  prSetStatus(sh, 17, 'RUNNING');
}

/* ============================================================
   W0 — PUSH 5 FIELDS FROM RUNNER PASTE AREA (row 19)
============================================================ */
function prPushW0Fields() {
  var sh       = getPRSheet();
  var rawOutput = String(sh.getRange(19, PR_COL_CONTENT).getValue() || '').trim();
  if (!rawOutput) {
    SpreadsheetApp.getUi().alert('Paste the ChatGPT Stage 0 output into row 19 first.');
    return;
  }
  var res = pushStage0FieldsToActiveRow(rawOutput);
  if (res.success) {
    var d = getActiveRowDataMap();
    prWriteRow(sh, 20, 'PRIMARY SEARCH TERM', d['Primary Search Term'] || '', PR_COLOURS.result);
    prWriteRow(sh, 21, 'STRATEGIC REASONING', d['Strategic Reasoning'] || '', PR_COLOURS.result, true);
    prWriteRow(sh, 22, 'TOPICAL INTENT', d['Topical Intent'] || '', PR_COLOURS.result);
    prWriteRow(sh, 23, 'MATRIX ROLE / KEY DECISIONS', (d['Matrix Role'] || '') + ' | ' + (d['Key Decisions Explained'] || ''), PR_COLOURS.result, true);
    prSetStatus(sh, 17, 'DONE');
  } else {
    prWriteRow(sh, 20, 'PUSH RESULT', '❌ ' + res.message, PR_COLOURS.fail);
    prSetStatus(sh, 17, 'FAIL');
  }
}

/* ============================================================
   W1A-W1D — GENERATE PROMPTS
============================================================ */
function prGenerateW1A() { _prGenerateStage1('buildStage1APrompt', 25, 26); }
function prGenerateW1B() { _prGenerateStage1('buildStage1BPrompt', 28, 29); }
function prGenerateW1C() { _prGenerateStage1('buildStage1CPrompt', 31, 32); }
function prGenerateW1D() { _prGenerateStage1('buildStage1DPrompt', 34, 35); }

function _prGenerateStage1(fnName, headerRow, promptRow) {
  var sh = getPRSheet();
  prSetStatus(sh, headerRow, 'RUNNING');
  try {
    var prompt = this[fnName]();
    prWriteRow(sh, promptRow, 'PROMPT', prompt, PR_COLOURS.prompt, true);
    sh.setRowHeight(promptRow, 200);
    prSetStatus(sh, headerRow, 'RUNNING');
  } catch(e) {
    prWriteRow(sh, promptRow, 'PROMPT', '❌ Error: ' + e.toString(), PR_COLOURS.fail, true);
    prSetStatus(sh, headerRow, 'FAIL');
  }
}

/* ============================================================
   W1.5A — GENERATE AUDIT PROMPT
============================================================ */
function prGenerateW15A() {
  var sh = getPRSheet();
  prSetStatus(sh, 38, 'RUNNING');
  var prompt = buildStage15APrompt ? buildStage15APrompt() : 'buildStage15APrompt not found';
  prWriteRow(sh, 39, 'PROMPT', prompt, PR_COLOURS.prompt, true);
  sh.setRowHeight(39, 200);
  prSetStatus(sh, 38, 'RUNNING');
}

/* ============================================================
   W1.5B — GENERATE SECTION PLAN PROMPT
============================================================ */
function prGenerateW15B() {
  var sh = getPRSheet();
  prSetStatus(sh, 43, 'RUNNING');
  var prompt = buildStage15BPrompt ? buildStage15BPrompt() : 'buildStage15BPrompt not found';
  prWriteRow(sh, 44, 'PROMPT', prompt, PR_COLOURS.prompt, true);
  sh.setRowHeight(44, 200);
  prSetStatus(sh, 43, 'RUNNING');
}

/* ============================================================
   W1.5B — SAVE SECTION PLAN FROM RUNNER ROW 45
============================================================ */
function prSaveSectionPlan() {
  var sh   = getPRSheet();
  var plan = String(sh.getRange(45, PR_COL_CONTENT).getValue() || '').trim();
  if (!plan) {
    SpreadsheetApp.getUi().alert('Paste the approved section plan into row 45 first.');
    return;
  }
  var res = saveSectionPlan(plan);
  if (res && res.success !== false) {
    prWriteRow(sh, 46, 'STORED PLAN', plan, PR_COLOURS.result, true);
    sh.setRowHeight(46, 80);
    prSetStatus(sh, 43, 'DONE');
  } else {
    prWriteRow(sh, 46, 'STORED PLAN', '❌ Save failed — ' + (res ? res.message : 'unknown error'), PR_COLOURS.fail);
    prSetStatus(sh, 43, 'FAIL');
  }
}

/* ============================================================
   W2A — GENERATE PRE-WRITE PROMPT
============================================================ */
function prGenerateW2A() {
  var sh = getPRSheet();
  prSetStatus(sh, 49, 'RUNNING');
  var data   = getStage2BData ? getStage2BData() : null;
  var prompt = data ? data.part1 : 'getStage2BData not found';
  prWriteRow(sh, 50, 'PROMPT', prompt, PR_COLOURS.prompt, true);
  sh.setRowHeight(50, 200);
  prSetStatus(sh, 49, 'RUNNING');
}

/* ============================================================
   W2B — GENERATE EXECUTE PROMPTS (parts 1, 2, 3)
============================================================ */
function prGenerateW2B() {
  var sh   = getPRSheet();
  prSetStatus(sh, 53, 'RUNNING');
  var data = getStage2BData ? getStage2BData() : null;
  if (!data) {
    prWriteRow(sh, 54, 'PART 1', '❌ getStage2BData not found', PR_COLOURS.fail);
    prSetStatus(sh, 53, 'FAIL');
    return;
  }
  prWriteRow(sh, 54, 'PART 1', data.part1 || '', PR_COLOURS.prompt, true);
  sh.setRowHeight(54, 200);
  prWriteRow(sh, 55, 'PART 2', data.part2 || '', PR_COLOURS.prompt, true);
  sh.setRowHeight(55, 200);
  prWriteRow(sh, 56, 'PART 3', data.part3 || '', PR_COLOURS.prompt, true);
  sh.setRowHeight(56, 200);
  prSetStatus(sh, 53, 'RUNNING');
}

/* ============================================================
   W2B — SAVE HTML FROM RUNNER ROW 57
============================================================ */
function prSaveHTML() {
  var sh  = getPRSheet();
  var html = String(sh.getRange(57, PR_COL_CONTENT).getValue() || '').trim();
  if (!html) {
    SpreadsheetApp.getUi().alert('Paste the ChatGPT HTML output into row 57 first.');
    return;
  }
  // Write to posts sheet col 98 (New HTML) — also store figure inventory
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var posts  = ss.getSheetByName('posts');
  var row    = posts.getActiveRange().getRow();
  var headers = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0]
                     .map(function(h) { return String(h).trim(); });
  var htmlIdx = headers.indexOf('New HTML');
  if (htmlIdx > -1) {
    posts.getRange(row, htmlIdx + 1).setValue(html);
  }
  // Store figure inventory
  var figures = extractFigureInventory(html);
  var invIdx  = headers.indexOf('Figure Inventory (JSON)');
  if (invIdx > -1 && figures.length > 0) {
    posts.getRange(row, invIdx + 1).setValue(JSON.stringify(figures));
  }
  var wordCount = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().split(' ').length;
  prWriteRow(sh, 58, 'STORED HTML', html.substring(0, 300) + '...', PR_COLOURS.result, true);
  prWriteRow(sh, 60, 'WORD COUNT', wordCount + ' words', PR_COLOURS.result);
  prSetStatus(sh, 53, 'DONE');
}

/* ============================================================
   W4B — LOAD FIGURE INVENTORY
============================================================ */
function prLoadFigures() {
  var sh      = getPRSheet();
  var figures = getFigureInventory ? getFigureInventory() : [];
  if (!figures || figures.length === 0) {
    prWriteRow(sh, 63, 'FIGURE INVENTORY', 'No figures found — run W2B Save HTML first to extract inventory', PR_COLOURS.fail);
    prSetStatus(sh, 62, 'FAIL');
    return;
  }
  var display = figures.map(function(f, i) {
    return '[' + (i+1) + '] ' + (f.type || 'img') + ' — ' + (f.src || '').split('/').pop() + (f.alt ? ' | alt: ' + f.alt : '');
  }).join('\n');
  prWriteRow(sh, 63, 'FIGURE INVENTORY', display, PR_COLOURS.result, true);
  sh.setRowHeight(63, 20 + (figures.length * 16));
  prSetStatus(sh, 62, 'RUNNING');
}

/* ============================================================
   W4B — GENERATE ALT/CAPTION PROMPT
============================================================ */
function prGenerateW4B() {
  var sh  = getPRSheet();
  var res = buildAltCaptionPrompt ? buildAltCaptionPrompt() : { success: false, message: 'buildAltCaptionPrompt not found' };
  if (!res || !res.success) {
    prWriteRow(sh, 64, 'ALT/CAPTION PROMPT', '❌ ' + (res ? res.message : 'Function not found'), PR_COLOURS.fail, true);
    prSetStatus(sh, 62, 'FAIL');
    return;
  }
  prWriteRow(sh, 64, 'ALT/CAPTION PROMPT', res.prompt, PR_COLOURS.prompt, true);
  sh.setRowHeight(64, 200);
  prSetStatus(sh, 62, 'RUNNING');
}

/* ============================================================
   W3 — PRE-CHECK
============================================================ */
function prPreCheck() {
  var sh   = getPRSheet();
  var html = String(sh.getRange(69, PR_COL_CONTENT).getValue() || '').trim();
  if (!html) {
    prWriteRow(sh, 70, 'AUDIT RESULTS', '\u274c No HTML in row 69 — paste Stage 2B HTML first', PR_COLOURS.fail);
    return;
  }
  var report = runMechanicalAudit ? runMechanicalAudit() : 'Audit function not found';
  // Run against pasted HTML directly
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheet       = ss.getSheetByName('posts');
  var row         = sheet.getActiveRange().getRow();
  var articleType = row >= 2 ? String(sheet.getRange(row, 8).getValue() || 'General').trim() : 'General';
  var hubUrl      = row >= 2 ? String(sheet.getRange(row, 15).getValue() || '').trim() : '';
  report = runMechanicalAuditOnHTML(html, '', articleType, hubUrl);
  var hasFail = report.indexOf('FAIL') > -1;
  prWriteRow(sh, 70, 'AUDIT RESULTS', report, hasFail ? PR_COLOURS.fail : PR_COLOURS.result, true);
  sh.setRowHeight(70, Math.min(400, 20 + report.split('\n').length * 16));
  prWriteRow(sh, 71, 'PASS COUNT', hasFail ? '\u274c Failures found — see above' : '\u2714 All checks passed', hasFail ? PR_COLOURS.fail : PR_COLOURS.result);
  prSetStatus(sh, 68, hasFail ? 'FAIL' : 'PASS');
}



/* ============================================================
   W4 — RUN AUDIT ON HTML
============================================================ */
function prRunAudit() {
  var sh   = getPRSheet();
  var html = String(sh.getRange(69, PR_COL_CONTENT).getValue() || '').trim();
  if (!html) {
    SpreadsheetApp.getUi().alert('Paste HTML into row 69 first.');
    return;
  }
  var res = runMechanicalAudit ? runMechanicalAudit(html, false) : { results: ['Audit function not found'], passCount: 0, totalChecks: 0 };
  prWriteRow(sh, 70, 'AUDIT RESULTS', res.results.join('\n'), PR_COLOURS.result, true);
  sh.setRowHeight(70, 20 + (res.results.length * 16));
  prWriteRow(sh, 71, 'PASS COUNT', res.passCount + ' / ' + res.totalChecks + ' checks passed', res.passCount === res.totalChecks ? PR_COLOURS.result : PR_COLOURS.fail);
  if (res.passCount < res.totalChecks) {
    var corrPrompt = buildAuditCorrectionPrompt ? buildAuditCorrectionPrompt(res.results, html) : 'buildAuditCorrectionPrompt not found';
    prWriteRow(sh, 72, 'CORRECTION PROMPT', corrPrompt, PR_COLOURS.prompt, true);
    sh.setRowHeight(72, 100);
    prSetStatus(sh, 68, 'FAIL');
  } else {
    prWriteRow(sh, 72, 'CORRECTION PROMPT', '✔ All checks passed — no correction needed', PR_COLOURS.result);
    prSetStatus(sh, 68, 'PASS');
  }
}

/* ============================================================
   W4 — PUSH HTML TO POSTS SHEET
============================================================ */
function prPushHTML() {
  var sh   = getPRSheet();
  var html = String(sh.getRange(69, PR_COL_CONTENT).getValue() || '').trim();
  if (!html) {
    SpreadsheetApp.getUi().alert('No HTML in row 69 — paste Stage 2B HTML first.');
    return;
  }
  var res = pushHtmlToActiveRow ? pushHtmlToActiveRow(html) : { success: false, message: 'pushHtmlToActiveRow not found' };
  if (res && res.success) {
    prWriteRow(sh, 79, 'ACTION', '✔ HTML pushed to posts sheet', PR_COLOURS.result);
    prSetStatus(sh, 68, 'DONE');
  } else {
    prWriteRow(sh, 79, 'ACTION', '❌ Push failed — ' + (res ? res.message : 'unknown'), PR_COLOURS.fail);
  }
}

/* ============================================================
   W5 — GENERATE GOVERNANCE PROMPT
============================================================ */
function prGenerateW5() {
  var sh  = getPRSheet();
  prSetStatus(sh, 81, 'RUNNING');
  var res = buildGovernancePrompt ? buildGovernancePrompt() : { success: false, message: 'buildGovernancePrompt not found' };
  if (!res || !res.success) {
    prWriteRow(sh, 82, 'PROMPT', '❌ ' + (res ? res.message : 'Function not found'), PR_COLOURS.fail, true);
    prSetStatus(sh, 81, 'FAIL');
    return;
  }
  prWriteRow(sh, 82, 'PROMPT', res.prompt, PR_COLOURS.prompt, true);
  sh.setRowHeight(82, 200);
  prSetStatus(sh, 81, 'RUNNING');
}

/* ============================================================
   W5 — PUSH GOVERNANCE FIELDS FROM RUNNER ROW 83
============================================================ */
function prPushGovernance() {
  var sh     = getPRSheet();
  var raw    = String(sh.getRange(83, PR_COL_CONTENT).getValue() || '').trim();
  if (!raw) {
    SpreadsheetApp.getUi().alert('Paste ChatGPT governance output into row 83 first.');
    return;
  }
  var res = pushGovernanceFieldsToActiveRow ? pushGovernanceFieldsToActiveRow(raw) : { success: false, message: 'pushGovernanceFieldsToActiveRow not found' };
  if (res && res.success) {
    var d = getActiveRowDataMap();
    prWriteRow(sh, 84, 'NEW H1', d['New H1'] || '', PR_COLOURS.result);
    prWriteRow(sh, 85, 'NEW META TITLE', d['New Meta Title'] || '', PR_COLOURS.result);
    prWriteRow(sh, 86, 'NEW META DESC', d['New Meta Description'] || '', PR_COLOURS.result, true);
    prSetStatus(sh, 81, 'DONE');
  } else {
    prWriteRow(sh, 84, 'NEW H1', '❌ ' + (res ? res.message : 'unknown error'), PR_COLOURS.fail);
    prSetStatus(sh, 81, 'FAIL');
  }
}

/* ============================================================
   W5C — BUILD SCHEMA
============================================================ */
function prBuildSchema() {
  var sh  = getPRSheet();
  prSetStatus(sh, 89, 'RUNNING');
  var res = generateSchemaForActiveRow ? generateSchemaForActiveRow() : { success: false, message: 'generateSchemaForActiveRow not found' };
  if (!res) { res = { success: false, message: 'No response from schema builder' }; }

  if (res.preflight && !res.preflight.pass) {
    prWriteRow(sh, 90, 'PRE-FLIGHT', '⚠ ' + res.preflight.message, PR_COLOURS.paste, true);
    prSetStatus(sh, 89, 'FAIL');
    return;
  }
  prWriteRow(sh, 90, 'PRE-FLIGHT', '✔ Pre-flight passed', PR_COLOURS.result);
  if (res.success) {
    prWriteRow(sh, 91, 'SCHEMA OUTPUT', res.schema || '', PR_COLOURS.result, true);
    sh.setRowHeight(91, 120);
    prSetStatus(sh, 89, 'DONE');
    // Auto-run W5B audit
    prRunSchemaAudit();
  } else {
    prWriteRow(sh, 91, 'SCHEMA OUTPUT', '❌ ' + res.message, PR_COLOURS.fail, true);
    prSetStatus(sh, 89, 'FAIL');
  }
}

/* ============================================================
   W5B — RUN SCHEMA AUDIT
============================================================ */
function prRunSchemaAudit() {
  var sh  = getPRSheet();
  var res = runSchemaAudit ? runSchemaAudit() : { results: ['runSchemaAudit not found'], passCount: 0, totalChecks: 0 };
  if (!res) return;
  prWriteRow(sh, 92, 'W5B AUDIT RESULT', res.results ? res.results.join('\n') : String(res), PR_COLOURS.result, true);
  sh.setRowHeight(92, 60);
  if (res.correctionPrompt) {
    prWriteRow(sh, 93, 'CORRECTION PROMPT', res.correctionPrompt, PR_COLOURS.prompt, true);
    sh.setRowHeight(93, 80);
  }
}

/* ============================================================
   PUSH ALL — Push all results from runner to posts sheet
   Reads rows 19 (W0), 45 (section plan), 57 (W2B HTML),
   71 (HTML), 83 (W5 governance) and pushes.
============================================================ */
function prPushAll() {
  var sh  = getPRSheet();
  var ui  = SpreadsheetApp.getUi();
  var log = [];

  // W0 fields
  var w0paste = String(sh.getRange(19, PR_COL_CONTENT).getValue() || '').trim();
  if (w0paste) {
    var w0res = pushStage0FieldsToActiveRow(w0paste);
    log.push('W0: ' + (w0res.success ? '✔ ' + w0res.message : '❌ ' + w0res.message));
  }

  // Section plan
  var plan = String(sh.getRange(45, PR_COL_CONTENT).getValue() || '').trim();
  if (plan) {
    saveSectionPlan(plan);
    log.push('W1.5B Section Plan: ✔ saved');
  }

  // W2B HTML
  var rawHtml = String(sh.getRange(57, PR_COL_CONTENT).getValue() || '').trim();
  if (rawHtml) {
    prSaveHTML();
    log.push('W2B HTML: ✔ saved');
  }

  // HTML
  var rehydHtml = String(sh.getRange(69, PR_COL_CONTENT).getValue() || '').trim();
  if (rehydHtml) {
    var pushRes = pushHtmlToActiveRow ? pushHtmlToActiveRow(rehydHtml) : null;
    log.push('W4 HTML: ' + (pushRes && pushRes.success ? '✔ pushed' : '❌ failed'));
  }

  // W5 governance
  var govRaw = String(sh.getRange(83, PR_COL_CONTENT).getValue() || '').trim();
  if (govRaw) {
    var govRes = pushGovernanceFieldsToActiveRow ? pushGovernanceFieldsToActiveRow(govRaw) : null;
    log.push('W5 Governance: ' + (govRes && govRes.success ? '✔ pushed' : '❌ failed'));
  }

  ui.alert('Push All Complete\n\n' + log.join('\n'));
}

/* ============================================================
   PIPELINE RUNNER SIDEBAR
   Opens alongside the Pipeline Runner sheet.
   Reads prompt from active row, provides copy + paste interface.
============================================================ */

function openPipelineRunnerSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('PipelineRunnerSidebar')
    .setTitle('Pipeline — Copy / Paste')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ── prGetCurrentPrompt ──
   Called by sidebar on load and refresh.
   Reads the active row in Pipeline Runner sheet.
   Returns the prompt text and paste target info.
*/
function prGetCurrentPrompt() {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss.getSheetByName('Pipeline Runner');
    if (!sh) return { prompt: null, stage: 'Pipeline Runner sheet not found' };

    var activeRow = sh.getActiveRange() ? sh.getActiveRange().getRow() : 0;
    if (activeRow < 1) return { prompt: null, stage: 'Select a row in Pipeline Runner' };

    // Read the label column (A) and content column (D) of active row
    var label   = String(sh.getRange(activeRow, 1).getValue() || '').trim();
    var content = String(sh.getRange(activeRow, 4).getValue() || '').trim();

    // Only return if this is a PROMPT row
    if (label !== 'PROMPT' && label !== 'PART 1' && label !== 'PART 2' &&
        label !== 'PART 3' && label !== 'FIX PROMPT' && label !== 'ALT/CAPTION PROMPT' &&
        label !== 'LOCATION CONTEXT PROMPT') {
      return { prompt: null, stage: 'Select a PROMPT row (label column = PROMPT)' };
    }

    if (!content) return { prompt: null, stage: 'Prompt not yet generated — use Pipeline menu first' };

    // Determine paste target — which row to write output back to
    // Convention: paste area is always the next row after the prompt
    var pasteRow   = activeRow + 1;
    var pasteLabel = String(sh.getRange(pasteRow, 1).getValue() || '').trim();

    // Only set paste target if next row is a paste area
    var pasteTarget = null;
    if (pasteLabel === 'PASTE OUTPUT' || pasteLabel === 'PASTE HTML OUTPUT' ||
        pasteLabel === 'PASTE LLM OUTPUT' || pasteLabel === 'PASTE CHATGPT OUTPUT' ||
        pasteLabel === 'PASTE SECTION PLAN' || pasteLabel === 'PASTE AUDIT OUTPUT' ||
        pasteLabel === 'PASTE OUTPUT' || pasteLabel === 'LLM RESPONSE') {
      pasteTarget = { sheet: 'Pipeline Runner', row: pasteRow, col: 4 };
    }

    // Determine stage name from surrounding header
    var stageName = label;
    for (var r = activeRow; r >= Math.max(1, activeRow - 5); r--) {
      var cellBg = sh.getRange(r, 1).getBackground();
      var cellVal = String(sh.getRange(r, 1).getValue() || '').trim();
      if (cellBg === '#1565c0' || cellBg === '#1a237e' || cellBg === '#00695c') {
        stageName = cellVal.split('|')[0].trim();
        break;
      }
    }

    return {
      prompt:      content,
      stage:       stageName,
      pasteTarget: pasteTarget,
      promptRow:   activeRow
    };

  } catch(e) {
    return { prompt: null, stage: 'Error: ' + e.toString() };
  }
}

/* ── prSubmitOutput ──
   Writes paste output to the target cell in Pipeline Runner.
   Then triggers the appropriate push function based on the row label.
*/
function prSubmitOutput(pasteTarget, output) {
  try {
    if (!pasteTarget || !output) {
      return { success: false, message: 'No paste target or empty output' };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sh    = ss.getSheetByName(pasteTarget.sheet);
    if (!sh) return { success: false, message: pasteTarget.sheet + ' sheet not found' };

    // Write to paste cell
    var cell = sh.getRange(pasteTarget.row, pasteTarget.col, 1,
                           sh.getMaxColumns() - pasteTarget.col + 1);
    // Find the merged extent
    sh.getRange(pasteTarget.row, pasteTarget.col).setValue(output);

    // Check label of paste row to determine if auto-push is needed
    var pasteLabel = String(sh.getRange(pasteTarget.row, 1).getValue() || '').trim();

    var autoMsg = '';
    if (pasteLabel === 'PASTE OUTPUT' || pasteLabel === 'PASTE CHATGPT OUTPUT') {
      // Check what stage this is by looking at nearby header
      for (var r = pasteTarget.row; r >= Math.max(1, pasteTarget.row - 8); r--) {
        var hdrVal = String(sh.getRange(r, 1).getValue() || '').trim();
        if (hdrVal.indexOf('W0') === 0 && hdrVal.indexOf('FORENSIC') > -1) {
          // Auto-push W0 fields
          var res = pushStage0FieldsToActiveRow(output);
          autoMsg = res.success ? ' | W0 fields pushed to posts sheet' : ' | W0 push failed: ' + res.message;
          break;
        }
        if (hdrVal.indexOf('W5') === 0 && hdrVal.indexOf('GOVERNANCE') > -1) {
          // Auto-push governance fields
          var gRes = pushGovernanceFieldsToActiveRow ? pushGovernanceFieldsToActiveRow(output) : null;
          autoMsg = gRes && gRes.success ? ' | Governance fields pushed' : ' | Governance push failed';
          break;
        }
      }
    }

    if (pasteLabel === 'PASTE SECTION PLAN') {
      var sRes = saveSectionPlan(output);
      autoMsg = ' | Section plan saved to posts sheet';
    }

    if (pasteLabel === 'PASTE HTML OUTPUT') {
      // Store in runner only — user runs W2B Save HTML separately
      autoMsg = ' | HTML stored — run Pipeline → W2B Save HTML to Posts when ready';
    }

    return {
      success: true,
      message: 'Output saved to row ' + pasteTarget.row + autoMsg
    };

  } catch(e) {
    return { success: false, message: e.toString() };
  }
}
function openW7Sidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ce_Sidebar_W7')
    .setTitle('W7 — Video Prompt')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ============================================================
   W2C — GENERATE IMAGE SUGGESTION PROMPT
============================================================ */
function prGenerateW2C() {
  var sh  = getPRSheet();
  prSetStatus(sh, 98, 'RUNNING');
  var res = buildImageSuggestionPrompt ? buildImageSuggestionPrompt() : { success: false, message: 'buildImageSuggestionPrompt not found' };
  if (!res || !res.success) {
    prWriteRow(sh, 99, 'PROMPT', '❌ ' + (res ? res.message : 'Function not found'), PR_COLOURS.fail, true);
    prSetStatus(sh, 98, 'FAIL');
    return;
  }
  prWriteRow(sh, 99, 'PROMPT', res.prompt, PR_COLOURS.prompt, true);
  sh.setRowHeight(99, 200);
  prSetStatus(sh, 98, 'RUNNING');
}

/* ============================================================
   W2C — SAVE SUGGESTIONS FROM RUNNER ROW 100 TO POSTS COLUMN FK (166)
============================================================ */
function prSaveW2C() {
  var sh  = getPRSheet();
  var raw = String(sh.getRange(100, PR_COL_CONTENT).getValue() || '').trim();
  if (!raw) {
    SpreadsheetApp.getUi().alert('Paste the ChatGPT JSON output into row 100 first.');
    return;
  }

  // Strip markdown code fences if present
  var cleaned = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/, '').trim();

  // Validate JSON before saving
  try {
    JSON.parse(cleaned);
  } catch (e) {
    prWriteRow(sh, 100, 'PASTE LLM OUTPUT', '❌ Invalid JSON — ' + e.toString() + '\n\nOriginal paste:\n' + raw, PR_COLOURS.fail, true);
    prSetStatus(sh, 98, 'FAIL');
    return;
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();
  posts.getRange(row, 166).setValue(cleaned); // FK — Image Suggestions (JSON)

  prWriteRow(sh, 100, 'PASTE LLM OUTPUT', '✔ Saved to column FK (166)\n\n' + cleaned, PR_COLOURS.result, true);
  prSetStatus(sh, 98, 'DONE');
}