function onOpen(e) {
  try { DriveApp.getRootFolder(); } catch(e) {}
  try { DocumentApp.getActiveDocument(); } catch(e) {}
  const ui = SpreadsheetApp.getUi();

  // 2. ENTITY GOVERNANCE & STRATEGIC PROTECTION
  ui.createMenu("🚀 Row Analysis")
    .addItem('📋 Content Intelligence', 'openGovernancePromptAssembler')
    .addItem('✂️ Restore lost NewHtml etc', 'restoreActiveRowFromSiteExport')
    .addItem('📋 Batch Build Authority Briefs', 'bc_batchBuildAuthorityBriefs')
    .addSeparator()
    .addItem('🤖 Content Engine (ChatGPT) Stage 1 — ', 'openBuildContentEnginePromptFromActiveRowChatGPTSidebar')
    .addItem('New Article AC', 'da_openNewArticleSidebar')
    .addItem('🛒 Generate Azon Product HTML', 'generateAzonProductHTML')
    .addSeparator()
    .addItem("🪨 Select Stone & Update Benchmarks", "openStoneSelector")
    .addItem('🛡️ Run Strategic Action Evaluation', 'evaluateActiveRowStrategicLogic')
    .addItem('🛠️ Active Row TDF Export For CLAUDE to Analyse Meta, Shema, Make Suggestions', 'openActiveRowTDFExportSidebar')
    .addItem('🤖 Active Row TDF Export For ChatGPT to Analyse Meta, Shema, Make Suggestions', 'openActiveRowTDFExportSidebarGPT')
    .addSeparator()
    .addItem('🎨 Generate New H1 & Meta', 'GENERATE_DYNAMIC_SEO_PROMPT')
    .addItem('🗣️ Run Heading Audit (Active Row)', 'runHeadingAuditFromActiveRow')
    .addItem('🔬 Format Cell Text for better reading (Selection Only)', 'formatAbbeySelection')
    .addToUi();

  // 1. UTILITIES
  ui.createMenu('Utilities')
    .addItem('🚀Transpose Active Row to *posts-active-row* sheet', 'updateActiveRowSheet')
    .addItem('🔀 Filter by Canonical Material…', 'openCanonicalMaterialSidebar')
    .addSeparator()
    .addItem('🔎<h2>,<h3> | Extract Current Headings', 'extractCurrentHeadingsToPosts')
    .addSeparator()
    .addItem('🔄 📅 Calculate Earliest Rewrite', 'updateVisibleRewriteDates')
    .addSeparator()
    .addItem('🗃️ Open Sheets Sidebar', 'openNavigator')
    .addItem('📊 Generate Post Analysis Prompt', 'showPostAnalysisExportSidebar')
    .addSeparator()
    .addItem('🅱️ Convert markdown to bold', 'convertMarkdownBold')
    .addSeparator()
    .addItem('📤 Open Post Exporter Sidebar', 'openPostExporterSidebar')
    .addSeparator()
    .addItem('📝 Open Wordpress Post Update Sidebar', 'openPostUpdateSidebar')
    .addSeparator()
    .addItem('📈 GSC - Copy Row + Headers for Analysis', 'copyRowWithHeadersToClipboard')
    .addSeparator()
    .addItem('Sort Sheet By Entity Role', 'sortByEntityRolePreserveLinks')
    .addSeparator()
    .addItem("🔀 Sort Full Sheet By Stone Type", "openStoneTypeSidebar")
    .addItem("Show All Rows (Clear Stone Filter)", "showAllStoneRows")
    .addSeparator()
    .addItem('Open selected row in viewer sheet', 'openSelectedRowInViewerSheet')
    .addItem('🎨 Map Colored Sheet Index', 'generateColoredSheetList')
    .addItem('⚠️ Search gs files for hard coded columns', 'auditHardcodedColumns')
    .addSeparator()
    .addItem('📊 Open GSC Data Sheet', 'openGSCSheet')
    .addItem('Open Code Bundler', 'codeshowSidebar')
    .addItem('Update Sheet Index', 'generateSheetIndex')
    .addSeparator()
    // ================= GSC INTERACTIVE TRIGGER =================
    .addItem('🔄 Run GSC Dynamic Sync', 'openGSCTimescaleModal')
    .addItem('⚡ Run GSC Full Sync + Row Stats (Active Row)', 'runGSCFullSyncThenRowStats')
    .addItem('🔄 Run GSC Full Sync — All Sheets (Active Row)', 'openGSCFullSyncModal')
    .addItem('📊 Run GSC Row Stats — Active Row', 'openGSCRowStatsModal')
    .addItem('📥 Sync Live Data to... (Select Sheet)', 'syncGSCData')
    // ==========================================================
    .addToUi();
}

function openGSCSheet() {
  var html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:Arial,sans-serif;padding:20px;">' +
    '<h3>GSC Data Sheet</h3>' +
    '<p>This sheet allows you to run Google Search Console data on <strong>one or multiple URLs</strong>.</p>' +
    '<p>Use the <strong>GSC Data</strong> menu in that sheet to:</p>' +
    '<ul>' +
    '<li><strong>Fetch This Row</strong> — runs GSC data for the active row URL</li>' +
    '<li><strong>Fetch All Rows</strong> — runs GSC data for all URLs in column A</li>' +
    '</ul>' +
    '<a href="https://docs.google.com/spreadsheets/d/1em6KwGAsvjjPMQrPw57gN1JKnBCM3furwOgwGcH0ecI/edit?gid=0#gid=0" target="_blank">' +
    '<button style="background:#1a73e8;color:white;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;font-size:14px;">Open GSC Data Sheet</button>' +
    '</a>' +
    '<p style="margin-top:20px;"><small>Click the button above to open the sheet in a new tab.</small></p>' +
    '</body></html>'
  ).setWidth(400).setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, 'GSC Data Sheet');
}

/**
 * Opens the unified Dashboard sidebar
 */
function showDashboard() {
  const html = HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Active Row Analysis')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Wrapper — opens the bc_ Governance Prompt Assembler sidebar.
 */
function openGovernancePromptAssembler() {
  bc_showGovernancePromptAssembler();
}

/**
 * AUTOMATION: Sequence runner for the "Run All" feature
 */
function runFullSequence() {
  try {
    evaluateActiveRowStrategicLogic();
    return "Full Sequence Completed Successfully!";
  } catch (e) {
    return "Error in sequence: " + e.toString();
  }
}

/**
 * DATA BRIDGE: Step 4 (Human Action) needs the combined prompt.
 */
function getHumanActionData() {
  return generateHumanActionPrompt();
}

function openStep1Sidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ce_Step1_Sidebar')
    .setTitle('Step 1 — Structure Audit')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}