/**
 * ============================================================
 * bc_GovernancePromptAssembler_Dashboard.gs
 * Abbey Floor Care — Governance Prompt Assembler
 * Sidebar Launcher
 *
 * Prefix: bc_ (loads before ce_ files in Apps Script editor)
 * Version: 2.0 — Three-prompt staged generation
 * Depends on: bc_GovernancePromptAssembler_Config.gs
 * ============================================================
 */

function bc_showGovernancePromptAssembler() {
  checkAndShowPipelineUpdateNotice();

  const html = HtmlService
    .createTemplateFromFile('bc_GovernancePromptAssembler_Dashboard-html')
    .evaluate()
    .setTitle('Governance Prompt Assembler')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(600)
    .setHeight(900);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Content Intelligence');
}

/**
 * ONE-TIME NOTICE: reminds the user to re-run P3, W0, and SERP Entity
 * Bridge on all unpublished articles following a script update. Shows
 * once per Content Intelligence / Content Engine open until dismissed
 * via dismissPipelineUpdateNotice(), after which it never shows again.
 */
function checkAndShowPipelineUpdateNotice() {
  var props = PropertiesService.getScriptProperties();
  var dismissed = props.getProperty('PIPELINE_UPDATE_NOTICE_DISMISSED');
  if (dismissed === 'true') return;

  var html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:Arial,sans-serif;padding:20px;background:#fff3e0;">' +
    '<h3 style="color:#e65100;margin-top:0">⚠ Script Update — Action Required</h3>' +
    '<p>You have to rerun <strong>P3</strong>, <strong>W0</strong>, and <strong>SERP Entity Bridge</strong> again on all articles that are not yet published, owing to a script update.</p>' +
    '<button onclick="if(confirm(\'Are you sure? All unpublished rows must already be re-run through P3, W0, and SERP Entity Bridge. This notice will not show again once dismissed.\')){google.script.run.withSuccessHandler(function(){google.script.host.close();}).dismissPipelineUpdateNotice();}" ' +
    'style="background:#e65100;color:white;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;font-size:14px;margin-top:10px;">' +
    "I've updated all rows — dismiss this notice</button>" +
    '</body></html>'
  ).setWidth(420).setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, 'Pipeline Update Notice');
}

/**
 * Called by the dismiss button in the notice above — permanently stops
 * the notice from showing again.
 */
function dismissPipelineUpdateNotice() {
  PropertiesService.getScriptProperties().setProperty('PIPELINE_UPDATE_NOTICE_DISMISSED', 'true');
  return { success: true };
}

/**
 * TEMPORARY: run this once from the Apps Script editor (select the
 * function name in the dropdown, then click Run) to re-arm the notice.
 * Delete this function afterwards.
 */
function resetPipelineUpdateNotice_TEMP() {
  PropertiesService.getScriptProperties().deleteProperty('PIPELINE_UPDATE_NOTICE_DISMISSED');
}