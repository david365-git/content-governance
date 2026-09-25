/**
 * GLOBAL CONFIGURATION
 *
 * D60 GOVERNANCE CONTROLLER — DATA PACKAGER
 *
 * This script performs the following steps:
 * 1. Validates Benchmark and Monitor GSC data presence.
 * 2. Confirms user update before execution.
 * 3. Calculates performance deltas (CTR, Position, Impressions).
 * 4. Applies threshold detection logic.
 * 5. Prioritises CTR and Impression signals over Position movement.
 * 6. Packages Strategic Matrix, Schema Validation, and Entity Mapping data.
 * 7. Injects mapping data into D60 prompt at FINAL OUTPUT PROTOCOL anchor.
 * 8. Appends GSC, Governance, Performance Delta, and Site Export context.
 * 9. Returns a structured Master Prompt data package.
 *
 * IMPORTANT:
 * Do NOT modify or remove this summary section.
 * Do NOT rewrite historic change record entries.
 * Only append new entries when required.
 *
 * CHANGE RECORD:
 * 26/02/2026 10:32 GMT — Added Performance Delta Logic (Monitor vs Benchmark with threshold detection + reason output)
 * 26/02/2026 11:58 GMT — Adjusted delta logic: CTR + Impressions override Position.
 */

const CONFIG = {
  POSTS_GID: 835712004,
  prompts: "Prompts",
  matrix: "Strategic Protection Threshold Matrix Sheet",
  schema: "Schema Alignment Validation Sheet",
  export: "site-export"
};

/**
 * Opens sidebar
 */
function openD60GovernanceSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('d60-sidebar-html')
    .setTitle('D60 Governance Controller')
    .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Helper: Get sheet by GID
 */
function getSheetByGid(gid) {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.find(s => s.getSheetId() == gid);
}

/**
 * MAIN DATA PACKAGER
 */
function getCustomDataPackage() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // =============================
  // POSTS SHEET
  // =============================
  const postsSheet = getSheetByGid(CONFIG.POSTS_GID);
  if (!postsSheet) {
    return { isError: true, message: "Posts sheet not found." };
  }

  const activeRowIndex = postsSheet.getActiveSelection().getRow();
  if (activeRowIndex < 2) {
    return { isError: true, message: "Select a data row (Row 2+) in the Posts sheet." };
  }

  const fullData = postsSheet.getDataRange().getValues();
  const rawHeaders = fullData[0];
  const headers = rawHeaders.map(h => String(h).trim().toLowerCase());
  const activeRowData = fullData[activeRowIndex - 1];

  const getVal = (headerName) => {
    const idx = headers.indexOf(headerName.trim().toLowerCase());
    return idx !== -1 ? activeRowData[idx] : "";
  };

  /* ===============================
     PRE-RUN GSC DATA CONFIRMATION
  =============================== */

  const benDate = String(getVal("Benmrk Date") || "").trim();
  const benTime = String(getVal("Benmrk Timescale") || "").trim();
  const monDate = String(getVal("Monitor Date") || "").trim();
  const monTime = String(getVal("Monitoe Timescale") || "").trim();

  if (!benDate || !monDate) {
    ui.alert(
      "GSC Data Missing",
      "Benchmark or Monitor Date is empty.\n\nUpdate GSC data before running D60.",
      ui.ButtonSet.OK
    );
    return { isError: true, message: "Missing Benchmark or Monitor date." };
  }

  let warningMessage = "";

  if (benDate === monDate) {
    warningMessage = "\n\n⚠ WARNING: Benchmark and Monitor dates are identical.\nYou may be comparing the same window.";
  }

  const confirmRun = ui.alert(
    "Confirm GSC Data Update",
    "Benchmark:\n" +
    benDate + " (" + benTime + " month window)\n\n" +
    "Monitor:\n" +
    monDate + " (" + monTime + " month window)\n" +
    warningMessage +
    "\n\nHave you updated these values from GSC?",
    ui.ButtonSet.OK_CANCEL
  );

  if (confirmRun !== ui.Button.OK) {
    return { isError: true, message: "User cancelled D60 run." };
  }

  /* ===============================
     PERFORMANCE DELTA LOGIC
     CTR + IMPRESSIONS OVERRIDE POSITION
  =============================== */

  const parseNumber = (v) => parseFloat(String(v).replace(/[% ,]/g, "")) || 0;

  const benCTR = parseNumber(getVal("Benmrk CTR"));
  const monCTR = parseNumber(getVal("Montr CTR"));

  const benPos = parseNumber(getVal("Benmrk Pos"));
  const monPos = parseNumber(getVal("Montr Position"));

  const benImpr = parseNumber(getVal("Benmrk Impr"));
  const monImpr = parseNumber(getVal("Montr Impr"));

  let deltaStatus = "Stable";
  let deltaReason = "No material change detected.";

  const ctrDiff = monCTR - benCTR;
  const posDiff = monPos - benPos;
  const imprDiffPercent = benImpr > 0 ? ((monImpr - benImpr) / benImpr) * 100 : 0;

  const CTR_THRESHOLD = 1.0;
  const POS_THRESHOLD = 3;
  const IMPR_THRESHOLD = 15;

  const ctrTrigger = Math.abs(ctrDiff) > CTR_THRESHOLD;
  const imprTrigger = Math.abs(imprDiffPercent) > IMPR_THRESHOLD;

  if (ctrTrigger) {

    if (ctrDiff > 0) {
      deltaStatus = "Improved";
      deltaReason = `CTR increased by ${ctrDiff.toFixed(2)}%.`;
    } else {
      deltaStatus = "Declined";
      deltaReason = `CTR decreased by ${Math.abs(ctrDiff).toFixed(2)}%.`;
    }

  } else if (imprTrigger) {

    if (imprDiffPercent > 0) {
      deltaStatus = "Improved";
      deltaReason = `Impressions increased by ${imprDiffPercent.toFixed(1)}%.`;
    } else {
      deltaStatus = "Declined";
      deltaReason = `Impressions decreased by ${Math.abs(imprDiffPercent).toFixed(1)}%.`;
    }

  } else if (Math.abs(posDiff) > POS_THRESHOLD) {

    if (posDiff < 0) {
      deltaStatus = "Improved";
      deltaReason = `Average position improved by ${Math.abs(posDiff).toFixed(1)} places.`;
    } else {
      deltaStatus = "Declined";
      deltaReason = `Average position dropped by ${posDiff.toFixed(1)} places.`;
    }
  }

const deltaBlock = `
--- PERFORMANCE DELTA ANALYSIS ---
Status: ${deltaStatus}
Reason: ${deltaReason}
(Comparison based on Monitor vs. Benchmark columns)
`;

  // =============================
  // MATRIX DATA
  // =============================
  const matrixSheet = ss.getSheetByName(CONFIG.matrix);
  let matrixDataText = "### STRATEGIC PROTECTION THRESHOLD MATRIX ###\n";
  if (matrixSheet && matrixSheet.getLastRow() > 0) {
    const matrixValues = matrixSheet.getDataRange().getValues();
    matrixDataText += matrixValues.map(row => row.join(",")).join("\n");
  }

  // =============================
  // SCHEMA ALIGNMENT DATA
  // =============================
  const schemaSheet = ss.getSheetByName(CONFIG.schema);
  let schemaDataText = "\n### SCHEMA ALIGNMENT VALIDATION ###\n";
  if (schemaSheet && schemaSheet.getLastRow() > 0) {
    const schemaValues = schemaSheet.getDataRange().getValues();
    schemaDataText += schemaValues.map(row => row.join(",")).join("\n");
  }

  // =============================
  // ENTITY CONSTRAINT MAPPING DATA
  // =============================
  const mappingSheetName = "Entity Constraint Mapping Sheet";
  const mappingSheet = ss.getSheetByName(mappingSheetName);
  if (!mappingSheet) {
    throw "Entity Constraint Mapping Sheet not found.";
  }

  const mappingValues = mappingSheet.getDataRange().getValues();
  const mappingHeaders = mappingValues[0] || [];

  if (!mappingHeaders.includes("Entity Governance Status Prefix") || !mappingHeaders.includes("Entity Constraint Status")) {
    throw "Entity Constraint Mapping Sheet missing required headers.";
  }

  let mappingDataText = "### ENTITY CONSTRAINT MAPPING SHEET ###\n";
  mappingDataText += mappingValues.map(row => row.join(",")).join("\n");

  // =============================
  // D60 PROMPT
  // =============================
  const pSheet = ss.getSheetByName(CONFIG.prompts);
  let d60Prompt = pSheet ? String(pSheet.getRange("D60").getValue()) : "Prompt D60 not found.";

  const anchor = "FINAL OUTPUT PROTOCOL";
  if (d60Prompt.indexOf(anchor) === -1) {
    throw "FINAL OUTPUT PROTOCOL anchor not found.";
  }

  const promptParts = d60Prompt.split(anchor);
  d60Prompt = promptParts[0] + mappingDataText + "\n" + anchor + promptParts[1];

 // =============================
  // GSC DATA (COMPARATIVE D60 VIEW)
  // =============================
  const gscData = `--- GSC PERFORMANCE COMPARISON ---
Title: ${getVal("Title") || getVal("Post Title") || "Unknown"}

[MONITOR WINDOW (Current Situation)]
Date: ${monDate} (${monTime} months)
Clicks: ${getVal("Montr Clicks") || 0}
Impr: ${getVal("Montr Impr") || 0}
CTR: ${getVal("Montr CTR") || "0%"}
Pos: ${getVal("Montr Position") || 0}
Queries (Monitor):
${getVal("Monitor Queries") || "None"}

[BENCHMARK WINDOW (Historical Reference)]
Date: ${benDate} (${benTime} months)
Clicks: ${getVal("Benmrk Clicks") || 0}
Impr: ${getVal("Benmrk Impr") || 0}
CTR: ${getVal("Benmrk CTR") || "0%"}
Pos: ${getVal("Benmrk Pos") || 0}
Queries (Benchmark):
${getVal("Benmrk Queries") || "None"}`;

  // =============================
  // GOVERNANCE FIELDS
  // =============================
  const governanceBlock = `
--- GOVERNANCE FIELDS ---
Article Type: ${getVal("Article Type") || "Unknown"}
Drift Status: ${getVal("Drift Status") || "Unknown"}
Entity Governance Status: ${getVal("Entity Governance Status") || "Unknown"}
Feeds Hub: ${getVal("Feeds Hub") || "Unknown"}
`;

  // =============================
  // SITE EXPORT
  // =============================
  const exportSheet = ss.getSheetByName(CONFIG.export);
  let customExportText = "\n--- SITE EXPORT DETAILS ---";

  if (exportSheet) {
    const exportData = exportSheet.getDataRange().getValues();
    const exportHeaders = exportData[0];

    const postId = String(getVal("Post ID") || "").trim();
    const matchRow = postId
      ? exportData.find(row => String(row[0]).trim() === postId)
      : null;

    if (matchRow) {
      const cols = [0, 1, 2, 3, 4, 9, 10, 11, 12, 14];
      const hRow = "Headers: " + cols.map(i => exportHeaders[i] || "").join(" | ");
      const vRow = "Values:  " + cols.map(i => matchRow[i] || "").join(" | ");
      customExportText += "\n" + hRow + "\n" + vRow;
    } else {
      customExportText += "\n[No matching Post ID found in site-export sheet]";
    }
  }

  const finalPrompt =
    matrixDataText +
    "\n" +
    schemaDataText +
    "\n\n### INSTRUCTIONAL PROMPT ###\n" +
    d60Prompt +
    "\n\n### STRATEGIC CONTEXT & DATA TABLES ###\n" +
    gscData +
    governanceBlock +
    deltaBlock +
    customExportText +
    "\n\n--- END OF DATA PACKAGE ---";

  return {
    isError: false,
    masterPrompt: finalPrompt
  };
}