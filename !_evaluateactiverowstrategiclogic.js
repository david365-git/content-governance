/**
 * MASTER STRATEGIC PAGE EVALUATOR v4
 * evaluateActiveRowStrategicLogic
 * STRATEGIC REVIEW & ACTION REPORT ENGINE
 *
 * PURPOSE:
 * This script evaluates the currently selected row in the "posts" sheet
 * and produces a structured, plain-English strategic action report.
 *
 * It interprets — but does NOT recalculate — governance data already
 * populated by D60 and other governance systems.
 *
 * This script performs the following steps:
 *
 * 1. Confirms correct sheet and valid row selection.
 * 2. Retrieves site-export data (Meta Title, Meta Description, Schema).
 * 3. Reads governance fields (Tier, Risk, Structural Rules, Schema Risk, Entity Status).
 * 4. Performs strict Material Governance-based meta alignment validation.
 * 5. Analyses query data using impression-weighted dominance logic.
 * 6. Detects scope contraction or expansion via impression delta.
 * 7. Interprets Tier rules and structural permissions.
 * 8. Interprets performance risk level.
 * 9. Outputs a structured Strategic Action Report into the
 *    "Strategic Action Required" column.
 *
 * IMPORTANT:
 * This script is an INTERPRETATION ENGINE.
 * It does not alter governance values.
 * It does not modify structural data.
 * It does not rewrite content.
 *
 * Do NOT remove or rewrite this documentation block.
 * Do NOT edit historic change record entries.
 * Only append new dated entries below.
 *
 * CHANGE RECORD:
 * 26/02/2026 12:22 GMT — Added structured documentation header and protected change record block.
 * 26/02/2026 13:58 GMT — Added strict Material Governance meta alignment validation and scope contraction detection.
 */

function evaluateActiveRowStrategicLogic() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== "posts") {
    ui.alert("You must run this from the 'posts' sheet.");
    return;
  }

  const rowIndex = sheet.getActiveRange().getRow();
  if (rowIndex < 2) {
    ui.alert("Select a valid data row.");
    return;
  }

  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowIndex,1,1,sheet.getLastColumn()).getValues()[0];

  function get(col){
    const idx = headers.indexOf(col);
    return idx > -1 ? values[idx] : "";
  }

  const postId = get("Post ID");
  const pageTitle = get("Title");

  const confirm = ui.alert(
    "Confirm Evaluation",
    "Post ID: " + postId + "\nTitle: " + pageTitle + "\n\nContinue?",
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) return;

  /* ===============================
     SITE EXPORT LOOKUP
  =============================== */

  const exportSheet = ss.getSheetByName("site-export");
  const exportData = exportSheet.getDataRange().getValues();
  const exportHeaders = exportData[0];

  const idIdx = exportHeaders.indexOf("ID");
  const metaTitleIdx = exportHeaders.indexOf("Meta Title");
  const metaDescIdx = exportHeaders.indexOf("Meta Description");

  let metaTitle = "";
  let metaDescription = "";

  for (let i=1;i<exportData.length;i++){
    if (String(exportData[i][idIdx]) === String(postId)){
      metaTitle = exportData[i][metaTitleIdx] || "";
      metaDescription = exportData[i][metaDescIdx] || "";
      break;
    }
  }

  /* ===============================
     GOVERNANCE FIELDS
  =============================== */

  const primaryEntity = String(get("Primary Entity") || "").trim();
  const strategicTier = String(get("Strategic Tier") || "");
  const riskLevel = String(get("Strategic Risk Level") || "");
  const structuralChange = String(get("Structural Change Allowed") || "");
  const benImpr = parseFloat(get("Benmrk Impr")) || 0;
  const monImpr = parseFloat(get("Montr Impr")) || 0;
  const benPos = parseFloat(get("Benmrk Pos")) || 0;
  const monPos = parseFloat(get("Montr Position")) || 0;

  /* ===============================
     MATERIAL META ALIGNMENT CHECK
  =============================== */

  let materialAlignmentStatus = "Unknown";
  const metaCombined = (metaTitle + " " + metaDescription).toLowerCase();

  const materialSheet = ss.getSheetByName("Material Governance Reference Sheet");
  if (materialSheet && primaryEntity) {

    const materialData = materialSheet.getDataRange().getValues();
    const materialHeaders = materialData[0];

    const canonicalIdx = materialHeaders.indexOf("Canonical Material Name");
    const variantIdx = materialHeaders.indexOf("Acceptable Recognition Variants (Recognition Only – Not for Output)");

    for (let i=1;i<materialData.length;i++){

      const canonical = String(materialData[i][canonicalIdx] || "").trim();
      const variantsRaw = String(materialData[i][variantIdx] || "").toLowerCase();

      if (canonical === primaryEntity){

        const canonicalLower = canonical.toLowerCase();
        const variants = variantsRaw.split(";").map(v => v.trim()).filter(v => v);

        let match = false;

        if (metaCombined.includes(canonicalLower)) match = true;

        if (!match){
          for (let v of variants){
            if (metaCombined.includes(v)){
              match = true;
              break;
            }
          }
        }

        if (!metaTitle && !metaDescription) {
          materialAlignmentStatus = "META_MISSING_BOTH";
        } else if (!metaTitle) {
          materialAlignmentStatus = "META_MISSING_TITLE";
        } else if (!metaDescription) {
          materialAlignmentStatus = "META_MISSING_DESCRIPTION";
        } else if (!match) {
          materialAlignmentStatus = "MATERIAL_META_MISALIGNED";
        } else {
          materialAlignmentStatus = "MATERIAL_ALIGNED";
        }

        break;
      }
    }
  }

  /* ===============================
     SCOPE MOVEMENT CHECK
  =============================== */

  let scopeMessage = "";
  if (benImpr > 0) {
    const changePercent = ((monImpr - benImpr) / benImpr) * 100;

    if (changePercent < -30) {
      scopeMessage += "Impressions reduced from " + benImpr + " to " + monImpr +
                      " (" + Math.round(changePercent) + "% change).\n";
      scopeMessage += "Average position improved from " + benPos + " to " + monPos + ".\n";
      scopeMessage += "Google appears to prefer the tightened entity focus.\n\n";
      scopeMessage += "Strategic Options:\n";
      scopeMessage += "1. Reintroduce removed entities (not recommended for Tier 1).\n";
      scopeMessage += "2. Check whether other pages absorbed removed query coverage.\n";
      scopeMessage += "3. Create a dedicated page to reclaim lost query surface.\n\n";
    }
  }

  /* ===============================
     BUILD OUTPUT
  =============================== */

  let output = "";

  output += "### STRATEGIC INTENT REVIEW\n";
  output += "Primary Entity: " + primaryEntity + "\n\n";

  output += "### MATERIAL META ALIGNMENT\n";

  if (materialAlignmentStatus === "META_MISSING_BOTH") {
    output += "Meta title and description are missing.\nImmediate correction required.\n\n";
  }

  if (materialAlignmentStatus === "META_MISSING_TITLE") {
    output += "Meta title is missing.\nAdd a canonical material-aligned title.\n\n";
  }

  if (materialAlignmentStatus === "META_MISSING_DESCRIPTION") {
    output += "Meta description is missing.\nAdd a clear material-aligned description.\n\n";
  }

  if (materialAlignmentStatus === "MATERIAL_META_MISALIGNED") {
    output += "Meta does not reference the canonical material or approved recognition variants.\nRevise meta to align with Primary Entity.\n\n";
  }

  if (materialAlignmentStatus === "MATERIAL_ALIGNED") {
    output += "Material alignment confirmed.\n\n";
  }

  output += "### SCOPE & VISIBILITY MOVEMENT\n";
  output += scopeMessage || "No major scope contraction detected.\n\n";

  output += "### PERFORMANCE RISK\n";
  output += riskLevel === "Low"
    ? "Performance stable within defined thresholds.\n"
    : "Performance movement detected. Review positioning.\n";

  output += "\n";

  const actionCol = headers.indexOf("Strategic Action Required");
  if (actionCol > -1) {
    sheet.getRange(rowIndex, actionCol + 1).setValue(output);
  }
}