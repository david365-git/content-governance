function bc_batchBuildAuthorityBriefs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("posts");
  if (!sheet) throw new Error("posts sheet not found.");

  // Get material from active row
  const activeRow = sheet.getActiveRange().getRow();
  if (activeRow < 2) throw new Error("Select a data row in the posts sheet first.");

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).trim());

  const colIndex = name => {
    const idx = headers.indexOf(name);
    if (idx === -1) throw new Error("Column not found: " + name);
    return idx;
  };

  const stoneTypeIdx      = colIndex("Stone Type");
  const articleTypeIdx    = colIndex("Article Type");
  const primaryClusterIdx = colIndex("Primary Query Cluster Owned");
  const localityIdx       = headers.indexOf("Locality");
  const problemAngleIdx   = colIndex("Problem Angle");
  const authorityBriefIdx = colIndex("Authority Brief");

  const activeMaterial = String(sheet.getRange(activeRow, stoneTypeIdx + 1).getValue() || "").trim();
  if (!activeMaterial) throw new Error("Stone Type is empty on the active row — select a row with a material first.");

  // Confirmation dialog
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Run Authority Brief Batch",
    "This will build Problem Angle and Authority Brief for all " + activeMaterial + " rows that have completed entity governance and do not yet have a Problem Angle set.\n\nRun only after entity governance is complete for " + activeMaterial + ".\n\nProceed?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ANGLE_MAP = {
    "Hub Page":           "WHY",
    "Diagnostic Guide":   "ID",
    "Method Guide":       "FIX",
    "Service Page":       "FIX",
    "Geo Service Page":   "LOCAL",
    "Case Study":         "PROVE",
    "Maintenance Entity": "PREVENT"
  };

  const DECIDE_SIGNALS = ["worth", "should", "cost", "value", "decide", "whether", "is it", "how much"];
  const WHY_SIGNALS    = ["why", "what", "how", "explain", "understand", "behaviour", "behavior", "causes"];
  const CHOOSE_SIGNALS = ["who", "find", "hire", "best", "choose", "contractor", "recommend", "select"];

  const ANGLE_DEFS = {
    "WHY":     { label: "Material Explanation",       readerQ: "Why is this happening to my floor?",                      template: "explains why [PROBLEM] occurs at material level, so you understand what you are actually dealing with" },
    "ID":      { label: "Diagnosis",                  readerQ: "Which specific condition do I have?",                     template: "helps you identify which specific condition is causing [PROBLEM], so you choose the right solution" },
    "FIX":     { label: "Method / Process",           readerQ: "How is this professionally corrected?",                   template: "explains how [PROBLEM] is professionally corrected, stage by stage" },
    "PREVENT": { label: "Maintenance",                readerQ: "How do I stop this happening again?",                     template: "explains how to prevent [PROBLEM] recurring through correct ongoing maintenance" },
    "CHOOSE":  { label: "Contractor Selection",       readerQ: "Who should I hire and how do I evaluate them?",           template: "helps you find and evaluate the right specialist to correct [PROBLEM]" },
    "LOCAL":   { label: "Geo Service",                readerQ: "Can someone fix this near me?",                           template: "offers professional correction of [PROBLEM] for homeowners in [LOCALITY]" },
    "PROVE":   { label: "Case Study / Proof",         readerQ: "Has this actually been fixed for someone like me?",       template: "shows how [PROBLEM] was corrected on a real floor, with before and after evidence" },
    "DECIDE":  { label: "Value / Comparison",         readerQ: "Is professional intervention worth it for my situation?", template: "helps you decide whether professional intervention is the right choice for [PROBLEM]" }
  };

  const data = sheet.getDataRange().getValues();
  let written = 0;
  let skipped = 0;
  let flagged = 0;

  for (let i = 1; i < data.length; i++) {
    const articleType    = String(data[i][articleTypeIdx] || "").trim();
    const primaryCluster = String(data[i][primaryClusterIdx] || "").trim().toLowerCase();
    const locality       = localityIdx > -1 ? String(data[i][localityIdx] || "").trim() : "";
    const existingAngle  = String(data[i][problemAngleIdx] || "").trim();
    const rowMaterial    = String(data[i][stoneTypeIdx] || "").trim();

    if (!articleType) continue;
    if (rowMaterial.toLowerCase() !== activeMaterial.toLowerCase()) continue;
    if (existingAngle) { skipped++; continue; }

    // Determine angle code
    let angleCode = ANGLE_MAP[articleType] || "";

    if (!angleCode) {
      if (articleType === "Educational Guide") {
        const hasDecide = DECIDE_SIGNALS.some(w => primaryCluster.indexOf(w) > -1);
        const hasWhy    = WHY_SIGNALS.some(w => primaryCluster.indexOf(w) > -1);
        if (hasDecide && !hasWhy) {
          angleCode = "DECIDE";
        } else if (hasWhy || !hasDecide) {
          angleCode = "WHY";
        } else {
          Logger.log("Row " + (i + 1) + " — Educational Guide: ambiguous angle — flagged for manual selection");
          sheet.getRange(i + 1, problemAngleIdx + 1).setValue("MANUAL SELECTION REQUIRED — Educational Guide: WHY or DECIDE");
          flagged++;
          continue;
        }
      } else if (articleType === "Buyer Guide") {
        const hasDecide = DECIDE_SIGNALS.some(w => primaryCluster.indexOf(w) > -1);
        const hasChoose = CHOOSE_SIGNALS.some(w => primaryCluster.indexOf(w) > -1);
        if (hasDecide && !hasChoose) {
          angleCode = "DECIDE";
        } else if (hasChoose || !hasDecide) {
          angleCode = "CHOOSE";
        } else {
          Logger.log("Row " + (i + 1) + " — Buyer Guide: ambiguous angle — flagged for manual selection");
          sheet.getRange(i + 1, problemAngleIdx + 1).setValue("MANUAL SELECTION REQUIRED — Buyer Guide: CHOOSE or DECIDE");
          flagged++;
          continue;
        }
      } else {
        Logger.log("Row " + (i + 1) + " — Unknown article type: " + articleType + " — skipped");
        continue;
      }
    }

    const def = ANGLE_DEFS[angleCode];
    if (!def) continue;

    // Build angle statement
    const problem = String(data[i][primaryClusterIdx] || "this problem").trim();
    const loc     = locality || "your area";
    const stmt    = "This article " + def.template
      .replace("[PROBLEM]",  problem)
      .replace("[LOCALITY]", loc) + ".";

    const fullBlock = angleCode + " — " + def.label + "\n→ " + def.readerQ + "\n" + stmt;

    // Write Problem Angle
    sheet.getRange(i + 1, problemAngleIdx + 1).setNumberFormat("@").setValue(fullBlock);

    // Build and write Authority Brief
    sheet.setActiveRange(sheet.getRange(i + 1, 1));
    const result = buildAuthorityBrief();
    if (result.success) {
      sheet.getRange(i + 1, authorityBriefIdx + 1).setNumberFormat("@").setValue(result.brief);
      written++;
    } else {
      Logger.log("Row " + (i + 1) + " — Authority Brief failed: " + result.message);
    }
  }

  Logger.log("bc_batchBuildAuthorityBriefs complete. Written: " + written + " | Skipped: " + skipped + " | Flagged for manual: " + flagged);
}
