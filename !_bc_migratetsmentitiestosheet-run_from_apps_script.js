function bc_migrateTSMEntitiesToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const techSheet = ss.getSheetByName("technical");
  if (!techSheet) throw new Error("technical sheet not found.");

  // Create or clear the target sheet
  let entitiesSheet = ss.getSheetByName("technical_entities");
  if (!entitiesSheet) {
    entitiesSheet = ss.insertSheet("technical_entities");
  } else {
    entitiesSheet.clearContents();
  }

  // Write headers
  entitiesSheet.getRange(1, 1, 1, 5).setValues([[
    "Material", "Entity", "Type", "Primary Application", "Semantic Fingerprint Vocabulary"
  ]]);

  const techData = techSheet.getDataRange().getValues();
  const rows = [];
  let entityCount = 0;
  let materialCount = 0;

  for (let i = 0; i < techData.length; i++) {
    const material = String(techData[i][1] || "").trim();
    const tsmText  = String(techData[i][2] || "").trim();

    if (!material || !tsmText) continue;
    if (material.toLowerCase() === "material") continue; // skip header row

    const startMarker = "RECOGNISED TECHNICAL ENTITIES";
    const endMarker   = "--- END RECOGNISED TECHNICAL ENTITIES ---";

    const startIdx = tsmText.indexOf(startMarker);
    const endIdx   = tsmText.indexOf(endMarker);

    if (startIdx === -1 || endIdx === -1) {
      Logger.log("No entity section found for: " + material);
      continue;
    }

    const entitySection = tsmText.substring(startIdx + startMarker.length, endIdx);
    const lines = entitySection.split("\n");

    let materialEntities = 0;

    for (let j = 0; j < lines.length; j++) {
      const line = lines[j].trim();
      if (!line) continue;

     // Skip header row and separator lines
if (line.startsWith("=") || line.startsWith("-") || line.startsWith("9.") || line.startsWith("RECOGNISED")) continue;
if (line.toLowerCase().startsWith("entity") && line.indexOf("|") > -1) {
  const firstPart = line.split("|")[0].trim().toLowerCase();
  if (firstPart === "entity") continue;
}

const parts = line.split("|");
if (parts.length < 2) continue;

      const entity      = String(parts[0] || "").trim();
      const type        = String(parts[1] || "").trim();
      const application = String(parts[2] || "").trim();
      const fingerprint = String(parts[3] || "").trim();

      if (!entity) continue;

      rows.push([material, entity, type, application, fingerprint]);
      materialEntities++;
      entityCount++;
    }

    if (materialEntities > 0) {
      materialCount++;
      Logger.log(material + ": " + materialEntities + " entities extracted.");
    }
  }

  // Write all rows in one operation
  if (rows.length > 0) {
    entitiesSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  Logger.log("Migration complete. " + entityCount + " entities across " + materialCount + " materials.");
}