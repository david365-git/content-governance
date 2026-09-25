/* ============================================================
 ce_Utilities.gs
============================================================ */

function determineTierFromSheet(material, impressions, tierLevelsSheet) {
  const tierData = tierLevelsSheet.getDataRange().getValues();
  const tierHeaders = tierData[0];
  const stoneColIndex = tierHeaders.indexOf(material);
  if (stoneColIndex === -1) return { tier: 0, label: "UNKNOWN - Stone col not found." };
  const tierLabels = tierData.slice(1,5).map(row => row[0]);
  const rangeRows  = tierData.slice(1,5).map(row => row[stoneColIndex]);
  for (let i = 0; i < rangeRows.length; i++) {
    const cell  = String(rangeRows[i]);
    const match = cell.match(/RANGE:\s*([^\|]+)/i);
    if (!match) continue;
    const rangeText = match[1].trim();
    if (rangeText.startsWith(">")) {
      if (impressions > Number(rangeText.replace(/[^\d]/g,""))) return { tier: i+1, label: tierLabels[i] };
    } else if (rangeText.startsWith("<")) {
      if (impressions < Number(rangeText.replace(/[^\d]/g,""))) return { tier: i+1, label: tierLabels[i] };
    } else if (rangeText.includes("-")) {
      const p = rangeText.split("-");
      if (impressions >= Number(p[0].replace(/[^\d]/g,"")) && impressions <= Number(p[1].replace(/[^\d]/g,""))) return { tier: i+1, label: tierLabels[i] };
    }
  }
  return { tier: 0, label: "NO MATCH" };
}

function getCleaningLimitNotes(matrixSheet) {
  const data = matrixSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "Tier 1" && data[i][1] === "Cleaning Limit & Self-Diagnostic Threshold") {
      return `\nMANDATORY CLEANING LIMIT:\n${data[i][2]}\n`;
    }
  }
  return "";
}

/* ============================================================
   DATA MAPPING & DIAGNOSTICS (CORE ENGINE)
============================================================ */

function getActiveRowDataMap() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("posts");
  const row   = sheet.getActiveRange().getRow();
  if (row < 2) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const vals    = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  let map = {};
  headers.forEach((h, i) => map[h] = vals[i]);
  return map;
}
function getAcValuesForActiveRow() {
  try {
    var d = getActiveRowDataMap();
    var material = String(d["Stone Type"] || "").trim();
    if (!material) return { success: false, message: "No Stone Type found for active row." };

    var articleType = String(d["Article Type"] || "").trim();
    var acResult = getAcValuesForMaterial(material, articleType);
    var values = acResult.values || {};
    var labels = acResult.labels || {};
    if (!values || Object.keys(values).length === 0) {
      return { success: false, message: "No AC values found for material: " + material };
    }

    var hubDefaults = getHubPageDefaults();
    var existingValues = {};
    var axisKeys = ['ac_entry_condition','ac_homeowner_perception','ac_material_behaviour','ac_constraint','ac_process_emphasis','ac_result_type','ac_narrative_archetype'];
    axisKeys.forEach(function(key) { existingValues[key] = String(d[key] || '').trim(); });
    return { success: true, material: material, articleType: articleType, values: values, labels: labels, hubDefaults: hubDefaults, ac_entry_condition: existingValues.ac_entry_condition, ac_homeowner_perception: existingValues.ac_homeowner_perception, ac_material_behaviour: existingValues.ac_material_behaviour, ac_constraint: existingValues.ac_constraint, ac_process_emphasis: existingValues.ac_process_emphasis, ac_result_type: existingValues.ac_result_type, ac_narrative_archetype: existingValues.ac_narrative_archetype };

  } catch(e) {
    return { success: false, message: "getAcValuesForActiveRow error: " + e.message };
  }
}

function checkAcDuplication(selected) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { isDuplicate: false };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    function getIdx(name) { return headers.indexOf(name); }

    var colMap = {
      ac_entry_condition:      getIdx("ac_entry_condition"),
      ac_homeowner_perception: getIdx("ac_homeowner_perception"),
      ac_material_behaviour:   getIdx("ac_material_behaviour"),
      ac_constraint:           getIdx("ac_constraint"),
      ac_process_emphasis:     getIdx("ac_process_emphasis"),
      ac_result_type:          getIdx("ac_result_type"),
      ac_narrative_archetype:  getIdx("ac_narrative_archetype")
    };

    var stoneTypeIdx = getIdx("Stone Type");
    var activeStone  = String(sheet.getRange(row, stoneTypeIdx + 1).getValue()).trim().toLowerCase();

    var lastRow = sheet.getLastRow();
    var allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    for (var i = 0; i < allData.length; i++) {
      var checkRow = i + 2;
      if (checkRow === row) continue;

      var rowStone = String(allData[i][stoneTypeIdx] || "").trim().toLowerCase();
      if (rowStone !== activeStone) continue;

      var matchCount   = 0;
      var matchingCols = [];

      Object.keys(colMap).forEach(function(col) {
        var idx = colMap[col];
        if (idx === -1) return;
        var existingVal = String(allData[i][idx] || "").trim().toLowerCase();
        var selectedVal = String(selected[col]   || "").trim().toLowerCase();
        if (existingVal && selectedVal && existingVal === selectedVal) {
          matchCount++;
          matchingCols.push(col);
        }
      });

      if (matchCount >= 4) {
        var dupRowValues = {};
        Object.keys(colMap).forEach(function(col) {
          var idx = colMap[col];
          if (idx > -1) {
            dupRowValues[col] = String(allData[i][idx] || "").trim();
          }
        });
        return {
          isDuplicate:  true,
          matchCount:   matchCount,
          matchedRow:   checkRow,
          matchingCols: matchingCols,
          dupRowValues: dupRowValues
        };
      }
    }

    return { isDuplicate: false };

  } catch(e) {
    Logger.log("checkAcDuplication error: " + e.message);
    return { isDuplicate: false };
  }
}

function saveAcValuesToSheet(selected) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var written = [];
    var missing = [];
    var gaps    = [];

    Object.keys(selected).forEach(function(col) {
      var idx = headers.indexOf(col);
      if (idx === -1) {
        missing.push(col);
        return;
      }
      var entry = selected[col];
      var valueToSave = "";

      // Support both plain string values and {value, gap, suggested} objects
      if (typeof entry === "object" && entry !== null) {
        if (entry.gap) {
          // No good existing match — use the suggested wording so the
          // field is never left blank, and flag it for manual review.
          valueToSave = entry.suggested_option
            ? "[GAP — REVIEW] " + entry.suggested_option
            : "[GAP — NO SUGGESTION PROVIDED]";
          gaps.push(col);
        } else {
          valueToSave = entry.selected || entry.value || "";
        }
      } else {
        valueToSave = entry;
      }

      var cell = sheet.getRange(row, idx + 1);
      cell.setNumberFormat("@");
      cell.setValue(valueToSave);
      written.push(col);
    });

    var message = "Saved " + written.length + "/7 axes to row " + row + ".";
    if (missing.length > 0) message += " Columns not found: " + missing.join(", ");
    if (gaps.length > 0) message += " GAPS FLAGGED (review before publishing): " + gaps.join(", ");

    if (written.length > 0) logPipelineResume("AC — Article Classification", "");
    return { success: written.length > 0, message: message, gaps: gaps };

  } catch(e) {
    return { success: false, message: "saveAcValuesToSheet error: " + e.message };
  }
}

function getActiveRowDiagnosticInfo() {
  const d = getActiveRowDataMap();
  return { postID: d["Post ID"] || "N/A", title: d["Stone Type"] || "No Material" };
}

function getActiveRowUrl() {
  const d = getActiveRowDataMap();
  return {
    url:         String(d["URL"] || d["Canonical URL"] || "").trim(),
    postID:      String(d["Post ID"] || "").trim(),
    title:       String(d["Title"] || "").trim(),
    articleType: String(d["Article Type"] || "").trim()
  };
}

function cleanHtmlForLLM(html) {
  if (!html) return "";
  return html.replace(/<span[^>]*class="textannotation[^>]*>(.*?)<\/span>/gis, "$1")
             .replace(/<p>\s*<\/p>/gi, "").replace(/\n\s*\n/g, "\n").trim();
}

function buildArticleTypeControlBlockGPT(ss) {
  const sheet = ss.getSheetByName("Article Type Control Sheet");
  if (!sheet) return "CONTROL SHEET NOT FOUND.";
  const data  = sheet.getDataRange().getValues();
  let block   = "ARTICLE TYPE CONTROL LOCK:\n";
  for (let i = 1; i < data.length; i++) {
    block += `\n${i}. ${data[i][0]}\nFocus: ${data[i][1]}\nExpectations: ${data[i][3]}\n`;
  }
  return block;
}

/* ============================================================
   TECH DNA
============================================================ */

function getTechDNA(material) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const dnaSheet = ss.getSheetByName("Tech_DNA");
  if (!dnaSheet) return null;
  const data    = dnaSheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const rowMaterial = String(data[i][0]).trim().toLowerCase();
    if (rowMaterial === material.trim().toLowerCase()) {
      let map = {};
      headers.forEach((h, j) => map[String(h).trim()] = data[i][j]);
      return map;
    }
  }
  return null;
}

/* ============================================================
   SCHEMA ALIGNMENT VALIDATION SHEET READER
   Reads the Schema Alignment Validation Sheet for the given
   article type and returns a structured rules block for the
   Stage 2C schema generation prompt.
============================================================ */
function getSchemaAlignmentRules(articleType) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Schema Alignment Validation Sheet");
  if (!sheet) return "Schema Alignment Validation Sheet not found.";

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h).trim(); });

  // Find the row matching this article type
  let matchRow = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === articleType.trim().toLowerCase()) {
      matchRow = data[i];
      break;
    }
  }

  if (!matchRow) {
    return "No schema rules found for article type: " + articleType + ". Default to Article schema.";
  }

  // Build rules block from matched row
  var block = "--- SCHEMA ALIGNMENT RULES FOR: " + articleType + " ---\n";
  headers.forEach(function(header, idx) {
    if (header && matchRow[idx] !== undefined && String(matchRow[idx]).trim() !== "") {
      block += header + ": " + String(matchRow[idx]).trim() + "\n";
    }
  });

  return block;
}


/* ============================================================
   GET STONE PROMPT RULES
   Reads the PROMPT GOVERNANCE RULES section from the TSM
   for the active stone type.
   Returns a structured block for Stage 2B injection.
   Falls back to generic language if section not found.
============================================================ */
function getStonePromptRules(material) {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var techSheet = ss.getSheetByName("technical");
    if (!techSheet) return null;

    var techData = techSheet.getRange("B2:C100").getValues();
    var tsmContent = "";

    // Find the TSM content block for this material
    for (var i = 0; i < techData.length; i++) {
      var materialList = String(techData[i][0]).trim();
      if (materialList.toLowerCase().indexOf(material.trim().toLowerCase()) > -1) {
        tsmContent = String(techData[i][1] || "").trim();
        break;
      }
    }

    if (!tsmContent) return null;

    // Find PROMPT GOVERNANCE RULES section within TSM content
    var sectionStart = tsmContent.indexOf("PROMPT GOVERNANCE RULES");
    if (sectionStart === -1) return null;

    var sectionText = tsmContent.substring(sectionStart);

    // Parse key-value pairs from section
    var rules = {
      surfaceFinishDistinction:  "",
      processChemistryNote:      "",
      prohibitedGeneralisations: [],
      mechanicalPolishEligible:  "No",
      topicalFinishTypes:        [],
      sealersUsed:               ""
    };

    function extractField(text, fieldName) {
      var re = new RegExp(fieldName + ":\s*([^\n]+)", "i");
      var m = text.match(re);
      return m ? m[1].trim() : "";
    }

    rules.surfaceFinishDistinction  = extractField(sectionText, "Surface Finish Distinction");
    rules.processChemistryNote      = extractField(sectionText, "Process Chemistry Note");
    rules.mechanicalPolishEligible  = extractField(sectionText, "Mechanical Polish Eligible");
    rules.sealersUsed               = extractField(sectionText, "Sealers Used");

    var prohib = extractField(sectionText, "Prohibited Generalisations");
    rules.prohibitedGeneralisations = prohib
      ? prohib.split(",").map(function(s) { return s.trim(); }).filter(Boolean)
      : [];

    var topical = extractField(sectionText, "Topical Finish Types");
    rules.topicalFinishTypes = topical
      ? topical.split(",").map(function(s) { return s.trim(); }).filter(Boolean)
      : [];

    return rules;

  } catch(e) {
    return null;
  }
}

/* ============================================================
   BUILD STONE PROMPT RULES BLOCK
   Formats getStonePromptRules() output as a prompt-ready block
   for injection into Stage 2B part2.
============================================================ */
function buildStonePromptRulesBlock(material) {
  var rules = getStonePromptRules(material);

  if (!rules) {
    return "------------------------------------------------------------\n" +
           "STONE TYPE RULES: No Prompt Governance Rules found in TSM for " + material + ".\n" +
           "Apply general UK stone care principles.\n" +
           "------------------------------------------------------------\n";
  }

  var block = "";
  block += "------------------------------------------------------------\n";
  block += "STONE-SPECIFIC PROMPT RULES (from TSM — Hard Lock):\n";
  block += "These rules override any generic process descriptions.\n\n";

  if (rules.surfaceFinishDistinction) {
    block += "SURFACE FINISH DISTINCTION (use this exact framing):\n";
    block += rules.surfaceFinishDistinction + "\n\n";
  }

  if (rules.processChemistryNote) {
    block += "PROCESS CHEMISTRY (one sentence — include in mechanism explanation):\n";
    block += rules.processChemistryNote + "\n\n";
  }

  if (rules.mechanicalPolishEligible) {
    block += "MECHANICAL POLISH ELIGIBLE: " + rules.mechanicalPolishEligible + "\n";
    if (rules.mechanicalPolishEligible.toLowerCase() === "no") {
      block += "Do NOT describe mechanical polishing as an outcome for this stone type.\n";
      block += "Do NOT use the phrase mirror-like finish or mirror-like clarity.\n";
    }
    block += "\n";
  }

  if (rules.topicalFinishTypes && rules.topicalFinishTypes.length > 0) {
    block += "APPLICABLE TOPICAL FINISHES: " + rules.topicalFinishTypes.join(", ") + "\n\n";
  }

  if (rules.sealersUsed) {
    block += "SEALER GUIDANCE FOR THIS STONE:\n";
    block += rules.sealersUsed + "\n\n";
  }

  if (rules.prohibitedGeneralisations && rules.prohibitedGeneralisations.length > 0) {
    block += "PROHIBITED TERMS (must not appear in this article):\n";
    block += rules.prohibitedGeneralisations.join(", ") + "\n";
    block += "If any of the above terms appear in your output — remove them before responding.\n";
  }

  block += "------------------------------------------------------------\n";
  return block;
}

/* ============================================================
   BUILD SEMANTIC FINGERPRINT BLOCK
   Extracts entities with semantic fingerprints from TSM and
   formats them for injection into Stage 2B prompt.
   Implements LSI entity translation rule from IR textbook.
============================================================ */
function buildSemanticFingerprintBlock(material) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var entitiesSheet = ss.getSheetByName("technical_entities");
    if (!entitiesSheet) return "";

    var data = entitiesSheet.getDataRange().getValues();
    if (data.length < 2) return "";

    var headers = data[0].map(function(h) { return String(h).trim(); });
    var materialIdx    = headers.indexOf("Material");
    var entityIdx      = headers.indexOf("Entity");
    var typeIdx        = headers.indexOf("Type");
    var appIdx         = headers.indexOf("Primary Application");
    var fingerprintIdx = headers.indexOf("Semantic Fingerprint Vocabulary");

    if (materialIdx === -1 || entityIdx === -1 || fingerprintIdx === -1) return "";

    var materialLower = material.trim().toLowerCase();
    var entities = [];

    for (var i = 1; i < data.length; i++) {
      var rowMaterial = String(data[i][materialIdx] || "").trim().toLowerCase();
      if (rowMaterial !== materialLower) continue;

      var entityName  = String(data[i][entityIdx]      || "").trim();
      var entityType  = typeIdx      > -1 ? String(data[i][typeIdx]      || "").trim() : "";
      var primaryApp  = appIdx       > -1 ? String(data[i][appIdx]       || "").trim() : "";
      var fingerprint = String(data[i][fingerprintIdx] || "").trim();

      if (entityName && fingerprint) {
        entities.push({
          name:        entityName,
          type:        entityType,
          application: primaryApp,
          fingerprint: fingerprint
        });
      }
    }

    // If no entities with fingerprints found, return empty
    if (entities.length === 0) return "";

    // Build the injection block
    var block = "";
    block += "------------------------------------------------------------\n";
    block += "LSI ENTITY TRANSLATION RULE (Hard Lock — Semantic Fingerprint)\n";
    block += "------------------------------------------------------------\n";
    block += "Each governed entity MUST be surrounded by its required semantic vocabulary.\n";
    block += "Entity name appears 1-2x explicitly across the entire article.\n";
    block += "All other references use semantic fingerprint vocabulary.\n\n";
    block += "DISTRIBUTION PATTERN:\n";
    block += "- Opening section: 2-3 fingerprint instances\n";
    block += "- Main content: 3-5 fingerprint instances\n";
    block += "- Conclusion: 2-3 fingerprint instances\n";
    block += "- Target: 8-12 total fingerprint term instances per entity across article\n\n";
    block += "GOVERNED ENTITIES FOR THIS MATERIAL (" + material + "):\n\n";

    // Get Recovery Blueprint entities for priority ordering
    var blueprintEntities = bc_getRecoveryBlueprintEntities();
    var priorityNames = blueprintEntities ? (blueprintEntities.priority || []) : [];
    var supportingNames = blueprintEntities ? (blueprintEntities.supporting || []) : [];

    // Sort entities by priority tier
    var priorityEntities  = [];
    var supportingEntities = [];
    var otherEntities     = [];

    for (var j = 0; j < entities.length; j++) {
      var e = entities[j];
      var nameLower = e.name.toLowerCase();
      var isPriority   = priorityNames.some(function(p)  { return p.toLowerCase() === nameLower; });
      var isSupporting = supportingNames.some(function(s) { return s.toLowerCase() === nameLower; });
      if (isPriority)        priorityEntities.push(e);
      else if (isSupporting) supportingEntities.push(e);
      else                   otherEntities.push(e);
    }

    var orderedEntities = priorityEntities.concat(supportingEntities).concat(otherEntities);

    // Add each entity's fingerprint in priority order
    for (var k = 0; k < orderedEntities.length; k++) {
      var e = orderedEntities[k];
      var tier = priorityEntities.indexOf(e) > -1 ? " [PRIORITY]" :
                 supportingEntities.indexOf(e) > -1 ? " [SUPPORTING]" : "";
      block += "ENTITY: " + e.name + tier + "\n";
      block += "REQUIRED VOCABULARY: " + e.fingerprint + "\n\n";
    }

    block += "CRITICAL RULES:\n";
    block += "- Entity names must appear in LOWERCASE in body text (e.g., 'capillary action', not 'Capillary Action')\n";
    block += "- Capitalize only if proper noun or sentence start\n";
    block += "- Use fingerprint vocabulary naturally — no keyword stuffing\n";
    block += "- Fingerprint terms should appear throughout the article, not clustered in one section\n";
    block += "- Every governed entity in the Recovery Blueprint MUST use its fingerprint vocabulary\n";
    block += "- COMMA LIST PROHIBITION (Hard Lock): Never place fingerprint vocabulary terms in a\n";
    block += "  comma-separated list with no verb between items. A sentence reading\n";
    block += "  'term1, term2, term3, term4, term5 help explain...' is a CRITICAL FAILURE.\n";
    block += "  Each fingerprint term must appear inside a sentence that contains a verb\n";
    block += "  explaining what that term does, causes, prevents or produces.\n";
    block += "  WRONG: 'Layered structure, mineral planes, cleavage separation, delamination risk\n";
    block += "  and structural characteristic help explain why a visible break must be read carefully.'\n";
    block += "  CORRECT: 'Slate forms in distinct mineral layers, and those cleavage planes mean\n";
    block += "  a surface crack can originate from the tile itself, from subfloor movement, or from\n";
    block += "  a weak boundary between layers — each cause points to a different repair decision.'\n";
    block += "------------------------------------------------------------\n";
    block += "\n";
    block += "PRE-OUTPUT COMMA LIST SCAN (Hard Lock — run this before writing your first token):\n";
    block += "Before outputting any HTML, scan every paragraph you have written.\n";
    block += "For each <p> block, apply this test:\n";
    block += "Count the number of commas in the paragraph.\n";
    block += "If a paragraph contains 4 or more commas, check whether a verb appears between each comma-separated segment.\n";
    block += "If any segment between two consecutive commas contains no verb — that paragraph is a COMMA LIST FAILURE.\n";
    block += "REWRITE the failing paragraph before outputting.\n";
    block += "A paragraph that reads as a list of nouns or technical terms joined by commas is always a failure regardless of how many words it contains.\n";
    block += "Do not output until every paragraph passes this test.\n";
    block += "------------------------------------------------------------\n";

    return block;

  } catch(e) {
    Logger.log("buildSemanticFingerprintBlock error: " + e.message);
    return "";
  }
}

function bc_getRecoveryBlueprintEntities() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return null;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idx     = headers.indexOf("Recovery Blueprint");
    if (idx === -1) return null;

    var raw = String(sheet.getRange(row, idx + 1).getValue() || "").trim();
    if (!raw || raw.indexOf("RECOVERY BLUEPRINT MISSING") > -1) return null;

    var result = { priority: [], supporting: [], peripheral: [], gscGap: [] };

    var lines = raw.split(/\r?\n/);
    var currentTier = null;

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;

      if (line.indexOf("PRIORITY ENTITIES") > -1)   { currentTier = "priority";   continue; }
      if (line.indexOf("SUPPORTING ENTITIES") > -1)  { currentTier = "supporting";  continue; }
      if (line.indexOf("PERIPHERAL ENTITIES") > -1)  { currentTier = "peripheral";  continue; }
      if (line.indexOf("GSC GAP ENTITIES") > -1)     { currentTier = "gscGap";      continue; }

      if (currentTier && line.indexOf(":") === -1) {
        var names = line.split(",").map(function(n) { return n.trim(); }).filter(Boolean);
        names.forEach(function(name) {
          if (result[currentTier].indexOf(name) === -1) result[currentTier].push(name);
        });
      }
    }

    return result;

  } catch(e) {
    Logger.log("bc_getRecoveryBlueprintEntities error: " + e.message);
    return null;
  }
}
function bc_getSerpValidationEntities() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return null;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idx     = headers.indexOf("SERP Validation");
    if (idx === -1) return null;

    var raw = String(sheet.getRange(row, idx + 1).getValue() || "").trim();
    if (!raw || raw === "NONE — no new entities found") return null;

    var entities = [];

    var firstLine = raw.split(/\r?\n/)[0].trim();
    if (firstLine && firstLine.indexOf("|") === -1 && firstLine.indexOf("ENTITY:") === -1) {
      entities = firstLine.split(",").map(function(n) { return n.trim(); }).filter(Boolean);
    }

    if (entities.length === 0) {
      var lines = raw.split(/\r?\n/);
      for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (!line.startsWith("ENTITY:")) continue;
        var nameMatch = line.match(/ENTITY:\s*([^|]+)/i);
        if (nameMatch) entities.push(nameMatch[1].trim());
      }
    }

    return entities.length > 0 ? entities : null;

  } catch(e) {
    Logger.log("bc_getSerpValidationEntities error: " + e.message);
    return null;
  }
}

function bc_getTfidfTerms() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return null;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idx     = headers.indexOf("tfidf_terms");
    if (idx === -1) return null;

    var raw = String(sheet.getRange(row, idx + 1).getValue() || "").trim();
    if (!raw) return null;

    var terms = [];
    var parts = raw.split(",");
    for (var i = 0; i < parts.length; i++) {
      var pair = parts[i].trim();
      if (!pair) continue;
      var colon = pair.indexOf(":");
      if (colon === -1) {
        terms.push({ term: pair, score: 0 });
      } else {
        var term  = pair.substring(0, colon).trim();
        var score = parseFloat(pair.substring(colon + 1).trim()) || 0;
        if (term) terms.push({ term: term, score: score });
      }
    }

    return terms.length > 0 ? terms : null;

  } catch(e) {
    Logger.log("bc_getTfidfTerms error: " + e.message);
    return null;
  }
}

function bc_getRecallGapTerms() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return null;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idx     = headers.indexOf("recall_notes");
    if (idx === -1) return null;

    var raw = String(sheet.getRange(row, idx + 1).getValue() || "").trim();
    if (!raw) return null;

    var gapTerms = [];

    var notInMatch = raw.match(/Not in PQC:\s*(.+)/i);
    if (notInMatch) {
      gapTerms = notInMatch[1].split(",").map(function(t) { return t.trim(); }).filter(Boolean);
    }

    return gapTerms.length > 0 ? gapTerms : null;

  } catch(e) {
    Logger.log("bc_getRecallGapTerms error: " + e.message);
    return null;
  }
}

function bc_getCompetitorVocabGap() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return null;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idx     = headers.indexOf("Competitor Vocab Gap");
    if (idx === -1) return null;

    var raw = String(sheet.getRange(row, idx + 1).getValue() || "").trim();
    if (!raw) return null;

    var terms = raw.split(",").map(function(t) { return t.trim(); }).filter(Boolean);
    return terms.length > 0 ? terms : null;

  } catch(e) {
    Logger.log("bc_getCompetitorVocabGap error: " + e.message);
    return null;
  }
}

function bc_getCompetitorWordCount() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return null;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idx     = headers.indexOf("Competitor Word Count");
    if (idx === -1) return null;

    var raw = String(sheet.getRange(row, idx + 1).getValue() || "").trim();
    if (!raw) return null;

    var count = parseInt(raw);
    return isNaN(count) ? null : count;

  } catch(e) {
    Logger.log("bc_getCompetitorWordCount error: " + e.message);
    return null;
  }
}

function getTechnicalMatrix(material) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const techSheet = ss.getSheetByName("technical");
  if (!techSheet) return "No Technical Matrix Available.";
  const techData  = techSheet.getRange("B2:C100").getValues();
  for (let i = 0; i < techData.length; i++) {
    const materialList  = String(techData[i][0]).trim();
    const matrixContent = techData[i][1];
    if (materialList.toLowerCase().includes(material.trim().toLowerCase())) {
      return matrixContent;
    }
  }
  return "No specific technical matrix found for " + material + ". Use general stone care principles.";
}

/* ============================================================
   UTILITIES: REHYDRATION & PUSHING
============================================================ */

/* ============================================================
   GET MATERIAL LIST
   Reads col A from Material Governance Reference Sheet.
   Returns array of material names for the image library picker.
============================================================ */
function getMaterialList() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Material Governance Reference Sheet");
    if (!sheet) return { success: false, message: "Material Governance Reference Sheet not found." };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var materials = [];
    data.forEach(function(row) {
      var name = String(row[0] || "").trim();
      if (name) materials.push(name);
    });
    return { success: true, materials: materials };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/* ============================================================
   IMAGE LIBRARY
   Bulk extracts first 3 non-decorative images from col 98 HTML
   for every row in the posts sheet and writes to cols DA-DI:
     col 105 (DA) — img1_src
     col 106 (DB) — img1_alt
     col 107 (DC) — img1_caption
     col 108 (DD) — img2_src
     col 109 (DE) — img2_alt
     col 110 (DF) — img2_caption
     col 111 (DG) — img3_src
     col 112 (DH) — img3_alt
     col 113 (DI) — img3_caption
   Run once before processing articles. Re-run any time to refresh.
============================================================ */

function buildImageLibrary(stoneType) {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var sheet    = ss.getSheetByName("posts");
    var lastRow  = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No data rows found." };

    stoneType = String(stoneType || "").trim();
    if (!stoneType) {
      return { success: false, message: "No stone type provided." };
    }

    var processed = 0;
    var skipped   = 0;
    var errors    = [];

    for (var row = 2; row <= lastRow; row++) {
      try {
        // Only process rows matching the active stone type
        var rowStone = String(sheet.getRange(row, 7).getValue() || "").trim();
        if (rowStone.toLowerCase() !== stoneType.toLowerCase()) continue;

        var html = String(sheet.getRange(row, 98).getValue() || "").trim();
        if (!html) { skipped++; continue; }

        var images = extractImagesForLibrary(html);

        // Write up to 3 images into DA-DI (cols 105-113)
        for (var i = 0; i < 3; i++) {
          var baseCol = 105 + (i * 3); // DA=105, DD=108, DG=111
          if (i < images.length) {
            var img = images[i];
            sheet.getRange(row, baseCol).setValue(img.src);
            sheet.getRange(row, baseCol + 1).setValue(img.alt);
            sheet.getRange(row, baseCol + 2).setValue(img.caption);
          } else {
            sheet.getRange(row, baseCol).setValue("");
            sheet.getRange(row, baseCol + 1).setValue("");
            sheet.getRange(row, baseCol + 2).setValue("");
          }
        }
        processed++;

        // Flush every 20 rows to avoid timeout
        if (processed % 20 === 0) {
          SpreadsheetApp.flush();
        }
      } catch(rowErr) {
        errors.push("Row " + row + ": " + rowErr.toString());
      }
    }

    SpreadsheetApp.flush();
    var msg = stoneType + " image library built — " + processed + " rows processed, " +
              skipped + " skipped (no HTML).";
    if (errors.length > 0) msg += " Errors: " + errors.slice(0, 3).join(" | ");
    return { success: true, message: msg, processed: processed, skipped: skipped, stoneType: stoneType };

  } catch(e) {
    return { success: false, message: "BUILD ERROR: " + e.toString() };
  }
}

function extractImagesForLibrary(html) {
  var images = [];
  var seen   = [];

  // Decorative pattern — skip these src URLs
  var decorativeRe = /logo|icon|avatar|sprite|pixel|tracking|banner|badge/i;

  function addImage(src, alt, caption) {
    src = src.trim();
    alt = (alt || "").trim();
    caption = (caption || "").trim();
    if (!src || seen.indexOf(src) > -1) return;
    if (decorativeRe.test(src)) return;
    seen.push(src);
    images.push({ src: src, alt: alt, caption: caption });
  }

  // ── Pass 1: <figure> blocks — preferred format ──
  var figRe = /<figure[\s\S]*?<\/figure>/gi;
  var m;
  while ((m = figRe.exec(html)) !== null && images.length < 3) {
    var block = m[0];
    var srcM   = block.match(/src="([^"]+)"/i);
    var altM   = block.match(/alt="([^"]*)"/i);
    var capM   = block.match(/<figcaption[^>]*>([\s\S]*?)<\/figcaption>/i);
    var caption = capM ? capM[1].replace(/<[^>]+>/g, "").trim() : "";
    if (srcM) addImage(srcM[1], altM ? altM[1] : "", caption);
  }

  // ── Pass 2: image-ninja-pro-wrapper divs ──
  var ninjaRe = /<div[^>]*image-ninja-pro-wrapper[^>]*>[\s\S]*?<\/div>/gi;
  while ((m = ninjaRe.exec(html)) !== null && images.length < 3) {
    var block = m[0];
    var srcM  = block.match(/src="([^"]+)"/i);
    var altM  = block.match(/alt="([^"]*)"/i);
    if (srcM) addImage(srcM[1], altM ? altM[1] : "", "");
  }

  // ── Pass 3: bare <img> tags not already captured ──
  var imgRe = /<img[^>]+>/gi;
  while ((m = imgRe.exec(html)) !== null && images.length < 3) {
    var tag  = m[0];
    var srcM = tag.match(/src="([^"]+)"/i);
    var altM = tag.match(/alt="([^"]*)"/i);
    if (srcM) addImage(srcM[1], altM ? altM[1] : "", "");
  }

  return images;
}

function getImageSuggestions(stoneType, currentRow) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var lastRow = sheet.getLastRow();
    var suggestions = [];

    for (var row = 2; row <= lastRow; row++) {
      if (row === currentRow) continue; // skip current article
      var rowStone = String(sheet.getRange(row, 7).getValue() || "").trim().toLowerCase();
      if (rowStone !== stoneType.trim().toLowerCase()) continue;

      for (var i = 0; i < 3; i++) {
        var baseCol = 105 + (i * 3);
        var src = String(sheet.getRange(row, baseCol).getValue() || "").trim();
        var alt = String(sheet.getRange(row, baseCol + 1).getValue() || "").trim();
        var cap = String(sheet.getRange(row, baseCol + 2).getValue() || "").trim();
        if (src && alt && alt.length > 15) {
          suggestions.push({ src: src, alt: alt, caption: cap, postRow: row });
        }
      }
      if (suggestions.length >= 6) break; // cap at 6 candidates
    }

    return suggestions;
  } catch(e) {
    return [];
  }
}

/* ============================================================
   FIGURE INVENTORY
   Extracts all <figure> blocks from HTML, stores them in col 99 (CU)
   as a JSON array. Injected into Stage 2B as a hard-coded paste list
   so the LLM cannot omit or alter any figure.

   Col 99 (CU) — Figure Inventory (JSON)
============================================================ */

function extractFigureInventory(html) {
  // Extracts minimal inventory: src, alt, caption only
  var figures = [];
  var seen    = [];

  if (!html) return figures;

  var IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.avif'];
  var SKIP_PATTERNS    = ['logo', 'icon', 'avatar', 'sprite', 'pixel', 'tracking'];

  function shouldSkip(src) {
    var lower = src.toLowerCase();
    // Skip non-site images — Amazon product images and any other external domains
    if (lower.indexOf('abbeyfloorcare.co.uk') === -1 && lower.indexOf('http') === 0) return true;
    for (var s = 0; s < SKIP_PATTERNS.length; s++) {
      if (lower.indexOf(SKIP_PATTERNS[s]) !== -1) return true;
    }
    if (/\-\d+x\d+(@\d+x)?\.(jpg|jpeg|png|webp|gif|avif)$/i.test(src)) return true;
    return false;
  }

  function extractAttr(context, attrName) {
    var search = attrName + '="';
    var start  = context.toLowerCase().indexOf(search.toLowerCase());
    if (start === -1) return "";
    start += search.length;
    var end = context.indexOf('"', start);
    return end !== -1 ? context.substring(start, end).trim() : "";
  }

  function extractCaption(context, imgEnd) {
    var afterImg = context.substring(imgEnd).replace(/^\s*\/?>?\s*/, "");
    var stopTag  = afterImg.search(/<|<\/|\[\/caption\]/i);
    var cap      = stopTag > 0 ? afterImg.substring(0, stopTag).trim() : "";
    var figCapM  = context.match(/<figcaption[^>]*>([\s\S]*?)<\/figcaption>/i);
    if (figCapM) cap = figCapM[1].replace(/<[^>]+>/g, "").trim();
    return cap;
  }

  var pos = 0;
  while (pos < html.length) {
    var srcDouble = html.indexOf('src="', pos);
    var srcSingle = html.indexOf("src='", pos);
    var srcTag = -1;
    var quoteChar = '"';
    if (srcDouble !== -1 && (srcSingle === -1 || srcDouble <= srcSingle)) {
      srcTag = srcDouble; quoteChar = '"';
    } else if (srcSingle !== -1) {
      srcTag = srcSingle; quoteChar = "'";
    }
    if (srcTag === -1) break;

    var urlStart = srcTag + 5;
    var urlEnd   = html.indexOf(quoteChar, urlStart);
    if (urlEnd === -1) { pos = urlStart; continue; }

    var src   = html.substring(urlStart, urlEnd).trim();
    var lower = src.toLowerCase();

    var isImage = false;
    for (var e = 0; e < IMAGE_EXTENSIONS.length; e++) {
      if (lower.indexOf(IMAGE_EXTENSIONS[e]) !== -1) { isImage = true; break; }
    }

    if (isImage && !shouldSkip(src) && seen.indexOf(src) === -1) {
      seen.push(src);
      var ctxStart  = Math.max(0, srcTag - 600);
      var ctxEnd    = Math.min(html.length, urlEnd + 400);
      var context   = html.substring(ctxStart, ctxEnd);
      var relSrcPos = srcTag - ctxStart;
      if (context.indexOf("<iframe") !== -1) { pos = urlEnd + 1; continue; }
      var alt = extractAttr(context, "alt");
      var imgTagEnd = context.indexOf('>', relSrcPos);
      var caption   = imgTagEnd !== -1 ? extractCaption(context, imgTagEnd + 1) : "";
      figures.push({ src: src, alt: alt, caption: caption });
    }

    pos = urlEnd + 1;
  }

  // ── Safety pass: scan for any image URLs missed by src= search ──
  var domainSearch = 'wp-content/uploads/';
  var scanPos = 0;
  while (scanPos < html.length) {
    var domainIdx = html.indexOf(domainSearch, scanPos);
    if (domainIdx === -1) break;
    var urlEndChars = ['"', "'", ' ', ']', ')', String.fromCharCode(10), String.fromCharCode(13), '<'];
    var urlE = html.length;
    for (var uc = 0; uc < urlEndChars.length; uc++) {
      var idx = html.indexOf(urlEndChars[uc], domainIdx);
      if (idx !== -1 && idx < urlE) urlE = idx;
    }
    var urlS = html.lastIndexOf('http', domainIdx);
    if (urlS === -1 || urlS < domainIdx - 200) { scanPos = domainIdx + 1; continue; }
    var fullUrl = html.substring(urlS, urlE).trim();
    var fullLower = fullUrl.toLowerCase();
    var isImg = false;
    for (var ex = 0; ex < IMAGE_EXTENSIONS.length; ex++) {
      if (fullLower.indexOf(IMAGE_EXTENSIONS[ex]) !== -1) { isImg = true; break; }
    }
    if (isImg && !shouldSkip(fullUrl) && seen.indexOf(fullUrl) === -1) {
      seen.push(fullUrl);
      var ctxS   = Math.max(0, urlS - 600);
      var ctxE   = Math.min(html.length, urlE + 400);
      var ctx    = html.substring(ctxS, ctxE);
      if (ctx.indexOf('<iframe') === -1) {
        var alt     = extractAttr(ctx, 'alt');
        var imgEnd  = ctx.indexOf('>', ctx.indexOf(fullUrl.substring(fullUrl.length - 20)));
        var caption = imgEnd !== -1 ? extractCaption(ctx, imgEnd + 1) : '';
        figures.push({ src: fullUrl, alt: alt, caption: caption });
      }
    }
    scanPos = domainIdx + domainSearch.length;
  }

  // ── Video embeds: find iframes separately ──
  var iframeRe = /<iframe[^>]+src="([^"]+)"[^>]*>[\s\S]*?<\/iframe>/gi;
  var m;
  while ((m = iframeRe.exec(html)) !== null) {
    var vsrc = m[1];
    if (/youtube|vimeo|wistia|loom|embed/i.test(vsrc) && seen.indexOf(vsrc) === -1) {
      seen.push(vsrc);
      figures.push({ src: vsrc, alt: "", caption: "", html: m[0] });
    }
  }

  return figures;
}

/* ============================================================
   STAGE 1.5A AUDIT RESULTS STORAGE
   Stores the Stage 1.5A nine-check audit output in col 99 (CU).
   Col 99 (CU) — Stage 1.5A Audit Results (plain text)
============================================================ */

function saveStage15AAudit(auditText) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };
    var cell = sheet.getRange(row, 99);
    cell.setNumberFormat("@");
    cell.setValue(auditText);
    logPipelineResume("W1.5A — Audit Save", "");
    return { success: true, message: "Stage 1.5A audit saved to row " + row };
  } catch(e) {
    return { success: false, message: "SAVE ERROR: " + e.toString() };
  }
}

function getStage15AAudit() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return null;
    var raw = String(sheet.getRange(row, 99).getValue() || "").trim();
    return raw || null;
  } catch(e) {
    return null;
  }
}

/* ============================================================
   SECTION PLAN STORAGE
   Stores the approved Stage 1.5B section plan in col 100 (CV).
   Col 100 (CV) — Approved Section Plan (plain text)
   NOTE: getSectionPlan() reads from col 103 (CY) first — W1.5E
   Google Optimization JSON — then falls back to col 100 (CV).
============================================================ */

function saveSectionPlan(planText) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };
    var cell = sheet.getRange(row, 100);
    cell.setNumberFormat("@");
    cell.setValue(planText);
    logPipelineResume("W1.5B — Section Plan", "");
    return { success: true, message: "Section plan saved to row " + row };
  } catch(e) {
    return { success: false, message: "SAVE ERROR: " + e.toString() };
  }
}

function storeFigureInventory() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    var originalHtml = "";
    var source = "";

    // PRIMARY: read from col 104 (CZ) — humanised HTML (last pipeline output)
    var col104 = String(sheet.getRange(row, 104).getValue() || "").trim();
    if (col104) {
      originalHtml = col104;
      source = "col 104 (CZ — Humanised HTML)";
    }

    // FALLBACK 1: read from col 98 (CT) — New HTML (pre-humanisation governed output)
    if (!originalHtml) {
      var col98 = String(sheet.getRange(row, 98).getValue() || "").trim();
      if (col98) {
        originalHtml = col98;
        source = "col 98 (CT — New HTML)";
      }
    }

    // FALLBACK 2: read from site-export sheet if both cols are empty
    if (!originalHtml) {
      var postId = String(sheet.getRange(row, 1).getValue()).trim();
      var exportSheet = ss.getSheetByName("site-export");
      if (exportSheet) {
        var exportData = exportSheet.getDataRange().getValues();
        var idIdx   = exportData[0].indexOf("ID");
        var htmlIdx = exportData[0].indexOf("Full Post HTML");
        if (idIdx > -1 && htmlIdx > -1) {
          for (var i = 1; i < exportData.length; i++) {
            if (String(exportData[i][idIdx]).trim() === postId) {
              originalHtml = String(exportData[i][htmlIdx] || "");
              source = "site-export sheet";
              break;
            }
          }
        }
      }
    }

    if (!originalHtml) {
      return { success: false, message: "No HTML found — CZ (104) and CT (98) are empty and post not found in site-export. Run W3 humanisation first." };
    }

    var figures = extractFigureInventory(originalHtml);
    var cell = sheet.getRange(row, 148); // col ER — Figure Inventory (JSON)
    cell.setNumberFormat("@");
    cell.setValue(JSON.stringify(figures));

    return {
      success: true,
      count: figures.length,
      ids: figures.map(function(f) { return f.src; }),
      message: figures.length + " figure(s) stored from " + source + " in col 148 (ER) for row " + row
    };
  } catch(e) {
    return { success: false, message: "FIGURE STORE ERROR: " + e.toString() };
  }
}

function getFigureInventory() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return null;
    var raw = String(sheet.getRange(row, 148).getValue() || "").trim();
    if (!raw) return null;
    return JSON.parse(raw);
  } catch(e) {
    return null;
  }
}

function pushHtmlToActiveRow(cleanHTML) {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheet     = ss.getSheetByName("posts");
    const activeRow = sheet.getActiveRange().getRow();
    if (activeRow < 2) return "ERROR: Select a valid row.";

    // ── BIO BOX INJECTION ─────────────────────────────────────────────
    // Extract BIO_PARAGRAPH: marker and inject complete bio box HTML
    // The LLM only provides the paragraph text — script owns all structure
    var bioParaMatch = cleanHTML.match(/BIO_PARAGRAPH:\s*([\s\S]*?)(?=\s*<footer)/i);
    if (!bioParaMatch) {
      bioParaMatch = cleanHTML.match(/BIO_PARAGRAPH:\s*([\s\S]*?)$/i);
    }
    if (bioParaMatch) {
      var bioParagraph = bioParaMatch[1].trim().replace(/<[^>]+>/g, '').trim();
      var bioBoxHtml =
        '\n<div class="abbey-bio-box" style="border:1px solid #ddd;border-radius:8px;padding:16px;background:#f9f9f9;margin:24px 0;display:flex;gap:16px;align-items:flex-start">\n' +
        '  <div style="flex:0 0 80px">\n' +
        '    <img src="https://www.abbeyfloorcare.co.uk/wp-content/uploads/David_Allen.jpg" alt="David Allen, marble and stone restoration specialist" style="width:80px;height:auto;border-radius:6px" />\n' +
        '  </div>\n' +
        '  <div style="flex:1">\n' +
        '    <p style="margin:0 0 6px 0;font-weight:700">David Allen \u2014 <a href="https://maps.app.goo.gl/W8GSsZUiWoxYPQ1Y6" target="_blank" rel="noopener noreferrer">Abbey Floor Care</a></p>\n' +
        '    <p style="margin:0;font-size:0.95em;color:#444">' + bioParagraph + '</p>\n' +
        '  </div>\n' +
        '</div>\n';
      cleanHTML = cleanHTML.replace(/BIO_PARAGRAPH:[\s\S]*?(?=\s*<footer)/i, bioBoxHtml);
      cleanHTML = cleanHTML.replace(/BIO_PARAGRAPH:[\s\S]*$/i, bioBoxHtml);
    }

    // ── SAFETY NET ────────────────────────────────────────────────────
    // Remove any David_Allen.jpg outside the bio box
    var htmlWithoutBioBox = cleanHTML.replace(/<div[^>]*abbey-bio-box[^>]*>[\s\S]*?<\/div>/i, '');
    if (/David_Allen\.jpg/i.test(htmlWithoutBioBox)) {
      cleanHTML = cleanHTML.replace(/<figure[^>]*>[\s\S]*?David_Allen\.jpg[\s\S]*?<\/figure>/gi, '');
      cleanHTML = cleanHTML.replace(/<img[^>]*David_Allen\.jpg[^>]*\/?>/gi, '');
    }

    // ── BIO BOX GUARANTEE ─────────────────────────────────────────────
    // If no abbey-bio-box is present after injection attempt, build a default one
    var hasBioBox = /<div[^>]*abbey-bio-box/i.test(cleanHTML);
    if (!hasBioBox) {
      try {
        var bioSheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("posts");
        var bioRow     = bioSheet.getActiveRange().getRow();
        var bioHeaders = bioSheet.getRange(1, 1, 1, bioSheet.getLastColumn()).getValues()[0];
        function getBioVal(name) {
          var idx = bioHeaders.indexOf(name);
          return idx > -1 ? String(bioSheet.getRange(bioRow, idx + 1).getValue() || "").trim() : "";
        }
        var bioArticleType = getBioVal("Article Type");
        var bioStoneType   = getBioVal("Stone Type").toLowerCase();
        var bioLocality    = getBioVal("Locality");
        var bioParentArea  = getBioVal("Parent Area");
        var bioLocation    = bioLocality ? bioLocality + (bioParentArea ? ", " + bioParentArea : "") : bioParentArea;

        var defaultBioPara = "";
        if (bioArticleType === "Case Study") {
          defaultBioPara = "David Allen has spent more than 30 years restoring " + bioStoneType + " floors across the UK through Abbey Floor Care" +
            (bioLocation ? ", including projects in " + bioLocation : "") +
            ". This case study documents a real restoration visit where on-site assessment shaped every decision.";
        } else if (bioArticleType === "Geo Service Page") {
          defaultBioPara = "David Allen brings more than 30 years of practical experience restoring " + bioStoneType + " floors across the UK through Abbey Floor Care" +
            (bioLocation ? ", with projects completed across " + bioLocation + " and the surrounding area" : "") +
            ". His guidance reflects hands-on knowledge of local building stock, period floor conditions and the restoration decisions that produce lasting results.";
        } else if (bioArticleType === "Method Guide") {
          defaultBioPara = "David Allen has restored " + bioStoneType + " floors across the UK for over 30 years through Abbey Floor Care. His approach to this method reflects direct experience of how " + bioStoneType + " surfaces respond under real working conditions.";
        } else if (bioArticleType === "Diagnostic Guide") {
          defaultBioPara = "David Allen has diagnosed and corrected " + bioStoneType + " floor conditions across the UK for over 30 years through Abbey Floor Care. His diagnostic approach is grounded in direct observation of how these conditions develop and respond to professional intervention.";
        } else {
          defaultBioPara = "David Allen has worked with " + bioStoneType + " floors across the UK for over 30 years through Abbey Floor Care. His practical experience with material behaviour, restoration sequencing and long-term floor care informs every article published under the Abbey Floor Care name.";
        }

        var defaultBioBox =
          '\n<div class="abbey-bio-box" style="border:1px solid #ddd;border-radius:8px;padding:16px;background:#f9f9f9;margin:24px 0;display:flex;gap:16px;align-items:flex-start">\n' +
          '  <div style="flex:0 0 80px">\n' +
          '    <img src="https://www.abbeyfloorcare.co.uk/wp-content/uploads/David_Allen.jpg" alt="David Allen, marble and stone restoration specialist" style="width:80px;height:auto;border-radius:6px" />\n' +
          '  </div>\n' +
          '  <div style="flex:1">\n' +
          '    <p style="margin:0 0 6px 0;font-weight:700">David Allen \u2014 <a href="https://maps.app.goo.gl/W8GSsZUiWoxYPQ1Y6" target="_blank" rel="noopener noreferrer">Abbey Floor Care</a></p>\n' +
          '    <p style="margin:0;font-size:0.95em;color:#444">' + defaultBioPara + '</p>\n' +
          '  </div>\n' +
          '</div>\n';

        if (/<footer/i.test(cleanHTML)) {
          cleanHTML = cleanHTML.replace(/(<footer)/i, defaultBioBox + '$1');
        } else {
          cleanHTML = cleanHTML + defaultBioBox;
        }
      } catch(bioErr) {
        Logger.log("Bio box guarantee fallback error: " + bioErr.message);
      }
    }


    // col 98 — New HTML
    const cell = sheet.getRange(activeRow, 98);
    cell.setNumberFormat('@STRING@');
    cell.setValue(cleanHTML);
    logPipelineResume("W4 — HTML Push", "");

    // Auto-recalculate TF-IDF and recall score now that New HTML has changed —
    // avoids relying on remembering to click Recalc TF-IDF separately.
    var tfidfResult = bc_recalculateTfIdfForActiveRow();
    var tfidfNote = tfidfResult && tfidfResult.success
      ? " " + tfidfResult.message
      : " (TF-IDF recalculation skipped: " + (tfidfResult ? tfidfResult.message : "unknown error") + ")";

    return "SUCCESS: Content pushed to 'New HTML' in Row " + activeRow + "." + tfidfNote;
  } catch (e) {
    return "PUSH ERROR: " + e.toString();
  }
}

/* ============================================================
   GET HTML FROM ACTIVE ROW
   Reads New HTML column (col 98) for the active row and
   returns it to the sidebar W4 textarea via loadHtmlFromSheet().
============================================================ */
function getHtmlFromActiveRow() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheet     = ss.getSheetByName("posts");
    const activeRow = sheet.getActiveRange().getRow();
    if (activeRow < 2) return "ERROR: Select a valid data row first.";
    // col 98 — New HTML (CV)
    var html = String(sheet.getRange(activeRow, 104).getValue() || "").trim();
    if (!html) html = String(sheet.getRange(activeRow, 98).getValue() || "").trim();
    if (!html) return "ERROR: No HTML found in col 104 (CZ) or col 98 (CT) for Row " + activeRow + ". Push re-hydrated HTML first.";
    return String(html);
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}

/* ============================================================
   STAGE 2C — H1, META & SCHEMA PUSH
   Accepts output from getMetaPrompt() — H1, Meta Title, Meta Description only.
   Schema is now built deterministically by W5C (generateSchemaForActiveRow).
   Supports both old prompt labels (H1:, Title:, Description:)
   and new prompt labels (New H1:, New Meta Title:, New Meta Description:)
   Maps to sheet columns: New H1, New Meta Title, New Meta Description.
============================================================ */
/* ============================================================
   GOVERNANCE FIELD SANITISER
   Fixes common LLM output issues in H1, Meta Title, Meta Description:
   - Adds terminal punctuation where missing
   - Fixes mid-sentence capitals on common stone/service nouns
   - Trims excess whitespace
============================================================ */
function sanitiseGovernanceField(value, fieldName) {
  if (!value) return value;
  var v = value.trim();

  // Strip emoji
  try {
    v = v.replace(/[\u{1F000}-\u{1FFFF}]/gu, '')
         .replace(/[\u{2600}-\u{27BF}]/gu, '')
         .replace(/[\u{1F300}-\u{1F9FF}]/gu, '')
         .replace(/\u2728|\u2605|\u2764/gu, '');
  } catch(e) { /* emoji regex not supported — skip */ }

  // Strip year references e.g. (2024 Guide), (2025 Safe Guide), [2026]
  v = v.replace(/\s*[\(\[]\d{4}[^\)\]]*[\)\]]/g, '');
  // Strip year-only in brackets e.g. "(2026)"
  v = v.replace(/\s*\(\d{4}\)\s*/g, ' ');
  // Strip non-evergreen pipe suffixes e.g. "| Updated", "| 2025"
  v = v.replace(/\s*[-–|]\s*(Updated|New|Latest|Current|\d{4}).*$/i, '');
  // Strip trailing year e.g. "Marble Cleaning Tips 2026"
  v = v.replace(/\s+\d{4}\s*$/, '');
  // Strip UK phone numbers
  v = v.replace(/\s+0[\d\s]{10,15}$/, '');
  // Strip "By Abbey Floor Care" suffix
  v = v.replace(/\s+[Bb]y\s+Abbey\s+Floor\s+Care\s*$/i, '');

  // Fix multiple spaces
  v = v.replace(/  +/g, ' ').trim();
  // Strip redundant closing phrases the LLM adds
  v = v.replace(/\s+(today now|now today|today today|right now today|today\.?\s*now)\.?$/i, '.');

  // ── Field-specific fixes ──

  // H1: remove trailing full stop
  if (fieldName === 'New H1') {
    v = v.replace(/\.+$/, '').trim();
  }

  // Meta Description: ensure terminal punctuation
  if (fieldName === 'New Meta Description') {
    if (!/[.!?]$/.test(v)) {
      v = v + '.';
    }
  }

  // Fix run-on sentences before known sentence-starting words
  if (fieldName === 'New Meta Description' || fieldName === 'New H1') {
    var sentenceStarters = ['Learn', 'Book', 'Understand', 'Find', 'Discover',
                            'See', 'Read', 'Get', 'Know', 'Check', 'Explore',
                            'Arrange', 'Contact', 'Call', 'Visit', 'Ask',
                            'Choose', 'Avoid', 'Prevent', 'Restore', 'Protect',
                            'What', 'How', 'Why', 'When', 'Where', 'Which',
                            "What's", "How's", "Why's"];
    var conditionEndings = ['dull','etched','scratched','stained','damaged','patchy',
                             'cloudy','worn','dirty','faded','cracked','chipped','uneven'];
    sentenceStarters.forEach(function(word) {
      v = v.replace(new RegExp('([a-z]+) (' + word + ' )', 'g'), function(m, before, after) {
        var isQuestion  = /^(What|How|Why|When|Where|Which)/.test(after);
        var isCondition = conditionEndings.indexOf(before.toLowerCase()) > -1;
        return before + ((isQuestion || isCondition) ? '? ' : '. ') + after.trimLeft();
      });
    });
    var agents   = ['restorer','specialist','contractor','technician','expert',
                    'company','professional','team','service','provider','fitter',
                    'installer','cleaner','polisher','tradesperson'];
    var problems = ['dull','etched','scratched','stained','damaged','patchy','cloudy',
                    'worn','dirty','faded','discoloured','cracked','chipped','uneven',
                    'marble','travertine','limestone','slate','granite','terrazzo',
                    'sandstone','flagstone','porcelain','victorian'];
    v = v.replace(/(\w+)\s+(who)\b/gi, function(match, precedingWord, who) {
      var wordLower = precedingWord.toLowerCase();
      if (problems.indexOf(wordLower) > -1 && agents.indexOf(wordLower) === -1) {
        return precedingWord + '? ' + who.charAt(0).toUpperCase() + who.slice(1).toLowerCase();
      }
      return match;
    });
    v = v.replace(/([dtyg]) (what'?s?|how'?s?|why|when|where|which) /gi, function(m, before, qword) {
      return before + '? ' + qword.charAt(0).toUpperCase() + qword.slice(1).toLowerCase() + ' ';
    });
  }

  // Fix mid-sentence capitals on stone/service nouns (Meta Description only)
  if (fieldName === 'New Meta Description') {
    var midCapWords = ['Marble','Travertine','Limestone','Slate','Granite','Terrazzo',
                       'Sandstone','Quartzite','Porcelain','Flagstone','Victorian',
                       'Cleaning','Polishing','Honing','Sealing','Restoration'];
    midCapWords.forEach(function(word) {
      v = v.replace(new RegExp('(?<=[a-z] )' + word + '(?= )', 'g'), word.toLowerCase());
    });
  }

  return v.trim();
}

function pushGovernanceFieldsToActiveRow(rawOutput) {
  try {
    if (!rawOutput || rawOutput.trim() === '') {
      return { success: false, message: 'ERROR: No output provided.' };
    }

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheet     = ss.getSheetByName('posts');
    const activeRow = sheet.getActiveRange().getRow();
    if (activeRow < 2) {
      return { success: false, message: 'ERROR: Select a valid data row first.' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                         .getValues()[0]
                         .map(function(h) { return String(h).trim(); });

    // H1 + Meta only — schema is handled by W5C
    const FIELDS = [
      { labels: ['New H1', 'H1'],                       col: 'New H1'              },
      { labels: ['New Meta Title', 'Title'],             col: 'New Meta Title'       },
      { labels: ['New Meta Description', 'Description'], col: 'New Meta Description' }
    ];

    // Normalise output — LLM sometimes returns all fields on one line
    var normOutput = rawOutput
      .replace(/\s+(New\s+H1:|New\s+Meta\s+Title:|New\s+Meta\s+Description:|(?<![A-Za-z])H1:|(?<![A-Za-z])Title:|(?<![A-Za-z])Description:)/g, '\n$1');
    const lines  = normOutput.split(/\r?\n/);
    const parsed = {};
    let currentField = null;

    lines.forEach(function(line) {
      var trimmed = line.trim();

      // Skip schema blocks entirely — not expected, but safe to ignore
      if (trimmed.indexOf('<script') === 0) { currentField = null; return; }
      if (/^SCHEMA TYPE SELECTED:/i.test(trimmed)) return;

      var matched = false;
      FIELDS.forEach(function(f) {
        f.labels.forEach(function(lbl) {
          var prefix = lbl + ':';
          if (!matched && trimmed.indexOf(prefix) === 0) {
            currentField = f.col;
            parsed[f.col] = trimmed.substring(prefix.length).trim();
            matched = true;
          }
        });
      });

      // Accumulate continuation lines
      if (!matched && currentField && trimmed.length > 0) {
        parsed[currentField] = (parsed[currentField] || '') + trimmed;
      }
    });

    const written  = [];
    const missing  = [];
    const notFound = [];

    FIELDS.forEach(function(f) {
      const colIndex = headers.indexOf(f.col);
      if (colIndex === -1) { notFound.push(f.col); return; }
      const value = parsed[f.col];
      if (!value || value.trim() === '') { missing.push(f.col); return; }
      const cleaned = sanitiseGovernanceField(stripMarkdown(value), f.col);
      const cell = sheet.getRange(activeRow, colIndex + 1);
      try {
        cell.setPlainTextValue(cleaned);
      } catch(plainErr) {
        cell.setNumberFormat('@');
        cell.setValue(cleaned);
      }
      written.push(f.col);
    });

    var message = 'PUSHED ' + written.length + '/3 fields to Row ' + activeRow + '.';
    if (missing.length > 0)  message += '\nNot found in output: ' + missing.join(', ');
    if (notFound.length > 0) message += '\nColumn missing from sheet: ' + notFound.join(', ');

    if (written.length > 0) logPipelineResume("W5 — H1 Meta Push", "");
    return {
      success:  written.length > 0,
      message:  message,
      written:  written,
      missing:  missing,
      notFound: notFound
    };

  } catch(e) {
    return { success: false, message: 'PUSH ERROR: ' + e.toString() };
  }
}

/* ============================================================
   MARKDOWN STRIPPER
============================================================ */
function logPipelineResume(stageName, pipeline) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Pipeline Resume Log");
    if (!logSheet) return;

    var postsSheet = ss.getSheetByName("posts");
    var row = postsSheet.getActiveRange().getRow();
    if (row < 2) return;

    var headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];
    var postIdIdx = headers.indexOf("Post ID");
    var titleIdx  = headers.indexOf("Title");

    var postId = postIdIdx > -1 ? String(postsSheet.getRange(row, postIdIdx + 1).getValue() || "").trim() : "";
    var title  = titleIdx  > -1 ? String(postsSheet.getRange(row, titleIdx  + 1).getValue() || "").trim() : "";
    var timestamp = new Date();

    // Determine pipeline from posts_progress sheet
    var pipelineLabel = "Unknown";
    try {
      var progressSheet = ss.getSheetByName("posts_progress");
      if (progressSheet && postId) {
        var progressData = progressSheet.getDataRange().getValues();
        var pHeaders = progressData[0];
        var colIndex = {};
        for (var i = 0; i < pHeaders.length; i++) { colIndex[pHeaders[i]] = i; }
        for (var r = 1; r < progressData.length; r++) {
          if (String(progressData[r][colIndex["Post ID"]]) === postId) {
            var ce1date = progressData[r][colIndex["CE-1-date"]];
            var wpDate  = progressData[r][colIndex["WP-date"]];
            if (wpDate) {
              pipelineLabel = "Complete";
            } else if (ce1date) {
              pipelineLabel = "Pipeline 2";
            } else {
              pipelineLabel = "Pipeline 1";
            }
            break;
          }
        }
      }
    } catch(pe) {
      pipelineLabel = "Unknown";
    }

    // Insert new row at row 2 pushing existing entries down
    logSheet.insertRowBefore(2);
    logSheet.getRange(2, 1).setValue(postId);
    logSheet.getRange(2, 2).setValue(title);
    logSheet.getRange(2, 3).setValue(stageName);
    logSheet.getRange(2, 4).setValue(pipelineLabel);
    logSheet.getRange(2, 5).setValue(timestamp);

  } catch(e) {
    Logger.log("logPipelineResume error: " + e.message);
  }
}

function stripMarkdown(value) {
  if (!value) return value;

  // Decode URL-encoded characters (%22 → ", etc)
  if (value.indexOf('%22') > -1 || value.indexOf('%3A') > -1 || value.indexOf('%26') > -1) {
    value = value
      .replace(/%22/g, '"').replace(/%3A/g, ':').replace(/%2F/g, '/')
      .replace(/%7B/g, '{').replace(/%7D/g, '}').replace(/%5B/g, '[')
      .replace(/%5D/g, ']').replace(/%2C/g, ',').replace(/%20/g, ' ')
      .replace(/%26/g, '&');
  }

  // Decode HTML entities that Sheets may have introduced
  value = value.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'");

  // Remove bracket-wrapped URLs: ["https://... → "https://...
  value = value.replace(/\[("?https?:\/\/[^\]"]*"?)\]/g, '$1');
  value = value.replace(/\[(https?:\/\/)/g, '$1');

  // If this looks like JSON or a schema block, return now — don't run markdown stripping
  var t = value.trim();
  if (t.charAt(0) === '{' || t.indexOf('<script') === 0) return t;

  return value
    .replace(/\*\*(.+?)\*\*/g, '$1')   // **bold**
    .replace(/\*(.+?)\*/g, '$1')        // *italic*
    .replace(/__(.+?)__/g, '$1')        // __bold__
    .replace(/_(.+?)_/g, '$1')          // _italic_
    .replace(/`(.+?)`/g, '$1')          // `code`
    .replace(/^#{1,6}\s+/gm, '')        // # headings
    .replace(/^[\-\*\•]\s+/gm, '')      // - * • list markers
    .trim();
}

function getUrlByPostID(postID) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("posts");
  const data  = sheet.getDataRange().getValues();
  const headers  = data[0];
  const idIndex  = headers.indexOf("Post ID");
  const urlIndex = headers.indexOf("URL");
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(postID)) {
      return data[i][urlIndex];
    }
  }
  return null;
}

/* ============================================================
   getGovernanceFieldsForAudit()
   Called by W7 Run Audit button.
============================================================ */
function getGovernanceFieldsForAudit() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName("posts");
    const row     = sheet.getActiveRange().getRow();
    if (row < 2) return { error: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    function getCol(name) {
      const idx = headers.indexOf(name);
      if (idx === -1) return "";
      let val = sheet.getRange(row, idx + 1).getValue();
      return val ? String(val) : "";
    }

    function sanitiseSchema(raw) {
      if (!raw) return raw;
      const m = raw.match(/<script[^>]*>([\s\S]*?)<\/scr\x69pt>/i);
      if (m) raw = m[1].trim();
      raw = raw.replace(/%22/g, '"').replace(/%3A/g, ':').replace(/%2F/g, '/')
               .replace(/%7B/g, '{').replace(/%7D/g, '}').replace(/%26/g, '&')
               .replace(/%5B/g, '[').replace(/%5D/g, ']')
               .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
      raw = raw.replace(/\[("?https?:\/\/[^\]"]*"?)\]/g, '$1');
      raw = raw.replace(/\[(https?:\/\/)/g, '$1');
      return raw;
    }

    return {
      h1:        getCol("New H1"),
      metaTitle: getCol("New Meta Title"),
      metaDesc:  getCol("New Meta Description"),
      schema:    sanitiseSchema(getCol("Schema (JSON-LD)"))
    };
  } catch(e) {
    return { error: "Could not read sheet: " + e.toString() };
  }
}

/* ============================================================
   STAGE 0 — ANALYTICAL FIELDS PUSH
   Reads the five labelled lines from Stage 0 LLM output and
   writes them to the active row.
   Expected labels (exact):
     Primary Search Term:
     Strategic Reasoning:
     Topical Intent:
     Matrix Role:
     Key Decisions Explained:
============================================================ */

/* ============================================================
   SIDEBAR — CONTENT ENGINE (CHATGPT) MAIN INTERFACE
   Opens the modeless sidebar for building Stage 1-5 prompts
   and pushing results back to the sheet.
============================================================ */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function openBuildContentEnginePromptFromActiveRowChatGPTSidebar() {
  checkAndShowPipelineUpdateNotice();

  const d = getActiveRowDataMap();
  const preload = {
    postID:         d['Post ID']                    || '',
    title:          d['Title']                       || '',
    url:            d['URL']                         || '',
    articleType:    d['Article Type']                || '',
    stoneType:      d['Stone Type']                  || '',
    primaryCluster: d['Primary Query Cluster Owned'] || '',
    problemAngle:   d['Problem Angle']               || '',
    authorityBrief: d['Authority Brief']             || '',
    rewriteBrief:   d['Page Rewrite Brief']          || '',
    rewriteStatus:  d['Rewrite Status']              || '',
    locality:       d['Locality']                    || '',
    parentArea:     d['Parent Area']                 || '',
    locationCtx:    d['Location Context']            || '',
    newH1:          d['New H1']                      || '',
    newMetaTitle:   d['New Meta Title']              || '',
    newMetaDesc:    d['New Meta Description']        || ''
  };

  const w1 = runW1ReadinessCheck(d);

  const tmpl = HtmlService.createTemplateFromFile('ce_Sidebar_Main');
  tmpl.preload = JSON.stringify(preload);
  tmpl.w1Check = JSON.stringify(w1 || {});

  const html = tmpl.evaluate()
    .setWidth(600)
    .setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Content Engine');
}

/* ============================================================
   TODO: ADD GUARDRAILS EXTRACTION FUNCTIONS
   
   These functions must be added manually:
   
   1. extractGuardrailsFromTSM(material)
      - Extracts Technical Guardrails section from TSM
      - Returns guardrails text or empty string
      
   2. extractProhibitedPatternsOnly(guardrailsText)
      - Extracts prohibited language patterns from guardrails
      - Used for W4B image alt text generation
      
   See Pipeline_Architecture_Analysis_v76.md for implementation.
============================================================ */

/**
 * Get TSM requirements for a specific tier and article type
 */
function getTierStructuralRequirements(material, tierLabel, articleType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tsmSheet = ss.getSheetByName("Tier Structural Coverage Matrix");
  if (!tsmSheet) return [];

  const data = tsmSheet.getDataRange().getValues();
  const headers = data[0];

  const tierCol = headers.indexOf("Tier");
  const labelCol = headers.indexOf("Label");
  const reqCol = headers.indexOf("Structural Requirement");
  const orderCol = headers.indexOf("Order");
  const wordCountCol = headers.indexOf("Word Count");
  const permittedCol = headers.indexOf("Permitted Article Types");

  const articleTypeLower = articleType
    ? String(articleType).trim().toLowerCase()
    : null;

  let requirements = [];

  for (let i = 1; i < data.length; i++) {
    const rowTier = String(data[i][tierCol] || "").trim();

    if (rowTier !== tierLabel) continue;

    if (
      rowTier === "GLOBAL RULE" ||
      rowTier === "GOVERNANCE RULE"
    ) {
      continue;
    }

    if (articleTypeLower && permittedCol > -1) {
      const permittedRaw =
        String(data[i][permittedCol] || "").trim();

      const permittedList = permittedRaw
        .split(",")
        .map(function(s) {
          return s.trim().toLowerCase();
        });

      if (
        permittedList.indexOf(articleTypeLower) === -1
      ) {
        continue;
      }
    }

    const order = data[i][orderCol];
    const name = String(data[i][reqCol] || "").trim();
    const wc = data[i][wordCountCol];

    if (!order || isNaN(Number(order))) continue;

    requirements.push({
      order: Number(order),
      name: name,
      wordCount:
        wc && !isNaN(Number(wc))
          ? Number(wc)
          : 120
    });
  }

  requirements.sort(function(a, b) {
    return a.order - b.order;
  });

  return requirements;
}

/**
 * Get total word count for a tier by summing TSM requirement word counts
 */
function getTierWordCount(tierLabel, articleType) {
  const requirements = getTierStructuralRequirements(null, tierLabel, articleType);
  if (requirements.length === 0) {
    return 900;
  }
  let total = 0;
  requirements.forEach(function(req) {
    total += req.wordCount;
  });
  return total > 0 ? total : 900;
}

/**
 * SHARED SOURCE OF TRUTH — total article word count ceiling by Article Type.
 * Sums the Tier Structural Coverage Matrix's Word Count column for every
 * row whose Permitted Article Types includes this Article Type, then
 * applies the standard ±15% band.
 * Used by both W2B generation (buildTotalWordCountTargetBlock) and the
 * W1.5F semantic alignment fix prompt — keep this the ONLY place that
 * defines the ceiling calculation so the two never drift out of sync.
 * Returns null if no matching Tier Matrix rows are found.
 */
function getTierWordCountCeiling(articleType) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var matrixSheet = ss.getSheetByName("Tier Structural Coverage Matrix");
    if (!matrixSheet) return null;

    var data = matrixSheet.getDataRange().getValues();
    if (data.length < 2) return null;
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var permittedIdx = headers.indexOf("Permitted Article Types");
    var wordCountIdx = headers.indexOf("Word Count");
    if (permittedIdx === -1 || wordCountIdx === -1) return null;

    var atLower = String(articleType || "").trim().toLowerCase();
    var total = 0;
    var matchedRows = 0;

    for (var i = 1; i < data.length; i++) {
      var permitted = String(data[i][permittedIdx] || "");
      var permittedList = permitted.split(",").map(function(s) { return s.trim().toLowerCase(); });
      if (permittedList.indexOf(atLower) > -1) {
        var wc = parseInt(data[i][wordCountIdx], 10);
        if (!isNaN(wc)) {
          total += wc;
          matchedRows++;
        }
      }
    }

    if (total === 0 || matchedRows === 0) return null;

    return {
      total: total,
      minTarget: Math.round(total * 0.85),
      maxTarget: Math.round(total * 1.15)
    };
  } catch(e) {
    return null;
  }
}

/**
 * Get article type expectations from Article Type Control Sheet
 */
function getArticleTypeExpectations(articleType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Article Type Control Sheet");
  if (!sheet) return "No article type expectations found.";
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const typeCol = headers.indexOf("Canonical Article Type Label");
  const expectCol = headers.indexOf("On-Page Content Expectations");
  
  if (typeCol === -1 || expectCol === -1) {
    return "ERROR: Article Type Control Sheet missing required columns";
  }
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][typeCol]).trim() === articleType) {
      return data[i][expectCol] || "No expectations defined for this article type.";
    }
  }
  
  return "Article type not found in Article Type Control Sheet.";
}

/**
 * Get section plan from Column CX (102) — W1.5D Final Plan.
 * This is the last approved stage before W2B execution.
 * W2B reads optimization data directly from Column CY (103) via ce_Stage2_ContentGeneration.gs.
 */
function getSectionPlan() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row = sheet.getActiveRange().getRow();
    if (row < 2) return null;
    var plan = String(sheet.getRange(row, 102).getValue() || "").trim();
    if (!plan) return null;
    // Strip markdown-wrapped URLs (e.g. "[url](url)") that may have been
    // introduced by an LLM stage — Header Video and Internal link parsing
    // downstream require raw URLs, not markdown syntax.
    plan = plan.replace(/\[(https?:\/\/[^\]]+)\]\([^)]+\)/g, '$1');
    return plan || null;
  } catch(e) {
    return null;
  }
}
function getActivePostId() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  const row = sheet.getActiveRange().getRow();

  // 🔴 CHANGE COLUMN NUMBER to your Post ID column
  const postId = sheet.getRange(row, 1).getValue();

  return postId;
}

function bc_recalculateTfIdfForActiveRow() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    function getIdx(name) { return headers.indexOf(name); }

    const col104Idx      = getIdx("tfidf_terms");
    const tfidfTermsIdx  = getIdx("tfidf_terms");
    const tfidfSummaryIdx = getIdx("tfidf_summary");
    const recallScoreIdx = getIdx("recall_score");
    const recallNotesIdx = getIdx("recall_notes");
    const pqcIdx         = getIdx("Primary Query Cluster Owned");

    if (tfidfTermsIdx === -1)   return { success: false, message: "tfidf_terms column not found." };
    if (tfidfSummaryIdx === -1) return { success: false, message: "tfidf_summary column not found." };

    // Read new HTML — prefer col 104 (humanised), fall back to col 98 (new HTML)
    var html = String(sheet.getRange(row, 104).getValue() || "").trim();
    if (!html) html = String(sheet.getRange(row, 98).getValue() || "").trim();
    if (!html) return { success: false, message: "No HTML found in col 104 or col 98 for this row." };

    // Strip HTML tags and extract plain text words
    var plainText = html.replace(/<[^>]+>/g, " ").replace(/&[a-z]+;/gi, " ").toLowerCase();
    var allWords  = plainText.split(/\s+/);

    const STOPWORDS = new Set([
      "the","and","for","with","how","can","you","your","from","that",
      "this","are","was","have","has","not","use","used","using",
      "what","why","when","will","does","best","get","all","its",
      "more","been","into","our","out","their","they","who","which",
      "but","also","than","then","them","these","those","some","any",
      "one","two","per","new","old","top","way","need","make",
      "take","much","many","long","very","well","good","able",
      "just","like","even","still","often","around","always",
      "never","every","each","both","here","there","where","while",
      "about","after","before","should","could","would","other",
      "first","last","next","may","might","most","only","over",
      "such","same","back","keep","look","work","give","move"
    ]);

    function stemWord(word) {
      if (word.length < 4) return word;
      if (/ing$/.test(word) && word.length > 6)  return word.replace(/ing$/, "");
      if (/tion$/.test(word) && word.length > 6) return word.replace(/tion$/, "");
      if (/tions$/.test(word) && word.length > 7) return word.replace(/tions$/, "");
      if (/ing$/.test(word) && word.length > 5)  return word.replace(/ing$/, "");
      if (/ness$/.test(word))  return word.replace(/ness$/, "");
      if (/ment$/.test(word))  return word.replace(/ment$/, "");
      if (/ated$/.test(word))  return word.replace(/ated$/, "ate");
      if (/ical$/.test(word))  return word.replace(/ical$/, "ic");
      if (/ical$/.test(word))  return word.replace(/ical$/, "");
      if (/ally$/.test(word))  return word.replace(/ally$/, "al");
      if (/ously$/.test(word)) return word.replace(/ously$/, "ous");
      if (/ives$/.test(word))  return word.replace(/ives$/, "ive");
      if (/ive$/.test(word) && word.length > 5) return word.replace(/ive$/, "");
      if (/ised$/.test(word))  return word.replace(/ised$/, "ise");
      if (/ized$/.test(word))  return word.replace(/ized$/, "ize");
      if (/ers$/.test(word) && word.length > 5)  return word.replace(/ers$/, "er");
      if (/ed$/.test(word) && word.length > 5)   return word.replace(/ed$/, "");
      if (/es$/.test(word) && word.length > 4)   return word.replace(/es$/, "");
      if (/s$/.test(word) && word.length > 4)    return word.replace(/s$/, "");
      return word;
    }

    var words = [];
    allWords.forEach(function(word) {
      var clean = word.replace(/[^a-z]/g, "");
      if (clean.length >= 3 && !STOPWORDS.has(clean)) words.push(stemWord(clean));
    });

    if (words.length === 0) return { success: false, message: "No usable words found in HTML." };

    // Calculate TF
    var tf = {};
    words.forEach(function(w) { tf[w] = (tf[w] || 0) + 1; });

    // For IDF we use the existing corpus df from all rows
    const postsData = sheet.getDataRange().getValues();
    const postHeaders = postsData[0];
    const gscIdx = postHeaders.indexOf("GSC Query List");
    var docCount = 0;
    var df = {};

    if (gscIdx > -1) {
      for (var i = 1; i < postsData.length; i++) {
        var drValue = String(postsData[i][gscIdx] || "").trim();
        if (!drValue) continue;
        docCount++;
        var docWords = new Set();
        drValue.split(",").forEach(function(q) {
          q.trim().toLowerCase().split(/\s+/).forEach(function(w) {
            var c = w.replace(/[^a-z]/g, "");
            if (c.length >= 3 && !STOPWORDS.has(c)) docWords.add(c);
          });
        });
        docWords.forEach(function(w) { df[w] = (df[w] || 0) + 1; });
      }
    }

    if (docCount === 0) docCount = 1;

    // Score terms
    var scores = {};
    Object.keys(tf).forEach(function(word) {
      var termTf  = tf[word] / words.length;
      var termIdf = Math.log(docCount / (df[word] || 1));
      scores[word] = termTf * termIdf;
    });

    var ranked = Object.keys(scores).sort(function(a, b) {
      return scores[b] - scores[a];
    }).slice(0, 15);

    var tfidfTerms   = ranked.map(function(w) { return w + ":" + scores[w].toFixed(3); }).join(", ");
    var tfidfSummary = ranked.slice(0, 8).join(", ");

    // Save previous recall score and notes before overwriting
    if (recallScoreIdx > -1 && recallNotesIdx > -1) {
      var prevScoreIdx = headers.indexOf("recall_score_previous");
      var prevNotesIdx = headers.indexOf("recall_notes_previous");
      if (prevScoreIdx > -1) {
        sheet.getRange(row, prevScoreIdx + 1).setValue(
          sheet.getRange(row, recallScoreIdx + 1).getValue()
        );
      }
      if (prevNotesIdx > -1) {
        sheet.getRange(row, prevNotesIdx + 1).setValue(
          sheet.getRange(row, recallNotesIdx + 1).getValue()
        );
      }
    }

    sheet.getRange(row, tfidfTermsIdx + 1).setValue(tfidfTerms);
    sheet.getRange(row, tfidfSummaryIdx + 1).setValue(tfidfSummary);

    // Recalculate recall score if PQC column exists
    var recallResult = "";
    if (pqcIdx > -1 && recallScoreIdx > -1 && recallNotesIdx > -1) {
      var pqc = String(sheet.getRange(row, pqcIdx + 1).getValue() || "").toLowerCase();
      var pqcWords = new Set();
    pqc.split(/\s+/).forEach(function(w) {
      var c = w.replace(/[^a-z]/g, "");
      if (c.length >= 3 && !STOPWORDS.has(c)) {
        var stemmed = stemWord(c);
        if (!STOPWORDS.has(stemmed)) pqcWords.add(stemmed);
      }
    });

      var matched   = [];
      var unmatched = [];
      var pqcOriginal = {};
      pqc.split(/\s+/).forEach(function(w) {
        var c = w.replace(/[^a-z]/g, "");
        if (c.length >= 3 && !STOPWORDS.has(c)) {
          var stemmed = stemWord(c);
          if (!STOPWORDS.has(stemmed)) pqcOriginal[stemmed] = c;
        }
      });
      pqcWords.forEach(function(w) {
        if (tf[w]) matched.push(pqcOriginal[w] || w);
        else unmatched.push(pqcOriginal[w] || w);
      });

      var recallScore = pqcWords.size > 0
        ? (matched.length / pqcWords.size).toFixed(2)
        : "0.00";

      var recallNotes = "Matched: " + matched.join(", ") +
                        ". Not in PQC: " + unmatched.join(", ");

      sheet.getRange(row, recallScoreIdx + 1).setValue(recallScore);
      sheet.getRange(row, recallNotesIdx + 1).setValue(recallNotes);
      var articleTypeForRecall = String(sheet.getRange(row, headers.indexOf("Article Type") + 1).getValue() || "").trim();
      var recallRecommendation = buildRecallRecommendation(parseFloat(recallScore), articleTypeForRecall, matched, unmatched);
      sheet.getRange(row, recallNotesIdx + 1).setValue(recallNotes + "\n\nRECOMMENDATION:\n" + recallRecommendation);
      recallResult = " Recall score: " + recallScore + ". " + recallRecommendation;
    }

    return {
      success: true,
      message: "TF-IDF recalculated for row " + row + "." + recallResult
    };

  } catch(e) {
    return { success: false, message: "bc_recalculateTfIdfForActiveRow error: " + e.message };
  }
}
function showVideoManagerPanel() {
  const html = HtmlService.createHtmlOutputFromFile('ce_Sidebar_VideoManager')
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Video Manager');
}
function getAcValuesForMaterial(material, articleType) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "ac-library-hub";
    if (articleType) {
      var typeMap = {
        "Case Study":        "ac-library-case-study",
        "Method Guide":      "ac-library-method-guide",
        "Diagnostic Guide":  "ac-library-diagnostic-guide",
        "Educational Guide": "ac-library-educational-guide",
        "Buyer Guide":       "ac-library-buyer-guide",
        "Service Page":      "ac-library-service-page",
        "Geo Service Page":  "ac-library-geo-service-page",
        "Hub Page":          "ac-library-hub"
      };
      if (typeMap[articleType]) sheetName = typeMap[articleType];
    }
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.getSheetByName("ac-library-hub");
    if (!sheet) return { values: {}, labels: {} };

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });

    var dropdownCol  = headers.indexOf("dropdown");
    var valueCol     = headers.indexOf("Value") > -1 ? headers.indexOf("Value") : headers.indexOf("value");
    var universalCol = headers.indexOf("Universal");
    var labelCol     = headers.indexOf("Label");

    if (dropdownCol === -1 || valueCol === -1 || universalCol === -1) return {};

    var materialLower = String(material || "").trim().toLowerCase();
    var isCatchAll = (materialLower === "tile" || materialLower === "stone");

    var materialCol = -1;
    if (!isCatchAll) {
      for (var h = 0; h < headers.length; h++) {
        if (headers[h].toLowerCase() === materialLower) {
          materialCol = h;
          break;
        }
      }
    }

    var result = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var axis  = String(row[dropdownCol] || "").trim();
      var value = String(row[valueCol]    || "").trim();
      if (!axis || !value) continue;

      var isUniversal    = String(row[universalCol] || "").trim().toLowerCase() === "y";
      var isMaterialMatch = isCatchAll ||
                            (materialCol > -1 && String(row[materialCol] || "").trim().toLowerCase() === "y");

      if (isUniversal || isMaterialMatch) {
        if (!result[axis]) result[axis] = [];
        result[axis].push(value);
      }
    }

    // Build label map from Label column
    var labels = {};
    if (labelCol > -1) {
      for (var li = 1; li < data.length; li++) {
        var lAxis  = String(data[li][dropdownCol] || "").trim();
        var lLabel = String(data[li][labelCol]    || "").trim();
        if (lAxis && lLabel && !labels[lAxis]) {
          labels[lAxis] = lLabel;
        }
      }
    }

    return { values: result, labels: labels };

  } catch(e) {
    Logger.log("getAcValuesForMaterial error: " + e.message);
    return { values: {}, labels: {} };
  }
}

function getHubPageDefaults() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sheet   = ss.getSheetByName("ac-library-hub");
    if (!sheet) return {};

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });

    var dropdownIdx  = headers.indexOf("dropdown");
    var valueIdx     = headers.indexOf("Value");
    var hubDefaultIdx = headers.indexOf("Hub Default");

    if (dropdownIdx === -1 || valueIdx === -1 || hubDefaultIdx === -1) return {};

    var hubDefaults = {};
    for (var i = 1; i < data.length; i++) {
      var dropdown = String(data[i][dropdownIdx] || "").trim();
      var value    = String(data[i][valueIdx]    || "").trim();
      var isHub    = String(data[i][hubDefaultIdx] || "").trim().toLowerCase() === "y";
      if (dropdown && value && isHub) {
        hubDefaults[dropdown] = value;
      }
    }
    return hubDefaults;

  } catch(e) {
    Logger.log("getHubPageDefaults error: " + e.message);
    return {};
  }
}
function buildFeaturedImagePrompt() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    var d = getActiveRowDataMap();

    var material    = String(d["Stone Type"]          || "").trim();
    var articleType = String(d["Article Type"]        || "").trim();
    var locality    = String(d["Locality"]            || "").trim();
    var parentArea  = String(d["Parent Area"]         || "").trim();
    var h1          = String(d["New H1"]              || "").trim();
    var header      = String(sheet.getRange(row, 98).getValue() || "").trim();

    var entryCondition    = String(d["ac_entry_condition"]      || "").trim();
    var homPerception     = String(d["ac_homeowner_perception"]  || "").trim();
    var matBehaviour      = String(d["ac_material_behaviour"]    || "").trim();
    var constraint        = String(d["ac_constraint"]            || "").trim();
    var processEmphasis   = String(d["ac_process_emphasis"]      || "").trim();
    var resultType        = String(d["ac_result_type"]           || "").trim();
    var storyDirection    = String(d["ac_narrative_archetype"]   || "").trim();

    // Extract header text from HTML
    var headerText = "";
    if (header) {
      var headerMatch = header.match(/<header[\s\S]*?<\/header>/i);
      if (headerMatch) {
        headerText = headerMatch[0].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().substring(0, 300);
      }
    }

    // Build location string
    var locationStr = locality ? locality + (parentArea ? ", " + parentArea : "") : parentArea;

    // Build hook phrase from H1 or fallback
    var hookPhrase = h1 || (material + " — " + articleType);

    // Build scene description from ac_ fields and article content
    var sceneElements = [];

    if (entryCondition) sceneElements.push("The floor condition: " + entryCondition);
    if (matBehaviour)   sceneElements.push("What the floor was doing: " + matBehaviour);
    if (constraint)     sceneElements.push("The governing constraint: " + constraint);
    if (resultType)     sceneElements.push("The outcome: " + resultType);
    if (headerText)     sceneElements.push("Article opening context: " + headerText);

    // Build story mood from story direction
    var mood = "";
    var dirLower = storyDirection.toLowerCase();
    if (dirLower.includes("discovery"))           mood = "Mood: discovery and revelation — something beautiful found after years hidden.";
    else if (dirLower.includes("transformation")) mood = "Mood: dramatic contrast — before and after, neglect becoming feature.";
    else if (dirLower.includes("rescue"))         mood = "Mood: tension and recovery — floor on the edge of being lost, then saved.";
    else if (dirLower.includes("preservation"))   mood = "Mood: careful respect — original character protected throughout.";
    else if (dirLower.includes("reassurance"))    mood = "Mood: calm expertise — careful professional approach proving safety.";
    else if (dirLower.includes("sympathetic"))    mood = "Mood: respectful craft — period character respected in every decision.";
    else if (dirLower.includes("maintenance"))    mood = "Mood: relief and simplicity — impossible floor made manageable again.";
    else if (dirLower.includes("recovery"))       mood = "Mood: correction and stability — previous damage put right.";
    else if (dirLower.includes("extending"))      mood = "Mood: longevity and care — floor given another generation of life.";
    else if (dirLower.includes("correction"))     mood = "Mood: professional correction — wrong treatment put right by expert.";
    else                                           mood = "Mood: professional expertise and care.";

    // Build the full prompt — art direction format
    var p = "";

    // ── SECTION 1: IMAGE PURPOSE AND STORY ──
    p += "CREATE A PREMIUM EDITORIAL FEATURE IMAGE\n\n";
    p += "IMAGE PURPOSE:\n";
    p += "Featured image for a professional UK stone and tile restoration article.\n\n";
    p += "PRIMARY VISUAL STORY:\n";
    if (storyDirection.toLowerCase().includes("authority")) {
      p += "The room has already been beautifully modernised, but the " + material.toLowerCase() + " floor is still failing.\n";
      p += "The viewer should immediately think:\n";
      p += "\"Everything else looks expensive and complete — why does the floor still look tired?\"\n";
    } else if (storyDirection.toLowerCase().includes("transformation")) {
      p += "A floor that was written off as beyond saving has been transformed into the standout feature of the room.\n";
      p += "The viewer should immediately think:\n";
      p += "\"I would never have believed that floor could look this good.\"\n";
    } else if (storyDirection.toLowerCase().includes("rescue")) {
      p += "A " + material.toLowerCase() + " floor that looked beyond saving has been recovered through careful professional intervention.\n";
      p += "The viewer should immediately think:\n";
      p += "\"That floor looked lost — but it came back.\"\n";
    } else if (storyDirection.toLowerCase().includes("sympathetic")) {
      p += "An original " + material.toLowerCase() + " floor has been carefully restored without erasing its age or character.\n";
      p += "The viewer should immediately think:\n";
      p += "\"The history of that floor has been respected, not removed.\"\n";
    } else if (storyDirection.toLowerCase().includes("discovery")) {
      p += "A hidden " + material.toLowerCase() + " floor has been uncovered after years of concealment.\n";
      p += "The viewer should immediately think:\n";
      p += "\"Something beautiful was hidden here all along.\"\n";
    } else {
      p += "The " + material.toLowerCase() + " floor in this room is the subject of professional restoration.\n";
      p += "The viewer should immediately think:\n";
      p += "\"This floor needed expert attention — and it got it.\"\n";
    }
    p += "\nThe emotional tension between the aspirational modern room and the worn " + material.toLowerCase() + " floor is the entire image concept.\n\n";
    p += "ARTICLE CONTEXT:\n";
    p += "Material: " + material + "\n";
    p += "Article type: " + articleType + "\n";
    if (locationStr) p += "Location: " + locationStr + "\n";
    p += "Tone: calm authority, professional expertise, realistic restoration judgement\n\n";

    // ── SECTION 2: FLOOR DIRECTION ──
    p += "FLOOR CONDITION:\n";
    p += "Use the attached floor photograph as the exact visual reference for tile pattern, colour variation, grout tone, surface wear, and ingrained soiling.\n";
    p += "The floor must NOT look generic. It should clearly show the specific condition described below.\n\n";
    if (entryCondition) p += "Entry condition: " + entryCondition + "\n";
    if (matBehaviour)   p += "What the floor was doing: " + matBehaviour + "\n";
    if (constraint)     p += "Governing constraint: " + constraint + "\n";
    if (resultType)     p += "Outcome: " + resultType + "\n";
    if (headerText)     p += "Article context: " + headerText + "\n";
    p += "\nThe floor must show:\n";
    p += "— embedded contamination and ingrained soiling\n";
    p += "— worn or inconsistent surface texture\n";
    p += "— dark grout lines and uneven sheen\n";
    p += "— signs that normal cleaning is no longer improving the appearance\n\n";
    p += "The renovation is complete except for the floor.\n";
    p += "The floor is the hero of the image. Everything else in the room exists to make the floor condition visible.\n\n";
    p += "VISUAL PRIORITY ORDER:\n";
    p += "1. Floor condition — must be immediately readable\n";
    p += "2. Room brightness contrast — the modern room makes the worn floor more visible\n";
    p += "3. Typography readability — clear at thumbnail scale\n";
    p += "4. Premium magazine mood — restrained and professional\n";
    p += "5. Architectural styling — supports the story without distracting from it\n\n";
    // ── SECTION 3: ROOM DIRECTION ──
    p += "ROOM DESIGN DIRECTION:\n";
    p += "The room itself must look bright, modern, expensive and recently renovated.\n";
    p += "Use warm off-whites, greige, or muted sage wall colours inspired by current British interior design magazines.\n";
    p += "Include contrast skirting boards, brushed brass or matte black details, soft natural daylight.\n";
    p += "Add subtle lifestyle details — a potted olive tree or plant, clean hallway styling, natural textures.\n\n";
    if (articleType === "Case Study" && locality) {
      p += "Setting: a period property entrance hallway in " + locality + " that has been modernised throughout except for the floor.\n\n";
    } else if (articleType === "Method Guide") {
      p += "Setting: a close-up professional restoration environment — bright, clean, well-lit workroom atmosphere.\n\n";
    } else if (articleType === "Diagnostic Guide") {
      p += "Setting: a domestic floor surface under clear diagnostic lighting — bright and clinical, not atmospheric.\n\n";
    } else if (articleType === "Hub Page") {
      p += "Setting: an elegant wide-angle hallway or kitchen composition showing the full floor surface across multiple conditions.\n\n";
    } else {
      p += "Setting: a bright modern domestic interior where the floor is clearly the only unresolved element.\n\n";
    }
    p += "DO NOT make the room:\n";
    p += "— dark, gloomy or atmospheric\n";
    p += "— Victorian, rustic or derelict\n";
    p += "— cluttered or overly decorated\n";
    p += "— a luxury hotel or fantasy interior\n";
    p += "Only the floor should appear aged or worn.\n\n";
    // ── SECTION 4: TYPOGRAPHY ──
    p += "TYPOGRAPHY OVERLAY:\n";
    p += "Leave the left third of the image visually clean for typography.\n";
    p += "Place text directly over the wall area — dark charcoal or near-black text reads clearly against light walls without any panel or overlay.\n\n";
    if (locationStr) {
      p += "Top small label (gold caps): " + locationStr.toUpperCase() + "\n";
    }
    var articleTypePhraseMap = {
      "Hub Page":          "Before You Touch The Floor",
      "Buyer Guide":       "What To Ask Before You Commit",
      "Educational Guide": "Why Your Floor Behaves This Way",
      "Method Guide":      "What Actually Happens Next",
      "Diagnostic Guide":  "When Cleaning Stops Helping",
      "Case Study":        "Not A Showroom Floor",
      "Service Page":      "Real Work On Real Floors",
      "Geo Service Page":  "Assessed And Restored Locally"
    };
    var middleTitle = articleTypePhraseMap[articleType] || "Real Work On Real Floors";
    p += "Middle title (elegant dark serif): " + middleTitle + "\n";
    // Store middleTitle for sidebar pre-population
    var _middleTitle = middleTitle;
    p += "Bottom hook phrase (bold condensed, newspaper headline weight): " + hookPhrase.toUpperCase() + "\n\n";
    p += "Typography rules:\n";
    p += "— elegant serif for the material and article type label\n";
    p += "— bold condensed sans-serif for the hook phrase\n";
    p += "— muted gold for location label and any decorative divider lines\n";
    p += "— dark charcoal or near-black text colour — no white text, no dark overlay panel\n";
    p += "— all text must be clearly readable at thumbnail scale\n\n";
    // ── SECTION 5: NEGATIVE RULES ──
    p += "NEGATIVE RULES — DO NOT INCLUDE:\n";
    p += "— cleaning equipment, buckets, mops, or spray bottles\n";
    p += "— workers, technicians, or people posing\n";
    p += "— cartoon faces or exaggerated expressions\n";
    p += "— dramatic cracked or broken tiles\n";
    p += "— generic marble or stone patterns unrelated to " + material.toLowerCase() + "\n";
    p += "— AI fantasy interiors or luxury hotel styling\n";
    p += "— overly glossy or artificially polished floor surfaces\n";
    p += "— orange HDR lighting or artificial glow effects\n";
    p += "— dark overlay panels or gradient boxes behind typography\n";
    p += "— any element that makes the room look derelict, abandoned or neglected\n\n";
    p += "The image must feel grounded, realistic and professionally restrained.\n\n";
    // ── SECTION 6: OUTPUT SPECIFICATION ──
    p += "OUTPUT SPECIFICATION:\n";
    p += "Width: exactly 1270px\n";
    p += "Format: webp\n";
    p += "File size: under 100kb after compression\n";
    p += "Aspect ratio: 16:9 or 4:3 landscape\n";
    p += "Style: semi-realistic painterly illustration — not photographic, not cartoon\n";
    p += "Realistic " + material.toLowerCase() + " surface texture detail throughout\n";
    p += "Warm cinematic natural lighting with realistic shadows\n";
    p += "Premium architectural restoration magazine aesthetic\n\n";
    p += "SEED IMAGE:\n";
    p += "Attach a photograph of the actual " + material.toLowerCase() + " floor when sending this prompt to ChatGPT.\n";
    p += "Use it as the exact reference for tile pattern, colour, grout tone and surface condition.\n";
    p += "The generated floor must reflect the real surface — not a generic " + material.toLowerCase() + " pattern.\n";
    p += "If no photograph is attached, generate a realistic " + material.toLowerCase() + " surface appropriate to the article context.\n";

    // Build governed filename
    var materialSlug    = material.toLowerCase().replace(/\s+/g, "-").replace(/[^a-z0-9-]/g, "");
    var articleTypeSlug = articleType.toLowerCase().replace(/\s+/g, "-").replace(/[^a-z0-9-]/g, "");
    var locationSlug    = locality.toLowerCase().replace(/\s+/g, "-").replace(/[^a-z0-9-]/g, "");
    var filename = [materialSlug, articleTypeSlug, locationSlug].filter(Boolean).join("-") + ".webp";

    // Build alt text
    var altText = material + " " + articleType.toLowerCase();
    if (entryCondition) altText += " — " + entryCondition.toLowerCase();
    if (locationStr)    altText += " in " + locationStr;

    return {
      success:     true,
      prompt:      p,
      filename:    filename,
      altText:     altText,
      middleTitle: _middleTitle
    };

  } catch(e) {
    return { success: false, message: "buildFeaturedImagePrompt error: " + e.message };
  }
}

function saveFeaturedImageToSheet(url, filename, altText) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });

    var saved   = [];
    var missing = [];

    function writeCol(name, value) {
      var idx = headers.indexOf(name);
      if (idx === -1) { missing.push(name); return; }
      var cell = sheet.getRange(row, idx + 1);
      cell.setNumberFormat("@");
      cell.setValue(value);
      saved.push(name);
    }

    writeCol("Featured Image URL",      url);
    writeCol("Featured Image Filename", filename);
    writeCol("Featured Image Alt Text", altText);

    var message = "Saved " + saved.length + " field(s) to row " + row + ".";
    if (missing.length > 0) message += " Columns not found — add to posts sheet: " + missing.join(", ");

    return { success: saved.length > 0 || missing.length > 0, message: message };

  } catch(e) {
    return { success: false, message: "saveFeaturedImageToSheet error: " + e.message };
  }
}
function acBuildClassifyPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName('posts');
  const row = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getRow(); // your existing row-fetch function
  const headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];

  const articleTypeCol = headers.indexOf('Article Type');
  const stoneTypeCol = headers.indexOf('Stone Type');
  const rowData = postsSheet.getRange(row, 1, 1, postsSheet.getLastColumn()).getValues()[0];

  const articleType = rowData[articleTypeCol];
  const stoneType = rowData[stoneTypeCol];

  const sheetMap = {
    "Case Study":        "ac-library-case-study",
    "Method Guide":      "ac-library-method-guide",
    "Diagnostic Guide":  "ac-library-diagnostic-guide",
    "Educational Guide": "ac-library-educational-guide",
    "Buyer Guide":       "ac-library-buyer-guide",
    "Service Page":      "ac-library-service-page",
    "Geo Service Page":  "ac-library-geo-service-page",
    "Hub Page":          "ac-library-hub"
  };

  const acSheetName = sheetMap[articleType];
  if (!acSheetName) return { error: 'No AC sheet found for article type: ' + articleType };

  const acSheet = ss.getSheetByName(acSheetName);
  if (!acSheet) return { error: 'AC sheet not found: ' + acSheetName };

  const data = acSheet.getDataRange().getValues();
  const headerRow = data[0];

  const universalCol = headerRow.indexOf('Universal');
  const stoneCol = headerRow.indexOf(stoneType);

  // Group options by dropdown key
  const fields = {};
  const fieldLabels = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = row[0];
    const label = row[1];
    const value = row[2];
    const isUniversal = row[universalCol] === 'y';
    const isStone = stoneCol > -1 && row[stoneCol] === 'y';
    if (!key || !value) continue;
    if (isUniversal || isStone) {
      if (!fields[key]) { fields[key] = []; fieldLabels[key] = label; }
      if (!fields[key].includes(value)) fields[key].push(value);
    }
  }

  // Build prompt
  let prompt = `You are an article classifier for a specialist floor care content pipeline.\n\n`;
  prompt += `Article Type: ${articleType}\n`;
  prompt += `Stone Type: ${stoneType}\n\n`;
  prompt += `Read the article below and for each field, select the single closest matching option from the list provided.\n\n`;
  prompt += `Output your answer as a JSON array only. No preamble, no explanation, no markdown code fences.\n\n`;
  prompt += `For each field, only select an existing option if it is a genuinely close match to what the article describes. Do NOT force a selection if none of the options fit well — set "gap" to true instead and suggest new wording.\n\n`;
  prompt += `Format:\n[\n  {"index": 0, "field": "[field name]", "gap": false, "selected": "[exact option text]", "suggested_option": "", "reason": "[one sentence]"},\n  ...\n]\n\n`;
  prompt += `If gap is true, "selected" should be empty string and "suggested_option" should contain your proposed new option wording that better fits the article.\n\n`;
  prompt += `The index must match the field number as listed below. Use the exact option text as listed when gap is false. Do not output anything else.\n\n`;
  prompt += `---\n\nFIELDS AND OPTIONS:\n\n`;

  const fieldOrder = [
    'ac_what_they_found',
    'ac_why_they_called',
    'ac_what_the_floor_was_doing',
    'ac_what_forced_careful_decisions',
    'ac_what_governed_the_approach',
    'ac_what_changed_for_the_homeowner',
    'ac_story_direction'
  ];

  let fieldIndex = 0;
  fieldOrder.forEach(function(key) {
    if (!fields[key]) return;
    prompt += `FIELD ${fieldIndex}: ${fieldLabels[key]}\n`;
    fields[key].forEach(function(opt) { prompt += `- ${opt}\n`; });
    prompt += `\n`;
    fieldIndex++;
  });

  // Fetch article content from site-export sheet
  const exportSheet = ss.getSheetByName('site-export');
  let articleContent = '[Article content not found]';
  if (exportSheet) {
    const exportHeaders = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn()).getValues()[0];
    const exportIdCol = exportHeaders.indexOf('ID');
    const exportHtmlCol = exportHeaders.indexOf('Full Post HTML');
    const postIdCol = headers.indexOf('Post ID');
    const postId = String(rowData[postIdCol]).trim();
    if (exportIdCol > -1 && exportHtmlCol > -1 && postId) {
      const exportData = exportSheet.getRange(2, 1, exportSheet.getLastRow() - 1, exportSheet.getLastColumn()).getValues();
      for (let i = 0; i < exportData.length; i++) {
        if (String(exportData[i][exportIdCol]).trim() === postId) {
          articleContent = exportData[i][exportHtmlCol];
          break;
        }
      }
    }
  }

  prompt += `---\n\nARTICLE CONTENT:\n\n${articleContent}`;

  return { prompt: prompt, articleType: articleType, stoneType: stoneType };
}
function fiSuggestTitlePrompt() {
  try {
    var d = getActiveRowDataMap();
    var material       = String(d["Stone Type"]              || "").trim();
    var articleType    = String(d["Article Type"]            || "").trim();
    var locality       = String(d["Locality"]                || "").trim();
    var h1             = String(d["New H1"]                  || "").trim();
    var storyDir       = String(d["ac_narrative_archetype"]  || "").trim();
    var entryCondition = String(d["ac_entry_condition"]      || "").trim();

    var prompt = "I need a punchy editorial middle title for a featured image on a UK floor restoration website.\n\n";
    prompt += "Rules:\n";
    prompt += "- 4 to 6 words maximum\n";
    prompt += "- Magazine cover line style — not a keyword or SEO phrase\n";
    prompt += "- Must intrigue the reader and hint at the story without giving it away\n";
    prompt += "- No quotes, no punctuation at the end\n\n";
    prompt += "Article details:\n";
    prompt += "Material: " + material + "\n";
    prompt += "Article type: " + articleType + "\n";
    if (locality)       prompt += "Location: " + locality + "\n";
    if (h1)             prompt += "H1: " + h1 + "\n";
    if (storyDir)       prompt += "Narrative direction: " + storyDir + "\n";
    if (entryCondition) prompt += "Floor condition: " + entryCondition + "\n";
    prompt += "\nSuggest 5 options. Return only the five title options, one per line, numbered 1 to 5.";

    return { prompt: prompt };

  } catch(e) {
    return { error: 'fiSuggestTitlePrompt error: ' + e.message };
  }
}
function saveW2BHtmlToSheet(html) {
  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const sheet =
      ss.getSheetByName("posts");

    const row =
      sheet.getActiveRange().getRow();

    if (row < 2) {
      return {
        success: false,
        message: "Select a data row first."
      };
    }

    const sectionPlan =
      String(
        sheet.getRange(row, 102).getValue() || ""
      ).trim();

    if (!sectionPlan) {
      return {
        success: false,
        message:
          "Column CX is empty — W1.5D must complete before W2B."
      };
    }

    const optimizations =
      String(
        sheet.getRange(row, 103).getValue() || ""
      ).trim();

    if (!optimizations) {
      return {
        success: false,
        message:
          "Column CY is empty — W1.5E must complete before W2B."
      };
    }

    let cleanedHtml =
      ce_stripHtmlCodeFences_(
        html
      );

    if (!cleanedHtml) {
      return {
        success: false,
        message:
          "W2B returned empty HTML — EU not changed."
      };
    }

    const errors = [];

    // --------------------------------------------
    // BASIC HTML STRUCTURE
    // --------------------------------------------

    if (
      !/^<header[\s>]/i.test(
        cleanedHtml
      )
    ) {
      errors.push(
        "W2B HTML does not begin with a <header> block."
      );
    }

    if (
      !/<\/header>/i.test(
        cleanedHtml
      )
    ) {
      errors.push(
        "W2B HTML is missing the closing </header> tag."
      );
    }

    // --------------------------------------------
    // GOVERNED SECTION SET FROM CX
    // --------------------------------------------

    const expectedSections = [];
    const expectedH2s = {};

    const planRegex =
      /^SECTION\s+(\d+):\s*(?:Heading H2:\s*)?(.+)$/gmi;

    let planMatch;

    while (
      (
        planMatch =
          planRegex.exec(
            sectionPlan
          )
      ) !== null
    ) {

      const number =
        Number(
          planMatch[1]
        );

      const heading =
        String(
          planMatch[2] || ""
        ).trim();

      expectedSections.push(
        number
      );

      expectedH2s[number] =
        heading;
    }

    if (
      expectedSections.length === 0
    ) {
      return {
        success: false,
        message:
          "Could not determine governed sections from CX."
      };
    }

    // --------------------------------------------
    // EXTRACT GENERATED SECTIONS + H2s
    // --------------------------------------------

    const generatedSections = [];

    const sectionRegex =
      /<section\b[^>]*>[\s\S]*?<\/section>/gi;

    let sectionMatch;

    while (
      (
        sectionMatch =
          sectionRegex.exec(
            cleanedHtml
          )
      ) !== null
    ) {

      const block =
      sectionMatch[0];

    // Hub Page navigation block is intentionally not a governed H2 section.
    if (
      /<section\b[^>]*\bid=["']hub-intro["']/i.test(block)
    ) {
      continue;
    }

    const h2Match =
      block.match(
        /<h2\b[^>]*>([\s\S]*?)<\/h2>/i
      );

      if (!h2Match) {
        errors.push(
          "A generated <section> block is missing its H2."
        );

        continue;
      }

      const heading =
        String(
          h2Match[1] || ""
        )
          .replace(/<[^>]+>/g, "")
          .replace(/&amp;/g, "&")
          .replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'")
          .replace(/&nbsp;/g, " ")
          .trim();

      generatedSections.push(
        heading
      );
    }

    if (
      generatedSections.length !==
      expectedSections.length
    ) {
      errors.push(
        "Expected " +
        expectedSections.length +
        " governed section(s), but W2B generated " +
        generatedSections.length +
        "."
      );
    }

    // --------------------------------------------
    // H2 LOCK
    // --------------------------------------------

    expectedSections.forEach(
      function(number, index) {

        if (
          typeof generatedSections[index] ===
          "undefined"
        ) {
          return;
        }

        if (
          generatedSections[index] !==
          expectedH2s[number]
        ) {
          errors.push(
            "SECTION " +
            number +
            " H2 does not exactly match the governed CX heading."
          );
        }
      }
    );

    // --------------------------------------------
    // INTERNAL LINK PRESERVATION
    // --------------------------------------------

    const governedLinks = [];

    sectionPlan
      .split("\n")
      .forEach(
        function(line) {

          const match =
            String(line)
              .trim()
              .match(
                /^Internal link:\s*(.+)$/i
              );

          if (!match) {
            return;
          }

          const value =
            String(
              match[1] || ""
            ).trim();

          if (
            value &&
            value.toLowerCase() !== "none"
          ) {
            governedLinks.push(
              value
            );
          }
        }
      );

    governedLinks.forEach(
      function(url) {

        if (
          cleanedHtml.indexOf(
            url
          ) === -1
        ) {
          errors.push(
            "Governed internal link missing from W2B HTML: " +
            url
          );
        }
      }
    );

    // --------------------------------------------
    // REJECT BEFORE WRITING EU
    // --------------------------------------------

    if (
      errors.length > 0
    ) {
      return {
        success: false,
        message:
          "W2B HTML rejected — EU not changed. " +
          errors.join(" | ")
      };
    }

    const headers =
      sheet
        .getRange(
          1,
          1,
          1,
          sheet.getLastColumn()
        )
        .getValues()[0]
        .map(function(h) {
          return String(h).trim();
        });

    const colIdx =
      headers.indexOf(
        "W2B Raw HTML"
      );

    if (colIdx === -1) {
      return {
        success: false,
        message:
          "Column 'W2B Raw HTML' not found in posts sheet."
      };
    }

    const cell =
      sheet.getRange(
        row,
        colIdx + 1
      );

    cell.setNumberFormat("@");
    cell.setValue(
      cleanedHtml
    );

    logPipelineResume(
      "W2B — Raw HTML Saved",
      ""
    );

    return {
      success: true,
      message:
        "Validated W2B Raw HTML saved to row " +
        row +
        " — " +
        expectedSections.length +
        " governed section(s) confirmed."
    };

  } catch (e) {

    return {
      success: false,
      message:
        "saveW2BHtmlToSheet error: " +
        e.message
    };
  }
}

function getW2BHtmlFromSheet() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName("posts");
    const row     = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                         .map(function(h) { return String(h).trim(); });
    const colIdx  = headers.indexOf("W2B Raw HTML");
    if (colIdx === -1) return { success: false, message: "Column 'W2B Raw HTML' not found in posts sheet." };

    const html = String(sheet.getRange(row, colIdx + 1).getValue() || "").trim();
    if (!html) return { success: false, message: "No W2B Raw HTML found for this row — save from W2B first." };

    return { success: true, html: html };
  } catch(e) {
    return { success: false, message: "getW2BHtmlFromSheet error: " + e.message };
  }
}

function pushLocationContextToSheet(locality, parentArea, locationContext) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                         .map(function(h) { return String(h).trim(); });

    var saved = [];

    var locIdx = headers.indexOf("Locality");
    if (locIdx > -1 && locality) {
      var locCell = sheet.getRange(row, locIdx + 1);
      locCell.setNumberFormat("@");
      locCell.setValue(locality.trim());
      saved.push("Locality");
    }

    var parIdx = headers.indexOf("Parent Area");
    if (parIdx > -1 && parentArea) {
      var parCell = sheet.getRange(row, parIdx + 1);
      parCell.setNumberFormat("@");
      parCell.setValue(parentArea.trim());
      saved.push("Parent Area");
    }

    var ctxIdx = headers.indexOf("Location Context");
    if (ctxIdx > -1 && locationContext) {
      var ctxCell = sheet.getRange(row, ctxIdx + 1);
      ctxCell.setNumberFormat("@");
      ctxCell.setValue(locationContext.trim());
      saved.push("Location Context");
    }

    if (saved.indexOf("Location Context") > -1) logPipelineResume("W0B — Location Context", "");
    return { success: saved.length > 0, message: "Saved: " + saved.join(", ") };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/* ============================================================
   RECALL RECOMMENDATION BY ARTICLE TYPE
   Reads threshold rules from Article Type Control Sheet.
   Returns a recommendation string appropriate for the page type.
============================================================ */
function buildRecallRecommendation(score, articleType, matched, unmatched) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Article Type Control Sheet");
    if (!sheet) return "Recall recommendation unavailable — Article Type Control Sheet not found.";

    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });

    var typeCol    = headers.indexOf("Canonical Article Type Label");
    var measureCol = headers.indexOf("Recall Measure");
    var logicCol   = headers.indexOf("Recall Recommendation Logic");

    if (typeCol === -1 || measureCol === -1 || logicCol === -1) {
      return "Recall recommendation unavailable — columns not found in Article Type Control Sheet.";
    }

    var measure = "";
    var logic   = "";

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][typeCol]).trim() === articleType) {
        measure = String(data[i][measureCol] || "").trim();
        logic   = String(data[i][logicCol]   || "").trim();
        break;
      }
    }

    if (!logic) return "No recall recommendation logic found for article type: " + articleType;

    var unmatchedNote = unmatched && unmatched.length > 0
      ? " Missing terms: " + unmatched.join(", ") + "."
      : "";

    return "RECALL MEASURE: " + measure + " | " +
           "SCORE: " + score + " | " +
           logic + unmatchedNote;

  } catch(e) {
    return "buildRecallRecommendation error: " + e.message;
  }
}

/* ============================================================
   ENTITY DUMP DETECTOR
   Scans paragraphs for comma-separated noun sequences of 5+
   items with no verb between commas — signals keyword stuffing.
   Used by W7 Run Analysis and W8A Edit Brief.
============================================================ */
function detectEntityDumps(html) {
  var dumps = [];

  var VERB_PATTERN = /\b(is|are|was|were|has|have|had|does|do|did|makes|make|made|causes|cause|prevents|prevent|reduces|reduce|creates|create|allows|allow|requires|require|affects|affect|helps|help|gives|give|means|mean|produces|produce|involves|involve|explains|explain|shows|show|needs|need|changes|change|adds|add|keeps|keep|holds|hold|traps|trap|leaves|leave|forms|form|builds|build|breaks|break|moves|move|stays|stay|comes|come|goes|go|looks|look|feels|feel|works|work|contains|contain|includes|include|shapes|shape|appears|appear|can|could|will|would|should|may|might|must)\b/i;

  var sectionRe = /<section[^>]+id="([^"]+)"[^>]*>([\s\S]*?)<\/section>/gi;
  var sm;
  while ((sm = sectionRe.exec(html)) !== null) {
    var sectionId      = sm[1];
    var sectionContent = sm[2];

    var paraRe = /<p[^>]*>([\s\S]*?)<\/p>/gi;
    var pm;
    while ((pm = paraRe.exec(sectionContent)) !== null) {
      var paraHtml  = pm[1];
      var plainText = paraHtml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

      // Split into sentences first — a verb anywhere in the sentence
      // clears the whole sentence, even if it arrives after a comma list.
      var sentences = plainText.match(/[^.!?]+[.!?]*/g) || [plainText];

      sentences.forEach(function(sentence) {
        var segments = sentence.split(',');
        if (segments.length < 5) return;

        // If the sentence has a verb ANYWHERE, it is not a bare noun dump —
        // list-then-verb constructions ("A, B, C and D can all appear...")
        // are normal English, not keyword stuffing.
        if (VERB_PATTERN.test(sentence)) return;

        dumps.push({
          sectionId: sectionId,
          snippet:   sentence.trim().substring(0, 120)
        });
      });
    }
  }

  return dumps;
}

function checkSectionLoad(wordBudgetMax, hasImage, visualPattern, entityCount, hasLink) {
  var used = 0;
  if (hasImage) used += 40;
  if (visualPattern && visualPattern !== "None") used += 30;
  used += entityCount * 15;
  if (hasLink) used += 25;

  var remaining = wordBudgetMax - used;

  return {
    remaining: remaining,
    overloaded: remaining < 60
  };
}

function bc_getActiveRowNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  if (!sheet) return { success: false, message: 'posts sheet not found.' };
  const row = sheet.getActiveCell().getRow();
  return { success: true, row: row };
}
function bc_getActiveRowNumberAndTitle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  if (!sheet) return { success: false, message: 'posts sheet not found.' };
  const row = sheet.getActiveCell().getRow();
  const title = String(sheet.getRange('B' + row).getValue() || '').trim();
  return { success: true, row: row, title: title };
}

function getArticleAngleOptions(articleType) {

  const libraries = {

    'Case Study': [
      'Diagnostic-Cause-Led',
      'Chronological/Step-by-Step',
      'Contrast/Before-After-Led',
      'Assessment-Pathway/Self-Test-Led',
      'Finish-Recovery/Shine-Led'
    ],

    'Method Guide': [
      'Sequential-Workflow-Led',
      'Risk-Control-Led',
      'Decision-Gate-Led',
      'Preparation-to-Completion-Led',
      'Problem-Solution-Workflow-Led'
    ],

    'Diagnostic Guide': [
      'Symptom-to-Cause-Led',
      'Cause-Elimination-Led',
      'Severity-Escalation-Led',
      'Pattern-Recognition-Led',
      'Diagnosis-to-Remedy-Led'
    ],

    'Buyer Guide': [
      'Comparison-Led',
      'Decision-Criteria-Led',
      'Risk-Trade-Off-Led',
      'Suitability-Pathway-Led',
      'Cost-Value-Led'
    ],

    'Educational Guide': [
      'Explanation-Led',
      'Question-Led',
      'Misconception-Correction-Led',
      'Behaviour-Consequence-Led',
      'Problem-Understanding-Led'
    ],

    'Service Page': [
      'Problem-to-Service-Led',
      'Assessment-Pathway-Led',
      'Suitability-Scope-Led',
      'Outcome-Pathway-Led',
      'Trust-Process-Led'
    ],

    'Geo Service Page': [
      'Local-Problem-to-Service-Led',
      'Local-Assessment-Pathway-Led',
      'Local-Suitability-Scope-Led',
      'Local-Trust-Delivery-Led',
      'Local-Condition-Outcome-Led'
    ],

    'Hub Page': [
      'Topic-Map-Led',
      'Problem-Family-Led',
      'Lifecycle-Led',
      'Decision-Navigation-Led',
      'Material-Behaviour-Map-Led'
    ]
  };

  return libraries[articleType] || [];
}


function getArticleAngleForActiveRow() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  const row = sh.getActiveCell().getRow();

  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  const angleCol = headers.indexOf('Article Angle');
  const typeCol = headers.indexOf('Article Type');

  if (angleCol === -1) {
    return {
      success: false,
      message: 'Article Angle column not found.'
    };
  }

  if (typeCol === -1) {
    return {
      success: false,
      message: 'Article Type column not found.'
    };
  }

  const articleType = String(
    sh.getRange(row, typeCol + 1).getValue() || ''
  ).trim();

  const value = String(
    sh.getRange(row, angleCol + 1).getValue() || ''
  ).trim();

  return {
    success: true,
    value: value,
    articleType: articleType,
    options: getArticleAngleOptions(articleType)
  };
}


function getArticleAngleUsageForStoneType(row) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  row = Number(row);

  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  const colArticleType = headers.indexOf('Article Type');
  const colStoneType = headers.indexOf('Stone Type');
  const colAngle = headers.indexOf('Article Angle');

  if (
    colArticleType === -1 ||
    colStoneType === -1 ||
    colAngle === -1
  ) {
    return {
      success: false,
      message: 'Could not find Article Type, Stone Type, or Article Angle columns.'
    };
  }

  const rowData = sh
    .getRange(row, 1, 1, sh.getLastColumn())
    .getValues()[0];

  const thisArticleType = String(
    rowData[colArticleType] || ''
  ).trim();

  const thisStoneType = String(
    rowData[colStoneType] || ''
  ).trim();

  if (!thisArticleType) {
    return {
      success: false,
      message: 'Article Type is empty for this row.'
    };
  }

  if (!thisStoneType) {
    return {
      success: false,
      message: 'Stone Type is empty for this row.'
    };
  }

  const allShapes =
    getArticleAngleOptions(thisArticleType);

  if (!allShapes.length) {
    return {
      success: false,
      message:
        'No Article Angle library exists for Article Type: ' +
        thisArticleType
    };
  }

  const usageCounts = {};

  allShapes.forEach(function(shape) {
    usageCounts[shape] = 0;
  });

  const allData =
    sh.getDataRange().getValues();

  for (let i = 1; i < allData.length; i++) {

    if (i + 1 === row) {
      continue;
    }

    const r = allData[i];

    const articleType = String(
      r[colArticleType] || ''
    ).trim();

    const stoneType = String(
      r[colStoneType] || ''
    ).trim();

    const angle = String(
      r[colAngle] || ''
    ).trim();

    if (articleType !== thisArticleType) {
      continue;
    }

    if (stoneType !== thisStoneType) {
      continue;
    }

    if (!angle) {
      continue;
    }

    if (usageCounts.hasOwnProperty(angle)) {
      usageCounts[angle]++;
    }
  }

  let minCount = Infinity;

  allShapes.forEach(function(shape) {
    if (usageCounts[shape] < minCount) {
      minCount = usageCounts[shape];
    }
  });

  const leastUsed =
    allShapes.filter(function(shape) {
      return usageCounts[shape] === minCount;
    });

  return {
    success: true,
    articleType: thisArticleType,
    stoneType: thisStoneType,
    usageCounts: usageCounts,
    recommended: leastUsed
  };
}



function buildArticleAngleSuggestionPrompt(row) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  row = Number(row);

  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  function getVal(col) {
    var idx = headers.indexOf(col);
    if (idx === -1) return '';
    return String(
      sh.getRange(row, idx + 1).getValue() || ''
    ).trim();
  }

  const articleType = getVal('Article Type');

  if (!articleType) {
    return {
      success: false,
      message: 'Article Type is not set for this row.'
    };
  }

  // =========================================================
  // SAVED AC CLASSIFICATION
  // =========================================================

  const acValues = {
    'What They Found':
      getVal('ac_entry_condition'),

    'Why They Called':
      getVal('ac_homeowner_perception'),

    'What The Floor Was Doing':
      getVal('ac_material_behaviour'),

    'What Forced Careful Decisions':
      getVal('ac_constraint'),

    'What Governed The Approach':
      getVal('ac_process_emphasis'),

    'What Changed For The Homeowner':
      getVal('ac_result_type'),

    'Story Direction':
      getVal('ac_narrative_archetype')
  };

  const hasValues =
    Object.values(acValues).some(function(v) {
      return String(v || '').trim() !== '';
    });

  if (!hasValues) {
    return {
      success: false,
      message:
        'No AC classification values are saved for this row. Run AC first.'
    };
  }

  // =========================================================
  // GOVERNED ARTICLE-ANGLE LIBRARIES
  // =========================================================

  const shapeLibraries = {

    'Case Study': [
      {
        name: 'Diagnostic-Cause-Led',
        def:
          'Open by identifying the technical cause of the problem before showing how it was corrected.'
      },
      {
        name: 'Chronological/Step-by-Step',
        def:
          'Follow the actual project sequence from assessment through treatment to completed result.'
      },
      {
        name: 'Contrast/Before-After-Led',
        def:
          'Structure the story around the contrast between the starting condition and the finished result.'
      },
      {
        name: 'Assessment-Pathway/Self-Test-Led',
        def:
          'Lead with how the problem was assessed and how the evidence determined the treatment pathway.'
      },
      {
        name: 'Finish-Recovery/Shine-Led',
        def:
          'Centre the story on recovering appearance, finish, reflectivity or surface quality.'
      }
    ],

    'Method Guide': [
      {
        name: 'Sequential-Workflow-Led',
        def:
          'Take the reader through the method in the correct practical order from preparation to completion.'
      },
      {
        name: 'Risk-Control-Led',
        def:
          'Organise the guide around what can go wrong, how to avoid damage and when to stop.'
      },
      {
        name: 'Decision-Gate-Led',
        def:
          'Structure the method around checks and decisions that determine the next safe step.'
      },
      {
        name: 'Preparation-to-Completion-Led',
        def:
          'Move from preparation and testing through treatment, rinsing, drying and final inspection.'
      },
      {
        name: 'Problem-Solution-Workflow-Led',
        def:
          'Begin with the practical problem and then explain the complete method that resolves it safely.'
      }
    ],

    'Diagnostic Guide': [
      {
        name: 'Symptom-to-Cause-Led',
        def:
          'Start with what the homeowner can see or feel, then connect those symptoms to likely causes.'
      },
      {
        name: 'Cause-Elimination-Led',
        def:
          'Work through possible causes systematically so the reader can rule them in or out.'
      },
      {
        name: 'Severity-Escalation-Led',
        def:
          'Organise the article by severity, showing what is minor, what requires caution and what needs professional assessment.'
      },
      {
        name: 'Pattern-Recognition-Led',
        def:
          'Help the reader identify the problem from location, appearance, recurrence and behaviour patterns.'
      },
      {
        name: 'Diagnosis-to-Remedy-Led',
        def:
          'Move from recognising the condition to understanding the likely cause and appropriate remedy pathway.'
      }
    ],

    'Buyer Guide': [
      {
        name: 'Comparison-Led',
        def:
          'Compare the available options directly across the factors that matter to the buyer.'
      },
      {
        name: 'Decision-Criteria-Led',
        def:
          'Structure the guide around the criteria the reader should evaluate before choosing.'
      },
      {
        name: 'Risk-Trade-Off-Led',
        def:
          'Show the advantages, limitations and risks associated with each choice.'
      },
      {
        name: 'Suitability-Pathway-Led',
        def:
          'Help the reader determine which option is appropriate for their material, condition, room or priorities.'
      },
      {
        name: 'Cost-Value-Led',
        def:
          'Organise the article around price, durability, maintenance burden and long-term value.'
      }
    ],

    'Educational Guide': [
      {
        name: 'Explanation-Led',
        def:
          'Build understanding from the underlying material behaviour or principle through its practical consequences.'
      },
      {
        name: 'Question-Led',
        def:
          'Organise the article around the main questions a homeowner is trying to answer.'
      },
      {
        name: 'Misconception-Correction-Led',
        def:
          'Address common misunderstandings first, then replace them with the correct explanation.'
      },
      {
        name: 'Behaviour-Consequence-Led',
        def:
          'Explain how the material behaves and connect each behaviour to what the homeowner sees or experiences.'
      },
      {
        name: 'Problem-Understanding-Led',
        def:
          'Start with the reader-recognised problem and build the knowledge needed to understand it correctly.'
      }
    ],

    'Service Page': [
      {
        name: 'Problem-to-Service-Led',
        def:
          'Start with the homeowner problem, explain what professional assessment identifies and show how the service addresses it.'
      },
      {
        name: 'Assessment-Pathway-Led',
        def:
          'Organise the page around professional assessment, scope definition and the appropriate treatment pathway.'
      },
      {
        name: 'Suitability-Scope-Led',
        def:
          'Explain which problems the service is suitable for, what falls within scope and when another approach is required.'
      },
      {
        name: 'Outcome-Pathway-Led',
        def:
          'Connect the existing condition to the professional work required and the realistic outcome.'
      },
      {
        name: 'Trust-Process-Led',
        def:
          'Build confidence by explaining how the work is assessed, planned, carried out and checked.'
      }
    ],

    'Geo Service Page': [
      {
        name: 'Local-Problem-to-Service-Led',
        def:
          'Start with the local homeowner problem and connect it directly to the professional service available in that locality.'
      },
      {
        name: 'Local-Assessment-Pathway-Led',
        def:
          'Explain how local properties are assessed and how the condition determines the service pathway.'
      },
      {
        name: 'Local-Suitability-Scope-Led',
        def:
          'Explain which problems are suitable for the service locally and when professional escalation is required.'
      },
      {
        name: 'Local-Trust-Delivery-Led',
        def:
          'Structure the page around local service delivery, assessment, communication and realistic scope.'
      },
      {
        name: 'Local-Condition-Outcome-Led',
        def:
          'Connect common local property conditions with the professional work needed and expected outcome.'
      }
    ],

    'Hub Page': [
      {
        name: 'Topic-Map-Led',
        def:
          'Organise the page as a clear map of the major subjects the reader may need to explore.'
      },
      {
        name: 'Problem-Family-Led',
        def:
          'Group related homeowner problems into recognisable families and direct readers to the relevant guidance.'
      },
      {
        name: 'Lifecycle-Led',
        def:
          'Structure the hub around the material lifecycle from selection and installation through care, damage, repair and restoration.'
      },
      {
        name: 'Decision-Navigation-Led',
        def:
          'Help readers decide which type of guidance or service they need before directing them deeper into the site.'
      },
      {
        name: 'Material-Behaviour-Map-Led',
        def:
          'Organise the hub around major material behaviours and their practical implications.'
      }
    ]
  };

  const shapeDefs = shapeLibraries[articleType];

  if (!shapeDefs) {
    return {
      success: false,
      message:
        'No Article Angle library exists for Article Type: ' +
        articleType
    };
  }

  // =========================================================
  // BUILD SAVED-AC EVIDENCE BLOCK
  // =========================================================

  let acBlock = '';

  Object.keys(acValues).forEach(function(label) {
    acBlock +=
      label +
      ': ' +
      (acValues[label] || '(not set)') +
      '\n';
  });

  let shapeBlock = '';

  shapeDefs.forEach(function(s) {
    shapeBlock +=
      '- ' +
      s.name +
      ': ' +
      s.def +
      '\n';
  });

  // =========================================================
  // BUILD PROMPT
  // =========================================================

  const prompt =
`You are choosing the structural narrative shape ("Article Angle") for a governed content article.

ARTICLE TYPE:
${articleType}

SAVED AC CLASSIFICATION:
${acBlock}

PERMITTED ARTICLE ANGLE SHAPES FOR THIS ARTICLE TYPE:
${shapeBlock}

TASK:

Evaluate ONLY the permitted shapes listed above.

Base the decision on the saved AC classification, especially:
- What They Found
- What The Floor Was Doing
- What Forced Careful Decisions
- What Governed The Approach
- Story Direction

The Article Angle controls structural shape only.

It must NOT:
- change the Article Type;
- change the governed topic;
- invent facts;
- introduce a new service or intent;
- override the saved AC classification;
- broaden the page beyond its governed role.

For EACH permitted shape output:

SHAPE: [exact permitted shape name]
FIT: YES / NO / PARTIAL
REASON: [one sentence grounded directly in the saved AC classification]

Then output exactly one final line:

RECOMMENDED: [exact permitted shape name]

The RECOMMENDED value MUST exactly match one of the permitted shape names above.

Return no preamble and no markdown.`;

  return {
    success: true,
    articleType: articleType,
    allowedShapes: shapeDefs.map(function(s) {
      return s.name;
    }),
    prompt: prompt
  };
}

function saveArticleAngleForActiveRow(value, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('posts');
  row = Number(row);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });  const col = headers.indexOf('Article Angle');
  if (col === -1) return { success: false, message: "Article Angle column not found." };
  sh.getRange(row, col + 1).setValue(value);
  if (typeof logPipelineResume === 'function') {
    logPipelineResume("Article Angle set — " + value, "");
  }
  return { success: true, message: "Article Angle saved: " + value };
}

/* ============================================================
   SHARED GOVERNANCE — CONCRETE LANGUAGE STANDARD
   Returns the banned-verb/banned-phrase/concrete-language block
   for injection into any fix or generation prompt across the
   pipeline. Single source of truth — edit here only.
============================================================ */
function ce_getConcreteLanguageStandard() {
  return `LANGUAGE STANDARD (Hard Lock): Use concrete, active verbs naming the actual physical task performed: inspect, assess, diagnose, test, grind, hone, polish, seal, grout, patch, fill, repair, replace, strip, reseal, buff, install, protect, maintain, restore. Show how value is delivered through the specific tangible step taken, not a claimed result. Do not use abstract corporate buzzwords, hollow marketing fluff, or vague outcome-claims standing in for a real action.

Never use these banned verbs: boost, elevate, evolving, dominate, transform, leverage, utilize, maximize, optimize, enhance, revolutionize, streamline, unlock, empower, harness, curate, deliver (as in "deliver results"), drive (as in "drive value").

Never use these banned phrases: world-class, cutting-edge, state-of-the-art, best-in-class, tailored solutions, holistic approach, seamless experience, unparalleled results, passionate about, dedicated to excellence, peace of mind (as a floating claim with no action attached).

Examples:
- WEAK: "We transform tired floors." → STRONG: "We grind out lippage, hone the surface flat, and polish to the finish level you choose."
- WEAK: "We optimize your floor's protection." → STRONG: "We test porosity, apply a penetrating sealer, and re-seal grout lines."
- WEAK: "We elevate the look of your travertine." → STRONG: "We fill open pores, hone away etch marks, and polish back the natural shine."
- WEAK: "Our team delivers a seamless restoration experience." → STRONG: "We assess the damage on-site, grind out deep scratches, and buff to match the surrounding floor."
- WEAK: "We revolutionize how homeowners maintain natural stone." → STRONG: "We show you which cleaner to use, how often to reseal, and what to avoid."

Rule of thumb: if a sentence describes an outcome or benefit without naming the specific tool, technique, or material step used to achieve it, it is too abstract — rewrite it around the actual physical action performed on the stone or grout.`;
}

/* ============================================================
   SHARED GOVERNANCE — FIX-BATCH INTEGRITY CHECK
   Prevents redundancy introduced WITHIN a single fix response
   (two separate fixes each restating the same step) and prevents
   orphaned/empty paragraphs left behind when flagged content is
   removed. Single source of truth — edit here only.
============================================================ */
function ce_getFixBatchIntegrityCheck() {
  return `FIX-BATCH INTEGRITY CHECK (Hard Lock): You are likely proposing more than one fix in this response. Before finalising the full set, re-read ALL of your own proposed "new" replacements together, not just each one against the original article in isolation. If two or more of your own fixes would each independently restate the same technical step, fact, or claim (e.g. two separate fixes both adding a sentence about pressure-rinsing, or both adding a sentence about assessing a cavity), keep the fullest, best-placed instance and remove or shorten the others so the corrected article does not end up repeating the same point across nearby paragraphs.

ORPHANED CONTENT CHECK (Hard Lock): If removing or shortening flagged content would leave a dangling lead-in sentence with no follow-through (e.g. "Ordinary soil and coating wear required different responses." with nothing after it), or would leave an empty or near-empty <p> tag (e.g. "<p> </p>"), your "new" replacement must either complete the thought with genuinely supported content, merge the remaining fragment into an adjacent paragraph, or remove the empty/incomplete paragraph entirely — never leave a dangling sentence or blank paragraph tag in the corrected article.`;
}

function ce_getContentPreservationRules() {
  return `CONTENT PRESERVATION RULES (Hard Lock — apply before finalising any fix):
Every article is required to contain three specific elements. If your fix touches a sentence containing any of these, you must preserve the required content — rewrite around the flagged issue, never delete the required content wholesale.

1. UNIVERSAL BENEFIT STATEMENTS — the article must state: (a) the floor will look significantly better, in many cases better than when installed; (b) a professionally restored and correctly sealed floor is significantly easier to clean and maintain than a worn floor; (c) correct ongoing maintenance — pH-neutral cleaning, grit removal before wet mopping, resealing at the right interval — is the single most important factor in extending the floor's life, with a contextual internal link to the maintenance article.

2. MAINTENANCE EDUCATION ELEMENT — the article must include: one sentence on what correct ongoing maintenance achieves for this stone type; one sentence naming one specific thing to avoid and why; a contextual internal link to the maintenance or cleaning tips article at that point.

3. PROFESSIONAL DECISION REASONING — whenever the article states a specific material, product, method, technique, or sealer type was chosen for this job, it must state the reason for that choice, grounded in the floor's condition, homeowner input, or the environment — never a bare fact with no reason. The reason must NOT compare against or disparage an alternative product/method by name.
  WRONG (bare fact): "A film-forming sealer was applied across the stone."
  WRONG (disparages alternative): "A film-forming sealer was chosen rather than the invisible, no-film protection an impregnating sealer gives."
  CORRECT: "We chose a film-forming sealer because the worn surface could not be mechanically honed to a consistent finish, creating a protective topical layer that makes routine maintenance easier."

REFERENCE EXAMPLES (Hard Lock — for Professional Decision Reasoning only):
These are a pattern bank, not text to insert verbatim. Use them to understand the STYLE of reasoning required — active voice, condition-grounded, no comparison — then write a reason that fits the actual material, defect, and decision already established in the surrounding text you are editing. Do not select the closest-sounding example and drop it in; the reasoning must match the specific sentence and paragraph context.

Sealer type:
- "We chose a film-forming sealer because the worn surface could not be mechanically honed to a consistent finish, creating a protective topical layer that makes routine maintenance easier."
- "We chose an impregnating, vapour-permeable sealer because the porous stone retained moisture requiring continued vapour transmission to prevent efflorescence."
- "We chose a satin topical sealer because the porous stone required a surface barrier and the homeowner requested a low-sheen appearance."

Filler type:
- "We chose a two-part resin filler for the cavities because it polishes flush with the surrounding stone to form a seamless repair."
- "We chose colour-matched cementitious grout for the void repairs because the wide, shallow surface holes required a matte finish matching the original jointing."
- "We chose a polyester resin filler for the crack repairs because its structural bond accommodates minimal sub-floor movement without telegraphing through the stone."

Honing grit sequence:
- "We worked diamond pads from a coarse 50 grit up through fine grits to grind out heavy surface scratches and achieve a uniform finish."
- "We chose a 50-grit starting diamond pad to remove deep traffic scratches before beginning surface refinement."
- "We stepped the honing sequence through progressively finer diamond grits to close the stone and generate an even surface for polishing."

Cleaning chemistry:
- "We chose a heavy-duty alkaline degreaser because the surface held ingrained grease that needed a stronger cleaning agent to break down and lift away."
- "We used a fresh-water extraction rinse for the final pass to remove chemical residues and restore a neutral pH balance to the stone."
- "We used an extended dwell time to help emulsify compacted soil sitting in the surface."

Extraction method:
- "We used wet-vacuum extraction because the loosened contamination needed to be removed before it resettled in the recesses."
- "We used multiple extraction passes because deep-set soil in the textured crevices required repeated flushing to clear fully."
- "We chose enclosed high-pressure rotary extraction because the deep-set grout joints held compacted particulate matter within their textured edges."

Repair technique:
- "We used masking around each repair to keep the material confined and reduce abrasion to sound stone."
- "We struck the wet grout joints slightly recessed to match the profile of the original installation."
- "We filled the crack before honing so the repair could be honed along with the floor."

If your fix would remove or shorten a sentence carrying any of the three required elements above, rewrite it to preserve the required content — do not delete it because it sits next to the flagged issue.`;
}

function ce_getSourceEvidencePriorityLock() {
  return `SOURCE-EVIDENCE PRIORITY LOCK (Hard Lock):
The Rewrite Brief defines required topics and editorial direction only. It is not factual evidence that a named product, grit, method, material, sequence, location or outcome was used on the project.

Before adding or confirming any technical detail, verify that it appears explicitly in the supplied article, source notes or project evidence. Never convert a Rewrite Brief label into a project fact.

If the brief names a detail that is not supported by the supplied evidence:
- do not insert it;
- do not infer it;
- do not invent a compatible version;
- mark it as unsupported or unresolved;
- preserve the nearest supported wording already present.

A RE-ANCHOR requirement means "check for supported coverage," not "manufacture coverage."

Evidence priority:
1. Original project notes or source material
2. Existing article statements supported by those notes
3. Rewrite Brief
4. General technical knowledge

Lower-priority material must never override or add facts absent from higher-priority evidence.

Example:
Rewrite Brief: "Diamond Abrasive Pads (400–3000)"
Project evidence: "Repairs were locally refined with diamond pads."
Permitted wording: "The repairs were locally refined with diamond abrasive pads."
Prohibited wording: "The repairs were honed from 400 through 3000 grit" unless that exact grit sequence is confirmed in the project evidence.`;
}

function ce_getCustomerFacingCopyOnlyLock() {
  return `CUSTOMER-FACING COPY ONLY (Hard Lock):
All returned article text must read as finished customer-facing website content.

Never mention:
- the Rewrite Brief;
- source notes, project evidence or supplied evidence;
- audits, compliance checks or failing items;
- whether a detail is recorded, confirmed, supported, unsupported or unresolved;
- editorial decisions, prompt instructions or factual-verification limits.

When a required detail is more specific than the available evidence, use only the nearest supported project wording. Do not explain the omission inside the article.

Example:
WRONG: "The project evidence confirms diamond refinement but does not record a 400–3000 grit sequence."
CORRECT: "Before sealing, the repaired areas were locally levelled with diamond abrasive pads rather than conventional floor sanding."

FINAL META-COMMENTARY CHECK: Before returning the HTML, remove any sentence that explains why wording was included, excluded, qualified or not verified. Every sentence must describe the floor, its condition, the work performed, the homeowner's experience or the completed result.`;
}

function ce_getProjectGroundedMaintenanceLock() {
  return `PROJECT-GROUNDED MAINTENANCE VS REUSABLE DIY ADVICE (Hard Lock):

The Maintenance Education requirement and the "no reusable DIY advice" drift boundary are both mandatory and must be satisfied together, never traded off against each other.

REUSABLE DIY ADVICE (must not appear):
generic imperative instruction that would read identically on any stone floor — e.g. "Remove loose grit before wet mopping, clean with a pH-neutral product and reseal at the correct interval."

PROJECT-GROUNDED MAINTENANCE (required, and compliant):
the same required information narrated as what was actually explained to this specific homeowner about this specific floor's condition, tied to what was just restored — e.g. "Because the new mid-sheen film is a sacrificial layer, I told the homeowner to sweep or vacuum loose grit before mopping rather than dragging it across the finish, and to expect a resealing visit once wear became visible on this floor."

The test:
does the sentence read as a standalone instruction list detachable from this project, or as narration of what was actually communicated about this floor? If detachable, rewrite as narration — do not delete the required content to escape the drift flag.`;
}

function ce_getLocationVarietyWithoutFabricationLock() {
  return `LOCATION VARIETY WITHOUT FABRICATION (Hard Lock):

When adjusting content to strengthen relevance to a target location or region, use variety in HOW the location is referenced — not fabricated specifics about what happened there.

PERMITTED:
- Real, verifiable sub-region and town names within the target area, used to vary phrasing instead of repeating one region name mechanically.
- General, uncontroversial geographic framing (e.g. "coastal and rural properties across the area") where genuinely true of the region.

PROHIBITED:
- Inventing specific environmental or technical claims about the region (water hardness, soil composition, climate specifics) and presenting them as established fact, unless sourced from verified project or reference data actually supplied in this prompt.
- Inventing specific property scenarios tied to a named town or sub-region (e.g. "a farm conversion near [town]") when this article has no real underlying project to draw that detail from — this creates a false impression of completed local work that did not happen.
- Repeating the literal target location phrase mechanically through the body text — this is the keyword-stuffing pattern to avoid, but the fix is genuine variety in real place references, never invented specifics.

If in doubt about whether a location-related detail is verifiable, omit it and rely on real town/sub-region names for variety instead.

VERIFIED EVIDENCE EXCEPTION: The prohibition above does not apply to local detail explicitly supplied in this prompt — e.g. the ORIGINAL SOURCE ARTICLE, Recovery Blueprint, or other project data actually provided. If a genuine local project, fact, or characteristic is present in the supplied material, it may be used and is not "invented." The exception never extends to filling a gap with a plausible-sounding detail that is not actually present in the supplied material — plausible is not verified.`;
}
