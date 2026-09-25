/**
 * ============================================================
 * bc_SerpEntityBridge.gs
 * Abbey Floor Care — Governance Prompt Assembler
 * SERP Entity Bridge — validates live SERP entities against
 * the governed Recovery Blueprint before content generation.
 *
 * Sits between BC governance (Prompt 2 Load) and CE content
 * pipeline. Run after Recovery Blueprint is built in col DJ.
 *
 * Contains:
 * - bc_buildSerpBridgePrompt()
 * - bc_pushSerpBridgeToSheet()
 * - bc_addSerpEntitiesToBlueprint()
 *
 * Depends on: bc_GovernancePromptAssembler_Config.gs
 *             bc_DataFetcher.gs
 * ============================================================
 */

function bc_buildSerpBridgePrompt() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(
      s => s.getSheetId() == BC_SHEET_CONFIG.posts
    );
    if (!sheet) return { success: false, message: "Posts sheet not found." };

    const row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());

    const getVal = name => {
      const idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
    };

    const material     = getVal("Stone Type");
    const articleType  = getVal("Article Type");
    const intent       = getVal("Topical Intent");
    const rawPrimaryTerm = getVal("Primary Search Term");

    let primaryTerm = rawPrimaryTerm;
    if (articleType === "Case Study") {
      const intentWords = intent.toLowerCase().replace(/[^a-z\s]/g, "").split(/\s+/).slice(0, 3).join(" ");
      const materialLower = material.toLowerCase();
      const intentClean = intentWords.replace(new RegExp('^' + materialLower + '\\s*', 'i'), '').trim();
        const confirmedIntent = getVal("Confirmed Primary Intent") || getVal("Primary Intent") || "cleaning";
        const fallback = confirmedIntent.toLowerCase().split(" ").slice(0, 2).join(" ");
        primaryTerm = materialLower + " " + (intentClean || fallback);
    }
    const blueprint    = getVal("Recovery Blueprint");
    const title        = getVal("Title");

    if (!primaryTerm) {
      return { success: false, message: "Primary Search Term is empty. Run Prompt 2 Load first, then run W0 Generate in the CE pipeline if needed." };
    }
    if (!blueprint || blueprint.indexOf("RECOVERY BLUEPRINT MISSING") > -1) {
      return { success: false, message: "Recovery Blueprint is empty. Run Prompt 2 Load first." };
    }

    // Extract full TSM for this material
    let tsmOverview = "";
    let tsmEntityNames = "";
    let tsmGuardrails = "";

    try {
      const techSheet = bc_getSheetByConfigKeys_(["techSpec"]);
      if (techSheet) {
        const data = techSheet.getDataRange().getValues();
        const materialLower = material.toLowerCase().trim();
        let matchRow = null;
        for (let i = 0; i < data.length; i++) {
          if (String(data[i][1] || "").trim().toLowerCase() === materialLower) {
            matchRow = data[i];
            break;
          }
        }
        if (matchRow) {
          const fullTSM = String(matchRow[2] || "");

          // Extract overview section
          const overviewStart = fullTSM.indexOf("OVERVIEW");
          const overviewEnd = fullTSM.indexOf("1. COMMON ISSUES");
          if (overviewStart > -1 && overviewEnd > -1) {
            tsmOverview = fullTSM.substring(overviewStart, overviewEnd).trim();
          }

          // Extract guardrails
          const guardrailStart = fullTSM.indexOf("PROMPT GOVERNANCE RULES");
          if (guardrailStart > -1) {
            tsmGuardrails = fullTSM.substring(guardrailStart).trim();
          }

          // Read entity names from technical_entities sheet
          try {
            const entitiesSheet = ss.getSheetByName("technical_entities");
            if (entitiesSheet) {
              const entitiesData = entitiesSheet.getDataRange().getValues();
              const entityNames = [];
              const materialLower = material.toLowerCase().trim();
              for (let i = 1; i < entitiesData.length; i++) {
                const rowMaterial = String(entitiesData[i][0] || "").trim().toLowerCase();
                if (rowMaterial === materialLower) {
                  const name = String(entitiesData[i][1] || "").trim();
                  if (name) entityNames.push(name);
                }
              }
              tsmEntityNames = entityNames.join(", ");
            }
          } catch(e) {
            tsmEntityNames = "";
          }
        }
      }
    } catch(e) {
      tsmEntityNames = "";
    }

    // Read SERP Entity Filter Rules from Tech_DNA
    let serpFilterRules = "";
    try {
      const techDnaSheet = bc_getSheetByConfigKeys_(["techDNA", "techDna"]);
      if (techDnaSheet) {
        const dnaData = techDnaSheet.getDataRange().getValues();
        const dnaHeaders = dnaData[0].map(h => String(h).trim());
        const stoneIdx = dnaHeaders.indexOf("Stone Type");
        if (stoneIdx === -1) dnaHeaders.indexOf("Material");
        const filterIdx = dnaHeaders.indexOf("SERP Entity Filter Rules");
        const materialLower = material.toLowerCase().trim();
        if (filterIdx > -1) {
          for (let i = 1; i < dnaData.length; i++) {
            const rowMaterial = String(dnaData[i][stoneIdx] || "").trim().toLowerCase();
            if (rowMaterial === materialLower) {
              serpFilterRules = String(dnaData[i][filterIdx] || "").trim();
              break;
            }
          }
        }
      }
    } catch(e) {
      serpFilterRules = "";
    }

    const serpFilterBlock = serpFilterRules
      ? `PRE-OUTPUT FILTER RULES (apply before deciding whether to include any entity):
${serpFilterRules}

`
      : "";

    const tsmOverviewBlock = tsmOverview
      ? `MATERIAL CONTEXT (use this to judge entity relevance):
${tsmOverview}

`
      : "";

    const tsmGuardrailBlock = tsmGuardrails
      ? `MATERIAL GUARDRAILS (entities or processes explicitly prohibited for this material):
${tsmGuardrails}

`
      : "";

    const tsmExclusionBlock = tsmEntityNames
      ? `FULL TSM ENTITY LIST (already governed — do not report any of these or functionally equivalent alternatives):
${tsmEntityNames}

`
      : "";

    const prompt = `SERP ENTITY BRIDGE — COMPREHENSIVE SERP ANALYSIS
ROLE: Senior SEO Analyst with material science knowledge

SEARCH CONTEXT:
Perform all searches using UK search context only.
Use google.co.uk results.
All competitor and entity signals must reflect UK search rankings only.
Ignore US, Australian, or other non-UK results.

TASK:
Retrieve the top 10 Google results for the search term below. Extract every named entity from every inspected page AND the technical vocabulary that surrounds each entity. Use the material context provided to judge relevance. Report all genuinely new entities not already covered by the governed lists.

MANDATORY: You MUST use your web search tool for this task. Do not use training data. If you cannot perform a live web search, output exactly: WEB SEARCH UNAVAILABLE — cannot complete this task without live search access.

PRIMARY SEARCH TERM:
${primaryTerm}

CONTEXT:
Material: ${material}
Article Type: ${articleType}
Intent: ${intent}

${tsmOverviewBlock}${tsmGuardrailBlock}RECOVERY BLUEPRINT ENTITIES (already covered — do not report these):
${blueprint}

${tsmExclusionBlock}RELEVANCE FRAMEWORK:
Use the material context above to judge whether each extracted entity is relevant to this material. An entity is relevant if it relates to any of the following for this specific material:
- Material behaviour, structure, or properties
- Diagnosis of conditions or defects
- Cleaning chemistry or process
- Restoration or repair methods
- Sealing or protection systems
- Maintenance or aftercare
- Tools or equipment used in professional work
- Structural features that affect treatment decisions

An entity is NOT relevant if it:
- Is sourced from abbeyfloorcare.co.uk — exclude this domain entirely from inspection
- Is generic to all floor types without material-specific significance
- Is explicitly prohibited by the material guardrails above
- Is a brand name or product name rather than a category entity
- Is a location name, company name, or person name
- Relates to structural building engineering, wall repairs, masonry anchoring, or load-bearing structures
- Applies to external stonework, cladding, or facades rather than interior floor surfaces
- Relates to building conservation planning constraints, listed building consent, or statutory regulatory frameworks
- Is a governance or regulatory framework rather than a named technical process, material, chemistry, or tool
- Applies to structural building integrity, load-bearing walls, masonry anchoring, or building movement unrelated to floor tile behaviour
- Is an abstract material science metric or engineering property without a direct content application (e.g. tensile strength, modulus of elasticity, compressive strength)
- Is a installation specification or building standard rather than a named professional process, tool, chemistry, or condition observable on a finished floor

ENTITY CATEGORIES (classify every new entity into one of these):
Chemistry | Coating | Tool/Process | Material Property | Common Issue | Structural Feature | Technical Constraint

SEMANTIC FINGERPRINT EXTRACTION:
For each new entity, extract the 8-12 technical terms that consistently appear within 50 words of that entity across the inspected sources. These terms are the "semantic fingerprint" that validates topical authority in LSI (Latent Semantic Indexing) models.

Example:
Entity: "Efflorescence"
Semantic Fingerprint: crystalline, powdery, dissolved salts, upward migration, surface deposits, evaporation, moisture movement, diagnostic signal, white deposits

${serpFilterBlock}INSTRUCTIONS:
1. Use web search to retrieve the current top 10 Google results for: ${primaryTerm}
   Exclude any results from abbeyfloorcare.co.uk — do not inspect Abbey Floor Care's own pages.
2. Cite each URL you inspect — do not infer from prior knowledge.
3. Extract EVERY named entity from EVERY inspected page that passes the relevance framework above.
4. For each entity, extract the 8-12 technical terms that appear within 50 words of it across sources.
5. Compare the full extracted list against both exclusion lists above — including functionally equivalent or alternative names for entities already listed.
6. Report only entities that are genuinely absent from both exclusion lists.
7. For each new entity assign an Action:
   TSM only — add to RECOGNISED TECHNICAL ENTITIES table, do not add to Recovery Blueprint
   TSM + Blueprint Priority — add to RECOGNISED TECHNICAL ENTITIES table AND Recovery Blueprint Priority tier
   TSM + Blueprint Supporting — add to RECOGNISED TECHNICAL ENTITIES table AND Recovery Blueprint Supporting tier

CRITICAL: Use web search. Do not rely on training data for live SERP analysis.
If web search is unavailable, state that explicitly rather than guessing.
Do not summarise your intentions or confirm your understanding before running. Proceed immediately and output the full results following the strict output format above. Do not ask for confirmation or approval to begin.

VOCABULARY BENCHMARK TASK (run alongside entity extraction — same pages):
While inspecting each page, also collect:
1. The word count of each inspected page body content (excluding navigation, headers, footers)
2. Every technical term that appears across 3 or more of the inspected pages

From this data calculate:
- Average word count across all inspected pages
- The top 20 most frequently occurring technical terms across the inspected pages that are NOT already in the Recovery Blueprint or TSM entity list above

OUTPUT THESE AS THREE ADDITIONAL BLOCKS AFTER THE ACTION REQUIRED BLOCK:

---COMPETITOR NODE MAP---
For each inspected URL, classify the page into one of these node types based on its structure and intent:
Evidence Node (case study / project proof)
Process Node (how-to / method guide)
Property Node (what is / educational explainer)
Decision Node (buyer guide / cost page)
Comparison Node (vs / alternative / which is best)
Location Node (geo service page)
Problem Node (diagnosis / fault finding)
Hub Node (silo hub / topic overview)

Output as:
URL | Node Type | Primary Intent (one phrase)
---END COMPETITOR NODE MAP---

---VOCAB BENCHMARK---
AVERAGE WORD COUNT: [number]
COMPETITOR VOCAB GAP: [comma-separated list of up to 20 technical terms that appear on 3+ competitor pages but are absent from the Recovery Blueprint and TSM entity list above]
---END VOCAB BENCHMARK---

OUTPUT FORMAT (STRICT — follow exactly):
Line 1: Comma-separated list of new entities only.
Line 2: Sources inspected (comma-separated URLs).
Line 3 onward — one line per new entity:
ENTITY: [name] | TYPE: [type] | SEMANTIC FINGERPRINT: [8-12 technical terms that appear within 50 words of this entity across inspected sources, comma-separated] | ACTION: [TSM only / TSM + Blueprint Priority / TSM + Blueprint Supporting] | REASON: [one sentence]

If no new entities found, output exactly:
NONE
---ACTION REQUIRED---
No new entities found. No technical sheet update needed. No blueprint update needed.
---END ACTION---

After the entity lines, output this action block. Fill in the actual entity rows — do not use placeholders:

---ACTION REQUIRED---
Open the technical sheet in your spreadsheet.
Find the row for ${material}.
Add the following rows to the bottom of the RECOGNISED TECHNICAL ENTITIES table using this exact pipe-separated format (4 columns):

[Output one line per new entity — fill in real values, not placeholders:]
Exact Entity Name | Type | One sentence describing primary application for this material | Semantic Fingerprint Vocabulary (8-12 terms, comma-separated)

After updating the technical sheet, click Add to Blueprint in the sidebar to automatically update the Recovery Blueprint for entities marked TSM + Blueprint Priority or TSM + Blueprint Supporting.
Then clear the pipeline columns for this row from Stage 0 onward and rerun.
---END ACTION---`;

    return { success: true, prompt: prompt, title: title };

  } catch(e) {
    return { success: false, message: "bc_buildSerpBridgePrompt error: " + e.message };
  }
}


function bc_pushSerpBridgeToSheet(raw, saveResponse) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(
      s => s.getSheetId() == BC_SHEET_CONFIG.posts
    );
    if (!sheet) return { success: false, message: "Posts sheet not found." };

    const row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());

    const colIdx = headers.indexOf("SERP Validation");
    if (colIdx === -1) return { success: false, message: "Column 'SERP Validation' not found." };

    const cleaned = String(raw || "").trim();
    if (!cleaned) return { success: false, message: "No output to push — paste LLM response first." };

    if (cleaned === "NONE") {
      sheet.getRange(row, colIdx + 1).setValue("NONE — no new entities found");
      return { success: true, message: "No new entities. SERP Validation updated." };
    }

    const cell = sheet.getRange(row, colIdx + 1);
    try {
      cell.setPlainTextValue(cleaned);
    } catch(e) {
      cell.setNumberFormat("@");
      cell.setValue(cleaned);
    }

    if (saveResponse) {
      const responseColIdx = headers.map(function(h) { return String(h).toLowerCase().trim(); }).indexOf('serp response');
      if (responseColIdx > -1) {
        sheet.getRange(row, responseColIdx + 1).setValue(String(raw).trim());
      }
    }

    // Parse and write Competitor Vocab Gap and Competitor Word Count
    const vocabMatch = cleaned.match(/---VOCAB BENCHMARK---([\s\S]*?)---END VOCAB BENCHMARK---/i);
    const nodeMapMatch = cleaned.match(/---COMPETITOR NODE MAP---([\s\S]*?)---END COMPETITOR NODE MAP---/i);
    if (nodeMapMatch) {
      const nodeMapIdx = headers.indexOf("Competitor Node Map");
      if (nodeMapIdx > -1) {
        const nodeMapCell = sheet.getRange(row, nodeMapIdx + 1);
        nodeMapCell.setNumberFormat("@");
        nodeMapCell.setValue(nodeMapMatch[1].trim());
      }
    }

    if (vocabMatch) {
      const vocabBlock = vocabMatch[1].trim();

      const wordCountMatch = vocabBlock.match(/AVERAGE WORD COUNT:\s*(\d+)/i);
      const vocabGapMatch  = vocabBlock.match(/COMPETITOR VOCAB GAP:\s*(.+)/i);

      const wordCountIdx = headers.indexOf("Competitor Word Count");
      const vocabGapIdx  = headers.indexOf("Competitor Vocab Gap");

      if (wordCountIdx > -1 && wordCountMatch) {
        sheet.getRange(row, wordCountIdx + 1).setValue(wordCountMatch[1].trim());
      }
      if (vocabGapIdx > -1 && vocabGapMatch) {
        const cell = sheet.getRange(row, vocabGapIdx + 1);
        cell.setNumberFormat("@");
        cell.setValue(vocabGapMatch[1].trim());
      }
    }

    logPipelineResume("SERP — Entity Bridge Push", "");

    // Write ELG-2 to posts_progress
    try {
      bc_writeELGToProgress(2);
    } catch(elgErr) {
      Logger.log("ELG-2 write failed: " + elgErr.message);
    }

    return { success: true, message: "SERP Validation saved. Competitor vocab and word count updated. ELG-2 progress recorded." };

  } catch(e) {
    return { success: false, message: "bc_pushSerpBridgeToSheet error: " + e.message };
  }
}


function bc_addSerpEntitiesToBlueprint() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(
      s => s.getSheetId() == BC_SHEET_CONFIG.posts
    );
    if (!sheet) return { success: false, message: "Posts sheet not found." };

    const row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());

    const serpIdx = headers.indexOf("SERP Validation");
    const blueprintIdx = headers.indexOf("Recovery Blueprint");

    if (serpIdx === -1) return { success: false, message: "Column 'SERP Validation' not found." };
    if (blueprintIdx === -1) return { success: false, message: "Column 'Recovery Blueprint' not found." };

    const serpRaw = String(sheet.getRange(row, serpIdx + 1).getValue() || "").trim();
    if (!serpRaw || serpRaw === "NONE — no new entities found") {
      return { success: false, message: "No SERP Validation output found. Push LLM output first." };
    }

    const blueprintRaw = String(sheet.getRange(row, blueprintIdx + 1).getValue() || "").trim();
    if (!blueprintRaw || blueprintRaw.indexOf("RECOVERY BLUEPRINT MISSING") > -1) {
      return { success: false, message: "Recovery Blueprint is empty. Run Prompt 2 Load first." };
    }

    // Parse ENTITY lines from SERP output
    const priorityEntities  = [];
    const supportingEntities = [];

    const lines = serpRaw.split(/\r?\n/);
    lines.forEach(function(line) {
      const trimmed = line.trim();
      if (!trimmed.startsWith("ENTITY:")) return;

      const nameMatch   = trimmed.match(/ENTITY:\s*([^|]+)/i);
     const actionMatch = trimmed.match(/ACTION:\s*([^|]+)/i);

    if (!nameMatch || !actionMatch) return;

    const name   = nameMatch[1].trim();
    // Strip any trailing markdown link references e.g. ([Source][1])
    const action = actionMatch[1].replace(/\s*\([^\)]*\)\s*$/, "").trim().toLowerCase();

      if (action.indexOf("priority") > -1) {
        priorityEntities.push(name);
      } else if (action.indexOf("supporting") > -1) {
        supportingEntities.push(name);
      }
    });

    if (priorityEntities.length === 0 && supportingEntities.length === 0) {
      return { success: false, message: "No Priority or Supporting entities found — nothing to add to Recovery Blueprint." };
    }

    // Update Recovery Blueprint
    let updated = blueprintRaw;

    if (priorityEntities.length > 0) {
      const priorityMarker = "PRIORITY ENTITIES (must appear — anchor co-occurrence targets):";
      const priorityIdx = updated.indexOf(priorityMarker);
      if (priorityIdx > -1) {
        const lineEnd = updated.indexOf("\n", priorityIdx + priorityMarker.length);
        const existingLine = lineEnd > -1
          ? updated.substring(priorityIdx + priorityMarker.length, lineEnd).trim()
          : updated.substring(priorityIdx + priorityMarker.length).trim();
        const newLine = existingLine
          ? existingLine + ", " + priorityEntities.join(", ")
          : priorityEntities.join(", ");
        updated = lineEnd > -1
          ? updated.substring(0, priorityIdx + priorityMarker.length) + "\n" + newLine + updated.substring(lineEnd)
          : updated.substring(0, priorityIdx + priorityMarker.length) + "\n" + newLine;
      }
    }

    if (supportingEntities.length > 0) {
      const supportingMarker = "SUPPORTING ENTITIES (use where contextually relevant):";
      const supportingIdx = updated.indexOf(supportingMarker);
      if (supportingIdx > -1) {
        const lineEnd = updated.indexOf("\n", supportingIdx + supportingMarker.length);
        const existingLine = lineEnd > -1
          ? updated.substring(supportingIdx + supportingMarker.length, lineEnd).trim()
          : updated.substring(supportingIdx + supportingMarker.length).trim();
        const newLine = existingLine
          ? existingLine + ", " + supportingEntities.join(", ")
          : supportingEntities.join(", ");
        updated = lineEnd > -1
          ? updated.substring(0, supportingIdx + supportingMarker.length) + "\n" + newLine + updated.substring(lineEnd)
          : updated.substring(0, supportingIdx + supportingMarker.length) + "\n" + newLine;
      }
    }

    const cell = sheet.getRange(row, blueprintIdx + 1);
    try {
      cell.setPlainTextValue(updated);
    } catch(e) {
      cell.setNumberFormat("@");
      cell.setValue(updated);
    }

    const added = [...priorityEntities, ...supportingEntities];
    return {
      success: true,
      message: "Recovery Blueprint updated. Added: " + added.join(", ")
    };

  } catch(e) {
    return { success: false, message: "bc_addSerpEntitiesToBlueprint error: " + e.message };
  }
}
function bc_pushEntitiesToTSM(newRowsText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const postsSheet = ss.getSheets().find(
      s => s.getSheetId() == BC_SHEET_CONFIG.posts
    );
    if (!postsSheet) return { success: false, message: "Posts sheet not found." };

    const row = postsSheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());

    const stoneIdx = headers.indexOf("Stone Type");
    if (stoneIdx === -1) return { success: false, message: "Stone Type column not found." };

    const material = String(postsSheet.getRange(row, stoneIdx + 1).getValue() || "").trim();
    if (!material) return { success: false, message: "Stone Type is empty for this row." };

    const entitiesSheet = ss.getSheetByName("technical_entities");
    if (!entitiesSheet) return { success: false, message: "technical_entities sheet not found." };

    // Get existing entity names for this material
    const entitiesData = entitiesSheet.getDataRange().getValues();
    const existingNames = [];
    const materialLower = material.toLowerCase().trim();

    for (let i = 1; i < entitiesData.length; i++) {
      const rowMaterial = String(entitiesData[i][0] || "").trim().toLowerCase();
      if (rowMaterial === materialLower) {
        const name = String(entitiesData[i][1] || "").trim().toLowerCase();
        if (name) existingNames.push(name);
      }
    }

    // Parse and validate new rows
    const newRows = String(newRowsText || "").trim();
    if (!newRows) return { success: false, message: "No entity rows to add." };

    const rowsToAdd = [];
    const skipped = [];

    newRows.split("\n").forEach(function(line) {
      const trimmed = line.trim();
      if (!trimmed) return;
      const parts = trimmed.split("|");
      if (parts.length < 2) return;
      const name = parts[0].trim().toLowerCase();
      if (existingNames.indexOf(name) > -1) {
        skipped.push(parts[0].trim());
      } else {
        rowsToAdd.push([
          material,
          String(parts[0] || "").trim(),
          String(parts[1] || "").trim(),
          String(parts[2] || "").trim(),
          String(parts[3] || "").trim()
        ]);
      }
    });

    if (rowsToAdd.length === 0) {
      return { success: false, message: "All entities already exist." + (skipped.length ? " Skipped: " + skipped.join(", ") : "") };
    }

    // Append new rows to technical_entities sheet
    const lastRow = entitiesSheet.getLastRow();
    entitiesSheet.getRange(lastRow + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);

    let message = rowsToAdd.length + " entity rows added to technical_entities for " + material + ".";
    if (skipped.length) message += " Skipped (already exist): " + skipped.join(", ");

    return { success: true, message: message };

  } catch(e) {
    return { success: false, message: "bc_pushEntitiesToTSM error: " + e.message };
  }
}

function bc_pushNodeMapOnly(raw) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(
      s => s.getSheetId() == BC_SHEET_CONFIG.posts
    );
    if (!sheet) return { success: false, message: "Posts sheet not found." };

    const row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());

    const cleaned = String(raw || "").trim();
    if (!cleaned) return { success: false, message: "No output to push." };

    const nodeMapMatch = cleaned.match(/---COMPETITOR NODE MAP---([\s\S]*?)---END COMPETITOR NODE MAP---/i);
    if (!nodeMapMatch) return { success: false, message: "No Competitor Node Map block found in output." };

    const nodeMapIdx = headers.indexOf("Competitor Node Map");
    if (nodeMapIdx === -1) return { success: false, message: "Column 'Competitor Node Map' not found." };

    const nodeMapCell = sheet.getRange(row, nodeMapIdx + 1);
    nodeMapCell.setNumberFormat("@");
    nodeMapCell.setValue(nodeMapMatch[1].trim());

    return { success: true, message: "Competitor Node Map updated. No other columns affected." };

  } catch(e) {
    return { success: false, message: "bc_pushNodeMapOnly error: " + e.message };
  }
}

function bc_getSerpBridgePrereqs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(
      s => s.getSheetId() == BC_SHEET_CONFIG.posts
    );
    if (!sheet) return {};

    const row = sheet.getActiveRange().getRow();
    if (row < 2) return {};

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());

    const getVal = name => {
      const idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
    };

    return {
      primaryTerm:        getVal("Primary Search Term"),
      strategicReasoning: getVal("Strategic Reasoning"),
      topicalIntent:      getVal("Topical Intent"),
      matrixRole:         getVal("Matrix Role"),
      keyDecisions:       getVal("Key Decisions Explained"),
      recoveryBlueprint:  getVal("Recovery Blueprint")
    };

  } catch(e) {
    return {};
  }
}