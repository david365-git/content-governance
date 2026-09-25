/**
 * ============================================================
 * bc_RowWriter.gs
 * Abbey Floor Care — Governance Prompt Assembler
 * LLM Output Writer — writes parsed LLM output back to the
 * active row in the posts sheet by matching field labels
 * to column headers.
 *
 * Version: 1.0 — Split from bc_GovernancePromptAssembler_DataFetcher.gs v3.4
 *
 * Contains:
 * - bc_writeLLMOutputToRow()
 *
 * Depends on: bc_GovernancePromptAssembler_Config.gs
 *             bc_DataFetcher.gs
 * ============================================================
 */


function bc_writeLLMOutputToRow(pairs, rawResponse, responseColumnName, lockedRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const postsSheet = ss.getSheets().find(
    s => s.getSheetId() == BC_SHEET_CONFIG.posts
  );
  if (!postsSheet) throw new Error("Posts sheet not found.");

  const activeRowIndex = lockedRow || ss.getActiveRange().getRow();
  if (activeRowIndex < 2)
    throw new Error("No valid row to write to.");

  const headers = postsSheet
    .getRange(1, 1, 1, postsSheet.getLastColumn())
    .getValues()[0]
    .map(function(h) { return String(h).trim(); });

  const normalise = function(str) {
    return String(str).toLowerCase().trim();
  };

  const headerMap = {};
  headers.forEach(function(h, i) {
    if (h) headerMap[normalise(h)] = i + 1;
  });

  let written = 0;
  let skipped = 0;
  const skippedLabels = [];

  const pairKeys = Object.keys(pairs);
  for (let i = 0; i < pairKeys.length; i++) {
    const label    = pairKeys[i];
    const value    = pairs[label];
    const colIndex = headerMap[normalise(label)];
    if (colIndex) {
      postsSheet.getRange(activeRowIndex, colIndex).setValue(value);
      written++;
    } else {
      skipped++;
      skippedLabels.push(label);
    }
  }

  // Write raw LLM response to storage column if provided
  if (rawResponse && responseColumnName) {
    const responseColIdx = headerMap[normalise(responseColumnName)];
    if (responseColIdx) {
      let cleanResponse = String(rawResponse).trim();
      if (cleanResponse.charAt(0) === '=') {
        cleanResponse = "'" + cleanResponse;
      }
      postsSheet.getRange(activeRowIndex, responseColIdx).setValue(cleanResponse);
    }
  }

  let summary = written + " fields written to row " + activeRowIndex + ".";
  if (skipped > 0) {
    summary += " " + skipped + " not matched: " + skippedLabels.join(", ");
  }
  return summary;
}

function bc_clearRowOutput() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  
  sheet.getRangeList([
    'H' + row + ':AG' + row,
    'AW' + row,
    'DJ' + row    // ← ADD THIS
  ]).clearContent();
  
  return "✅ Cleared columns H–AG, AW, and DJ in row " + row;
}
function bc_moveToNextRow() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    if (!sheet) return { success: false, message: "posts sheet not found." };

    const currentRow = sheet.getActiveRange().getRow();
    const articleTypeIdx = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim())
      .indexOf("Article Type");

    if (articleTypeIdx === -1) return { success: false, message: "Article Type column not found." };

    // Get all visible rows in the current filter view
    const lastRow = sheet.getLastRow();
    const visibleRows = [];

    for (let i = 2; i <= lastRow; i++) {
      if (!sheet.isRowHiddenByFilter(i)) {
        visibleRows.push(i);
      }
    }

    if (visibleRows.length === 0) return { success: false, message: "No visible rows found." };

    // Find the next visible row after current
    let nextRow = visibleRows[0]; // default to first visible row
    for (let i = 0; i < visibleRows.length; i++) {
      if (visibleRows[i] > currentRow) {
        nextRow = visibleRows[i];
        break;
      }
    }

    const articleTypeCol = articleTypeIdx + 1;
    sheet.setActiveRange(sheet.getRange(nextRow, articleTypeCol));
    return { success: true, message: "Moved to row " + nextRow };

  } catch(e) {
    return { success: false, message: "bc_moveToNextRow error: " + e.message };
  }
}
function bc_migrateP1AResponses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(
      s => s.getSheetId() == BC_SHEET_CONFIG.posts
    );
    if (!sheet) throw new Error("Posts sheet not found.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(h) { return String(h).trim(); });

    const normalise = function(str) {
      return String(str).toLowerCase().trim();
    };

    const getColIdx = function(name) {
      for (let i = 0; i < headers.length; i++) {
        if (normalise(headers[i]) === normalise(name)) return i + 1;
      }
      return -1;
    };

    const classificationContextCol = getColIdx("1A Classification Context");
    const p1aResponseCol           = getColIdx("P1A Response");
    const secondaryIntentCol       = getColIdx("Secondary Intent Decisions");

    if (classificationContextCol === -1) throw new Error("Column '1A Classification Context' not found.");
    if (p1aResponseCol === -1)           throw new Error("Column 'P1A Response' not found.");
    if (secondaryIntentCol === -1)       throw new Error("Column 'Secondary Intent Decisions' not found.");

    let processed = 0;
    let skipped = 0;

    for (let row = 2; row <= 271; row++) {
      const context = String(sheet.getRange(row, classificationContextCol).getValue() || "").trim();

      if (!context || context.indexOf("BLOCK 1 — GSC SIGNAL ASSESSMENT") !== 0) {
        skipped++;
        continue;
      }

      // Copy to P1A Response
      sheet.getRange(row, p1aResponseCol).setValue(context);

      // Parse secondary intent decisions from BLOCK 4 — RANKED INTENTS
      const rankedMatch = context.match(/BLOCK 4[^\n]*\n([\s\S]*?)(?:BLOCK 5|$)/);
      const rankedLines = rankedMatch
        ? rankedMatch[1].trim().split('\n').filter(function(l) { return l.trim(); })
        : [];

      // Parse confirmed primary intent
      const primaryIntentMatch = context.match(/CONFIRMED PRIMARY INTENT:\s*(.+)/);
      const confirmedPrimaryIntent = primaryIntentMatch ? primaryIntentMatch[1].trim().toLowerCase() : "";

      // Build secondary decisions — all ranked intents except primary = Y
      const secondaryDecisions = [];
      for (let i = 0; i < rankedLines.length; i++) {
        const intentName = rankedLines[i].replace(/^\d+\.\s*/, '').trim();
        if (!intentName) continue;
        if (intentName.toLowerCase() === confirmedPrimaryIntent) continue;
        secondaryDecisions.push({
          num: String(i + 1),
          name: intentName,
          answer: 'Y',
          reason: 'Confirmed in auto-parse.'
        });
      }

      if (secondaryDecisions.length > 0) {
        const lines = ["RETAINED SUPPORTING INTENTS:"];
        for (let i = 0; i < secondaryDecisions.length; i++) {
          lines.push(secondaryDecisions[i].num + ". " + secondaryDecisions[i].name + " — " + secondaryDecisions[i].reason);
        }
        const secondaryBlock = lines.join("\n");
        sheet.getRange(row, secondaryIntentCol).setValue(secondaryBlock);
      }

      processed++;
    }

    return "Done. Processed: " + processed + " rows. Skipped: " + skipped + " rows.";

  } catch(e) {
    throw new Error("bc_migrateP1AResponses: " + e.message);
  }
}

function bc_parsePairsServer(raw) {
  var pairs = {};
  var marker = '=== VERTICAL OUTPUT — PASTE INTO PANEL 2 ===';
  var start = raw.indexOf(marker);
  var text = start > -1 ? raw.substring(start + marker.length) : raw;
  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (line.indexOf('CLUSTERS:') === 0) {
      pairs['Observed Query Cluster'] = line.substring('CLUSTERS:'.length).trim();
      continue;
    }
    var colonIdx = line.indexOf(':');
    if (colonIdx === -1) continue;
    var label = line.substring(0, colonIdx).trim();
    var value = line.substring(colonIdx + 1).trim();
    if (!label) continue;
    if (label.charAt(0) === '=' || label.charAt(0) === '-') continue;
    if (label.indexOf('If ') === 0) continue;
    if (label.indexOf('Before ') === 0) continue;
    if (/^\d+\./.test(label)) continue;
    if (label === 'Rules') continue;
    if (label === 'After the vertical block output exactly') continue;
    if (label === 'Waiting for Input') continue;
    if (label.indexOf('CONFIRMED') === 0) continue;
    if (label.indexOf('AMBIGUITY') === 0) continue;
    if (label.indexOf('QUESTION') === 0) continue;
    if (label.indexOf('RANKED') === 0) continue;
    if (label.indexOf('SECONDARY') === 0) continue;
    if (label.indexOf('HARD RULE') === 0) continue;
    value = value.replace(/\[(https?:\/\/[^\]]+)\]\([^)]+\)/g, '$1');
    pairs[label] = value;
  }
  return pairs;
}