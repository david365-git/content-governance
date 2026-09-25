/**
 * ================================================================================
 * bc_GovernanceValidation.gs — POST-P3 COLUMN VALIDATION
 * ================================================================================
 *
 * Checks all governed columns written by P1A, P1B, P2, P3 for:
 *   - Empty values
 *   - Literal [value] placeholder text
 *
 * Called from the Governance Prompt Assembler sidebar after P3 Write to Row.
 *
 * Part of Abbey Floor Care Content Pipeline
 * ================================================================================
 */

function bc_validateGovernanceColumns(lockedRow) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = lockedRow || sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "No valid row to validate." };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(function(h) { return String(h).trim(); });

    var governed = [
      { col: "Article Type",                stage: "P1A" },
      { col: "Confirmed Primary Intent",    stage: "P1A" },
      { col: "Primary Entity",              stage: "P1B" },
      { col: "Material Entity",             stage: "P1B" },
      { col: "Entity Role",                 stage: "P1B" },
      { col: "Entity Type",                 stage: "P1B" },
      { col: "Supporting Entities Core",    stage: "P1B" },
      { col: "Peripheral Entities Link Out",stage: "P1B" },
      { col: "Feeds Hub",                   stage: "P1B" },
      { col: "Page Governance Summary",     stage: "P1B" },
      { col: "Entity Governance Status",    stage: "P1B" },
      { col: "Publish Justification",       stage: "P1B" },
      { col: "Observed Query Cluster",      stage: "P2"  },
      { col: "Primary Query Cluster Owned", stage: "P2"  },
      { col: "Drift Status",                stage: "P2"  },
      { col: "Recovery Blueprint",          stage: "P2"  },
      { col: "GSC Query List",              stage: "P2"  },
      { col: "Rewrite Status",              stage: "P3"  },
      { col: "Page Rewrite Brief",          stage: "P3"  },
      { col: "Cannibalisation Guardrail",   stage: "P3"  },
      { col: "Safe Handoff Pages",          stage: "P3"  },
      { col: "Rewrite Role Lock",           stage: "P3"  }
    ];

    var failures = [];
    var passes   = [];

    governed.forEach(function(field) {
      var idx = headers.indexOf(field.col);
      if (idx === -1) {
        failures.push({ col: field.col, stage: field.stage, reason: "Column not found in sheet", value: "" });
        return;
      }
      var val = String(sheet.getRange(row, idx + 1).getValue() || "").trim();
      if (!val) {
        failures.push({ col: field.col, stage: field.stage, reason: "Empty", value: "" });
      } else if (val === "[value]" || val.indexOf("[value]") > -1) {
        failures.push({ col: field.col, stage: field.stage, reason: "[value] placeholder not replaced", value: val.substring(0, 80) });
      } else {
        passes.push({ col: field.col, stage: field.stage, value: val.substring(0, 60) });
      }
    });

    return {
      success:      true,
      row:          row,
      totalChecked: governed.length,
      passCount:    passes.length,
      failCount:    failures.length,
      failures:     failures,
      passes:       passes,
      clean:        failures.length === 0
    };

  } catch(e) {
    return { success: false, message: "Validation error: " + e.toString() };
  }
}