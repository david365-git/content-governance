/**
 * ============================================================
 * bc_PromptBlocks.gs-v1.4
 * Abbey Floor Care — Governance Prompt Assembler
 * Reusable Prompt Block Builders
 *
 * Version: 1.4 — Prompt 1A governance input block builders added
 *
 * Contains:
 * - bc_buildArticleTypeClassificationBlock()
 * - bc_buildArticleTypeDataBlock()
 * - bc_getTSMEntitiesBlock()
 * - bc_getServiceAuthorityBlock()
 * - bc_buildRoleMappingBlock()
 * - bc_getTierStructuralCoverageBlock()
 * - bc_getEntityRoleControlBlock()
 * - bc_getSurfaceIssueMasterBlock()
 * - bc_getMaterialGovernanceBlock()
 * - bc_getMaterialIdentityBlock()
 * - bc_getTechDNABlock()
 *
 * Depends on: bc_GovernancePromptAssembler_Config.gs
 *             bc_DataFetcher.gs
 * ============================================================
 */


function bc_buildArticleTypeClassificationBlock() {
  try {
    const sheet = bc_getSheetByConfigKeys_(["articleTypeControl"]);
    if (!sheet) return "(Article Type Control sheet not found)";

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return "(Article Type Control sheet is empty)";

    const rows = data.slice(1).filter(function(row) {
      return (
        String(row[0]  || "").trim() !== "" &&
        String(row[10] || "").trim() !== "" &&
        String(row[11] || "").trim() !== "" &&
        !isNaN(parseInt(row[11], 10))
      );
    });

    rows.sort(function(a, b) {
      return parseInt(a[11], 10) - parseInt(b[11], 10);
    });

    const lines = [
      "===== ARTICLE TYPE CLASSIFICATION FRAMEWORK =====",
      "",
      "CLASSIFICATION METHOD:",
      "Do NOT classify from tone alone.",
      "Extract signals from five families:",
      "  1. Structural — heading hierarchy, section order, page modules",
      "  2. Organisational — how content is internally arranged",
      "  3. Conversion — evidence the page generates enquiries",
      "  4. URL / Metadata — contextual anchors that disambiguate structure",
      "  5. Content-scope — breadth and boundary of what the page covers",
      "",
      "Apply hard-locks in precedence order below.",
      "Stop at the first hard-lock that fires.",
      "If no hard-lock fires, use weighted signal scoring.",
      "If two types score within 2 points, apply precedence order as tie-break.",
      "",
      "SIGNAL WEIGHTS:",
      "Primary structural signal present:  +5",
      "Strong supporting signal present:   +3",
      "Supporting signal present:          +2",
      "Weak contextual anchor present:     +1",
      "Negative signal present:            -3",
      "Hard-lock contradiction:            -6",
      "",
      "HARD-LOCK PRECEDENCE ORDER:",
      "(Check in this order. Stop at first match.)",
      ""
    ];

    for (let i = 0; i < rows.length; i++) {
      const row        = rows[i];
      const label      = String(row[0]  || "").trim();
      const signals    = String(row[10] || "").trim();
      const precedence = String(row[11] || "").trim();
      if (!label || !signals) continue;
      lines.push("--- " + precedence + ". " + label.toUpperCase() + " ---");
      lines.push(signals);
      lines.push("");
    }

    lines.push(
      "MANDATORY OVERRIDE RULE:",
      "If a page has ALL THREE of the following:",
      "  1. A specific location identity (place name in URL, title, headings, or body content as a primary targeting entity)",
      "  2. A service-offering architecture",
      "  3. At least one conversion signal (phone, form, quote CTA, booking language, contact invitation)",
      "Then classify as Geo Service Page regardless of educational prose tone.",
      "Educational sections about cleaning, polishing, sealing, or maintenance do not override this combination.",
      "",
      "OUTPUT REQUIREMENT:",
      "Return exactly one article type from the ALLOWED ARTICLE TYPES list.",
      "===== END CLASSIFICATION FRAMEWORK ====="
    );

    return lines.join("\n");

  } catch (e) {
    return "(Article type classification block error: " + e.message + ")";
  }
}


function bc_buildArticleTypeDataBlock(rowData) {
  const lines = [
    "===== ARTICLE TYPE DATA =====",
    "Article Type: " + rowData.articleType,
    ""
  ];

  if (rowData.onPageExpectations) {
    lines.push(
      "On-Page Content Expectations:",
      rowData.onPageExpectations,
      ""
    );
  }

  if (rowData.reasoning) {
    lines.push(
      "Content Strategy:",
      rowData.reasoning,
      ""
    );
  }

  if (rowData.intentDecisionTree) {
    lines.push(
      "Intent Classification Decision Tree:",
      rowData.intentDecisionTree,
      ""
    );
  } else {
    lines.push(
      "Intent Classification Decision Tree:",
      "(No decision tree defined for this article type in column M of the Article Type Control Sheet.)",
      "Apply structural analysis only.",
      ""
    );
  }

  lines.push("===== END ARTICLE TYPE DATA =====");
  return lines.join("\n");
}


/* =========================================================================
 * Prompt 1A governance block helpers
 * ========================================================================= */

function bc_getSheetByConfigKeys_(keys) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const fallbackNamesByKey = {
    prompts: ["prompts", "Prompts"],
    intentTax: ["1 - Intent Taxonomy Table", "Intent Taxonomy Table"],
    legitMap: ["2 - Master Intent Legitimacy Map", "Master Intent Legitimacy Map"],
    articleTypeControl: ["Article Type Control Sheet"],
    entityTypeControl: ["Entity Type Control Sheet"],
    roleMapping: ["Role Mapping Sheet"],
    tierStructuralCoverage: ["Tier Structural Coverage Matrix"],
    tierStructuralCoverageMatrix: ["Tier Structural Coverage Matrix"],
    entityRoleControl: ["Entity Role Control Sheet"],
    entityRoleControlSheet: ["Entity Role Control Sheet"],
    surfaceIssueMaster: ["Surface Issue Master Control Sheet"],
    surfaceIssueMasterControl: ["Surface Issue Master Control Sheet"],
    surfaceIssueMasterControlSheet: ["Surface Issue Master Control Sheet"],
    materialGovernanceReference: ["Material Governance Reference Sheet"],
    materialGovernanceReferenceSheet: ["Material Governance Reference Sheet"],
    materialIdentityMapping: ["Material Identity Mapping sheet", "Material Identity Mapping Sheet"],
    materialIdentityMappingSheet: ["Material Identity Mapping sheet", "Material Identity Mapping Sheet"],
    techDNA: ["Tech_DNA"],
    techDna: ["Tech_DNA"],
    tech_DNA: ["Tech_DNA"],
    techSpec: ["technical", "Technical Matrix"],
    serviceAuthority: ["service-authority sheet", "Service Authority Sheet"],
    entityRole: ["Entity Role Control Sheet"]
  };

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (
      typeof BC_SHEET_CONFIG !== "undefined" &&
      BC_SHEET_CONFIG &&
      Object.prototype.hasOwnProperty.call(BC_SHEET_CONFIG, key)
    ) {
      const sheetId = BC_SHEET_CONFIG[key];
      const byId = ss.getSheets().find(function(s) {
        return s.getSheetId() == sheetId;
      });
      if (byId) return byId;
    }
  }

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const fallbackNames = fallbackNamesByKey[key] || [];
    for (let j = 0; j < fallbackNames.length; j++) {
      const wanted = bc_normalizeText_(fallbackNames[j]);
      const byName = ss.getSheets().find(function(s) {
        return bc_normalizeText_(s.getName()) === wanted;
      });
      if (byName) return byName;
    }
  }

  return null;
}


function bc_normalizeText_(value) {
  return String(value || "").trim().toLowerCase();
}


function bc_findHeaderIndex_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const needle = bc_normalizeText_(candidates[i]);
    for (let j = 0; j < headers.length; j++) {
      if (bc_normalizeText_(headers[j]) === needle) return j;
    }
  }
  return -1;
}


function bc_safeCell_(row, index) {
  if (index < 0 || index >= row.length) return "";
  return String(row[index] || "").trim();
}


function bc_limitLines_(text, maxLines) {
  return String(text || "")
    .split(/\r?\n/)
    .slice(0, maxLines)
    .join("\n");
}


function bc_buildGenericGovernanceBlock_(title, introLines, headers, rows, selectedIndexes) {
  const lines = ["===== " + title + " ====="];

  if (introLines && introLines.length) {
    for (let i = 0; i < introLines.length; i++) lines.push(introLines[i]);
    lines.push("");
  }

  if (!rows || rows.length === 0) {
    lines.push("(No matching governed rows found.)");
    lines.push("===== END " + title + " =====");
    return lines.join("\n");
  }

  const headerLine = selectedIndexes.map(function(idx) {
    return headers[idx];
  }).join(" | ");

  lines.push(headerLine);

  for (let i = 0; i < rows.length; i++) {
    lines.push(selectedIndexes.map(function(idx) {
      return bc_safeCell_(rows[i], idx);
    }).join(" | "));
  }

  lines.push("===== END " + title + " =====");
  return lines.join("\n");
}


function bc_getMaterialGovernanceRow_(stoneType) {
  try {
    const sheet = bc_getSheetByConfigKeys_([
      "materialGovernanceReference",
      "materialGovernance",
      "materialGovernanceReferenceSheet"
    ]);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;

    const headers = data[0];
    const matIdx  = bc_findHeaderIndex_(headers, [
      "Canonical Material Name",
      "Canonical Material",
      "Material",
      "Stone Type"
    ]);

    if (matIdx === -1) return null;

    const target = bc_normalizeText_(stoneType);

    for (let i = 1; i < data.length; i++) {
      if (bc_normalizeText_(data[i][matIdx]) === target) {
        return {
          headers: headers,
          row: data[i]
        };
      }
    }

    return null;
  } catch (e) {
    return null;
  }
}


function bc_getMaterialClassForStoneType_(stoneType) {
  const mg = bc_getMaterialGovernanceRow_(stoneType);
  if (!mg) return "";
  const classIdx = bc_findHeaderIndex_(mg.headers, [
    "Material Class Type",
    "Material Class"
  ]);
  return classIdx > -1 ? bc_safeCell_(mg.row, classIdx) : "";
}


function bc_getTierStructuralCoverageBlock(articleType) {
  try {
    const sheet = bc_getSheetByConfigKeys_([
      "tierStructuralCoverage",
      "tierLevels",
      "tierStructuralCoverageMatrix"
    ]);
    if (!sheet) {
      return [
        "===== GOVERNANCE INPUT — TIER STRUCTURAL COVERAGE MATRIX =====",
        "(Tier Structural Coverage Matrix sheet not found — add the relevant config key to BC_SHEET_CONFIG if needed.)",
        "===== END GOVERNANCE INPUT — TIER STRUCTURAL COVERAGE MATRIX ====="
      ].join("\n");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [
        "===== GOVERNANCE INPUT — TIER STRUCTURAL COVERAGE MATRIX =====",
        "(Tier Structural Coverage Matrix sheet is empty.)",
        "===== END GOVERNANCE INPUT — TIER STRUCTURAL COVERAGE MATRIX ====="
      ].join("\n");
    }

    const headers = data[0];
    const permittedIdx = bc_findHeaderIndex_(headers, [
      "Permitted Article Types"
    ]);
    const labelIdx = bc_findHeaderIndex_(headers, ["Label"]);
    const structuralReqIdx = bc_findHeaderIndex_(headers, ["Structural Requirement"]);
    const orderIdx = bc_findHeaderIndex_(headers, ["Order"]);
    const entityCatIdx = bc_findHeaderIndex_(headers, ["Entity Category"]);
    const primaryFocusIdx = bc_findHeaderIndex_(headers, ["Primary Focus Discriminator"]);
    const mustNotIdx = bc_findHeaderIndex_(headers, ["Must Not Contain"]);
    const notesIdx = bc_findHeaderIndex_(headers, ["Notes"]);

    const rows = data.slice(1).filter(function(row) {
      const permitted = bc_safeCell_(row, permittedIdx);
      return permitted && permitted.toLowerCase().indexOf(String(articleType || "").trim().toLowerCase()) > -1;
    });

    const selected = [
      labelIdx,
      structuralReqIdx,
      orderIdx,
      entityCatIdx,
      primaryFocusIdx,
      mustNotIdx,
      notesIdx
    ].filter(function(idx) { return idx > -1; });

    return bc_buildGenericGovernanceBlock_(
      "GOVERNANCE INPUT — TIER STRUCTURAL COVERAGE MATRIX",
      [
        "Use this as a page-level structural architecture validator only.",
        "Do not use it as the primary H2 intent classifier.",
        "Current candidate Article Type: " + articleType
      ],
      headers,
      rows,
      selected
    );
  } catch (e) {
    return [
      "===== GOVERNANCE INPUT — TIER STRUCTURAL COVERAGE MATRIX =====",
      "(Tier Structural Coverage Matrix fetch error: " + e.message + ")",
      "===== END GOVERNANCE INPUT — TIER STRUCTURAL COVERAGE MATRIX ====="
    ].join("\n");
  }
}


function bc_getEntityRoleControlBlock() {
  try {
    const sheet = bc_getSheetByConfigKeys_([
      "entityRoleControl",
      "entityRole",
      "entityRoleControlSheet"
    ]);
    if (!sheet) {
      return [
        "===== GOVERNANCE INPUT — ENTITY ROLE CONTROL SHEET =====",
        "(Entity Role Control Sheet not found — add the relevant config key to BC_SHEET_CONFIG if needed.)",
        "===== END GOVERNANCE INPUT — ENTITY ROLE CONTROL SHEET ====="
      ].join("\n");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [
        "===== GOVERNANCE INPUT — ENTITY ROLE CONTROL SHEET =====",
        "(Entity Role Control Sheet is empty.)",
        "===== END GOVERNANCE INPUT — ENTITY ROLE CONTROL SHEET ====="
      ].join("\n");
    }

    const headers = data[0];
    const selected = [
      bc_findHeaderIndex_(headers, ["Canonical Role Label"]),
      bc_findHeaderIndex_(headers, ["Role Definition (Governance Function Only)", "Role Definition"]),
      bc_findHeaderIndex_(headers, ["When To Use"]),
      bc_findHeaderIndex_(headers, ["When NOT To Use"]),
      bc_findHeaderIndex_(headers, ["Drift Risk Notes"]),
      bc_findHeaderIndex_(headers, ["Mapped Intent"])
    ].filter(function(idx) { return idx > -1; });

    return bc_buildGenericGovernanceBlock_(
      "GOVERNANCE INPUT — ENTITY ROLE CONTROL SHEET",
      [
        "Use this as a role-drift validator only.",
        "Do not use it as the primary H2 intent classifier."
      ],
      headers,
      data.slice(1),
      selected
    );
  } catch (e) {
    return [
      "===== GOVERNANCE INPUT — ENTITY ROLE CONTROL SHEET =====",
      "(Entity Role Control Sheet fetch error: " + e.message + ")",
      "===== END GOVERNANCE INPUT — ENTITY ROLE CONTROL SHEET ====="
    ].join("\n");
  }
}


function bc_getSurfaceIssueMasterBlock(stoneType) {
  try {
    const sheet = bc_getSheetByConfigKeys_([
      "surfaceIssueMaster",
      "surfaceIssueMasterControl",
      "surfaceIssueMasterControlSheet"
    ]);
    if (!sheet) {
      return [
        "===== GOVERNANCE INPUT — SURFACE ISSUE MASTER CONTROL SHEET =====",
        "(Surface Issue Master Control Sheet not found — add the relevant config key to BC_SHEET_CONFIG if needed.)",
        "===== END GOVERNANCE INPUT — SURFACE ISSUE MASTER CONTROL SHEET ====="
      ].join("\n");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [
        "===== GOVERNANCE INPUT — SURFACE ISSUE MASTER CONTROL SHEET =====",
        "(Surface Issue Master Control Sheet is empty.)",
        "===== END GOVERNANCE INPUT — SURFACE ISSUE MASTER CONTROL SHEET ====="
      ].join("\n");
    }

    const headers = data[0];
    const appliesIdx = bc_findHeaderIndex_(headers, ["Applies To Material Class"]);
    const issueIdx = bc_findHeaderIndex_(headers, ["Canonical Surface Issue Label"]);
    const variantIdx = bc_findHeaderIndex_(headers, ["Plain English Variant (Recognition Only)"]);
    const matrixIdx = bc_findHeaderIndex_(headers, ["Matrix Behaviour Category"]);
    const notesIdx = bc_findHeaderIndex_(headers, ["Notes"]);

    const materialClass = bc_getMaterialClassForStoneType_(stoneType);
    const normalizedClass = bc_normalizeText_(materialClass);
    const normalizedStone = bc_normalizeText_(stoneType);

    const rows = data.slice(1).filter(function(row) {
      const applies = bc_normalizeText_(row[appliesIdx]);
      if (!applies) return false;
      if (applies.indexOf("all materials") > -1) return true;
      if (normalizedClass && applies.indexOf(normalizedClass) > -1) return true;
      if (normalizedStone && applies.indexOf(normalizedStone) > -1) return true;
      if (normalizedStone.indexOf("victorian") > -1 && applies.indexOf("historic clay") > -1) return true;
      if (normalizedStone.indexOf("terracotta") > -1 && applies.indexOf("clay") > -1) return true;
      if (normalizedStone.indexOf("quarry") > -1 && applies.indexOf("clay") > -1) return true;
      return false;
    });

    const selected = [issueIdx, variantIdx, appliesIdx, matrixIdx, notesIdx]
      .filter(function(idx) { return idx > -1; });

    return bc_buildGenericGovernanceBlock_(
      "GOVERNANCE INPUT — SURFACE ISSUE MASTER CONTROL SHEET",
      [
        "Use this as diagnostic / defect support only.",
        "Do not use it as a full task classifier.",
        "RUN MATERIAL: " + stoneType,
        "Derived Material Class: " + (materialClass || "(not found)")
      ],
      headers,
      rows,
      selected
    );
  } catch (e) {
    return [
      "===== GOVERNANCE INPUT — SURFACE ISSUE MASTER CONTROL SHEET =====",
      "(Surface Issue Master Control Sheet fetch error: " + e.message + ")",
      "===== END GOVERNANCE INPUT — SURFACE ISSUE MASTER CONTROL SHEET ====="
    ].join("\n");
  }
}


function bc_getMaterialGovernanceBlock(stoneType) {
  try {
    const mg = bc_getMaterialGovernanceRow_(stoneType);
    if (!mg) {
      return [
        "===== GOVERNANCE INPUT — MATERIAL GOVERNANCE REFERENCE SHEET =====",
        "(No Material Governance row found for material: " + stoneType + ")",
        "===== END GOVERNANCE INPUT — MATERIAL GOVERNANCE REFERENCE SHEET ====="
      ].join("\n");
    }

    const headers = mg.headers;
    const row = mg.row;
    const selected = [
      bc_findHeaderIndex_(headers, ["Canonical Material Name"]),
      bc_findHeaderIndex_(headers, ["Material Class Type"]),
      bc_findHeaderIndex_(headers, ["Acceptable Recognition Variants (Recognition Only – Not for Output)", "Acceptable Recognition Variants"]),
      bc_findHeaderIndex_(headers, ["Primary Entity Stem Format"]),
      bc_findHeaderIndex_(headers, ["Approved Surface Issue Classes"]),
      bc_findHeaderIndex_(headers, ["Matrix Version Reference"]),
      bc_findHeaderIndex_(headers, ["Governance Boundary Notes"])
    ].filter(function(idx) { return idx > -1; });

    return bc_buildGenericGovernanceBlock_(
      "GOVERNANCE INPUT — MATERIAL GOVERNANCE REFERENCE SHEET",
      [
        "Use this as the material-specific constraint layer.",
        "Do not use it to infer article type or intent.",
        "RUN MATERIAL: " + stoneType
      ],
      headers,
      [row],
      selected
    );
  } catch (e) {
    return [
      "===== GOVERNANCE INPUT — MATERIAL GOVERNANCE REFERENCE SHEET =====",
      "(Material Governance Reference Sheet fetch error: " + e.message + ")",
      "===== END GOVERNANCE INPUT — MATERIAL GOVERNANCE REFERENCE SHEET ====="
    ].join("\n");
  }
}


function bc_getMaterialIdentityBlock(stoneType, searchEntity) {
  try {
    const sheet = bc_getSheetByConfigKeys_([
      "materialIdentityMapping",
      "materialIdentity",
      "materialIdentityMappingSheet"
    ]);
    if (!sheet) {
      return [
        "===== GOVERNANCE INPUT — MATERIAL IDENTITY MAPPING SHEET =====",
        "(Material Identity Mapping Sheet not found — add the relevant config key to BC_SHEET_CONFIG if needed.)",
        "===== END GOVERNANCE INPUT — MATERIAL IDENTITY MAPPING SHEET ====="
      ].join("\n");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [
        "===== GOVERNANCE INPUT — MATERIAL IDENTITY MAPPING SHEET =====",
        "(Material Identity Mapping Sheet is empty.)",
        "===== END GOVERNANCE INPUT — MATERIAL IDENTITY MAPPING SHEET ====="
      ].join("\n");
    }

    const headers = data[0];
    const hubIdx = bc_findHeaderIndex_(headers, ["Hub (Stone Type)", "Hub"]);
    const variantIdx = bc_findHeaderIndex_(headers, ["Variant (Search Entity)", "Variant"]);
    const matEntityIdx = bc_findHeaderIndex_(headers, ["Material Entity"]);
    const sourceIdx = bc_findHeaderIndex_(headers, ["Evidence Source"]);

    const normalizedStone = bc_normalizeText_(stoneType);
    const normalizedSearch = bc_normalizeText_(searchEntity);

    let rows = data.slice(1).filter(function(row) {
      return bc_normalizeText_(row[hubIdx]) === normalizedStone;
    });

    if (normalizedSearch) {
      const narrowed = rows.filter(function(row) {
        return (
          bc_normalizeText_(row[variantIdx]) === normalizedSearch ||
          bc_normalizeText_(row[matEntityIdx]) === normalizedSearch
        );
      });
      if (narrowed.length > 0) rows = narrowed;
    }

    rows = rows.slice(0, 20);

    const selected = [hubIdx, variantIdx, matEntityIdx, sourceIdx]
      .filter(function(idx) { return idx > -1; });

    return bc_buildGenericGovernanceBlock_(
      "GOVERNANCE INPUT — MATERIAL IDENTITY MAPPING SHEET",
      [
        "Use this for canonical material / variant mapping only.",
        "Do not use it to infer article type or intent.",
        "RUN MATERIAL: " + stoneType,
        "Search Entity / Context Term: " + (searchEntity || "(none)")
      ],
      headers,
      rows,
      selected
    );
  } catch (e) {
    return [
      "===== GOVERNANCE INPUT — MATERIAL IDENTITY MAPPING SHEET =====",
      "(Material Identity Mapping Sheet fetch error: " + e.message + ")",
      "===== END GOVERNANCE INPUT — MATERIAL IDENTITY MAPPING SHEET ====="
    ].join("\n");
  }
}


function bc_getTechDNABlock(stoneType) {
  try {
    const sheet = bc_getSheetByConfigKeys_([
      "techDNA",
      "techDna",
      "tech_DNA"
    ]);
    if (!sheet) {
      return [
        "===== GOVERNANCE INPUT — TECH_DNA =====",
        "(Tech_DNA sheet not found — add the relevant config key to BC_SHEET_CONFIG if needed.)",
        "===== END GOVERNANCE INPUT — TECH_DNA ====="
      ].join("\n");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [
        "===== GOVERNANCE INPUT — TECH_DNA =====",
        "(Tech_DNA sheet is empty.)",
        "===== END GOVERNANCE INPUT — TECH_DNA ====="
      ].join("\n");
    }

    const headers = data[0];
    const stoneIdx = bc_findHeaderIndex_(headers, ["Stone Type", "Material"]);
    const target = bc_normalizeText_(stoneType);

    let row = null;
    for (let i = 1; i < data.length; i++) {
      const candidate = bc_normalizeText_(data[i][stoneIdx]);
      if (
        candidate === target ||
        candidate.indexOf(target) > -1 ||
        target.indexOf(candidate) > -1
      ) {
        row = data[i];
        break;
      }
    }

    if (!row) {
      return [
        "===== GOVERNANCE INPUT — TECH_DNA =====",
        "(No Tech_DNA row found for material: " + stoneType + ")",
        "===== END GOVERNANCE INPUT — TECH_DNA ====="
      ].join("\n");
    }

    const selected = [
      bc_findHeaderIndex_(headers, ["Stone Type"]),
      bc_findHeaderIndex_(headers, ["Fragility Calibration"]),
      bc_findHeaderIndex_(headers, ["Transformation Confidence"]),
      bc_findHeaderIndex_(headers, ["Anchor Entities"]),
      bc_findHeaderIndex_(headers, ["Prohibited LLM Behaviours"]),
      bc_findHeaderIndex_(headers, ["UK Localisation Notes"]),
      bc_findHeaderIndex_(headers, ["TSM Version"])
    ].filter(function(idx) { return idx > -1; });

    return bc_buildGenericGovernanceBlock_(
      "GOVERNANCE INPUT — TECH_DNA",
      [
        "Use this as a material behaviour framing and technical rejection layer only.",
        "Do not use it to infer article type or primary intent.",
        "RUN MATERIAL: " + stoneType
      ],
      headers,
      [row],
      selected
    );
  } catch (e) {
    return [
      "===== GOVERNANCE INPUT — TECH_DNA =====",
      "(Tech_DNA fetch error: " + e.message + ")",
      "===== END GOVERNANCE INPUT — TECH_DNA ====="
    ].join("\n");
  }
}


function bc_filterTSMEntitiesByIntent(sec9Text, lockBlock) {
  if (!lockBlock) return sec9Text;

  const retainedIntents = [];
  const lines = String(lockBlock || "").split(/\r?\n/);
  let inRetained = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (line === "RETAINED SUPPORTING INTENTS:") {
      inRetained = true;
      continue;
    }

    if (!inRetained) continue;

    if (!line) break;
    if (line.startsWith("=====") || line.startsWith("HARD RULE:")) break;

    const m = line.match(/^\d+\.\s+(.+?)\s+—\s+/);
    if (m) retainedIntents.push(m[1].trim());
  }

  const excludedByAbsence = {
    "Sealing / Resealing": [
      "Impregnating Sealer",
      "Colour-Enhancing Impregnator",
      "Topical Satin Coating"
    ],
    "Grout Cleaning / Sealing": [],
    "Sealer Removal / Stripping": [
      "Impregnating Sealer",
      "Colour-Enhancing Impregnator",
      "Topical Satin Coating"
    ]
  };

  const excludedEntities = {};

  Object.keys(excludedByAbsence).forEach(function(intentName) {
    if (retainedIntents.indexOf(intentName) === -1) {
      const entities = excludedByAbsence[intentName];
      for (let i = 0; i < entities.length; i++) {
        excludedEntities[entities[i]] = true;
      }
    }
  });

  if (Object.keys(excludedEntities).length === 0) return sec9Text;

  return sec9Text
    .split(/\r?\n/)
    .filter(function(line) {
      const trimmed = line.trim();
      const entityNames = Object.keys(excludedEntities);
      for (let i = 0; i < entityNames.length; i++) {
        if (trimmed.indexOf(entityNames[i]) === 0) return false;
      }
      return true;
    })
    .join("\n");
}


function bc_getTSMEntitiesBlock(stoneType, lockBlock) {
  try {
    const sheet = bc_getSheetByConfigKeys_(["techSpec"]);
    if (!sheet) {
      return "(Technical specification sheet not found — check sheet ID " +
             BC_SHEET_CONFIG.techSpec + ")";
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return "(Technical specification sheet is empty)";

    const clean  = function(str) {
      return String(str || "").trim().toLowerCase();
    };
    const target = clean(stoneType);

    let matchRow = null;
    for (let i = 0; i < data.length; i++) {
      if (clean(data[i][1]) === target) {
        matchRow = data[i];
        break;
      }
    }

    if (!matchRow) {
      return [
        "===== GOVERNANCE INPUT — TECHNICAL SPECIFICATION MATRIX =====",
        "Material: " + stoneType,
        "(No entry found for this material.)",
        "===== END GOVERNANCE INPUT — TECHNICAL SPECIFICATION MATRIX ====="
      ].join("\n");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const entitiesSheet = ss.getSheetByName("technical_entities");

    if (!entitiesSheet) {
      return [
        "===== GOVERNANCE INPUT — TECHNICAL SPECIFICATION MATRIX =====",
        "Material: " + stoneType,
        "(technical_entities sheet not found.)",
        "===== END GOVERNANCE INPUT — TECHNICAL SPECIFICATION MATRIX ====="
      ].join("\n");
    }

    const entData = entitiesSheet.getDataRange().getValues();
    const entLines = [];
    entLines.push("RECOGNISED TECHNICAL ENTITIES");
    entLines.push("Entity | Type | Primary Application | Semantic Fingerprint Vocabulary");

    for (let i = 1; i < entData.length; i++) {
      if (clean(entData[i][0]) === target) {
        entLines.push([
          String(entData[i][1] || "").trim(),
          String(entData[i][2] || "").trim(),
          String(entData[i][3] || "").trim(),
          String(entData[i][4] || "").trim()
        ].join(" | "));
      }
    }

    if (entLines.length <= 2) {
      return [
        "===== GOVERNANCE INPUT — TECHNICAL SPECIFICATION MATRIX =====",
        "Material: " + stoneType,
        "(No entities found in technical_entities sheet for this material.)",
        "===== END GOVERNANCE INPUT — TECHNICAL SPECIFICATION MATRIX ====="
      ].join("\n");
    }

    let sec9Text = entLines.join("\n");
    sec9Text = bc_filterTSMEntitiesByIntent(sec9Text, lockBlock);

    return [
      "===== GOVERNANCE INPUT — TECHNICAL SPECIFICATION MATRIX =====",
      "Material: " + stoneType,
      "",
      "TECHNICAL ENTITY REFERENCE — VALIDATION ONLY (PROMPT 1B):\n" +
      "This table defines recognised technical entities for this material.\n" +
      "In Prompt 1B:\n" +
      "— Use this ONLY as a technical validation and rejection layer\n" +
      "— Do NOT use it to determine Article Type\n" +
      "— Do NOT use it to determine PRIMARY INTENT\n" +
      "— Do NOT use it to infer Entity Role or Entity Type\n" +
      "This table may only be used to:\n" +
      "— reject technically invalid interpretations\n" +
      "— confirm material-specific constraints\n" +
      "Do NOT invent entity names.\n" +
      "Do NOT use this table as a primary classifier.\n" +
      "",
      sec9Text,
      "",
      "===== END GOVERNANCE INPUT — TECHNICAL SPECIFICATION MATRIX ====="
    ].join("\n");

  } catch (e) {
    return "(Technical specification matrix fetch error: " + e.message + ")";
  }
}

function bc_getTechnicalMatrixBlock(stoneType) {
  try {
    const sheet = bc_getSheetByConfigKeys_(["techSpec"]);

    if (!sheet) {
      return [
        "===== GOVERNANCE INPUT — TECHNICAL MATRIX =====",
        "(Technical sheet not found.)",
        "===== END GOVERNANCE INPUT — TECHNICAL MATRIX ====="
      ].join("\n");
    }

    const data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return [
        "===== GOVERNANCE INPUT — TECHNICAL MATRIX =====",
        "(Technical sheet is empty.)",
        "===== END GOVERNANCE INPUT — TECHNICAL MATRIX ====="
      ].join("\n");
    }

    const clean = function(value) {
      return String(value || "").trim().toLowerCase();
    };

    const target = clean(stoneType);

    for (let i = 1; i < data.length; i++) {
      if (clean(data[i][1]) === target) {
        const matrixText = String(data[i][2] || "").trim();

        return [
          "===== GOVERNANCE INPUT — TECHNICAL MATRIX =====",
          "Material: " + stoneType,
          "",
          matrixText,
          "",
          "===== END GOVERNANCE INPUT — TECHNICAL MATRIX ====="
        ].join("\n");
      }
    }

    return [
      "===== GOVERNANCE INPUT — TECHNICAL MATRIX =====",
      "(No technical matrix found for material: " + stoneType + ")",
      "===== END GOVERNANCE INPUT — TECHNICAL MATRIX ====="
    ].join("\n");

  } catch (e) {
    return [
      "===== GOVERNANCE INPUT — TECHNICAL MATRIX =====",
      "(Technical matrix fetch error: " + e.message + ")",
      "===== END GOVERNANCE INPUT — TECHNICAL MATRIX ====="
    ].join("\n");
  }
}

function bc_getServiceAuthorityBlock(stoneType, articleType, location) {
  try {
    const sheet = bc_getSheetByConfigKeys_(["serviceAuthority"]);

    if (!sheet) {
      return [
        "===== STAGE 4 — SERVICE AUTHORITY ENTITY TABLE =====",
        "(Service authority sheet not found — check sheet ID " +
          BC_SHEET_CONFIG.serviceAuthority + ")",
        "===== END STAGE 4 ====="
      ].join("\n");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [
        "===== STAGE 4 — SERVICE AUTHORITY ENTITY TABLE =====",
        "(Service authority sheet is empty)",
        "===== END STAGE 4 ====="
      ].join("\n");
    }

    const clean  = function(str) {
      return String(str || "").trim().toLowerCase();
    };
    const target = clean(stoneType);

    let matchRow = null;
    for (let i = 1; i < data.length; i++) {
      if (clean(data[i][1]) === target) {
        matchRow = data[i];
        break;
      }
    }

    if (!matchRow) {
      return [
        "===== STAGE 4 — SERVICE AUTHORITY ENTITY TABLE: " +
          stoneType.toUpperCase() + " =====",
        "(No service authority entity table found for material: " +
          stoneType + ")",
        "===== END STAGE 4 ====="
      ].join("\n");
    }

    const entityTable = String(matchRow[3] || "").trim();

    if (!entityTable) {
      return [
        "===== STAGE 4 — SERVICE AUTHORITY ENTITY TABLE: " +
          stoneType.toUpperCase() + " =====",
        "(Service authority entity table is empty for material: " +
          stoneType + ")",
        "===== END STAGE 4 ====="
      ].join("\n");
    }

    const compositionRule =
      "Include service authority entities where the Article Type Weighting " +
      "column shows PRIMARY or INTEGRATED for this article type. Include " +
      "VALIDATING entities only where they directly validate a specific claim " +
      "present in the article content. Exclude all entities marked ABSENT for " +
      "this article type. CRITICAL: If an entity has no weighting entry at all " +
      "for this article type — treat it as ABSENT. Do NOT infer inclusion from " +
      "weightings assigned to other article types.";

    const geoInstruction =
      (String(articleType || "").trim() === "Geo Service Page")
        ? "\nGEO-LOCK INSTRUCTION:\n" +
          "Target location for this page: " + location + "\n" +
          "Match this location against the delivery entities in the table below.\n" +
          "Include only the single delivery entity that matches this location.\n" +
          "Do not include any other delivery entity.\n" +
          "If no delivery entity matches this location, do not include any delivery entity.\n"
        : "";

    return [
      "===== STAGE 4 — SERVICE AUTHORITY ENTITY TABLE: " +
        stoneType.toUpperCase() + " =====",
      "",
      "PURPOSE:",
      "This table is the governed source of service authority entity names for " +
        stoneType + " content.",
      "Column 6 (Supporting Entities Core) must contain a mix of material authority entities from Stage 3 and service authority entities from this table.",
      "The Article Type Weighting column in the table governs which service authority entities are candidates for this specific page.",
      "",
      "COLUMN 6 SELECTION RULE:",
      "Article Type: " + articleType,
      compositionRule,
      "",
      "WEIGHTING KEY:",
      "PRIMARY    — Include in Column 6. This entity anchors service authority.",
      "INTEGRATED — Include in Column 6 where relevant to the specific article intent.",
      "VALIDATING — Include only where it validates a specific claim in the article.",
      "ABSENT     — Do not include for this article type.",
      "",
      "ENTITY NAME RULE:",
      "Use entity names exactly as written in the Entity column below.",
      "Do not paraphrase, abbreviate, or combine entity names.",
      "Do not invent service authority entities not present in this table.",
      geoInstruction,
      "SERVICE AUTHORITY ENTITY TABLE:",
      "",
      entityTable,
      "",
      "===== END STAGE 4 ====="
    ].join("\n");

  } catch (e) {
    return [
      "===== STAGE 4 — SERVICE AUTHORITY ENTITY TABLE =====",
      "(Service authority fetch error: " + e.message + ")",
      "===== END STAGE 4 ====="
    ].join("\n");
  }
}
function bc_buildRecoveryBlueprint(cols1to13, rowData) {

  const coreRaw    = String(cols1to13["Supporting Entities Core"] || "").trim();
  const serviceRaw = String(cols1to13["Supporting Entities - Service Authority Layer"] || "").trim();
  const briefRaw   = String(cols1to13["Page Rewrite Brief"] || "").trim();
  const articleType   = String(cols1to13["Article Type"] || "").trim();
  const primaryEntity = String(cols1to13["Primary Entity"] || "").trim();
  const montrQueries  = String(rowData.queries || "").trim();
  const hasGSC = Number(String(rowData.impressions || "0").replace(/,/g, "")) > 0;

  if (!coreRaw) {
    return "RECOVERY BLUEPRINT MISSING — Supporting Entities Core is empty.";
  }

  const coreEntities = coreRaw.split(",").map(function(e) {
    return e.trim();
  }).filter(Boolean);

  // Set thresholds based on GSC availability
  const priorityThreshold   = hasGSC ? 3 : 2;
  const supportingThreshold = hasGSC ? 2 : 1;

  const scores = {};

  coreEntities.forEach(function(entity) {
    let score = 1;

    if (briefRaw.toLowerCase().includes(entity.toLowerCase())) {
      score++;
    }

    const entityLower   = entity.toLowerCase();
    const primaryLower  = primaryEntity.toLowerCase();
    if (
      primaryLower.includes(entityLower.split(" ")[0]) ||
      (articleType === "Hub Page" && entityLower.includes("porosity")) ||
      (articleType === "Case Study" && entityLower.includes("mastery"))
    ) {
      score++;
    }

    if (hasGSC && montrQueries) {
      const keyword = entityLower.split(" ")[0];
      if (montrQueries.toLowerCase().includes(keyword)) {
        score++;
      }
    }

    scores[entity] = score;
  });

  const priority   = coreEntities.filter(function(e) { return scores[e] >= priorityThreshold; });
  const supporting = coreEntities.filter(function(e) { return scores[e] === supportingThreshold; });
  const peripheral = serviceRaw.split(",").map(function(e) {
    return e.trim();
  }).filter(Boolean);

  const lines = [];
  lines.push("PRIORITY ENTITIES (must appear — anchor co-occurrence targets):");
  lines.push(priority.length ? priority.join(", ") : "(none scored " + priorityThreshold + "+)");
  lines.push("");
  lines.push("SUPPORTING ENTITIES (use where contextually relevant):");
  lines.push(supporting.length ? supporting.join(", ") : "(none scored " + supportingThreshold + ")");
  lines.push("");
  lines.push("PERIPHERAL ENTITIES (one appearance sufficient):");
  lines.push(peripheral.length ? peripheral.join(", ") : "(none)");

  if (hasGSC) {
    const gscGapEntities = coreEntities.filter(function(e) {
      const keyword = e.toLowerCase().split(" ")[0];
      return montrQueries.toLowerCase().includes(keyword) && scores[e] >= priorityThreshold;
    });
    if (gscGapEntities.length) {
      lines.push("");
      lines.push("GSC GAP ENTITIES (high impression, low rank — reinforce these):");
      lines.push(gscGapEntities.join(", "));
    }
  }

  return lines.join("\n");
}

