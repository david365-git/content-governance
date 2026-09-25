/**
 * ============================================================
 * MASTER GOVERNANCE CONTROLLER v4.3
 * Deterministic Ontology Injection — Manual Prompt Safe
 * Sidebar-Compatible Return Structure
 *
 * Revision: Search Entity Support Added
 * Revision Timestamp: 26/02/2026 14:32 (UK)
 * ============================================================
 */

const SHEET_CONFIG = {
  posts:               835712004,
  intentTax:           1313992653,
  techSpec:            1693163314,
  legitMap:            178752392,
  prompts:             1274348303,
  roleControl:         403284831,
  articleTypeControl:  702468283,
  entityTypeControl:   1213368809,
  materialMapping:     835308398,
  export:              "site-export" 
};


/**
 * ============================================================
 * SIDEBAR LOADER
 * ============================================================
 */
function prompt_plus_governance_control_sheets() {
  const html = HtmlService
    .createHtmlOutputFromFile('prompt+governance-control-sheets-html')
    .setTitle('Master Governance Controller v4.3')
    .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * ============================================================
 * CLEAN SHEET VALUE EXTRACTION
 * ============================================================
 */
function getCleanValues(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() == id);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 3);
  if (lastRow === 0) return [];

  return sheet
    .getRange(1, 1, lastRow, lastCol)
    .getValues()
    .map(row =>
      row.map(cell =>
        String(cell === null || cell === undefined ? "" : cell)
      )
    );
}


/**
 * ============================================================
 * FLATTEN FIRST COLUMN (ONTOLOGY LISTS)
 * ============================================================
 */
function flattenFirstColumn(data) {
  if (!data || data.length < 2) return [];

  return data
    .slice(1)
    .map(row => String(row[0] || "").trim())
    .filter(value => value !== "");
}


/**
 /**
 * ============================================================
 * FULL GOVERNANCE PACKAGE BUILDER
 * ============================================================
 */
function getFullGovernancePackage() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheets().find(
      s => s.getSheetId() == SHEET_CONFIG.prompts
    );

    const artOntologyRaw = getCleanValues(SHEET_CONFIG.articleTypeControl);
    const entOntologyRaw = getCleanValues(SHEET_CONFIG.entityTypeControl);

    const allowedArticleTypes = flattenFirstColumn(artOntologyRaw);
    const allowedEntityTypes  = flattenFirstColumn(entOntologyRaw);

    const articleTypeBlock =
      "ALLOWED ARTICLE TYPES (Stage 7 — EXACT STRING MATCH REQUIRED)\n" +
      allowedArticleTypes.map(v => "- " + v).join("\n");

    const entityTypeBlock =
      "ALLOWED ENTITY TYPES (Stage 8 — EXACT STRING MATCH REQUIRED)\n" +
      allowedEntityTypes.map(v => "- " + v).join("\n");

    const manualExecutionNotice =
      "IMPORTANT — MANUAL EXECUTION MODE\n" +
      "This prompt is generated for manual copy-paste use.\n" +
      "No runtime JavaScript objects exist.\n" +
      "Ontology validation MUST rely strictly on the visible whitelist blocks below.\n";

    // --- NEW INSERTION: ARTICLE TYPE DATA GATHERING (Rows 1-10) ---
    const artSheet = ss.getSheets().find(s => s.getSheetId() == "702468283");
    const artDataRange = artSheet.getRange(1, 1, 10, 7).getValues(); 
    
    const articleTypeReferenceBlock = 
      "--- STAGE 7: ARTICLE TYPE REFERENCE (Rows 1-10) ---\n" +
      artDataRange.map(row => row.slice(0, 5).join(" | ")).join("\n");

    const strategyLockBlock = 
      "--- GOVERNANCE STRATEGY LOCKS (Deterministic Flags Rows 1-9) ---\n" +
      artDataRange.slice(0, 9).map(row => {
        return `${row[0]} -> Comm: ${row[5]} | Geo: ${row[6]}`;
      }).join("\n");
    // --- END OF INSERTION ---

    // --- NEW: FETCH THE MAPPED DATA ---
    const rowData = getRowSpecificData();

    // --- NEW: BUILD THE MAPPING INSTRUCTION BLOCK ---
    const materialMappingBlock = 
      "DETERMINISTIC MATERIAL MAPPING (Columns 2 & 3)\n" +
      "PRIMARY ENTITY (Col 2): Use '" + rowData.stoneType + "' + [Detected Intent]\n" +
      "MATERIAL ENTITY (Col 3): Use '" + rowData.materialEntity + "' (STRICT MATCH REQUIRED)\n" +
      "REQUIRED govFlags (Col 9): " + rowData.govFlags; 

    const originalPromptJ16 = String(
      pSheet ? pSheet.getRange("J16").getValue() : ""
    );

    const originalPromptO16 = String(
      pSheet ? pSheet.getRange("O16").getValue() : ""
    );

    // --- ENHANCED PROMPT: ASSEMBLING WITH INSERTED BLOCKS ---
    const enhancedPromptJ16 =
      manualExecutionNotice +
      "\n\n" +
      materialMappingBlock + 
      "\n\n" +
      articleTypeReferenceBlock + 
      "\n\n" +
      strategyLockBlock +      
      "\n\n" +
      articleTypeBlock +
      "\n\n" +
      entityTypeBlock +
      "\n\n" +
      originalPromptJ16;

    return {
      framework: {
        promptJ16: enhancedPromptJ16,
        promptO16: originalPromptO16,
        roles: getCleanValues(SHEET_CONFIG.roleControl),
        taxonomy: getCleanValues(SHEET_CONFIG.intentTax),
        techSpec: getCleanValues(SHEET_CONFIG.techSpec),
        legitMap: getCleanValues(SHEET_CONFIG.legitMap)
      },
      row: rowData 
    };

  } catch (e) {
    throw new Error("Server Error: " + e.message);
  }
}


/**
 * ============================================================
 * ROW-SPECIFIC DATA EXTRACTION
 * ============================================================
 */
/**
 * REVISED: FULL GOVERNANCE PACKAGE BUILDER
 * Injected with Article Type Reference (Rows 1-10) and Strategy Locks (Rows 1-9)
 */
function getFullGovernancePackage() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheets().find(
      s => s.getSheetId() == SHEET_CONFIG.prompts
    );

    const artOntologyRaw = getCleanValues(SHEET_CONFIG.articleTypeControl);
    const entOntologyRaw = getCleanValues(SHEET_CONFIG.entityTypeControl);

    const allowedArticleTypes = flattenFirstColumn(artOntologyRaw);
    const allowedEntityTypes  = flattenFirstColumn(entOntologyRaw);

    const articleTypeBlock =
      "ALLOWED ARTICLE TYPES (Stage 7 — EXACT STRING MATCH REQUIRED)\n" +
      allowedArticleTypes.map(v => "- " + v).join("\n");

    const entityTypeBlock =
      "ALLOWED ENTITY TYPES (Stage 8 — EXACT STRING MATCH REQUIRED)\n" +
      allowedEntityTypes.map(v => "- " + v).join("\n");

    const manualExecutionNotice =
      "IMPORTANT — MANUAL EXECUTION MODE\n" +
      "This prompt is generated for manual copy-paste use.\n" +
      "No runtime JavaScript objects exist.\n" +
      "Ontology validation MUST rely strictly on the visible whitelist blocks below.\n";

    // --- NEW INSERTION: ARTICLE TYPE DATA GATHERING (Rows 1-10) ---
    const artSheet = ss.getSheets().find(s => s.getSheetId() == "702468283");
    const artDataRange = artSheet.getRange(1, 1, 10, 7).getValues(); 
    
    const articleTypeReferenceBlock = 
      "--- STAGE 7: ARTICLE TYPE REFERENCE (Rows 1-10) ---\n" +
      artDataRange.map(row => row.slice(0, 5).join(" | ")).join("\n");

    const strategyLockBlock = 
      "--- GOVERNANCE STRATEGY LOCKS (Deterministic Flags Rows 1-9) ---\n" +
      artDataRange.slice(0, 9).map(row => {
        return `${row[0]} -> Comm: ${row[5]} | Geo: ${row[6]}`;
      }).join("\n");
    // --- END OF INSERTION ---

    // --- FETCH THE MAPPED DATA ---
    const rowData = getRowSpecificData();

    // --- BUILD THE MAPPING INSTRUCTION BLOCK ---
    const materialMappingBlock = 
      "DETERMINISTIC MATERIAL MAPPING (Columns 2 & 3)\n" +
      "PRIMARY ENTITY (Col 2): Use '" + rowData.stoneType + "' + [Detected Intent]\n" +
      "MATERIAL ENTITY (Col 3): Use '" + rowData.materialEntity + "' (STRICT MATCH REQUIRED)\n" +
      "REQUIRED govFlags (Col 9): " + rowData.govFlags;

    const originalPromptJ16 = String(
      pSheet ? pSheet.getRange("J16").getValue() : ""
    );

    const originalPromptO16 = String(
      pSheet ? pSheet.getRange("O16").getValue() : ""
    );

    // --- ENHANCED PROMPT: ASSEMBLING WITH INSERTED BLOCKS ---
    const enhancedPromptJ16 =
      manualExecutionNotice +
      "\n\n" +
      materialMappingBlock + 
      "\n\n" +
      articleTypeReferenceBlock + 
      "\n\n" +
      strategyLockBlock +      
      "\n\n" +
      articleTypeBlock +
      "\n\n" +
      entityTypeBlock +
      "\n\n" +
      originalPromptJ16;

    return {
      framework: {
        promptJ16: enhancedPromptJ16,
        promptO16: originalPromptO16,
        roles: getCleanValues(SHEET_CONFIG.roleControl),
        taxonomy: getCleanValues(SHEET_CONFIG.intentTax),
        techSpec: getCleanValues(SHEET_CONFIG.techSpec),
        legitMap: getCleanValues(SHEET_CONFIG.legitMap)
      },
      row: rowData 
    };

  } catch (e) {
    throw new Error("Server Error: " + e.message);
  }
}

/**
 * REVISED: FULL GOVERNANCE PACKAGE BUILDER
 * Adjusted for Columns A, G, and H (Indices 0, 6, and 7)
 */
function getFullGovernancePackage() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheets().find(
      s => s.getSheetId() == SHEET_CONFIG.prompts
    );

    const artOntologyRaw = getCleanValues(SHEET_CONFIG.articleTypeControl);
    const entOntologyRaw = getCleanValues(SHEET_CONFIG.entityTypeControl);

    const allowedArticleTypes = flattenFirstColumn(artOntologyRaw);
    const allowedEntityTypes  = flattenFirstColumn(entOntologyRaw);

    const articleTypeBlock =
      "ALLOWED ARTICLE TYPES (Stage 7 — EXACT STRING MATCH REQUIRED)\n" +
      allowedArticleTypes.map(v => "- " + v).join("\n");

    const entityTypeBlock =
      "ALLOWED ENTITY TYPES (Stage 8 — EXACT STRING MATCH REQUIRED)\n" +
      allowedEntityTypes.map(v => "- " + v).join("\n");

    const manualExecutionNotice =
      "IMPORTANT — MANUAL EXECUTION MODE\n" +
      "This prompt is generated for manual copy-paste use.\n" +
      "No runtime JavaScript objects exist.\n" +
      "Ontology validation MUST rely strictly on the visible whitelist blocks below.\n";

    // --- NEW INSERTION: ARTICLE TYPE DATA GATHERING (Rows 1-10) ---
    const artSheet = ss.getSheets().find(s => s.getSheetId() == "702468283");
    // Fetching up to Column H (8 columns total)
    const artDataRange = artSheet.getRange(1, 1, 10, 8).getValues(); 
    
    const articleTypeReferenceBlock = 
      "--- STAGE 7: ARTICLE TYPE REFERENCE (Rows 1-10) ---\n" +
      artDataRange.map(row => row.slice(0, 5).join(" | ")).join("\n");

    const strategyLockBlock = 
      "--- GOVERNANCE STRATEGY LOCKS (Deterministic Flags Rows 1-9) ---\n" +
      artDataRange.slice(0, 9).map(row => {
        // Mapping Col A (0) to Col G (6) and Col H (7)
        return `${row[0]} -> Comm: ${row[6]} | Geo: ${row[7]}`;
      }).join("\n");
    // --- END OF INSERTION ---

    // --- FETCH THE MAPPED DATA ---
    const rowData = getRowSpecificData();

    // --- BUILD THE MAPPING INSTRUCTION BLOCK ---
    const materialMappingBlock = 
      "DETERMINISTIC MATERIAL MAPPING (Columns 2 & 3)\n" +
      "PRIMARY ENTITY (Col 2): Use '" + rowData.stoneType + "' + [Detected Intent]\n" +
      "MATERIAL ENTITY (Col 3): Use '" + rowData.materialEntity + "' (STRICT MATCH REQUIRED)\n" +
      "REQUIRED govFlags (Col 9): " + rowData.govFlags;

    const originalPromptJ16 = String(
      pSheet ? pSheet.getRange("J16").getValue() : ""
    );

    const originalPromptO16 = String(
      pSheet ? pSheet.getRange("O16").getValue() : ""
    );

    // --- ENHANCED PROMPT: ASSEMBLING WITH INSERTED BLOCKS ---
    const enhancedPromptJ16 =
      manualExecutionNotice +
      "\n\n" +
      materialMappingBlock + 
      "\n\n" +
      articleTypeReferenceBlock + 
      "\n\n" +
      strategyLockBlock +      
      "\n\n" +
      articleTypeBlock +
      "\n\n" +
      entityTypeBlock +
      "\n\n" +
      originalPromptJ16;

    return {
      framework: {
        promptJ16: enhancedPromptJ16,
        promptO16: originalPromptO16,
        roles: getCleanValues(SHEET_CONFIG.roleControl),
        taxonomy: getCleanValues(SHEET_CONFIG.intentTax),
        techSpec: getCleanValues(SHEET_CONFIG.techSpec),
        legitMap: getCleanValues(SHEET_CONFIG.legitMap)
      },
      row: rowData 
    };

  } catch (e) {
    throw new Error("Server Error: " + e.message);
  }
}

/**
 * REVISED: ROW-SPECIFIC DATA EXTRACTION
 * Adjusted to pull from Col G (6) and Col H (7)
 */
function getRowSpecificData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheets().find(
    s => s.getSheetId() == SHEET_CONFIG.posts
  );

  const activeRowIndex = ss.getActiveRange().getRow();

  if (activeRowIndex < 2)
    throw new Error("Select a data row in the Posts sheet.");

  const postsHeaders = postsSheet
    .getRange(1, 1, 1, postsSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).toLowerCase().trim());

  const activePostsRow = postsSheet
    .getRange(activeRowIndex, 1, 1, postsSheet.getLastColumn())
    .getValues()[0];

  const findPostVal = (name) => {
    const idx = postsHeaders.indexOf(name.toLowerCase().trim());
    return idx > -1 ? activePostsRow[idx] : "";
  };

  const postId =
    findPostVal("Post ID") ||
    postsSheet.getRange(activeRowIndex, 4).getValue();

  const articleType = String(findPostVal("Article Type") || "").trim();

  const exportSheet = ss.getSheetByName(SHEET_CONFIG.export);

  if (!exportSheet)
    throw new Error(
      "Sheet '" + SHEET_CONFIG.export + "' not found."
    );

  const exportData = exportSheet.getDataRange().getValues();
  const exportHeaders = exportData[0].map(h =>
    String(h).toLowerCase().trim()
  );

  const matchIndex = exportData.findIndex(
    row => String(row[0]) === String(postId)
  );

  if (matchIndex === -1)
    throw new Error(
      "ID " + postId + " not found in '" +
      SHEET_CONFIG.export + "'."
    );

  const matchRow = exportData[matchIndex];

  const findExportVal = (name) => {
    const idx = exportHeaders.indexOf(
      name.toLowerCase().trim()
    );
    return idx > -1 ? String(matchRow[idx] || "") : "";
  };

  const fullHtml = findExportVal("Full Post HTML");
  const hasAmazon = fullHtml.includes("amzn.to") || fullHtml.includes("amazon.co.uk");

  // --- DETERMINISTIC FLAG LOOKUP (Col A=0, Col G=6, Col H=7) ---
  const artSheet = ss.getSheets().find(s => s.getSheetId() == "702468283");
  const artData = artSheet.getDataRange().getValues();
  
  let commFlag = "[COMM:No]"; 
  let geoFlag = "[GEOCTA:No]";

  for (let i = 1; i < artData.length; i++) {
    if (String(artData[i][0]).trim() === articleType) {
      commFlag = artData[i][6] || "[COMM:No]"; // Column G (Index 6)
      geoFlag = artData[i][7] || "[GEOCTA:No]"; // Column H (Index 7)
      break;
    }
  }

  if (commFlag === "[COMM:BelowFold]" && !hasAmazon) {
    commFlag = "[COMM:No]";
  }

  const gscBlock =
`--- PAGE METADATA (site-export) ---
Post ID: ${postId}
Title: ${findExportVal("Title")}
URL: ${findExportVal("Canonical URL")}
Dates: Pub: ${findExportVal("Published Date")} | Upd: ${findExportVal("Last Updated Date")}
Meta Title: ${findExportVal("Meta Title") || "(none)"}
Meta Description: ${findExportVal("Meta Description") || "(none)"}
Schema: ${findExportVal("WPCode Header Schema") || "(none)"}

--- PERFORMANCE DATA (posts) ---
Clicks: ${findPostVal("Clicks")} | Impressions: ${findPostVal("Impressions")}
Top 5 Queries:
${String(findPostVal("Queries")).split('\n').slice(0, 6).join('\n')}

--- FULL POST CONTENT ---
${fullHtml}`;

  const searchEntity = String(findPostVal("Search Entity") || "").trim();
  const stoneType    = String(findPostVal("Stone Type") || "Unknown").trim();

  const materialEntity = getTargetMaterialEntity(stoneType, searchEntity);

  return {
    title: findExportVal("Title"),
    articleType: articleType,
    govFlags: commFlag + geoFlag,
    stoneType: stoneType,
    searchEntity: searchEntity,
    materialEntity: materialEntity, 
    gscBlock: gscBlock
  };
}
/**
 * ENGINE: TARGET MATERIAL ENTITY LOOKUP
 * Matches Hub (stoneType) and Variant (searchEntity) to find the target Col 3 value.
 */
function getTargetMaterialEntity(hub, variant) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(s => s.getSheetId() == SHEET_CONFIG.materialMapping);
    if (!sheet) return hub; 

    const data = sheet.getDataRange().getValues();
    
    // Helper to strip "s" and trim spaces
    const clean = (str) => String(str || "").trim().toLowerCase().replace(/s$/, "");

    const searchHub = clean(hub);
    const searchVar = clean(variant);

    const match = data.find(row => 
      clean(row[0]) === searchHub && clean(row[1]) === searchVar
    );

    return (match && match[2]) ? String(match[2]).trim() : hub;
  } catch (e) {
    return hub;
  }
}