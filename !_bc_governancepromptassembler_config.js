/**
 * ============================================================
 * bc_GovernancePromptAssembler_Config.gs
 * Abbey Floor Care — Governance Prompt Assembler
 * Shared Configuration and Utility Functions
 *
 * Prefix: bc_ (loads before ce_ files in Apps Script editor)
 * Version: 2.3 — serviceAuthority sheet ID added to BC_SHEET_CONFIG
 * ============================================================
 */

const BC_SHEET_CONFIG = {
  posts:              835712004,
  intentTax:          1313992653,
  techSpec:           1693163314,
  legitMap:           178752392,
  prompts:            1274348303,
  roleControl:        403284831,
  articleTypeControl: 702468283,
  entityTypeControl:  1213368809,
  materialMapping:    835308398,
  roleMapping:        641120685,
  serviceAuthority:   382558434,
  export:             "site-export"
};

// Prompt cell references in the Prompts sheet
const BC_PROMPT_CELLS = {
  section1: "J16",
  module1:  "D17",
  module2:  "D18",
  module3:  "D19",
  module4:  "D20"
};


/**
 * Get all values from a sheet by sheet ID.
 */
function bc_getCleanValues(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() == id);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 3);
  if (lastRow === 0) return [];
  return sheet
    .getRange(1, 1, lastRow, lastCol)
    .getValues()
    .map(row => row.map(cell => String(cell === null || cell === undefined ? "" : cell)));
}


/**
 * Flatten first column of a data array (skipping header row).
 */
function bc_flattenFirstColumn(data) {
  if (!data || data.length < 2) return [];
  return data.slice(1)
    .map(row => String(row[0] || "").trim())
    .filter(value => value !== "");
}


/**
 * Convert a 2D array to a tab-delimited string.
 */
function bc_toTDF(data) {
  if (!data || data.length === 0) return "(No data)";
  return data.map(row => Array.isArray(row) ? row.join('\t') : String(row)).join('\n');
}


/**
 * Filter a 2D data array to rows where column 0 contains
 * any of the supplied search terms (case-insensitive).
 * Always keeps the header row (index 0).
 */
function bc_filterByMaterial(rows, terms) {
  if (!rows || rows.length < 2) return "(Empty)";
  const filtered = rows.filter((r, i) =>
    i === 0 || (r[0] && terms.some(t => String(r[0]).toLowerCase().includes(t)))
  );
  return bc_toTDF(filtered);
}


/**
 * Build the pipe-separated 3-column role mapping lookup table
 * for the current page's Article Type.
 *
 * Reads the role-mapping sheet (641120685) and filters to rows
 * matching the supplied articleType. Returns a formatted block
 * ready for injection into Prompt 1.
 *
 * Format per row: Intent | Entity Role | Entity Type
 */
function bc_buildRoleMappingBlock(articleType) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() == BC_SHEET_CONFIG.roleMapping);

  if (!sheet) return "(Role mapping sheet not found — check sheet ID 641120685)";

  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return "(Role mapping sheet is empty)";

const rows = data.slice(1);
  const cleanAT = String(articleType || "").trim().toLowerCase();

  const matched = cleanAT
    ? rows.filter(row => String(row[1] || "").trim().toLowerCase() === cleanAT)
    : rows;

  if (cleanAT && matched.length === 0) {
    return "(No mapping rows found for Article Type: " + articleType + ")";
  }

  const lines = matched.map(row =>
    String(row[1] || "").trim() +
    " | " +
    String(row[0] || "").trim() +
    " | " +
    String(row[2] || "").trim() +
    " | " +
    String(row[3] || "").trim()
  );

  const filterNote = cleanAT
    ? "ARTICLE TYPE FILTER APPLIED: " + articleType
    : "NO ARTICLE TYPE PRE-ASSIGNED — FULL TABLE SUPPLIED (all Article Types). Once you confirm the Article Type, filter this table yourself to the matching Article Type row before selecting Entity Role and Entity Type.";

  return (
    "===== ENTITY ROLE LOOKUP (HARD LOCK) =====\n" +
    "This table is the ONLY permitted source for Entity Role and Entity Type values.\n" +
    "Any value not copied EXACTLY from this table is incorrect.\n" +
    "This is a NON-REASONING STEP. It is a direct lookup task.\n\n" +
    filterNote + "\n\n" +
    "Article Type | Intent | Entity Role | Entity Type\n" +
    lines.join("\n") +
    "\n\n" +
    "Execution steps:\n" +
    "1. Take the CONFIRMED ARTICLE TYPE and CONFIRMED PRIMARY INTENT from the lock block at the top of this prompt.\n" +
    "2. Find the EXACT matching Article Type + Intent row in the table above.\n" +
    "3. COPY Entity Role and Entity Type EXACTLY as written — no modification.\n" +
    "4. DO NOT infer from page content.\n" +
    "5. DO NOT use any value not present in this table.\n" +
    "6. If no match found → output: ERROR: INTENT NOT FOUND IN MAPPING TABLE\n\n" +
    "VALIDATION — check before proceeding:\n" +
    "- Entity Role EXACTLY matches a value from the table above\n" +
    "- Entity Type EXACTLY matches a value from the table above\n" +
    "- If not → STOP and correct using the table\n" +
    "===== END SECTION ====="
  );
}


/**
 * Read multiple governed column values from the active row.
 * Returns object keyed by column name.
 */
function bc_readActiveRowFields(fieldNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheets().find(s => s.getSheetId() == BC_SHEET_CONFIG.posts);
  if (!postsSheet) return {};

  const activeRowIndex = ss.getActiveRange().getRow();
  if (activeRowIndex < 2) return {};

  const headers = postsSheet
    .getRange(1, 1, 1, postsSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).toLowerCase().trim());

  const row = postsSheet
    .getRange(activeRowIndex, 1, 1, postsSheet.getLastColumn())
    .getValues()[0];

  const result = {};
  for (const name of fieldNames) {
    const idx = headers.indexOf(name.toLowerCase().trim());
    result[name] = idx > -1 ? String(row[idx] || "") : "";
  }
  return result;
}


/**
 * Open the Governance Prompt Assembler sidebar.
 */
function bc_showGovernancePromptAssembler() {
  const html = HtmlService
    .createTemplateFromFile('bc_GovernancePromptAssembler_Dashboard-html')
    .evaluate()
    .setTitle('Governance Prompt Assembler')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(560)
    .setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Governance Prompt Assembler');
}


/**
 * Add menu item on spreadsheet open.
 */
function bc_onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 BC Governance')
    .addItem('Open Governance Prompt Assembler', 'bc_showGovernancePromptAssembler')
    .addToUi();
}
