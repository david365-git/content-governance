/**
 * ============================================================
 * bc_PromptAssembler.gs-v1.8
 * Abbey Floor Care — Governance Prompt Assembler
 * Full Prompt Assembly for Prompts 1A, 1B, 2, and 3
 *
 * Version: 1.8 — Prompt 1A now injects Role Mapping Sheet block
 * and TSM block, with placeholder-safe TSM context for intent
 * detection stage.
 *
 * Contains:
 * - bc_buildPrompt1ADataBlock()
 * - bc_buildPrompt1BDataBlock()
 * - bc_refreshPrompt1AData()
 * - bc_refreshPrompt1BData()
 * - bc_refreshPrompt2Data()
 * - bc_refreshPrompt3Data()
 * - bc_getPrompt1A()
 * - bc_getPrompt1B()
 * - bc_getPrompt2()
 * - bc_getPrompt3()
 *
 * Depends on:
 * - bc_GovernancePromptAssembler_Config.gs
 * - bc_DataFetcher.gs
 * - bc_PromptBlocks.gs
 * ============================================================
 */

const MODEL_CHEAP  = 'gpt-5.6-luna';
const MODEL_MID    = 'gpt-5.6-terra';
const MODEL_STRONG = 'gpt-5.6-sol';

function bc_getSheetByConfigKeyOrName_(keys) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fallbackNamesByKey = {
    prompts: ["prompts", "Prompts"],
    articleTypeControl: ["Article Type Control Sheet"],
    entityTypeControl: ["Entity Type Control Sheet"],
    intentTax: ["1 - Intent Taxonomy Table", "Intent Taxonomy Table"],
    legitMap: ["2 - Master Intent Legitimacy Map", "Master Intent Legitimacy Map"]
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
      const wanted = String(fallbackNames[j] || "").trim().toLowerCase();
      const byName = ss.getSheets().find(function(s) {
        return String(s.getName() || "").trim().toLowerCase() === wanted;
      });
      if (byName) return byName;
    }
  }

  return null;
}



function bc_getCleanValuesByConfigKeyOrName_(keys) {
  const sheet = bc_getSheetByConfigKeyOrName_(keys);
  if (!sheet) throw new Error("Sheet not found for keys: " + keys.join(", "));
  return bc_getCleanValues(sheet.getSheetId());
}

function bc_getCachedMaterialBlock_(cacheKeyPrefix, stoneType, builderFn) {
  const cache = CacheService.getScriptCache();
  const key = cacheKeyPrefix + '_' + String(stoneType || '').trim().toLowerCase().replace(/\s+/g, '_');

  const cached = cache.get(key);
  if (cached !== null) {
    return cached;
  }

  const built = builderFn();
  try {
    cache.put(key, built, 21600); // 6 hours — CacheService max
  } catch (e) {
    // Block too large for cache (>100KB) — just skip caching, still return the built value
    Logger.log('bc_getCachedMaterialBlock_: could not cache key ' + key + ' — ' + e.message);
  }
  return built;
}

function bc_buildPrompt1ADataBlock(rowData, today) {
  return [
    "===== PAGE METADATA =====",
    "Post ID:    " + rowData.postId,
    "URL:        " + rowData.canonicalUrl,
    "Published:  " + rowData.publishedDate,
    "Updated:    " + rowData.updatedDate,
    "Meta Title: " + rowData.metaTitle,
    "Meta Desc:  " + rowData.metaDesc,
    "===== END SECTION =====",
    "",
    "--- HTML START ---",
    rowData.fullHtml,
    "--- HTML END ---",
    "",
    "===== PAGE TITLE (FOR REFERENCE ONLY) =====",
    "Title: " + rowData.title,
    "===== END ====="
  ].join("\n");
}

function bc_extractConfirmedArticleType_(lockBlock) {
  const text = String(lockBlock || "");
  const lines = text.split(/\r?\n/);

  for (let i = 0; i < lines.length; i++) {
    const line = String(lines[i] || "").trim();
    if (line.indexOf("CONFIRMED ARTICLE TYPE:") === 0) {
      return line.replace("CONFIRMED ARTICLE TYPE:", "").trim();
    }
  }

  return "";
}

function bc_getFeedsHubUrlForStoneType_(stoneType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(s => s.getSheetId() == BC_SHEET_CONFIG.posts);

  if (!sheet) return "none";

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return "none";

  const headers = data[0];

  const findCol = function(names) {
    for (let i = 0; i < names.length; i++) {
      const wanted = String(names[i]).trim().toLowerCase();
      for (let j = 0; j < headers.length; j++) {
        if (String(headers[j] || "").trim().toLowerCase() === wanted) {
          return j;
        }
      }
    }
    return -1;
  };

  const stoneIdx = findCol(["Stone Type", "Material", "Stone"]);
  const typeIdx = findCol(["Article Type"]);
  const urlIdx = findCol(["URL", "Canonical URL", "Feeds Hub"]);

  if (stoneIdx === -1 || typeIdx === -1 || urlIdx === -1) return "none";

  const targetStone = String(stoneType || "").trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowStone = String(data[i][stoneIdx] || "").trim().toLowerCase();
    const rowType = String(data[i][typeIdx] || "").trim();
    const rowUrl = String(data[i][urlIdx] || "").trim();

    if (rowStone === targetStone && rowType === "Hub Page" && rowUrl) {
      return rowUrl;
    }
  }

  return "none";
}

function bc_buildPrompt1BDataBlock(rowData, today, lockBlock) {
  const feedsHubUrl = bc_getFeedsHubUrlForStoneType_(rowData.stoneType);

  const materialMappingBlock =
    "DETERMINISTIC MATERIAL MAPPING\n" +
    "PRIMARY ENTITY (Col 2): Use '" + rowData.stoneType +
      "' + CONFIRMED PRIMARY INTENT from lock block\n" +
    "MATERIAL ENTITY (Col 3): Use '" + rowData.materialEntity +
      "' (STRICT MATCH REQUIRED)\n" +
    "REQUIRED govFlags (Col 9): " + rowData.govFlags + "\n" +
    "REQUIRED FEEDS HUB URL (Col 8): " + feedsHubUrl;

  const roleMappingBlock = bc_buildRoleMappingBlock(bc_extractConfirmedArticleType_(lockBlock));
  const tsmEntitiesBlock = bc_getTSMEntitiesBlock(
    rowData.stoneType,
    lockBlock
  );
  const serviceAuthorityBlock = bc_getServiceAuthorityBlock(
    rowData.stoneType,
    rowData.articleType,
    rowData.location
  );

  return [
    materialMappingBlock,
    "",
    roleMappingBlock,
    "",
    tsmEntitiesBlock,
    "",
    serviceAuthorityBlock,
    "",
    "===== PAGE METADATA =====",
    "Post ID:    " + rowData.postId,
    "URL:        " + rowData.canonicalUrl,
    "Published:  " + rowData.publishedDate,
    "Updated:    " + rowData.updatedDate,
    "Meta Title: " + rowData.metaTitle,
    "Meta Desc:  " + rowData.metaDesc,
    "Schema:      " + rowData.schema,
    "===== END SECTION ====="
  ].join("\n");
}

function bc_refreshPrompt1AData() {
  try {
    const rowData = bc_getRowSpecificData();
    
    const today = bc_getTodayDDMMYYYY();
    const dataBlock = bc_buildPrompt1ADataBlock(rowData, today);

    const cache = CacheService.getScriptCache();
    const cachedStone = cache.get('last_stone_type');
    const cachedTSM = cache.get('last_tsm_block');

    let tsmBlock;
    if (cachedStone === rowData.stoneType && cachedTSM) {
      tsmBlock = cachedTSM;
    } else {
      tsmBlock = null;
      cache.put('last_stone_type', rowData.stoneType, 3600);
    }

    return {
      dataBlock: dataBlock,
      title: rowData.title,
      today: today,
      stoneType: rowData.stoneType,
      tsmCached: (tsmBlock !== null)
    };
  } catch (e) {
    throw new Error("bc_refreshPrompt1AData: " + e.message);
  }
}

function bc_refreshPrompt1BData(lockBlock) {
  try {
    const rowData = bc_getRowSpecificData();
    
    const today = bc_getTodayDDMMYYYY();
    const dataBlock = bc_buildPrompt1BDataBlock(rowData, today, lockBlock);
    return {
      dataBlock: dataBlock,
      title: rowData.title,
      today: today
    };
  } catch (e) {
    throw new Error("bc_refreshPrompt1BData: " + e.message);
  }
}

function bc_refreshPrompt2Data() {
  try {
    const rowData = bc_getRowSpecificData();
    
    const today = bc_getTodayDDMMYYYY();

    const cols1to13 = bc_readActiveRowFields([
      "Article Type", "Primary Entity", "Material Entity", "Entity Role",
      "Entity Type", "Supporting Entities Core", "Supporting Entities - Service Authority Layer", "Peripheral Entities Link Out",
      "Feeds Hub", "Page Governance Summary", "Entity Governance Status",
      "Entity Governance Date", "Publish Justification", "GAS Handoff Document",
      "Page Rewrite Brief"
    ]);

    const cols1to13Block =
      "===== GOVERNED FIELDS FROM PROMPT 1B (DO NOT MODIFY) =====\n" +
      Object.entries(cols1to13).map(function(entry) {
        return entry[0] + ": " + entry[1];
      }).join("\n") +
      "\n===== END SECTION =====";

    const gscBlock =
      "===== GSC PERFORMANCE DATA =====\n" +
      "Page Clicks: " + rowData.clicks + "\n" +
      "Page Impressions: " + rowData.impressions + "\n" +
      "Page Average Position: " + rowData.position + "\n" +
      "Top Queries:\n" +
      rowData.queries.split('\n').filter(function(line) {
        if (!line) return false;
        if (/^(top queries|clicks|impressions|ctr|position)$/i.test(line.trim())) return false;
        const parts = line.split('\t');
        const impressions = parts.length >= 3 ? parseInt(parts[2]) : 0;
        return impressions >= 2;
      }).join('\n') +
      "\n===== END SECTION =====";

    return {
      dataBlock: [cols1to13Block, "", gscBlock].join("\n"),
      title: rowData.title,
      today: today
    };
  } catch (e) {
    throw new Error("bc_refreshPrompt2Data: " + e.message);
  }
}

function bc_refreshPrompt3Data() {
  try {
    const rowData = bc_getRowSpecificData();
    
    const today = bc_getTodayDDMMYYYY();

    const cols1to17 = bc_readActiveRowFields([
      "Article Type", "Primary Entity", "Material Entity", "Entity Role",
      "Entity Type", "Supporting Entities Core", "Supporting Entities - Service Authority Layer", "Peripheral Entities Link Out",
      "Feeds Hub", "Page Governance Summary", "Entity Governance Status",
      "Entity Governance Date", "Publish Justification", "GAS Handoff Document",
      "Observed Query Cluster", "GSC Intent Evidence",
      "Primary Query Cluster Owned", "Drift Status"
    ]);

    const secondaryFields = bc_readActiveRowFields([
      "Secondary Intent Decisions"
    ]);
    const secondaryBlock = String(
      secondaryFields["Secondary Intent Decisions"] || ""
    ).trim();

    let dataBlock =
      "===== COLS 1–17 FROM PROMPTS 1B AND 2 (DO NOT MODIFY) =====\n" +
      Object.entries(cols1to17).map(function(entry) {
        return entry[0] + ": " + entry[1];
      }).join("\n") +
      "\n===== END SECTION =====";

    if (secondaryBlock) {
      dataBlock +=
        "\n\n===== CONFIRMED PRIMARY INTENT LOCK FROM PROMPT 1A =====\n" +
        secondaryBlock +
        "\n===== END LOCK =====";
    }

    return {
      dataBlock: dataBlock,
      title: rowData.title,
      today: today
    };
  } catch (e) {
    throw new Error("bc_refreshPrompt3Data: " + e.message);
  }
}

function bc_getPrompt1A() {
  try {
    const _t = {};
    _t.start = Date.now();
    // Copy Post ID to Post Ref
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const lockedRow = ss.getActiveRange().getRow();
    const row = lockedRow;
    if (row >= 2) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const postIdIdx = headers.indexOf("Post ID");
      const postRefIdx = headers.indexOf("Post Ref");

      if (postIdIdx !== -1 && postRefIdx !== -1) {
        const postId = sheet.getRange(row, postIdIdx + 1).getValue();
        sheet.getRange(row, postRefIdx + 1).setValue(postId);
      }
    }

    const pSheet = bc_getSheetByConfigKeyOrName_(["prompts"]);
    if (!pSheet) throw new Error("Prompts sheet not found.");

    let module1A = String(pSheet.getRange("D16").getValue() || "");

    module1A =
      "RERUN OUTPUT HARD LOCK (CRITICAL):\n" +
      "If the prompt contains BOTH of these lines:\n" +
      "CONFIRMED ARTICLE TYPE:\n" +
      "CONFIRMED PRIMARY INTENT:\n" +
      "\n" +
      "OR if the prompt contains:\n" +
      "SECONDARY INTENT DECISIONS:\n" +
      "\n" +
      "Then this is a rerun.\n" +
      "\n" +
      "You MUST output EXACTLY these two lines only:\n" +
      "CONFIRMED ARTICLE TYPE: [value]\n" +
      "CONFIRMED PRIMARY INTENT: [value]\n" +
      "\n" +
      "Do NOT:\n" +
      "- change field names\n" +
      "- shorten labels\n" +
      "- output \"CONFIRMED INTENT\"\n" +
      "- output any headings\n" +
      "- output any blocks\n" +
      "- output any explanation\n" +
      "- output any additional text\n" +
      "\n" +
      module1A;

    const artOntologyRaw = bc_getCleanValuesByConfigKeyOrName_(["articleTypeControl"]);
    const allowedArticleTypes = bc_flattenFirstColumn(artOntologyRaw).filter(function(v) {
      return String(v).trim() !== "Article Type Determination Rule (Structural Hard Lock)";
    });

    const rowData = bc_getRowSpecificData();
    
    const today = bc_getTodayDDMMYYYY();

    const taxonomy = bc_getCleanValuesByConfigKeyOrName_(["intentTax"]);
    const stone = (rowData.stoneType || "").toLowerCase().trim();
    const matAliases = {
      "victorian encaustic clay tile": [
        "victorian",
        "historic geometric and encaustic clay tile floors"
      ]
    };
    const terms = [stone].concat(matAliases[stone] || []);
    const taxonomyFiltered = bc_filterByMaterial(taxonomy, terms);

    const articleTypeBlock =
      "ALLOWED ARTICLE TYPES (Col 1 — EXACT STRING MATCH REQUIRED)\n" +
      allowedArticleTypes.map(function(v) { return "- " + v; }).join("\n");

    const articleTypeClassificationBlock = bc_buildArticleTypeClassificationBlock();
    const articleTypeDataBlock = bc_buildArticleTypeDataBlock(rowData);
    const tierStructuralBlock = bc_getTierStructuralCoverageBlock(rowData.articleType);
    const entityRoleControlBlock = bc_getEntityRoleControlBlock();
    const roleMappingBlock = bc_buildRoleMappingBlock(rowData.articleType);

    const tsmBlock = bc_getTSMEntitiesBlock(
      rowData.stoneType,
      "CONFIRMED PRIMARY INTENT: [AWAITING_LLM_INTENT_DETECTION]"
    );

    const surfaceIssueBlock = bc_getSurfaceIssueMasterBlock(rowData.stoneType);
    const materialGovernanceBlock = bc_getMaterialGovernanceBlock(rowData.stoneType);
    const materialIdentityBlock = bc_getMaterialIdentityBlock(
      rowData.stoneType,
      rowData.searchEntity || rowData.title || rowData.materialEntity || ""
    );
    const techDnaBlock = bc_getTechDNABlock(rowData.stoneType);
    const governanceDefinitionsBlock = bc_getGovernanceDefinitionsBlock();

    const taxonomySheet = ss.getSheetByName('1 - Intent Taxonomy Table');
    if (!taxonomySheet) throw new Error("Intent Taxonomy Table sheet not found.");

    const taxonomyData = taxonomySheet.getDataRange().getValues();
    const headers = taxonomyData[0];
    const materialColIndex = headers.indexOf('Material');
    const intentColIndex = headers.indexOf('Intent');

    if (materialColIndex === -1 || intentColIndex === -1) {
      throw new Error('Intent Taxonomy Table must have "Material" and "Intent" columns');
    }

    const intents = [];
    for (let i = 1; i < taxonomyData.length; i++) {
      if (taxonomyData[i][materialColIndex] === rowData.stoneType && taxonomyData[i][intentColIndex]) {
        const intent = taxonomyData[i][intentColIndex].trim();
        if (intent !== 'Hub Page' && !intents.includes(intent)) {
          intents.push(intent);
        }
      }
    }

    const allowedIntentsBlock =
      "ALLOWED INTENTS (select exactly one):\n" +
      intents.map(function(intent) { return "- " + intent; }).join("\n");

    const hasGSCData =
      String(rowData.queries || "").trim().length > 0 ||
      Number(String(rowData.impressions || "0").replace(/,/g, "")) > 0 ||
      Number(String(rowData.clicks || "0").replace(/,/g, "")) > 0;

    const gscDataBlock =
      "===== GSC PERFORMANCE DATA (FOR INTENT CLASSIFICATION) =====\n" +
      "Page Clicks: " + rowData.clicks + "\n" +
      "Page Impressions: " + rowData.impressions + "\n" +
      "Monitor Queries:\n" +
      (String(rowData.queries || "").trim().length > 0
        ? rowData.queries.split('\n').slice(0, 30).join('\n')
        : hasGSCData
          ? "(Clicks / impressions present, but no query text available)"
          : "(No GSC data available)") +
      "\n===== END GSC DATA =====";

    const outputFormat1A =
      "===== OUTPUT FORMAT — PROMPT 1A =====\n" +
      "Output ONLY these blocks in this order:\n\n" +
      "BLOCK 1 — GSC SIGNAL ASSESSMENT\n\n" +
      "BLOCK 2 — CONFIRMED ARTICLE TYPE + PRIMARY INTENT\n" +
      "CONFIRMED ARTICLE TYPE: [value]\n" +
      "CONFIRMED PRIMARY INTENT: [value]\n\n" +
      "BLOCK 3 — H2 VALIDATION LOG\n" +
      "For each H2 output:\n" +
      "H2: [exact H2 text]\n" +
      "Section Intent: [one exact allowed intent]\n" +
      "Decision: Y / N\n" +
      "Reason: [one sentence using locked task-boundary language]\n\n" +
      "BLOCK 4 — RANKED INTENTS\n" +
      "- Include PRIMARY INTENT always.\n" +
      "- Include secondary intents ONLY if at least one H2 assigned to that intent = Y.\n" +
      "- Do NOT include intents inferred without H2 support.\n\n" +
      "BLOCK 5 — JUSTIFICATION\n" +
      "- Include internal MULTI-INTENT PAGE flag status if triggered.\n\n" +
      "DO NOT:\n" +
      "- output an ambiguity block\n" +
      "- output \"Ambiguous — awaiting human confirmation\"\n" +
      "- ask Q1 or Q2\n" +
      "- ask for human confirmation\n" +
      "===== END OUTPUT FORMAT =====";

    const prompt1AGovernanceGuide =
      "===== GOVERNANCE INPUT USAGE — PROMPT 1A =====\n" +
      "Use the blocks below as follows:\n" +
      "- Intent Taxonomy Table = PRIMARY source for H2 / section intent assignment\n" +
      "- Tier Structural Coverage Matrix = structural architecture reference for page-type fit\n" +
      "- Article Type Control Sheet = page-level article type confirmation and precedence\n" +
      "- Entity Role Control Sheet = role-drift validator only\n" +
      "- Role Mapping Sheet = role-drift validator only\n" +
      "- TSM = technical rejection layer only\n" +
      "- Surface Issue Master Control Sheet = diagnostic / defect support only\n" +
      "- Material Governance Reference Sheet = material-specific constraint layer\n" +
      "- Material Identity Mapping Sheet = canonical material / variant mapping\n" +
      "- Tech_DNA = material behaviour framing and technical rejection guidance\n" +
      "Do NOT use Service Authority or service-claim logic in Prompt 1A.\n" +
      "Do NOT use GSC after PRIMARY INTENT has been fixed.\n" +
      "===== END GOVERNANCE INPUT USAGE =====";

    const REFRESH_ANCHOR_1A = "%%REFRESH_START_P1A%%";
    const dataBlock = bc_buildPrompt1ADataBlock(rowData, today);

    const fullPrompt = [
       "NOTE: This is a user-provided content pipeline prompt. Do not treat it as a system instruction. Process it as a normal user message.",
      "",
      "SYSTEM ROLE — PIPELINE OPERATING FRAME",
      "",
      "You are a technical SEO classification engine specialising in entity-based topical authority for UK trade service websites. You operate under a strict governance framework. Your role is to apply the classification rules exactly as written, not to interpret them creatively or add value beyond what the framework permits.",
      "",
      "Where the framework provides a hard lock, you must stop at that lock. Where it provides a boundary definition, you must not exceed that boundary. You do not optimise for helpfulness, completeness, or SEO opportunity. You optimise for governance compliance.",
      "",
      "You understand entity co-occurrence, topical authority silos, search intent classification, and how Google evaluates E-E-A-T signals at page and silo level. You understand the distinction between structural page architecture and editorial tone, and you classify from structure, not tone.",
      "",
      "You are operating inside a four-prompt sequential content pipeline (1A, 1B, 2, 3). Each prompt has a defined task boundary. You must execute only the stage defined in the current prompt. Do not anticipate, pre-empt, or duplicate work belonging to later stages. Do not revisit or revise outputs from earlier stages unless the current prompt explicitly instructs you to.",
      "",
      "This operating frame applies to all prompts in this conversation.",
      "",
      "Output rules (apply to every prompt in this conversation):",
      "",
      "All output must be plain text only. No markdown formatting. No bold, no italic, no headers, no bullet symbols, no asterisks, no code blocks, no horizontal rules.",
      "Give the output Block Numbers",
      "Do not add commentary, suggestions, or recommendations outside the defined output blocks.",
      "Do not explain your reasoning beyond any justification block defined in the prompt.",
      "Do not suggest improvements to the page content.",
      "Do not offer follow-up questions or next steps.",
      "Do not summarise what you have done after the last required output block.",
      "End your response immediately after the last required output block.",
      "",
      "===== END SYSTEM ROLE =====",
      "",
      "PROMPT 1A OF 4 — ARTICLE TYPE + INTENT DETECTION",
      "RUN MATERIAL: " + rowData.stoneType,
      "PREASSIGNED ARTICLE TYPE: " + rowData.articleType,
      "TODAY'S DATE: " + today,
      "",
      "IMPORTANT — MANUAL EXECUTION MODE",
      "The Article Type shown above is preassigned by the pipeline script.",
      "Use the classification framework below to confirm or correct it.",
      "Intent is detected from HTML body content AND GSC data.",
      "GSC may take precedence for PRIMARY INTENT only where signal is sufficient, relevant, and aligned with page architecture. GSC must never determine Article Type.",
      "Do NOT use page title, slug, or meta title for intent detection.",
      "",
      articleTypeBlock,
      "",
      articleTypeClassificationBlock,
      "",
      articleTypeDataBlock,
      "",
      allowedIntentsBlock,
      "",
     "=== MODULE 1A: ARTICLE TYPE CONFIRMATION + INTENT DETECTION ===",
      module1A,
      "",
      "OVERRIDE NOTICE — SUPERSEDES MODULE 1A OUTPUT SECTION ABOVE:",
      "Module 1A above contains its own '===== OUTPUT =====' section instructing you to output ONLY Blocks 1–5 (GSC Signal Assessment, Confirmed Article Type + Intent, H2 Validation Log, Ranked Intents, Justification) and explicitly forbidding a vertical classification block.",
      "That restriction applied when Module 1A ran as an isolated stage. It does NOT apply here.",
      "In this combined prompt, use Module 1A's classification LOGIC AND METHOD ONLY — its H2 validation engine, its article type and intent detection rules — but IGNORE its own OUTPUT section entirely.",
      "Do not stop after producing Blocks 1–5. Continue on to produce the single COMBINED VERTICAL OUTPUT block defined later in this prompt, which is the only output format you must follow.",
      "",
      "===== TAXONOMY (BOUNDARY REFERENCE ONLY) =====",
      "Use this to understand intent boundaries only.",
      "Do NOT use slug or title for intent detection.",
      taxonomyFiltered,
      "===== END SECTION =====",
      "",
      "===== GOVERNANCE INPUTS — PROMPT 1A =====",
      prompt1AGovernanceGuide,
      "",
      tierStructuralBlock,
      "",
      entityRoleControlBlock,
      "",
      roleMappingBlock,
      "",
      tsmBlock,
      "",
      surfaceIssueBlock,
      "",
      materialGovernanceBlock,
      "",
      materialIdentityBlock,
      "",
      techDnaBlock,
      "",
      governanceDefinitionsBlock,
      "===== END GOVERNANCE INPUTS =====",
      "",
      outputFormat1A,
      "",
      REFRESH_ANCHOR_1A,
      gscDataBlock,
      "",
      dataBlock,
      "",
      "Please process the above and return your output now."
    ].join("\n");

    return { fullPrompt: fullPrompt, title: rowData.title, lockedRow: lockedRow };

  } catch (e) {
    throw new Error("bc_getPrompt1A: " + e.message);
  }
}

function bc_getPrompt1B(lockBlock) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lockedRow = ss.getActiveRange().getRow();

    const pSheet = bc_getSheetByConfigKeyOrName_(["prompts"]);
    if (!pSheet) throw new Error("Prompts sheet not found.");

    const module1B = String(
      pSheet.getRange(BC_PROMPT_CELLS.module1).getValue() || ""
    );

    const artOntologyRaw = bc_getCleanValuesByConfigKeyOrName_(["articleTypeControl"]);
    const entOntologyRaw = bc_getCleanValuesByConfigKeyOrName_(["entityTypeControl"]);
    const allowedArticleTypes = bc_flattenFirstColumn(artOntologyRaw).filter(function(v) {
      return String(v).trim() !== "Article Type Determination Rule (Structural Hard Lock)";
    });
    const allowedEntityTypes = bc_flattenFirstColumn(entOntologyRaw);

    const rowData = bc_getRowSpecificData();
    

    const lockedArticleTypeMatch = String(lockBlock || "").match(/CONFIRMED ARTICLE TYPE:\s*(.+)/);
    const lockedArticleType = lockedArticleTypeMatch ? lockedArticleTypeMatch[1].trim() : "";
    Logger.log("lockedArticleType = [" + lockedArticleType + "]");
    let lockedServiceAuthorityLayerRule = "Use weighting table";

    const serviceAuthorityLayerRule =
      String(rowData.serviceAuthorityLayerRule || "Use weighting table").trim();
    const today = bc_getTodayDDMMYYYY();

    const legitMap = bc_getCleanValuesByConfigKeyOrName_(["legitMap"]);
    const stone = (rowData.stoneType || "").toLowerCase().trim();
    const matAliases = {
      "victorian encaustic clay tile": [
        "victorian",
        "historic geometric and encaustic clay tile floors"
      ]
    };
    const terms = [stone].concat(matAliases[stone] || []);
    const legitFiltered = bc_filterByMaterial(legitMap, terms);

    const articleTypeBlock =
      "ALLOWED ARTICLE TYPES (Col 1 — EXACT STRING MATCH REQUIRED)\n" +
      allowedArticleTypes.map(function(v) { return "- " + v; }).join("\n");

    const entityTypeBlock =
      "ALLOWED ENTITY TYPES (Col 5 — EXACT STRING MATCH REQUIRED)\n" +
      allowedEntityTypes.map(function(v) { return "- " + v; }).join("\n");

    const section5_prompt1B =
      "===== SECTION 5 RULES — PROMPT 1B SCOPE (COLS 1–13 ONLY) =====\n" +
      "\n" +
      "CONFIRMED PRIMARY INTENT LOCK — NON-NEGOTIABLE:\n" +
      "The CONFIRMED ARTICLE TYPE and CONFIRMED PRIMARY INTENT values at the top of this\n" +
      "prompt are locked inputs from Prompt 1A. Do NOT redetect, reinterpret,\n" +
      "validate, or override them from HTML, title, slug, meta title, schema,\n" +
      "or any other source. Use them exactly as supplied for all field population.\n" +
      "\n" +
      "ONTOLOGY VALIDATION (COL 1 & 5):\n" +
      "Col 1 MUST exactly match one value from ALLOWED ARTICLE TYPES.\n" +
      "Col 5 MUST exactly match one value from ALLOWED ENTITY TYPES.\n" +
      "No inference. No synonyms. Exact string match required.\n" +
      "\n" +
      "PRIMARY ENTITY FORMAT (Col 2): [Material] [Confirmed Intent] — exactly one space.\n" +
      "Use the CONFIRMED PRIMARY INTENT from the lock block at the top of this prompt.\n" +
      "HUB PAGE EXCEPTION: If Article Type = Hub Page → Primary Entity = [Material] Hub\n" +
      "\n" +
      "LEGITIMACY PREFIX (COL 12): Must begin exactly with LEGIT:Allowed — or LEGIT:Conditional — or LEGIT:Disallowed —\n" +
      "\n" +
      "FLAGS (COL 9): Must end with the EXACT govFlags string from REQUIRED govFlags.\n" +
      "DO NOT recalculate. DO NOT detect. Copy exactly.\n" +
      "\n" +
      "CASE STUDY HARD LOCK — ENTITY TYPE:\n" +
      "Fires ONLY when Article Type = Case Study AND Entity Role = Authority Support Entity (both exact).\n" +
      "In that case ONLY: Entity Type = Material Behaviour.\n" +
      "For ALL other Article Types: Entity Type MUST come from the Role Mapping Table.\n" +
      "\n" +
      "DATE RULE (COL 11):\n" +
      "Today's date is: " + today + "\n" +
      "Use this EXACT date for Col 11.\n" +
      "DO NOT use the page published date or updated date.\n" +
      "\n" +
      "COL 8 (Feeds Hub):\n" +
        "If REQUIRED FEEDS HUB URL = none, output exactly: https://www.abbeyfloorcare.co.uk/" +
        String(rowData.stoneType || "").toLowerCase().trim().replace(/\s+/g, "-") +
        "/hub-page-temp/\n" +
        "Otherwise, output the EXACT URL supplied in REQUIRED FEEDS HUB URL.\n" +
        "Do NOT infer, detect, rewrite, shorten, or reformat it.\n" +
        "Use a RAW URL only.\n" +
        "Do NOT wrap it in markdown.\n" +
        "Do NOT use square brackets.\n" +
        "Do NOT use parentheses.\n" +
        "\n" +
      "COL 6 (Supporting Entities Core):\n" +
      "Populate ONLY with material authority entity names drawn from the Stage 3 TSM Recognised Entities table.\n" +
      "Do NOT include any service authority entities in this column.\n" +
      "\n" +
      "SUPPORTING ENTITIES – SERVICE AUTHORITY LAYER:\n" +
      "Populate ONLY with service authority entity names drawn from Stage 4.\n" +
      "Do NOT include any TSM material authority entities in this column.\n" +
      "SERVICE AUTHORITY LAYER RULE: " + serviceAuthorityLayerRule + "\n" +
      "If the rule above = Use weighting table:\n" +
      "→ use the Stage 4 weighting table normally.\n" +
      "If the rule above contains named entities:\n" +
      "→ restrict the Service Authority Layer to those exact entity names only.\n" +
      "→ ignore all other service authority entities even if the weighting table marks them PRIMARY or INTEGRATED.\n" +
      "→ then apply page relevance inside that restricted pool.\n" +
      "\n" +
      "===== END SECTION =====";

    const outputFormat =
      "===== OUTPUT FORMAT — VERTICAL ONLY (NO TSV) =====\n" +
      "Do NOT output a TSV row.\n" +
      "Output ONLY the vertical labelled block below.\n" +
      "One field per line. No blank lines between fields.\n" +
      "\n" +
      "=== VERTICAL OUTPUT — PASTE INTO PANEL 2 ===\n" +
      "Article Type: [value]\n" +
      "Primary Entity: [value]\n" +
      "Material Entity: [value]\n" +
      "Entity Role: [value]\n" +
      "Entity Type: [value]\n" +
      "Supporting Entities Core: [value]\n" +
      "Supporting Entities - Service Authority Layer: [value]\n" +
      "Peripheral Entities Link Out: [value]\n" +
      "Feeds Hub: [RAW URL only — no brackets, no markdown]\n" +
      "Page Governance Summary: [value]\n" +
      "Entity Governance Status: [value]\n" +
      "Entity Governance Date: [value]\n" +
      "Publish Justification: [value]\n" +
      "GAS Handoff Document: [value]\n" +
      "\n" +
      "After the vertical block output exactly:\n" +
      "Waiting for Input\n" +
      "===== END SECTION =====";

    const preOutputValidation =
      "===== PRE-OUTPUT VALIDATION (MANDATORY) =====\n" +
      "Before outputting, confirm ALL of the following:\n" +
      "1. Article Type matches CONFIRMED ARTICLE TYPE from lock block\n" +
      "2. If Hub Page: Primary Entity = [Material] Hub\n" +
      "3. If not Hub Page: Primary Entity uses CONFIRMED PRIMARY INTENT — NOT redetected\n" +
      "4. Entity Role exactly matches Role Mapping Table for CONFIRMED PRIMARY INTENT + Article Type\n" +
      "5. Entity Type exactly matches Role Mapping Table for CONFIRMED PRIMARY INTENT + Article Type\n" +
      "6. Entity Type exactly matches ALLOWED ENTITY TYPES\n" +
      "7. Supporting Entities Core contains ONLY material authority entity names drawn from the Stage 3 TSM Recognised Entities table\n" +
      "8. Supporting Entities - Service Authority Layer contains ONLY service authority entity names drawn from Stage 4\n" +
      "9. No entity is paraphrased, blended, or placed in the wrong column\n" +
      "10. Col 8 exactly matches REQUIRED FEEDS HUB URL when REQUIRED FEEDS HUB URL is not none. Output a raw URL only — no square brackets, no parentheses, no markdown link formatting\n" +
      "11. Col 9 ends with EXACT govFlags string from REQUIRED govFlags — NOT recalculated\n" +
      "12. Col 10 equals exactly one of: Governed, Compliant, Requires Review, Unclear\n" +
      "13. Col 11 is " + today + " — NOT the page published or updated date\n" +
      "14. Col 12 begins with LEGIT:Allowed — or LEGIT:Conditional — or LEGIT:Disallowed —\n" +
      "15. Col 13 follows format GAS-[MATERIAL]-[INTENT-SHORT]-[3-DIGIT-NUMBER]\n" +
      "If ANY check fails → correct internally before outputting.\n" +
      "===== END VALIDATION =====";

    const REFRESH_ANCHOR_1B = "%%REFRESH_START_P1B%%";
    const dataBlock = bc_buildPrompt1BDataBlock(rowData, today, lockBlock);

    const fullPrompt = [
      "PROMPT 1B OF 4 — ENTITY CLASSIFICATION (COLUMNS 1–13)",
      "RUN MATERIAL: " + rowData.stoneType,
      "TODAY'S DATE: " + today,
      "",
      "IMPORTANT — CONFIRMED PRIMARY INTENT LOCK BLOCK PREPENDED ABOVE THIS LINE BY SIDEBAR.",
      "Use the CONFIRMED ARTICLE TYPE and CONFIRMED PRIMARY INTENT from that block.",
      "Do NOT redetect from HTML, title, slug, or any other source.",
      "Do NOT fire the ambiguity rule. Classification is already confirmed.",
      "",
      articleTypeBlock,
      "",
      entityTypeBlock,
      "",
      section5_prompt1B,
      "",
      "=== MODULE 1B: ENTITY CLASSIFICATION (Columns 1–13) ===",
      module1B,
      "",
      "===== LEGITIMACY MAP (REFERENCE ONLY — for Col 12 only) =====",
      legitFiltered,
      "===== END SECTION =====",
      "",
      outputFormat,
      "",
      preOutputValidation,
      "",
      REFRESH_ANCHOR_1B,
      dataBlock
    ].join("\n");

    return { fullPrompt: fullPrompt, title: rowData.title, lockedRow: lockedRow };

  } catch (e) {
    throw new Error("bc_getPrompt1B: " + e.message);
  }
}

function bc_getPrompt2() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lockedRow = ss.getActiveRange().getRow();

    const pSheet = bc_getSheetByConfigKeyOrName_(["prompts"]);
    if (!pSheet) throw new Error("Prompts sheet not found.");

    const module2 = String(
      pSheet.getRange(BC_PROMPT_CELLS.module2).getValue() || ""
    );
    const module3 = String(
      pSheet.getRange(BC_PROMPT_CELLS.module3).getValue() || ""
    );

    const rowData = bc_getRowSpecificData();
    
    const today = bc_getTodayDDMMYYYY();

    const cols1to13 = bc_readActiveRowFields([
      "Article Type", "Primary Entity", "Material Entity", "Entity Role",
      "Entity Type", "Supporting Entities Core", "Supporting Entities - Service Authority Layer", "Peripheral Entities Link Out",
      "Feeds Hub", "Page Governance Summary", "Entity Governance Status",
      "Entity Governance Date", "Publish Justification", "GAS Handoff Document",
      "Page Rewrite Brief"
    ]);

    // --- GSC Query List extraction (returned to sidebar for column DR write) ---
    const gscQueryList = bc_extractGscQueryList_(rowData.queries, cols1to13["Material Entity"]);
    // --- End GSC Query List ---

    const cols1to13Block =
      "===== GOVERNED FIELDS FROM PROMPT 1B (DO NOT MODIFY) =====\n" +
      Object.entries(cols1to13).map(function(entry) {
        return entry[0] + ": " + entry[1];
      }).join("\n") +
      "\n===== END SECTION =====";

    const cleanedQueries = String(rowData.queries || "")
      .split('\n')
      .filter(function(line) {
        if (!line.trim()) return false;
        if (/^(top queries|clicks|impressions|ctr|position)$/i.test(line.trim())) return false;
        const parts = line.trim().split(/\s{2,}/);
        const impressions = parts.length >= 3 ? parseInt(parts[2]) : 0;
        return impressions >= 1;
      })
      .join('\n');

    const gscBlock =
      "===== GSC PERFORMANCE DATA =====\n" +
      "Page Clicks: " + rowData.clicks + "\n" +
      "Page Impressions: " + rowData.impressions + "\n" +
      "Page Average Position: " + rowData.position + "\n" +
      "Top Queries:\n" +
      (cleanedQueries || "(No valid query lines available)") +
      "\n===== END SECTION =====";

    const legitMap = bc_getCleanValuesByConfigKeyOrName_(["legitMap"]);
    const stone = (rowData.stoneType || "").toLowerCase().trim();
    const matAliases = {
      "victorian encaustic clay tile": [
        "victorian",
        "historic geometric and encaustic clay tile floors"
      ]
    };
    const terms = [stone].concat(matAliases[stone] || []);
    const legitFiltered = bc_filterByMaterial(legitMap, terms);

    // Fetch article HTML for zero-data pages
    const hasGscData = (rowData.clicks > 0 || rowData.impressions > 0);
    let articleContentBlock = "";

    if (!hasGscData) {
      const postIdField = bc_readActiveRowFields(["Post ID"]);
      const postId = postIdField["Post ID"];

      if (postId) {
        const postHtml = bc_getPostHtmlById_(postId);
        if (postHtml) {
          articleContentBlock =
            "\n===== ARTICLE CONTENT (FOR FALLBACK ANALYSIS ONLY) =====\n" +
            postHtml +
            "\n===== END ARTICLE CONTENT =====\n";
        }
      }
    }

    const section5_prompt2 =
      "===== SECTION 5 RULES — PROMPT 2 SCOPE (COLS 14–17 ONLY) =====\n" +
      "\n" +
      "COL 15 (GSC Intent Evidence) — HARD LOCK:\n" +
      "Use EXACTLY this format with ACTUAL numbers from the GSC data:\n" +
      "[X] clicks / [Y] impressions / avg position [Z]\n" +
      "Always use \"clicks\" (even for 0 or 1). Numbers only. No descriptive words.\n" +
      "Use Page Clicks and Page Impressions from the page-level GSC block, not individual query rows.\n" +
      "If query rows differ from page totals, page totals always win.\n" +
      "\n" +
      "COL 16 (Primary Query Cluster Owned):\n" +
      "Must begin with the canonical material name.\n" +
      "Must express the cluster as a READER PROBLEM or READER INTENT STATEMENT.\n" +
      "\n" +
      "COL 16 DERIVATION MODE (CRITICAL):\n" +
      "\n" +
      "STANDARD MODE (Page Clicks > 0 OR Page Impressions > 0):\n" +
      "Derive from actual GSC query data only.\n" +
      "The cluster must be expressed in language a real homeowner would type into Google.\n" +
      "HARD PROHIBITION: Do NOT use analytical vocabulary from the article, governance fields,\n" +
      "or entity names as the cluster. Words like 'suitability', 'doubts', 'governance',\n" +
      "'entity', 'legitimacy', 'diagnostic' must never appear in the PQC.\n" +
      "The PQC must read as a natural search phrase — not an internal classification label.\n" +
      "CORRECT: 'slate floor repair or replacement cost'\n" +
      "WRONG: 'slate kitchen floor suitability doubts'\n" +
      "Output format: [material] + [reader intent statement]\n" +
      "NO suffix symbol.\n" +
      "\n" +
      "FALLBACK MODE (Page Clicks = 0 AND Page Impressions = 0):\n" +
      "Derive Primary Query Cluster Owned by analyzing the article content provided below.\n" +
      "Derivation hierarchy:\n" +
      "1. ARTICLE CONTENT ANALYSIS (primary source)\n" +
      "   - Read the full article HTML\n" +
      "   - Identify the dominant reader problem or question the content addresses\n" +
      "   - Identify what diagnostic issue, material behaviour, or decision the reader is struggling with\n" +
      "   - Frame as what a reader searching for this content would be trying to understand or resolve\n" +
      "2. Page Governance Summary (if article content ambiguous)\n" +
      "3. Primary Entity + Publish Justification (if above unclear)\n" +
      "4. Title H-flag (final fallback)\n" +
      "\n" +
      "Output format: [material] + [reader intent statement] †\n" +
      "\n" +
      "The † symbol flags governance-derived intent (intended ownership)\n" +
      "rather than GSC-observed behaviour.\n" +
      "\n" +
      "Requirements remain unchanged:\n" +
      "- Must begin with canonical material name\n" +
      "- Must express as reader problem or intent statement (not topic label)\n" +
      "- Must contain NO performance metrics\n" +
      "\n" +
      "DRIFT CLARIFICATION RULE:\n" +
      "Drift is determined ONLY by: material mismatch, intent type mismatch, or legitimacy conflict.\n" +
      "If Material matches AND Intent remains Allowed → Drift Status = Aligned.\n" +
      "\n" +
      "===== END SECTION =====";

    const preOutputValidation =
      "===== PRE-OUTPUT VALIDATION (MANDATORY) =====\n" +
      "Before outputting, confirm ALL of the following:\n" +
      "1. Col 15 contains EXACT numeric values for clicks AND impressions AND position\n" +
      "2. Col 15 format is exactly: [X] clicks / [Y] impressions / avg position [Z]\n" +
      "3. Col 16 begins with the canonical material name\n" +
      "4. Col 16 is phrased as a reader problem or intent statement\n" +
      "5. If Page Clicks = 0 AND Page Impressions = 0, Col 16 ends with † symbol\n" +
      "6. If Page Clicks > 0 OR Page Impressions > 0, Col 16 has NO † symbol\n" +
      "7. Col 17 equals exactly one of: Aligned, Minor Drift, Significant Drift\n" +
      "If ANY check fails → correct internally before outputting.\n" +
      "===== END VALIDATION =====";

    const outputFormat =
      "===== OUTPUT FORMAT — VERTICAL ONLY (NO TSV) =====\n" +
      "Do NOT output a TSV row.\n" +
      "Output ONLY the vertical labelled block below.\n" +
      "One field per line. No blank lines between fields.\n" +
      "\n" +
      "=== VERTICAL OUTPUT — PASTE INTO PANEL 2 ===\n" +
      "Observed Query Cluster: [value]\n" +
      "GSC Intent Evidence: [value]\n" +
      "Primary Query Cluster Owned: [value]\n" +
      "Drift Status: [value]\n" +
      "\n" +
      "After the vertical block output exactly:\n" +
      "Waiting for Input\n" +
      "===== END SECTION =====";

    const REFRESH_ANCHOR_2 = "%%REFRESH_START_P2%%";

    const fullPrompt = [
      "PROMPT 2 OF 4 — GSC OBSERVATION + DRIFT STATUS (COLUMNS 14–17)",
      "RUN MATERIAL: " + rowData.stoneType,
      "TODAY'S DATE: " + today,
      "",
      "IMPORTANT — MANUAL EXECUTION MODE",
      "Use ONLY the supplied GSC data. No external data. No speculation.",
      "",
      section5_prompt2,
      "",
      "=== MODULE 2: GSC QUERY OBSERVATION (Columns 14–16) ===",
      module2,
      "",
      "=== MODULE 3: DRIFT STATUS DETERMINATION (Column 17) ===",
      module3,
      "",
      "===== LEGITIMACY MAP (REFERENCE ONLY) =====",
      legitFiltered,
      "===== END SECTION =====",
      "",
      outputFormat,
      "",
      preOutputValidation,
      "",
      REFRESH_ANCHOR_2,
      cols1to13Block,
      "",
      gscBlock,
      articleContentBlock
    ].join("\n");

    return {
      fullPrompt: fullPrompt,
      title: rowData.title,
      recoveryBlueprint: bc_buildRecoveryBlueprint(cols1to13, rowData),
      gscQueryList: gscQueryList,
      lockedRow: lockedRow
    };

  } catch (e) {
    throw new Error("bc_getPrompt2: " + e.message);
  }
}

function bc_getPrompt3() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lockedRow = ss.getActiveRange().getRow();

    const today = bc_getTodayDDMMYYYY();
    const rowData = bc_getRowSpecificData();

    const cols1to17 = bc_readActiveRowFields([
      "Article Type", "Primary Entity", "Material Entity", "Entity Role",
      "Entity Type", "Supporting Entities Core", "Supporting Entities - Service Authority Layer", "Peripheral Entities Link Out",
      "Feeds Hub", "Page Governance Summary", "Entity Governance Status",
      "Entity Governance Date", "Publish Justification", "GAS Handoff Document",
      "Observed Query Cluster", "GSC Intent Evidence",
      "Primary Query Cluster Owned", "Drift Status"
    ]);

    const secondaryFields = bc_readActiveRowFields(["Secondary Intent Decisions"]);
    const secondaryBlock = String(secondaryFields["Secondary Intent Decisions"] || "").trim();

    let cols1to17Block =
      "===== GOVERNED FIELDS (DO NOT MODIFY) =====\n" +
      Object.entries(cols1to17).map(function(entry) {
        return entry[0] + ": " + entry[1];
      }).join("\n") +
      "\n===== END SECTION =====";

    if (secondaryBlock) {
      cols1to17Block +=
        "\n\n===== CONFIRMED PRIMARY INTENT LOCK =====\n" +
        secondaryBlock +
        "\n===== END LOCK =====";
    }

    // Article HTML for content-grounding audit
    const postIdField = bc_readActiveRowFields(["Post ID"]);
    const postId = postIdField["Post ID"];
    let articleContentBlock = "";
    if (postId) {
      const postHtml = bc_getPostHtmlById_(postId);
      if (postHtml) {
        articleContentBlock =
          "\n===== ARTICLE CONTENT (FOR CONTENT-GROUNDING AUDIT) =====\n" +
          postHtml +
          "\n===== END ARTICLE CONTENT =====\n";
      }
    }

    const section5 =
      "===== PAGE REWRITE BRIEF ONLY — SCOPE =====\n" +
      "\n" +
      "TODAY'S DATE IS: " + today + "\n" +
      "Use this EXACT date for Rewrite Brief Date. DO NOT use any date from page metadata.\n" +
      "\n" +
      "SECONDARY INTENT DECISIONS — PAGE REWRITE BRIEF RULE:\n" +
      "If a CONFIRMED PRIMARY INTENT LOCK section with SECONDARY INTENT DECISIONS is present above,\n" +
      "each decision follows: [number]: [Y or N] — [intent name] — [reason]\n" +
      "For each intent marked Y: include an instruction to retain it as a supporting layer, with\n" +
      "the supplied reason appearing verbatim.\n" +
      "For each intent marked N: include an instruction to remove that layer entirely, with the\n" +
      "supplied reason appearing verbatim, and where relevant instruct an outbound link to a\n" +
      "dedicated page covering that intent instead.\n" +
      "If no such block is present, derive rewrite instructions from governed intent, material\n" +
      "entity, and entity coverage only.\n" +
      "Do NOT use Drift Status to determine editorial tone or rewrite instruction.\n" +
      "\n" +
      "PAGE REWRITE BRIEF — MANDATORY CONTENT-GROUNDING AUDIT (HARD LOCK):\n" +
      "Before writing the Page Rewrite Brief, you MUST read the actual article HTML supplied\n" +
      "above — not just the governed entity list.\n" +
      "For every material, product, technique, or intervention you intend to name in the\n" +
      "Rewrite Brief as a next step, escalation path, or alternative approach, you must first\n" +
      "check what THIS ARTICLE actually says about it — not what the entity list or your\n" +
      "general material knowledge assumes.\n" +
      "\n" +
      "Specifically check for and preserve, verbatim in effect:\n" +
      "— Whether the article restricts a material/technique to professionals only (e.g. due to\n" +
      "  cost, tools, smell, skill level, or time required) versus presenting it as a homeowner\n" +
      "  DIY option.\n" +
      "— Whether any section explicitly recommends AGAINST a method, product, or approach for\n" +
      "  homeowners, even if that same method/product is a governed entity for this page.\n" +
      "— Whether any section states a scope limit, comparison, or caveat that narrows how an\n" +
      "  entity should be framed (e.g. 'X is cheaper and easier, Y is for professionals').\n" +
      "\n" +
      "The Rewrite Brief must NEVER phrase two distinct options ambiguously as one combined\n" +
      "step (e.g. 'material X and professional intervention apply when Y' is AMBIGUOUS unless\n" +
      "the article confirms X and professional work are the same recommendation). If the article\n" +
      "content contradicts or narrows what the general entity list implies, the article's actual\n" +
      "stated position always wins — state it explicitly and unambiguously in the brief.\n" +
      "\n" +
      "This audit applies to every material and article type — do not skip it because an entity\n" +
      "seems self-evident from its name.\n" +
      "\n" +
      "PAGE REWRITE BRIEF — REQUIRED INTERNAL STRUCTURE (HARD LOCK):\n" +
      "Format the Page Rewrite Brief value using these exact labelled sub-sections, separated\n" +
      "by ' | ' on a single field line (downstream stages parse these labels specifically).\n" +
      "Every label below MUST appear even if its content is 'Not applicable' — omitting a label\n" +
      "entirely, for any reason, is a validation failure:\n" +
      "RE-ANCHOR FOCUS: [entities/topics this rewrite must strengthen]\n" +
      "SCOPE BOUNDARY: [what this page must not drift into, consistent with the Cannibalisation Guardrail]\n" +
      "ESCALATION BOUNDARY: [the exact point/action where DIY or basic intervention stops and\n" +
      "  professional work begins, stated as a single unambiguous fact grounded in the content-\n" +
      "  grounding audit above — write 'Not applicable' only if the article genuinely contains no\n" +
      "  such boundary, never omit the label]\n" +
      "QUERY PRECISION ACTION: Precision Score not calculated in this standalone run — if a score\n" +
      "  exists from a prior Combined Prompt run, note it here; otherwise write 'Not calculated in\n" +
      "  this run.'\n" +
      "Example: 'RE-ANCHOR FOCUS: Natural Void Structure, Factory Filler Weakness | SCOPE BOUNDARY:\n" +
      "localised repair only, no restoration or sealing | ESCALATION BOUNDARY: resin filler repair\n" +
      "requires professional application due to specialist tools and skill level — homeowners should\n" +
      "use sanded grout only | QUERY PRECISION ACTION: Not calculated in this run.'\n" +
      "\n" +
      "===== END SECTION =====";

    const outputFormat =
      "===== OUTPUT FORMAT — VERTICAL ONLY, TWO FIELDS ONLY =====\n" +
      "Output ONLY the two labelled lines below. Nothing else.\n" +
      "\n" +
      "Page Rewrite Brief: [value]\n" +
      "Rewrite Brief Date: [value]\n" +
      "\n" +
      "After the two lines output exactly:\n" +
      "Waiting for Input\n" +
      "===== END SECTION =====";

    const preOutputValidation =
      "===== PRE-OUTPUT VALIDATION (MANDATORY) =====\n" +
      "Before outputting, confirm ALL of the following:\n" +
      "1. Exactly 2 fields produced — nothing else.\n" +
      "2. Page Rewrite Brief contains all four required labels (RE-ANCHOR FOCUS, SCOPE BOUNDARY,\n" +
      "   ESCALATION BOUNDARY, QUERY PRECISION ACTION), each populated or explicitly 'Not applicable'.\n" +
      "3. Rewrite Brief Date is " + today + " — NOT any date from page metadata.\n" +
      "4. ESCALATION BOUNDARY content is grounded in the article HTML, not assumed from the entity name.\n" +
      "5. No ambiguous phrasing combining two distinct next-steps into one option.\n" +
      "If ANY check fails → correct internally before outputting.\n" +
      "===== END VALIDATION =====";

    const fullPrompt = [
      "PAGE REWRITE BRIEF UPDATE — SINGLE FIELD ONLY",
      "RUN MATERIAL: " + rowData.stoneType,
      "TODAY'S DATE: " + today,
      "",
      "IMPORTANT — MANUAL EXECUTION MODE",
      "Derive the rewrite brief strictly from the governed fields and article content supplied below.",
      "Do not introduce new entities, services, or claims not present in the governed fields or article.",
      "",
      section5,
      "",
      outputFormat,
      "",
      preOutputValidation,
      "",
      cols1to17Block,
      articleContentBlock
    ].join("\n");

    return { fullPrompt: fullPrompt, title: rowData.title, lockedRow: lockedRow };

  } catch (e) {
    throw new Error("bc_getPrompt3: " + e.message);
  }
}

function getGovernanceDefinitionsBlock() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Governance Definitions");
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const rows = data.slice(1);

  let block = "===== GOVERNANCE DEFINITIONS (BOUNDARY CONTROL) =====\n";

  rows.forEach(row => {
    const term = row[0];
    const definition = row[2];
    const boundary = row[3];

    if (term && definition && boundary) {
      block += `${term}:\n`;
      block += `Definition: ${definition}\n`;
      block += `Boundary: ${boundary}\n\n`;
    }
  });

  block += "===== END GOVERNANCE DEFINITIONS =====\n";

  return block;
}

function bc_getGovernanceDefinitionsBlock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Governance Definitions");

  if (!sheet) {
    return "(Governance Definitions sheet not found)";
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return "(Governance Definitions sheet is empty)";
  }

  const rows = data.slice(1);

  const lines = [];
  lines.push("===== GOVERNANCE DEFINITIONS (BOUNDARY CONTROL) =====");
  lines.push("Use this block as a hard intent-boundary enforcement layer.");
  lines.push("If structural interpretation conflicts with a Boundary definition, the Boundary overrides structural inference.");
  lines.push("Do NOT escalate an intent beyond its defined Boundary.");
  lines.push("");

  rows.forEach(function(row) {
    const term = String(row[0] || "").trim();
    const category = String(row[1] || "").trim();
    const definition = String(row[2] || "").trim();
    const boundary = String(row[3] || "").trim();

    if (!term || !definition || !boundary) return;

    if (category === "Intent Definition") {
      lines.push(term);
      lines.push("Definition: " + definition);
      lines.push("Boundary: " + boundary);
      lines.push("");
    }
  });

  lines.push("===== END GOVERNANCE DEFINITIONS =====");

  return lines.join("\n");
}

function bc_getPostHtmlById_(postId) {
  try {
    if (!postId) return "";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const siteExportSheet = ss.getSheetByName("site-export");

    if (!siteExportSheet) {
      Logger.log("site-export sheet not found");
      return "";
    }

    const data = siteExportSheet.getDataRange().getValues();
    const headers = data[0];

    const idColIndex = headers.indexOf("ID");
    const htmlColIndex = headers.indexOf("Full Post HTML");

    if (idColIndex === -1 || htmlColIndex === -1) {
      Logger.log("Required columns not found. ID: " + idColIndex + ", Full Post HTML: " + htmlColIndex);
      return "";
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(postId)) {
        return String(data[i][htmlColIndex] || "");
      }
    }

    Logger.log("Post ID " + postId + " not found in site-export");
    return "";

  } catch (e) {
    Logger.log("bc_getPostHtmlById_ error: " + e.message);
    return "";
  }
}


function bc_extractGscQueryList_(rawQueries, materialEntity) {
  if (!rawQueries) return "";

  const material = String(materialEntity || "").toLowerCase().trim();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Load material aliases from Material Governance Reference Sheet
  const matGovSheet = ss.getSheetByName("Material Governance Reference Sheet");
  const materialAliases = {};
  if (matGovSheet) {
    const matGovData = matGovSheet.getDataRange().getValues();
    for (let i = 1; i < matGovData.length; i++) {
      const canonical = String(matGovData[i][0] || "").toLowerCase().trim();
      const variants = String(matGovData[i][2] || "").toLowerCase().trim();
      if (canonical) {
        const terms = [canonical].concat(
          variants.split(';').map(function(v) { return v.trim(); }).filter(Boolean)
        );
        materialAliases[canonical] = terms;
      }
    }
  }
  const materialTerms = materialAliases[material] || [material.split(' ')[0]];

  // Load UK towns into a Set for geo filtering
  const townSheet = ss.getSheetByName("uk-towns");
  const townSet = new Set();
  if (townSheet) {
    const townData = townSheet.getRange(2, 1, townSheet.getLastRow() - 1, 1).getValues();
    townData.forEach(function(row) {
      if (row[0]) townSet.add(String(row[0]).toLowerCase().trim());
    });
  }

  // Load English indicator words from uk-words sheet
  const wordsSheet = ss.getSheetByName("uk-words");
  const englishWordSet = new Set();
  if (wordsSheet) {
    const wordsData = wordsSheet.getRange(2, 1, wordsSheet.getLastRow() - 1, 1).getValues();
    wordsData.forEach(function(row) {
      if (row[0]) englishWordSet.add(String(row[0]).toLowerCase().trim());
    });
  }

  const excludeBrand = /\b(abbey|tile doctor)\b/i;
  const excludeGeo = /\b(near me|near you|in my area|local|locally)\b/i;
  const excludeWrongMaterial = /\b(carpet tile|vinyl|laminate|porcelain|terracotta|wood floor|wooden floor|hardwood)\b/i;
  const excludeUnrelated = /\b(fireplace|countertop|worktop|dry lin|lift install|youtube|definition|cleaning job|surface repair|supplier|suppliers|merchant|merchants|showroom)\b/i;
  const excludeOperators = /inurl:|site:|"\w+.*"\s+\w+/i;

  const lines = String(rawQueries).split('\n');
  const parsed = [];

  lines.forEach(function(line) {
    if (!line.trim()) return;
    const parts = line.split(/\t/);
    const partsFallback = parts.length < 3 ? line.split(/\s{2,}/) : parts;
    const finalParts = partsFallback.length >= 3 ? partsFallback : parts;
    if (finalParts.length < 3) return;

    const query = finalParts[0].trim().toLowerCase();
    const clicks = parseInt(finalParts[1]) || 0;
    const impressions = parseInt(finalParts[2]) || 0;

    if (!query) return;
    //if (queryWords.length < 3) return;
    if (/^(query|top queries)$/i.test(query)) return;

    if (excludeBrand.test(query)) return;
    if (excludeGeo.test(query)) return;
    if (excludeUnrelated.test(query)) return;
    if (excludeOperators.test(query)) return;

    // Exclude queries with no recognised English words
    const queryWords = query.split(/\s+/);
    if (queryWords.length < 3) return;
    if (!queryWords.some(function(word) { return englishWordSet.has(word); })) return;

    // Exclude containing any UK town name
    const careTerms = /\b(clean|cleaning|restore|restoration|seal|sealing|polish|polishing|repair|tile|floor|stone|grout|hone|honing|strip|stripping|wax|waxing|care|maintenance)\b/i;
    const hasTown = queryWords.some(function(word) { return word.length >= 4 && townSet.has(word); });
    if (hasTown && !careTerms.test(query)) return;

    if (excludeWrongMaterial.test(query)) {
      const queryMentionsMaterial = materialTerms.some(function(term) { return query.indexOf(term) !== -1; });
      if (!queryMentionsMaterial) return;
    }

    if (clicks === 0 && impressions < 3) return;

    parsed.push({ query: query, clicks: clicks, impressions: impressions });
  });

  if (parsed.length === 0) return "";

  parsed.sort(function(a, b) {
    if (b.clicks !== a.clicks) return b.clicks - a.clicks;
    return b.impressions - a.impressions;
  });

  const kept = parsed.slice(0, 20).map(function(item) { return item.query; });

  return kept.join(", ");
}
function openPromptSplitter() {
  var html = HtmlService.createTemplateFromFile('PromptSplitter')
    .evaluate()
    .setTitle('Prompt Splitter');
  SpreadsheetApp.getUi().showSidebar(html);
}

function bc_sendPromptViaOpenAI(promptText, maxTokens, modelOverride) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY not set in Script Properties.');
  var model = modelOverride || MODEL_MID;
  var payload = {
    model: model,
    messages: [{ role: 'user', content: promptText }]
  };
  if (maxTokens) payload.max_completion_tokens = maxTokens;

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  var data = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    return { success: false, message: (data.error && data.error.message) || 'OpenAI API error' };
  }

  var promptTokens     = data.usage.prompt_tokens;
  var completionTokens = data.usage.completion_tokens;
  var PRICING = {
  'gpt-5.6-luna':  { input: 0.20, output: 1.20 },
  'gpt-5.6-terra': { input: 2.00, output: 12.00 },
  'gpt-5.6-sol':   { input: 4.00, output: 20.00 },
  'gpt-4.1-nano':  { input: 0.10, output: 0.40 },
  'gpt-4.1':       { input: 2.00, output: 8.00 }
};

var rates = PRICING[model] || PRICING[MODEL_MID];
  var cost = (promptTokens / 1000000 * rates.input) + (completionTokens / 1000000 * rates.output);
  var finishReason = data.choices[0].finish_reason;
  if (finishReason === 'length') {
    Logger.log('WARNING: OpenAI response truncated by max_tokens limit.');
  }

  return {
    success: true,
    text: data.choices[0].message.content,
    promptTokens: promptTokens,
    completionTokens: completionTokens,
    cost: cost,
    truncated: finishReason === 'length'
  };
}

function bc_getPromptCombined() {
  try {
    const pSheet = bc_getSheetByConfigKeyOrName_(["prompts"]);
    if (!pSheet) throw new Error("Prompts sheet not found.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const lockedRow = ss.getActiveRange().getRow();
    const row = lockedRow;
    if (row >= 2) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const postIdIdx = headers.indexOf("Post ID");
      const postRefIdx = headers.indexOf("Post Ref");
      if (postIdIdx !== -1 && postRefIdx !== -1) {
        const postId = sheet.getRange(row, postIdIdx + 1).getValue();
        sheet.getRange(row, postRefIdx + 1).setValue(postId);
      }
    }

    const rowData = bc_getRowSpecificData();
    const today = bc_getTodayDDMMYYYY();

    // Recovery Blueprint depends on classification fields (Article Type, Primary Entity, etc.)
    // which don't exist yet on a new row until Write to Row completes — so it's calculated
    // AFTER Write to Row instead, in bc_writeCombinedOutputToRow(), not here at Load time.
    const gscQueryList = bc_extractGscQueryList_(rowData.queries, rowData.materialEntity);

    const queryVectorPreview = bc_buildQueryVectorPreview_(gscQueryList, rowData.fullHtml);

    const queryVectorBlock =
      "===== QUERY VECTOR ANALYSIS (PRE-CALCULATED) =====\n" +
      "Precision Score: " + queryVectorPreview.precisionScore + "\n" +
      "(This is the percentage of meaningful terms from real GSC search queries that already appear in this article's HTML.)\n" +
      "Terms from real search queries ALREADY PRESENT in the article: " + (queryVectorPreview.presentTerms || []).join(", ") + "\n" +
      "Terms from real search queries MISSING from the article: " + (queryVectorPreview.missingTerms || []).join(", ") + "\n" +
      "===== END QUERY VECTOR ANALYSIS =====";

    // ---- P1A-equivalent blocks ----
    const module1A = String(pSheet.getRange("D16").getValue() || "");
    const artOntologyRaw = bc_getCleanValuesByConfigKeyOrName_(["articleTypeControl"]);
    const allowedArticleTypes = bc_flattenFirstColumn(artOntologyRaw).filter(function(v) {
      return String(v).trim() !== "Article Type Determination Rule (Structural Hard Lock)";
    });
    const stone = (rowData.stoneType || "").toLowerCase().trim();
    const matAliases = {
      "victorian encaustic clay tile": ["victorian", "historic geometric and encaustic clay tile floors"]
    };
    const terms = [stone].concat(matAliases[stone] || []);
    const taxonomyFiltered = bc_getCachedMaterialBlock_('taxonomyFiltered', rowData.stoneType, function() {
      const taxonomy = bc_getCleanValuesByConfigKeyOrName_(["intentTax"]);
      return bc_filterByMaterial(taxonomy, terms);
    });

    const articleTypeBlock =
      "ALLOWED ARTICLE TYPES (Col 1 — EXACT STRING MATCH REQUIRED)\n" +
      allowedArticleTypes.map(function(v) { return "- " + v; }).join("\n");

    const articleTypeClassificationBlock = bc_buildArticleTypeClassificationBlock();
    const articleTypeDataBlock = bc_buildArticleTypeDataBlock(rowData);
    const tierStructuralBlock = bc_getTierStructuralCoverageBlock(rowData.articleType);
    const entityRoleControlBlock = bc_getEntityRoleControlBlock();
    const roleMappingBlock1A = bc_buildRoleMappingBlock(rowData.articleType);
    const tsmBlock1A = bc_getCachedMaterialBlock_('tsmBlock1A', rowData.stoneType, function() {
      return bc_getTSMEntitiesBlock(rowData.stoneType, "CONFIRMED PRIMARY INTENT: [AWAITING_LLM_INTENT_DETECTION]");
    });
    const surfaceIssueBlock = bc_getCachedMaterialBlock_('surfaceIssueBlock', rowData.stoneType, function() {
      return bc_getSurfaceIssueMasterBlock(rowData.stoneType);
    });
    const materialGovernanceBlock = bc_getCachedMaterialBlock_('materialGovernanceBlock', rowData.stoneType, function() {
      return bc_getMaterialGovernanceBlock(rowData.stoneType);
    });
    const materialIdentityBlock = bc_getMaterialIdentityBlock(
      rowData.stoneType, rowData.searchEntity || rowData.title || rowData.materialEntity || ""
    );
    const techDnaBlock = bc_getCachedMaterialBlock_('techDnaBlock', rowData.stoneType, function() {
      return bc_getTechDNABlock(rowData.stoneType);
    });
    const governanceDefinitionsBlock = bc_getGovernanceDefinitionsBlock();

    const taxonomySheet = ss.getSheetByName('1 - Intent Taxonomy Table');
    if (!taxonomySheet) throw new Error("Intent Taxonomy Table sheet not found.");
    const taxonomyData = taxonomySheet.getDataRange().getValues();
    const tHeaders = taxonomyData[0];
    const materialColIndex = tHeaders.indexOf('Material');
    const intentColIndex = tHeaders.indexOf('Intent');
    if (materialColIndex === -1 || intentColIndex === -1) {
      throw new Error('Intent Taxonomy Table must have "Material" and "Intent" columns');
    }
    const intents = [];
    for (let i = 1; i < taxonomyData.length; i++) {
      if (taxonomyData[i][materialColIndex] === rowData.stoneType && taxonomyData[i][intentColIndex]) {
        const intent = taxonomyData[i][intentColIndex].trim();
        if (intent !== 'Hub Page' && !intents.includes(intent)) intents.push(intent);
      }
    }
    const allowedIntentsBlock =
      "ALLOWED INTENTS (select exactly one):\n" +
      intents.map(function(intent) { return "- " + intent; }).join("\n");

    const hasGSCData =
      String(rowData.queries || "").trim().length > 0 ||
      Number(String(rowData.impressions || "0").replace(/,/g, "")) > 0 ||
      Number(String(rowData.clicks || "0").replace(/,/g, "")) > 0;

    const gscDataBlock =
      "===== GSC PERFORMANCE DATA (FOR INTENT CLASSIFICATION) =====\n" +
      "Page Clicks: " + rowData.clicks + "\n" +
      "Page Impressions: " + rowData.impressions + "\n" +
      "Monitor Queries:\n" +
      (String(rowData.queries || "").trim().length > 0
        ? rowData.queries.split('\n').slice(0, 30).join('\n')
        : hasGSCData ? "(Clicks / impressions present, but no query text available)" : "(No GSC data available)") +
      "\n===== END GSC DATA =====";

    // ---- P1B-equivalent blocks (built after Article Type/Intent are known — deferred, see note below) ----
    const entOntologyRaw = bc_getCleanValuesByConfigKeyOrName_(["entityTypeControl"]);
    const allowedEntityTypes = bc_flattenFirstColumn(entOntologyRaw);
    const entityTypeBlock =
      "ALLOWED ENTITY TYPES (Col 5 — EXACT STRING MATCH REQUIRED)\n" +
      allowedEntityTypes.map(function(v) { return "- " + v; }).join("\n");

    const serviceAuthorityLayerRule = String(rowData.serviceAuthorityLayerRule || "Use weighting table").trim();
    const legitFiltered = bc_getCachedMaterialBlock_('legitFiltered', rowData.stoneType, function() {
      const legitMap = bc_getCleanValuesByConfigKeyOrName_(["legitMap"]);
      return bc_filterByMaterial(legitMap, terms);
    });

    const feedsHubUrl = bc_getFeedsHubUrlForStoneType_(rowData.stoneType);
    const tsmEntitiesBlock1B = bc_getCachedMaterialBlock_('tsmEntitiesBlock1B', rowData.stoneType, function() {
      return bc_getTSMEntitiesBlock(rowData.stoneType, "CONFIRMED PRIMARY INTENT: [TO BE DETECTED IN BLOCK 2 BELOW]");
    });
    const serviceAuthorityBlock = bc_getCachedMaterialBlock_('serviceAuthorityBlock_' + String(rowData.articleType || '').replace(/\s+/g,'_'), rowData.stoneType, function() {
      return bc_getServiceAuthorityBlock(rowData.stoneType, rowData.articleType, rowData.location);
    });
    const roleMappingBlock1B = bc_buildRoleMappingBlock(rowData.articleType);

    // ---- P2-equivalent blocks ----
    const cleanedQueries = String(rowData.queries || "")
      .split('\n')
      .filter(function(line) {
        if (!line.trim()) return false;
        if (/^(top queries|clicks|impressions|ctr|position)$/i.test(line.trim())) return false;
        const parts = line.trim().split(/\s{2,}/);
        const impressions = parts.length >= 3 ? parseInt(parts[2]) : 0;
        return impressions >= 1;
      }).join('\n');

    const gscBlockP2 =
      "===== GSC PERFORMANCE DATA (FOR COLS 14-17) =====\n" +
      "Page Clicks: " + rowData.clicks + "\n" +
      "Page Impressions: " + rowData.impressions + "\n" +
      "Page Average Position: " + rowData.position + "\n" +
      "Top Queries:\n" +
      (cleanedQueries || "(No valid query lines available)") +
      "\n===== END SECTION =====";

    const hasGscDataP2 = (rowData.clicks > 0 || rowData.impressions > 0);
    let articleContentBlock = "";
    if (!hasGscDataP2) {
      const postIdField = bc_readActiveRowFields(["Post ID"]);
      const postId = postIdField["Post ID"];
      if (postId) {
        const postHtml = bc_getPostHtmlById_(postId);
        if (postHtml) {
          articleContentBlock =
            "\n===== ARTICLE CONTENT (FOR FALLBACK ANALYSIS ONLY) =====\n" +
            postHtml + "\n===== END ARTICLE CONTENT =====\n";
        }
      }
    }

    const outputFormatCombined =
      "===== OUTPUT FORMAT — COMBINED VERTICAL OUTPUT (NO TSV) =====\n" +
      "Output ONLY the vertical labelled block below. One field per line. No blank lines between fields.\n" +
      "Detect Article Type and Primary Intent first, then use those values consistently for every field below — do not redetect partway through.\n" +
      "\n" +
      "=== COMBINED VERTICAL OUTPUT ===\n" +
      "Article Type: [value]\n" +
      "Confirmed Primary Intent: [value]\n" +
      "Secondary Intent Decisions: [all decisions on this ONE line only, separated by semicolons, format: N: Y or N — intent name — reason; N: Y or N — intent name — reason. MUST exclude the Confirmed Primary Intent — list only intents other than the primary one.]\n" +
      "Primary Entity: [value]\n" +
      "Material Entity: [value]\n" +
      "Entity Role: [value]\n" +
      "Entity Type: [value]\n" +
      "Supporting Entities Core: [value]\n" +
      "Supporting Entities - Service Authority Layer: [value]\n" +
      "Peripheral Entities Link Out: [value]\n" +
      "Feeds Hub: [RAW URL only — no brackets, no markdown]\n" +
      "Page Governance Summary: [value]\n" +
      "Entity Governance Status: [value]\n" +
      "Entity Governance Date: [value]\n" +
      "Publish Justification: [value]\n" +
      "GAS Handoff Document: [value]\n" +
      "Observed Query Cluster: [value]\n" +
      "GSC Intent Evidence: [value]\n" +
      "Primary Query Cluster Owned: [value]\n" +
      "Drift Status: [value]\n" +
      "Rewrite Status: [value]\n" +
      "Rewrite Governance Summary: [value]\n" +
      "Cannibalisation Guardrail: [value]\n" +
      "Safe Handoff Pages: [value]\n" +
      "Rewrite Role Lock: [value]\n" +
      "Page Rewrite Brief: [value]\n" +
      "Rewrite Brief Date: [value]\n" +
      "\n" +
      "After the vertical block output exactly:\n" +
      "Waiting for Input\n" +
      "===== END SECTION =====";

    const preOutputValidationCombined =
      "===== PRE-OUTPUT VALIDATION (MANDATORY) =====\n" +
      "1. Article Type matches one of ALLOWED ARTICLE TYPES exactly.\n" +
      "2. Primary Entity, Entity Role, Entity Type match Role Mapping Table for the detected Article Type + Primary Intent.\n" +
      "3. Entity Type matches ALLOWED ENTITY TYPES exactly.\n" +
      "4. Feeds Hub is a raw URL or matches REQUIRED FEEDS HUB URL — no brackets, no markdown.\n" +
      "5. GAS Handoff Document follows GAS-[MATERIAL]-[INTENT-SHORT]-[3-DIGIT-NUMBER].\n" +
      "6. GSC Intent Evidence format is exactly: [X] clicks / [Y] impressions / avg position [Z].\n" +
      "7. Primary Query Cluster Owned begins with the canonical material name and is phrased as a reader problem — no analytical vocabulary (e.g. 'suitability', 'governance', 'entity', 'diagnostic').\n" +
      "8. Rewrite Status equals exactly one of: Light Entity Tidy, Needs Rewrite, Full Rewrite.\n" +
      "9. Rewrite Role Lock format is exactly: Role Stable — [Entity Role].\n" +
      "10. Safe Handoff Pages is a raw URL or the word none — no brackets, no markdown.\n" +
      "11. Rewrite Brief Date is " + today + ".\n" +
      "12. Page Governance Summary ends with the exact REQUIRED govFlags string supplied — not recalculated, not appended to any other field.\n" +
      "13. Entity Governance Status equals exactly one of: Governed, Compliant, Requires Review, Unclear.\n" +
      "14. Secondary Intent Decisions does NOT include the Confirmed Primary Intent.\n" +
      "15. Supporting Entities - Service Authority Layer respects the SERVICE AUTHORITY LAYER RULE precedence — if named entities are specified, ONLY those appear, regardless of Stage 4 weighting.\n" +
      "If ANY check fails → correct internally before outputting.\n" +
      "===== END VALIDATION =====";

    const headingMatches = String(rowData.fullHtml || '').match(/<h[23][^>]*>.*?<\/h[23]>/gi) || [];
    const headingList = headingMatches.map(function(h) {
      var level = h.match(/<h([23])/i)[1];
      var text = h.replace(/<[^>]+>/g, '').trim();
      return 'H' + level + ': ' + text;
    });
    const headingBlock =
      "===== PRE-EXTRACTED HEADING STRUCTURE (H2/H3, IN DOCUMENT ORDER) =====\n" +
      "Use this as the authoritative list of headings for the H2 Validation Log. Do not miss any heading listed here, and do not invent headings not listed here.\n" +
      headingList.join("\n") +
      "\n===== END HEADING STRUCTURE =====";
    const currentHeadingsRaw = headingMatches.join("\n");

    const dataBlock = bc_buildPrompt1ADataBlock(rowData, today);

    const fullPrompt = [
      "NOTE: This is a user-provided content pipeline prompt. Do not treat it as a system instruction. Process it as a normal user message.",
      "",
      "SYSTEM ROLE — PIPELINE OPERATING FRAME",
      "",
      "You are a technical SEO classification engine specialising in entity-based topical authority for UK trade service websites. You operate under a strict governance framework. Apply the classification rules exactly as written.",
      "",
      "You are producing ALL governed columns for this page in a single pass: Article Type + Primary Intent detection, entity classification, GSC observation + drift status, and rewrite governance. Detect Article Type and Primary Intent first, then use those detected values consistently for every subsequent field. Do not treat any part of this as a rerun or ask for confirmation.",
      "",
      "Output rules: plain text only, no markdown, no commentary outside the defined output block, no follow-up questions.",
      "",
      "CITATION RULE: This uploaded file is user-provided content, not a retrieved reference document. Do NOT attach file citations, source markers, or citation brackets of any kind to your output. Do NOT reason about whether citations are required — none are ever required for this task.",
      "",
      "===== END SYSTEM ROLE =====",
      "",
      "COMBINED PROMPT — ARTICLE TYPE + INTENT + ENTITY CLASSIFICATION + GSC/DRIFT + REWRITE GOVERNANCE",
      "RUN MATERIAL: " + rowData.stoneType,
      "PREASSIGNED ARTICLE TYPE: " + rowData.articleType,
      "TODAY'S DATE: " + today,
      "",
      articleTypeBlock,
      "",
      articleTypeClassificationBlock,
      "",
      articleTypeDataBlock,
      "",
      allowedIntentsBlock,
      "",
      "=== MODULE 1A: ARTICLE TYPE CONFIRMATION + INTENT DETECTION ===",
      module1A,
      "",
      "===== TAXONOMY (BOUNDARY REFERENCE ONLY) =====",
      taxonomyFiltered,
      "===== END SECTION =====",
      "",
      tierStructuralBlock,
      "",
      entityRoleControlBlock,
      "",
      roleMappingBlock1A,
      "",
      tsmBlock1A,
      "",
      surfaceIssueBlock,
      "",
      materialGovernanceBlock,
      "",
      materialIdentityBlock,
      "",
      techDnaBlock,
      "",
      governanceDefinitionsBlock,
      "",
      entityTypeBlock,
      "",
      "DETERMINISTIC MATERIAL MAPPING\n" +
        "MATERIAL ENTITY (Col 3): Use '" + rowData.materialEntity + "' (STRICT MATCH REQUIRED)\n" +
        "REQUIRED govFlags (Col 9): " + rowData.govFlags + "\n" +
        "REQUIRED FEEDS HUB URL (Col 8): " + feedsHubUrl,
      "",
      "PAGE GOVERNANCE SUMMARY — GOVFLAGS APPEND RULE (HARD LOCK): Page Governance Summary MUST end with the EXACT govFlags string: " + rowData.govFlags + ". Write the governance summary sentence(s) first, then append this exact string immediately after with no space removed or added — e.g. '...within travertine care lifecycle." + rowData.govFlags + "'. DO NOT recalculate or detect this string. DO NOT append it to Publish Justification or any other field.",
      "",
      tsmEntitiesBlock1B,
      "",
      serviceAuthorityBlock,
      "SERVICE AUTHORITY LAYER RULE: " + serviceAuthorityLayerRule,
      "",
      "SERVICE AUTHORITY LAYER RULE — PRECEDENCE (HARD LOCK):",
      "If SERVICE AUTHORITY LAYER RULE above equals 'Use weighting table' → use the Stage 4 weighting table normally (PRIMARY/INTEGRATED/VALIDATING as defined).",
      "If SERVICE AUTHORITY LAYER RULE above contains named entities → restrict Supporting Entities - Service Authority Layer to ONLY those exact named entities.",
      "In that case, IGNORE all other Stage 4 service authority entities even if the weighting table marks them PRIMARY or INTEGRATED for this article type — the named restriction always overrides the weighting table.",
      "",
      "===== LEGITIMACY MAP (REFERENCE ONLY) =====",
      legitFiltered,
      "===== END SECTION =====",
      "",
      "===== GSC PERFORMANCE DATA (COLS 14-17) =====",
      gscBlockP2,
      articleContentBlock,
      "",
      queryVectorBlock,
      "",
      "===== COLUMN-SPECIFIC HARD LOCK RULES =====",
      "",
      "PRIMARY ENTITY FORMAT (Col 2): [Material] [Confirmed Primary Intent] — exactly one space.",
      "MULTI-WORD OR SLASHED INTENT NAMES: If the Confirmed Primary Intent contains a slash or multiple concept words (e.g. 'Repair / Chip / Crack', 'Sealing / Resealing', 'Honing / Polishing / Finish Recovery'), use ONLY the first word before the slash as the intent portion of Primary Entity. Example: Confirmed Primary Intent 'Repair / Chip / Crack' → Primary Entity 'Travertine Repair'. Example: Confirmed Primary Intent 'Sealing / Resealing' → Primary Entity 'Travertine Sealing'. Do NOT include the slash or the additional terms after it in Primary Entity.",
      "HUB PAGE EXCEPTION: If Article Type = Hub Page → Primary Entity = [Material] Hub",
      "",
      "LEGITIMACY PREFIX (Publish Justification / legitimacy field): Must begin exactly with LEGIT:Allowed — or LEGIT:Conditional — or LEGIT:Disallowed —",
      "Publish Justification must end after the LEGIT reasoning sentence. Do NOT append the govFlags string (e.g. [COMM:...][GEOCTA:...]) to Publish Justification — that value belongs ONLY at the end of Page Governance Summary.",
      "",
      "CASE STUDY HARD LOCK — ENTITY TYPE:",
      "Fires ONLY when Article Type = Case Study AND Entity Role = Authority Support Entity (both exact).",
      "In that case ONLY: Entity Type = Material Behaviour.",
      "For ALL other Article Types: Entity Type MUST come from the Role Mapping Table.",
      "",
      "DATE RULE (Entity Governance Date): Use " + today + " exactly. DO NOT use the page published date or updated date.",
      "",
      "FEEDS HUB: If REQUIRED FEEDS HUB URL = none, output exactly: https://www.abbeyfloorcare.co.uk/" +
        String(rowData.stoneType || "").toLowerCase().trim().replace(/\s+/g, "-") +
        "/hub-page-temp/\n" +
      "Otherwise output the EXACT URL supplied in REQUIRED FEEDS HUB URL. Do NOT infer, detect, rewrite, shorten, or reformat it. Raw URL only — no markdown, no brackets, no parentheses.\n" +
      "\n" +
      "ENTITY GOVERNANCE STATUS — HARD LOCK: Must equal exactly one of: Governed, Compliant, Requires Review, Unclear. No other value is permitted (e.g. 'Allowed' is NOT valid).\n" +
      "",
      "SUPPORTING ENTITIES CORE: Populate ONLY with material authority entity names from the TSM Recognised Entities table. Do NOT include service authority entities here.",
      "",
      "SUPPORTING ENTITIES - SERVICE AUTHORITY LAYER: Populate ONLY with service authority entity names from Stage 4. Do NOT include TSM material authority entities here.",
      "",
      "PERIPHERAL ENTITIES LINK OUT — HARD LOCK: This field does NOT contain named entities, page titles, or topics. It contains ONLY structural link-role category labels, separated by semicolons, chosen from this fixed set: 'Hub Page (Material)', 'Supporting Guide (Same Intent)', 'Supporting Guide (Adjacent Allowed Intent)', 'Case Study (Evidence) (Same Material)', 'FAQ (Same Material)'. Select only the categories that genuinely apply to this page's outbound linking strategy. Do NOT invent new category labels. Do NOT output specific entity names, sealer names, or product names in this field.",
      "",
      "GSC INTENT EVIDENCE — HARD LOCK: Use EXACTLY this format with ACTUAL numbers: [X] clicks / [Y] impressions / avg position [Z]. Always use \"clicks\" even for 0 or 1. Numbers only, no descriptive words. Use Page Clicks and Page Impressions from the page-level GSC block, not individual query rows.",
      "",
      "PRIMARY QUERY CLUSTER OWNED: Must begin with the canonical material name. Must express the cluster as a READER PROBLEM or READER INTENT STATEMENT — language a real homeowner would type into Google.",
      "HARD PROHIBITION: Do NOT use analytical vocabulary from the article, governance fields, or entity names as the cluster. Words like 'suitability', 'doubts', 'governance', 'entity', 'legitimacy', 'diagnostic' must never appear.",
      "STANDARD MODE (Page Clicks > 0 OR Page Impressions > 0): Derive from actual GSC query data only. NO suffix symbol.",
      "FALLBACK MODE (Page Clicks = 0 AND Page Impressions = 0): Derive from article content analysis — dominant reader problem the content addresses. Output format: [material] + [reader intent statement] † — the † symbol flags governance-derived intent.",
      "",
      "DRIFT STATUS: Determined ONLY by material mismatch, intent type mismatch, or legitimacy conflict. If Material matches AND Intent remains Allowed → Drift Status = Aligned.",
      "",
      "DIY / TASK-LED INSTRUCTION vs PRIMARY TASK — PRECEDENCE RULE (HARD LOCK): If the Confirmed Primary Intent is itself a hands-on task (e.g. Cleaning, Repair / Chip / Crack, Sealing / Resealing, Honing / Polishing / Finish Recovery), then any H2/H3 section giving step-by-step instructions for THAT SAME task is part of the primary task — it must NOT also be marked as a separate DIY / Task-Led Instruction intent in the H2 Validation Log or Secondary Intent Decisions. DIY / Task-Led Instruction only applies as a genuinely separate intent when the section covers a different self-service task than the one named in the Confirmed Primary Intent (e.g. a repair article containing a separate DIY cleaning tip that is not part of the repair steps).",
      "",
      "REWRITE ROLE LOCK — HARD LOCK: Format exactly: Role Stable — [Exact Entity Role Label]. No Article Type. No variation. No other prefix.",
      "",
      "PAGE REWRITE BRIEF — QUERY VECTOR RULE (HARD LOCK): You MUST reference the Precision Score from the QUERY VECTOR ANALYSIS block above in the Page Rewrite Brief. If the Precision Score is below 70%, select at least 3 genuinely useful MISSING terms from real search queries and instruct that they be worked naturally into the rewritten content. QUERY TERMS ARE SUBORDINATE TO GOVERNANCE: never require a missing term if it conflicts with the governed Material, Article Type, Primary Entity, SCOPE BOUNDARY, or governed locality. For Geo Service Pages and other locality-specific pages, a query term naming a different town, county, city, region or service area must be treated as cross-location query evidence only and MUST NOT be instructed into the rewritten body. Exclude generic/noise terms that do not materially improve the governed reader task. If fewer than 3 safe missing terms remain after this filtering, use only the safe terms available and explicitly state that conflicting terms were excluded by governance. If the Precision Score is 70% or above, state that the page's language is already well-aligned with real search demand and no additional term coverage is required.",
      "",
      "PAGE REWRITE BRIEF — MANDATORY CONTENT-GROUNDING AUDIT (HARD LOCK):",
      "Before writing the Page Rewrite Brief, you MUST read the actual article HTML supplied",
      "in the PAGE METADATA / HTML block below — not just the governed entity list.",
      "For every material, product, technique, or intervention you intend to name in the",
      "Rewrite Brief as a next step, escalation path, or alternative approach, you must first",
      "check what THIS ARTICLE actually says about it — not what the entity list or your",
      "general material knowledge assumes.",
      "",
      "Specifically check for and preserve, verbatim in effect:",
      "— Whether the article restricts a material/technique to professionals only (e.g. due to",
      "  cost, tools, smell, skill level, or time required) versus presenting it as a homeowner",
      "  DIY option.",
      "— Whether any section explicitly recommends AGAINST a method, product, or approach for",
      "  homeowners, even if that same method/product is a governed entity for this page.",
      "— Whether any section states a scope limit, comparison, or caveat that narrows how an",
      "  entity should be framed (e.g. 'X is cheaper and easier, Y is for professionals').",
      "",
      "The Rewrite Brief must NEVER phrase two distinct options ambiguously as one combined",
      "step (e.g. 'material X and professional intervention apply when Y' is AMBIGUOUS unless",
      "the article confirms X and professional work are the same recommendation). If the article",
      "content contradicts or narrows what the general entity list implies, the article's actual",
      "stated position always wins — state it explicitly and unambiguously in the brief.",
      "",
      "This audit applies to every material and article type — do not skip it because an entity",
      "seems self-evident from its name.",
      "",
      "PAGE REWRITE BRIEF — REQUIRED INTERNAL STRUCTURE (HARD LOCK):",
      "Format the Page Rewrite Brief value using these exact labelled sub-sections, separated",
      "by ' | ' on a single field line (downstream stages parse these labels specifically).",
      "Every label below MUST appear even if its content is 'Not applicable' — omitting a label",
      "entirely, for any reason, is a validation failure:",
      "RE-ANCHOR FOCUS: [entities/topics this rewrite must strengthen]",
      "SCOPE BOUNDARY: [what this page must not drift into, consistent with the Cannibalisation Guardrail]",
      "ESCALATION BOUNDARY: [the exact point/action where DIY or basic intervention stops and",
      "  professional work begins, stated as a single unambiguous fact grounded in the content-",
      "  grounding audit above — write 'Not applicable' only if the article genuinely contains no",
      "  such boundary, never omit the label]",
      "QUERY PRECISION ACTION: [terms to add per the Query Vector Rule, or the compliant statement",
      "  if Precision Score is 70% or above — this label must always be populated, never dropped]",
      "Example: 'RE-ANCHOR FOCUS: Natural Void Structure, Factory Filler Weakness | SCOPE BOUNDARY:",
      "localised repair only, no restoration or sealing | ESCALATION BOUNDARY: resin filler repair",
      "requires professional application due to specialist tools and skill level — homeowners should",
      "use sanded grout only | QUERY PRECISION ACTION: Precision Score 81%, no additional coverage required'",
      "",
      preOutputValidationCombined,
      "",
      gscDataBlock,
      "",
      headingBlock,
      "",
      dataBlock,
      "",
      "===== FINAL OUTPUT REQUIREMENT (THIS OVERRIDES ANY OTHER OUTPUT STRUCTURE) =====",
      "Your final answer MUST end with the exact marker line: === COMBINED VERTICAL OUTPUT ===",
      "Immediately after that marker, output ONLY these exact field labels, one per line, in this exact order, with no renaming, no merging, no omissions, and no BLOCK headers of any kind:",
      "Article Type: [value]",
      "Confirmed Primary Intent: [value]",
      "Secondary Intent Decisions: [all decisions on ONE line, semicolons, format: N: Y or N — intent name — reason]",
      "Primary Entity: [value]",
      "Material Entity: [value]",
      "Entity Role: [value]",
      "Entity Type: [value]",
      "Supporting Entities Core: [value]",
      "Supporting Entities - Service Authority Layer: [value]",
      "Peripheral Entities Link Out: [value]",
      "Feeds Hub: [raw URL only, no markdown, no brackets]",
      "Page Governance Summary: [value]",
      "Entity Governance Status: [value]",
      "Entity Governance Date: [value]",
      "Publish Justification: [value]",
      "GAS Handoff Document: [value]",
      "Observed Query Cluster: [value]",
      "GSC Intent Evidence: [value]",
      "Primary Query Cluster Owned: [value]",
      "Drift Status: [value]",
      "Rewrite Status: [value]",
      "Rewrite Governance Summary: [value]",
      "Cannibalisation Guardrail: [value]",
      "Safe Handoff Pages: [raw URL only, no markdown, no brackets]",
      "Rewrite Role Lock: [value]",
      "Page Rewrite Brief: [value]",
      "Rewrite Brief Date: [value]",
      "After the final field, output exactly: Waiting for Input",
      "Do NOT wrap any URL in markdown link syntax [text](url) — output the raw URL only.",
      "You may still show your GSC Signal Assessment, H2 Validation Log, Ranked Intents, and Justification reasoning BEFORE this final section, but the marker and field list above MUST appear, complete and unmodified, as the last part of your response.",
      "",
      "Please process the above and return your combined output now."
    ].join("\n");

    if (lockedRow >= 2 && currentHeadingsRaw) {
      sheet.getRange(lockedRow, 84).setValue(currentHeadingsRaw);
    }

    return { fullPrompt: fullPrompt, title: rowData.title };

  } catch (e) {
    throw new Error("bc_getPromptCombined: " + e.message);
  }
}

function bc_stripMarkdownUrl_(value) {
  if (!value) return value;
  value = String(value).trim();
  const selfLink = /^\[(https?:\/\/[^\]]+)\]\(\1\)$/;
  const mdLink = /^\[([^\]]*)\]\((https?:\/\/[^)\s"]+)[^)]*\)$/;
  const bareBracketUrl = /^\[(https?:\/\/[^\]]+)\]$/;
  const m1 = value.match(selfLink);
  if (m1) return m1[1];
  const m2 = value.match(mdLink);
  if (m2) return m2[2];
  const m3 = value.match(bareBracketUrl);
  if (m3) return m3[1];
  return value;
}

function bc_writeCombinedOutputToRow(raw, lockedRow) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const postsSheet = ss.getSheets().find(s => s.getSheetId() == BC_SHEET_CONFIG.posts);
    if (!postsSheet) throw new Error("Posts sheet not found.");

    const activeRowIndex = lockedRow || ss.getActiveRange().getRow();
    if (activeRowIndex < 2) throw new Error("No valid row to write to.");

    const marker = '=== COMBINED VERTICAL OUTPUT ===';
    const start = raw.indexOf(marker);
    const text = start > -1 ? raw.substring(start + marker.length) : raw;

    const pairs = {};

const existingValues = {};

const originalHeaders = postsSheet
  .getRange(1, 1, 1, postsSheet.getLastColumn())
  .getValues()[0]
  .map(function(h) {
    return String(h || "").trim();
  });

[
  "GAS Handoff Document",
  "Safe Handoff Pages"
].forEach(function(name) {
  const idx = originalHeaders.indexOf(name);

  if (idx > -1) {
    existingValues[name] = String(
      postsSheet
        .getRange(activeRowIndex, idx + 1)
        .getValue() || ""
    ).trim();
  }
});

text.split('\n').forEach(function(line) {
      const colonIdx = line.indexOf(':');
      if (colonIdx === -1) return;
      const label = line.substring(0, colonIdx).trim();
      let value = line.substring(colonIdx + 1).trim();
      if (!label) return;
      if (label === 'Waiting for Input') return;
      pairs[label] = value;
    });

// Normalise Secondary Intent Decisions numbering
if (pairs["Secondary Intent Decisions"]) {
  const items = String(pairs["Secondary Intent Decisions"])
    .split(";")
    .map(function(item) {
      return item.trim();
    })
    .filter(Boolean);

  const normalisedItems = items.map(function(item, index) {
    const match = item.match(/(?:\d+|[A-Za-z]+)?\s*:\s*([YN])\s*—\s*(.*)/i);

    if (match) {
      return (index + 1) + ": " +
        match[1].toUpperCase() + " — " +
        match[2].trim();
    }

    return (index + 1) + ": " + item;
  });

  pairs["Secondary Intent Decisions"] =
    normalisedItems.join("; ");
}


    const headers = postsSheet
      .getRange(1, 1, 1, postsSheet.getLastColumn())
      .getValues()[0]
      .map(function(h) { return String(h || "").trim().toLowerCase(); });

    let written = 0;
    const missing = [];

    Object.keys(pairs).forEach(function(label) {
      const idx = headers.indexOf(label.toLowerCase().trim());
      if (idx === -1) {
        missing.push(label);
        return;
      }
      postsSheet.getRange(activeRowIndex, idx + 1).setValue(bc_stripMarkdownUrl_(pairs[label]));
      written++;
    });

    let msg = "Wrote " + written + " column(s) to row " + activeRowIndex + ".";
    if (missing.length) {
      msg += " Column(s) not found in posts sheet: " + missing.join(", ") + ".";
    }

    // Populate GSC Query List for Query Vector analysis
  try {
    const rowDataForGsc = bc_getRowSpecificData();

    const gscQueryList = bc_extractGscQueryList_(
      rowDataForGsc.queries,
      rowDataForGsc.materialEntity
    );

    const gscHeaders = postsSheet
      .getRange(1, 1, 1, postsSheet.getLastColumn())
      .getValues()[0]
      .map(function(h) {
        return String(h || "").trim().toLowerCase();
      });

    const gscIdx = gscHeaders.indexOf("gsc query list");

    if (gscIdx > -1 && gscQueryList) {
      postsSheet
        .getRange(activeRowIndex, gscIdx + 1)
        .setValue(gscQueryList);

      msg += " GSC Query List written.";
    }
  } catch (gscErr) {
    msg += " GSC Query List error: " + gscErr.message;
  }

    // Now that classification fields are written, calculate and write Recovery Blueprint
    try {
      const rowDataForBlueprint = bc_getRowSpecificData();
      const cols1to13ForBlueprint = bc_readActiveRowFields([
        "Article Type", "Primary Entity", "Material Entity", "Entity Role",
        "Entity Type", "Supporting Entities Core", "Supporting Entities - Service Authority Layer", "Peripheral Entities Link Out",
        "Feeds Hub", "Page Governance Summary", "Entity Governance Status",
        "Entity Governance Date", "Publish Justification", "GAS Handoff Document",
        "Page Rewrite Brief"
      ]);
      const recoveryBlueprint = bc_buildRecoveryBlueprint(cols1to13ForBlueprint, rowDataForBlueprint);
      const rbHeaders = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || "").trim().toLowerCase(); });
      const rbIdx = rbHeaders.indexOf("recovery blueprint");
      if (rbIdx > -1) {
        postsSheet.getRange(activeRowIndex, rbIdx + 1).setValue(recoveryBlueprint);
        msg += " Recovery Blueprint written.";
      }
    } catch (rbErr) {
      msg += " Recovery Blueprint error: " + rbErr.message;
    }

    let qvMsg = "";
    try {
      qvMsg = " " + bc_calculateQueryVector(activeRowIndex);
    } catch (qvErr) {
      qvMsg = " Query vector error: " + qvErr.message;
    }

    let govResult;
    try {
      govResult = bc_validateGovernanceColumns(activeRowIndex);
    } catch (govErr) {
      govResult = { success: false, message: govErr.message };
    }

    return { message: msg + qvMsg, governance: govResult };

  } catch (e) {
    throw new Error("bc_writeCombinedOutputToRow: " + e.message);
  }
}

function bc_runStage2BAutomated() {
  var startTime = Date.now();
  var totalCost = 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select a data row first.');
  }

  // W2B must be built from the CURRENT validated W1.5D and W1.5E outputs.
  var currentCX =
    String(
      sheet.getRange(row, 102).getValue() || ''
    ).trim();

  var currentCY =
    String(
      sheet.getRange(row, 103).getValue() || ''
    ).trim();

  if (!currentCX) {
    throw new Error(
      'Column CX is empty — W1.5D must complete before W2B.'
    );
  }

  if (!currentCY) {
    throw new Error(
      'Column CY is empty — W1.5E must complete before W2B.'
    );
  }

  // Clear stale W2B Raw HTML before generating from current governance.
  var headers =
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

  var euIdx =
    headers.indexOf(
      'W2B Raw HTML'
    );

  if (euIdx === -1) {
    throw new Error(
      "Column 'W2B Raw HTML' not found in posts sheet."
    );
  }

  sheet
    .getRange(
      row,
      euIdx + 1
    )
    .clearContent();

  var data =
    getStage2BData();

  if (
    !data ||
    !data.part1 ||
    !data.part2 ||
    !data.part3
  ) {
    throw new Error(
      'W2B prompt data is incomplete.'
    );
  }

  var fullPrompt =
    data.part1 +
    '\n\n' +
    data.part2 +
    '\n\n' +
    data.part3;

  var apiResult =
    bc_sendPromptViaOpenAI(
      fullPrompt,
      8000,
      MODEL_CHEAP
    );

  if (!apiResult.success) {
    throw new Error(
      apiResult.message
    );
  }

  totalCost +=
    apiResult.cost;

  var saveResult =
    saveW2BHtmlToSheet(
      apiResult.text
    );

  bc_addToApiCostAndTime(
    totalCost,
    (Date.now() - startTime) / 1000
  );

  if (!saveResult.success) {

    bc_appendGovernancePipelineException(
      'W2B — Article Generation',
      saveResult.message
    );

    return {
      success: false,
      message:
        'W2B validation failed — EU left unchanged and failure recorded in GJ.',
      cost: totalCost
    };
  }

  return {
    success: true,
    message: saveResult.message,
    cost: totalCost
  };
}
