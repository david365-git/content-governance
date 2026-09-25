/**
 * ================================================================================
 * ce_Stage0_Forensic.gs - STAGE 0: FORENSIC ANALYSIS & ENTITY GOVERNANCE
 * ================================================================================
 *
 * Purpose: GSC data analysis, entity drift detection, cluster classification,
 *          recovery blueprints, and entity governance gate checking
 *
 * Key Functions:
 * - generateForensicPrompt()      → Builds forensic analysis prompt for ChatGPT
 * - checkGSCDataStatus()          → Validates GSC data availability
 * - checkEntityGovernanceDate()   → Governance freshness gate (3-day check)
 * - pushStage0FieldsToActiveRow() → Writes 5 analytical fields to sheet
 *
 * Outputs:
 * - Primary Search Term
 * - Strategic Reasoning
 * - Topical Intent
 * - Matrix Role
 * - Key Decisions Explained
 *
 * Version update: Governed entity profile injected into Stage 0 prompt.
 * - Supporting Entities Core read from active row — entities already governed
 *   are excluded from GAP line candidates.
 * - Secondary Intent Decisions read from active row — intents marked N
 *   are excluded from GAP line scope.
 * - Service Authority Entity Table injected alongside TSM.
 * - GAP lines reference both TSM and Service Authority governed lists.
 * - Material name matched against Column B of service-authority sheet.
 * - Entity table read from Column D.
 * - Geo-lock matching delegated to LLM via prompt instruction.
 * ================================================================================
 */
function getW0PostIdFromStatus() {
  var txt = document.getElementById('status0')?.textContent || '';
  var m = txt.match(/ID:\s*([^\s—]+)/);
  return m ? m[1] : 'unknown';
}

function checkGSCDataStatus() {
  const d = getActiveRowDataMap();
  const postID         = d["Post ID"] || "N/A";
  const monitorImpr    = d["Montr Impr"];
  const postIdCheck = String(d["Post ID"] || "").trim();
  const monitorQueries = postIdCheck ? bc_getQueriesFromSheet(postIdCheck, 'GSC Monitor') : String(d["Montr Queries"] || "").trim();
  const hasMonitorData = (monitorImpr !== "" && monitorImpr !== 0 && monitorQueries !== "");
  if (!hasMonitorData) {
    return {
      success: false,
      error:   "Missing Monitor Data (Impressions or Queries). Cannot analyze current state.",
      postID:  postID
    };
  }
  const articleType = String(d["Article Type"] || "").trim();
  return { success: true, error: null, postID: postID, articleType: articleType, title: String(d["Title"] || "").trim() };
}


function checkEntityGovernanceDate() {
  const d      = getActiveRowDataMap();
  const postID = d["Post ID"] || "N/A";
  const title  = d["Title"]   || d["Stone Type"] || "No Title";

  const govDate = d["Entity Governance Date"]
               || d["Governance Date"]
               || d["Gov Date"]
               || "";

  if (!govDate || govDate === "") {
    return {
      pass:      false,
      postID:    postID,
      title:     title,
      date:      "Not set",
      daysSince: null,
      message:   "Entity Governance Date is not set for this post. " +
                 "Running the pipeline without a current governance date risks " +
                 "producing content with stale entity alignment. " +
                 "Run Entity Governance first."
    };
  }

  let govDateObj;
  if (govDate instanceof Date) {
    govDateObj = govDate;
  } else {
    const s        = String(govDate).trim();
    const ukMatch  = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (ukMatch) {
      govDateObj = new Date(
        parseInt(ukMatch[3]),
        parseInt(ukMatch[2]) - 1,
        parseInt(ukMatch[1])
      );
    } else {
      govDateObj = new Date(s);
    }
  }

  if (isNaN(govDateObj.getTime())) {
    return {
      pass:      false,
      postID:    postID,
      title:     title,
      date:      String(govDate),
      daysSince: null,
      message:   "Entity Governance Date \"" + String(govDate) +
                 "\" could not be read as a valid date. " +
                 "Check the column format in the sheet."
    };
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  govDateObj.setHours(0, 0, 0, 0);

  const msPerDay   = 1000 * 60 * 60 * 24;
  const daysSince  = Math.round((today - govDateObj) / msPerDay);
  const dateStr    = govDateObj.toLocaleDateString("en-GB", {
    day: "numeric", month: "long", year: "numeric"
  });

  if (daysSince <= 3) {
    return {
      pass:      true,
      postID:    postID,
      title:     title,
      date:      dateStr,
      daysSince: daysSince,
      message:   "Entity Governance Date is current (" + dateStr + ", " +
                 (daysSince <= 0 ? "today" : daysSince + " day(s) ago") +
                 "). Proceed."
    };
  }

  const daysText = daysSince === 1 ? "1 day" : daysSince + " days";
  return {
    pass:      false,
    postID:    postID,
    title:     title,
    date:      dateStr,
    daysSince: daysSince,
    message:   "Entity Governance Date is " + dateStr + " — " + daysText +
               " ago. This exceeds the 3-day freshness window. " +
               "Entity alignment may be stale. Run Entity Governance before " +
               "proceeding, or confirm you want to continue anyway."
  };
}


/* ============================================================
   FULL TECHNICAL SPECIFICATION MATRIX LOOKUP
   Reads the technical sheet, finds the row matching the
   material name in column B (index 1), returns the full
   matrix from column C (index 2).
============================================================ */
function getFullTSM(material) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const techSheet = ss.getSheetByName("technical");
  if (!techSheet) {
    Logger.log("'technical' sheet not found");
    return "";
  }

  const data          = techSheet.getDataRange().getValues();
  const materialLower = material.toString().trim().toLowerCase();

  for (let i = 0; i < data.length; i++) {
    const cellVal = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
    if (cellVal === materialLower) {
      return data[i][2] ? data[i][2].toString() : "";
    }
  }

  Logger.log("No Technical Specification Matrix found for material: " + material);
  return "";
}


/* ============================================================
   ENTITY TABLE EXTRACTION
   Extracts Section 10 from a full Technical Specification
   Matrix string. Used to build the entity name constraint block.
============================================================ */
function extractEntityTable(tsmText) {
  if (!tsmText) return "";

  const sectionMarker = "RECOGNISED TECHNICAL ENTITIES";
  const sectionIndex  = tsmText.indexOf(sectionMarker);
  if (sectionIndex === -1) return "";

  const endMarker = "PROMPT GOVERNANCE RULES";
  const endIndex  = tsmText.indexOf(endMarker, sectionIndex);

  if (endIndex !== -1) {
    return tsmText.substring(sectionIndex, endIndex).trim();
  }
  return tsmText.substring(sectionIndex).trim();
}


/* ============================================================
   SERVICE AUTHORITY ENTITY TABLE LOOKUP
   Reads the service-authority sheet, finds the row matching
   the material name in column B (index 1), returns the
   entity table from column D (index 3).
============================================================ */
function getServiceAuthorityEntityTable(material) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find(
    function(s) { return s.getSheetId() == BC_SHEET_CONFIG.serviceAuthority; }
  );

  if (!sheet) {
    Logger.log("Service authority sheet not found.");
    return "";
  }

  const data          = sheet.getDataRange().getValues();
  const materialLower = material.toString().trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const cellVal = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
    if (cellVal === materialLower) {
      return data[i][3] ? data[i][3].toString() : "";
    }
  }

  Logger.log("No Service Authority Entity Table found for material: " + material);
  return "";
}


/* ============================================================
   TECHNICAL SPECIFICATION MATRIX PROMPT BLOCK BUILDER
   Injects the full matrix plus a detailed guidance header.
   Section 10 = material entity name constraint (governed list)
   Sections 1–9 = reasoning source only
   Used in both the no-GSC and has-GSC prompt paths.
============================================================ */
function buildTSMPromptBlock(material, articleType, primaryEntity) {
  const fullTSM = getFullTSM(material);

    if (!fullTSM) {
    return "--- TECHNICAL SPECIFICATION MATRIX: " + material.toUpperCase() + " ---\n" +
      "WARNING: No Technical Specification Matrix found for material \"" + material + "\".\n" +
      "Material GAP lines are NOT permitted without a governed Technical Specification Matrix.\n" +
      "Do not invent entity names, categories, section references, or substitute generic placeholders.\n" +
      "Examples of forbidden generic GAP names include: cleaning process, damage issue, surface problem, maintenance method, sealing system, restoration step.\n\n" +
      "CRITICAL CONSTRAINTS ON GAP LINES:\n" +
      "— TSM GAP lines cannot be produced unless a governed TSM is present.\n" +
      "— If no TSM exists, output no material GAP lines.\n" +
      "— Only Service Authority GAP lines may be produced if a governed Service Authority Entity Table exists.\n";
  }

  let entityTable = "";
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const entitiesSheet = ss.getSheetByName("technical_entities");
        if (entitiesSheet) {
          const entitiesData = entitiesSheet.getDataRange().getValues();
          const materialLower = material.toLowerCase().trim();
          const entityLines = [];
          entityLines.push("Entity                           | Type                  | Primary Application                                                      | Semantic Fingerprint Vocabulary");
          for (let i = 1; i < entitiesData.length; i++) {
            const rowMaterial = String(entitiesData[i][0] || "").trim().toLowerCase();
            if (rowMaterial === materialLower) {
              entityLines.push(
                String(entitiesData[i][1] || "").trim() + " | " +
                String(entitiesData[i][2] || "").trim() + " | " +
                String(entitiesData[i][3] || "").trim() + " | " +
                String(entitiesData[i][4] || "").trim()
              );
            }
          }
          entityTable = entityLines.join("\n");
        }
      } catch(e) {
        entityTable = extractEntityTable(fullTSM);
      }

  return "--- TECHNICAL SPECIFICATION MATRIX: " + material.toUpperCase() + " ---\n\n" +
    "HOW TO USE THIS MATRIX — READ THIS BEFORE PROCESSING ANYTHING ELSE:\n\n" +
    "This matrix is divided into two distinct roles within this prompt. " +
    "You must treat them differently.\n\n" +
    "ROLE 1 — SECTION 10 (RECOGNISED TECHNICAL ENTITIES): " +
    "THIS IS YOUR SOURCE OF MATERIAL ENTITY NAMES.\n" +
    "Section 10 contains the complete governed list of named technical entities for " +
    material + ". Every material entity name you output in a GAP line must be drawn " +
    "from this list exactly, character for character. You may not invent names, " +
    "paraphrase names, combine names, or use names from any other source including " +
    "your training knowledge. If a name does not appear in Section 10, it does not " +
    "exist for the purposes of this task. The entity list is reproduced separately " +
    "below the full matrix as an additional reference — use it to check your material " +
    "GAP line names before outputting.\n\n" +
    "ROLE 2 — SECTIONS 1–9: THESE ARE YOUR REASONING SOURCE ONLY.\n" +
    "Sections 1–9 contain the material science, diagnosis protocols, cleaning chemistry, " +
    "restoration processes, sealing systems, maintenance rules, pricing, and storytelling " +
    "frameworks for " + material + ". You must use these sections to write the REASON " +
    "component of each GAP line — explaining precisely why a specific entity is likely " +
    "absent from a thin " + articleType + " about " + primaryEntity + ", grounded in " +
    "what the matrix says about that entity's role and importance. Sections 1–9 do not " +
    "contribute entity names — they contribute reasoning depth only.\n\n" +
    "WHAT A GOOD MATERIAL GAP LINE LOOKS LIKE:\n" +
    "GAP: Acid Etching | SOURCE: TSM | TSM SECTION: 2 | REASON: A thin " +
    articleType + " about " + primaryEntity + " is unlikely to define acid etching " +
    "as a distinct damage mechanism separate from soiling, yet the matrix confirms " +
    "it is physical surface dissolution that cannot be corrected by cleaning alone, " +
    "so its absence weakens diagnostic accuracy and treatment legitimacy.\n\n" +
    "WHAT A BAD GAP LINE LOOKS LIKE:\n" +
    "GAP: Surface Gloss Clarity | SOURCE: TSM | TSM SECTION: 1 | REASON: ... " +
    "[REJECTED — this entity name does not appear in Section 10]\n" +
    "GAP: Marble Properties | SOURCE: TSM | TSM SECTION: 1 | REASON: ... " +
    "[REJECTED — this is a generic category name, not a governed entity]\n" +
    "GAP: Cleaning Process | SOURCE: TSM | TSM SECTION: 4 | REASON: ... " +
    "[REJECTED — this is a vague process label, not a governed material entity]\n" +
    "GAP: Damage Issue | SOURCE: TSM | TSM SECTION: 2 | REASON: ... " +
    "[REJECTED — this is a generic placeholder, not a named entity]\n\n" +
    "SECTION NUMBER ACCURACY:\n" +
    "When you assign a section number to a material GAP line, it must reflect the " +
    "section in this matrix where that entity's primary application is described. " +
    "Section 10 lists the primary application for each entity — use this to confirm " +
    "the correct section assignment before outputting.\n\n" +
    "THE FULL MATRIX FOLLOWS:\n\n" +
    fullTSM + "\n\n" +
    "--- END OF TECHNICAL SPECIFICATION MATRIX ---\n\n" +
    "--- GOVERNED MATERIAL ENTITY LIST (SECTION 10 EXTRACTED) ---\n" +
    "Cross-check every material GAP entity name you intend to output against " +
    "this list before writing your response.\n" +
    "If the name is not here exactly, do not use it.\n\n" +
    entityTable + "\n\n" +
    "--- END OF GOVERNED MATERIAL ENTITY LIST ---\n";
}


/* ============================================================
   SERVICE AUTHORITY PROMPT BLOCK BUILDER
   Injects the Service Authority Entity Table alongside the
   Technical Specification Matrix block. Service authority
   entities are a second governed source for GAP lines in
   Stage 0. Geo-lock matching is delegated to the LLM via
   the prompt instruction.
============================================================ */
function buildServiceAuthorityPromptBlock(material, articleType, location) {
  const entityTable = getServiceAuthorityEntityTable(material);

  if (!entityTable) {
    return "--- SERVICE AUTHORITY ENTITY TABLE: " + material.toUpperCase() + " ---\n" +
      "(No Service Authority Entity Table found for this material. " +
      "Service authority GAP lines cannot be produced.)\n" +
      "--- END SERVICE AUTHORITY ENTITY TABLE ---";
  }

  const geoInstruction = (String(articleType || "").trim() === "Geo Service Page")
    ? "\nGEO-LOCK INSTRUCTION:\n" +
      "Target location for this page: " + location + "\n" +
      "Match this location against the delivery entities in the table below.\n" +
      "Include only the single delivery entity that matches this location.\n" +
      "Do not include any other delivery entity.\n" +
      "If no delivery entity matches this location, " +
      "do not include any delivery entity.\n"
    : "";

  return "--- SERVICE AUTHORITY ENTITY TABLE: " + material.toUpperCase() + " ---\n\n" +
    "HOW TO USE THIS TABLE:\n\n" +
    "This table is the governed source of service authority entity names for " +
    material + " content.\n" +
    "Service authority entities are a second category of GAP line alongside " +
    "material entities from the Technical Specification Matrix. Both types of " +
    "GAP line use the same format but must reference their correct source.\n\n" +
    "WHAT A GOOD SERVICE AUTHORITY GAP LINE LOOKS LIKE:\n" +
    "GAP: 25-Year Stone Specialism | SOURCE: Service Authority | REASON: A thin " +
    articleType + " about " + material + " is unlikely to ground its service " +
    "recommendations in documented specialist tenure, yet the absence of this entity " +
    "leaves the page unable to differentiate Abbey Floor Care from a generic cleaning " +
    "contractor in Google's entity model.\n\n" +
    "GAP LINE FORMAT FOR SERVICE AUTHORITY ENTITIES:\n" +
    "GAP: [Exact Entity Name from table below] | SOURCE: Service Authority | " +
    "REASON: [one detailed sentence explaining why this entity is likely absent " +
    "and what its absence costs the page]\n\n" +
    "SELECTION RULE:\n" +
    "Include service authority entities where the Article Type Weighting column " +
    "shows PRIMARY or INTEGRATED for this article type. Include VALIDATING entities " +
    "only where they directly validate a specific claim present in the article content. " +
    "Exclude all entities marked ABSENT for this article type. " +
    "CRITICAL: If an entity has no weighting entry at all for this article type — " +
    "treat it as ABSENT. Do NOT infer inclusion from weightings assigned to other article types.\n" +
    geoInstruction + "\n" +
    "WEIGHTING KEY:\n" +
    "PRIMARY    — Include as a GAP line if absent. This entity anchors service authority.\n" +
    "INTEGRATED — Include as a GAP line if absent and relevant to the specific article intent.\n" +
    "VALIDATING — Include only where it validates a specific claim in the article.\n" +
    "ABSENT     — Do not include for this article type.\n\n" +
    "ENTITY NAME RULE:\n" +
    "Use entity names exactly as written in the Entity column below.\n" +
    "Do not paraphrase, abbreviate, or combine entity names.\n" +
    "Do not invent service authority entities not present in this table.\n\n" +
    "SERVICE AUTHORITY ENTITY TABLE:\n\n" +
    entityTable + "\n\n" +
    "--- END OF SERVICE AUTHORITY ENTITY TABLE ---\n";
}


/* ============================================================
   GOVERNED ENTITY PROFILE BLOCK BUILDER
   Reads Supporting Entities Core and Secondary Intent Decisions
   from the active row and builds a locked exclusion block.
   Stage 0 uses this to avoid producing GAP lines for entities
   already governed or intents already excluded.
============================================================ */
function buildGovernedEntityProfileBlock(supportingEntitiesCore, secondaryIntentDecisions) {
  if (!supportingEntitiesCore || !supportingEntitiesCore.trim()) {
    return "";
  }

  var block =
    "--- GOVERNED ENTITY PROFILE FOR THIS PAGE ---\n" +
    "The following entities have already been confirmed as present in the\n" +
    "governed entity profile for this page by the Entity Classification pipeline.\n" +
    "Do NOT produce GAP lines for any entity in this list.\n" +
    "They are already governed.\n\n" +
    "SUPPORTING ENTITIES CORE (already governed):\n" +
    supportingEntitiesCore.trim() + "\n";

  if (secondaryIntentDecisions && secondaryIntentDecisions.trim()) {
    block +=
      "\nSECONDARY INTENT BOUNDARIES:\n" +
      "The following secondary intent decisions have been confirmed by the " +
      "governance pipeline. GAP lines must respect these boundaries.\n" +
      "Intents marked Y are permitted supporting layers — GAP lines for entities\n" +
      "belonging to these intents are allowed.\n" +
      "Intents marked N are excluded from this page — do NOT produce GAP lines\n" +
      "for entities that belong exclusively to these excluded intents.\n\n" +
      secondaryIntentDecisions.trim() + "\n";
  }

  block += "\n--- END GOVERNED ENTITY PROFILE ---\n";
  return block;
}


/* ============================================================
   FORENSIC PROMPT GENERATOR
   Builds the full Stage 0 prompt for ChatGPT including:
   - Technical Specification Matrix block
   - Service Authority Entity Table block
   - Governed Entity Profile block (already-governed entities
     and secondary intent boundaries from the active row)
============================================================ */
function generateForensicPrompt() {
  const d = getActiveRowDataMap();

  const material      = d["Stone Type"]     || "UNKNOWN";
  const articleType   = d["Article Type"]   || "General";
  const primaryEntity = d["Primary Entity"] || material;
  var location = String(d["Locality"] || d["Location"] || "").trim();
  var parentArea = String(d["Parent Area"] || "").trim();
  var requiresLocality = (articleType === "Case Study" || articleType === "Geo Service Page");
  if (requiresLocality && !location) {
    return "STOP — Locality is empty for this " + articleType + ".\n" +
      "Run W0B first to generate Locality, Parent Area, and Location Context, " +
      "then return to W0.\n" +
      "The Primary Search Term for this article type must include the Locality.";
  }
  const competitorNodeMap = String(d["Competitor Node Map"] || "").trim();
  const benchmarkVal  = Number(d["Benmrk Impr"]) || 0;
  const isNewPost     = (benchmarkVal === 0);

  // Read governed entity profile from the active row
  const supportingEntitiesCore   = String(d["Supporting Entities Core"]   || "").trim();
  const secondaryIntentDecisions = String(d["Secondary Intent Decisions"] || "").trim();

  const postId = String(d["Post ID"] || "").trim();
const monitorQueries = postId ? bc_getQueriesFromSheet(postId, 'GSC Monitor') : String(d["Montr Queries"] || "").trim();
  const hasGSCData     = monitorQueries.length > 10;

  const governanceClusterMap = {
    "Hub Page":           "B",
    "Service Page":       "C",
    "Geo Service Page":   "C",
    "Buyer Guide":        "B",
    "Educational Guide":  "B",
    "Method Guide":       "B",
    "Diagnostic Guide":   "B",
    "Case Study":         "B"
  };

  const sectionTitle = isNewPost
    ? "STAGE 0 — NEW ENTITY SEEDING (GROWTH MODE)"
    : "STAGE 0 — FORENSIC ENTITY ALIGNMENT (RECOVERY MODE)";

  // Build all prompt blocks
  const tsmPromptBlock              = buildTSMPromptBlock(material, articleType, primaryEntity);
  const serviceAuthorityPromptBlock = buildServiceAuthorityPromptBlock(material, articleType, location);
  const governedEntityProfileBlock  = buildGovernedEntityProfileBlock(
    supportingEntitiesCore,
    secondaryIntentDecisions
  );

  const competitorNodeMapBlock = competitorNodeMap
    ? "--- COMPETITOR NODE MAP ---\n" +
      "The following competitor pages were found ranking for this page's primary search term.\n" +
      "Use this to understand what node types Google is rewarding for this intent.\n" +
      "Where multiple competitors use the same node type, that node type is confirmed as\n" +
      "authoritative for this intent. Where your page uses a different node type, ensure\n" +
      "the article structure justifies that differentiation.\n\n" +
      competitorNodeMap + "\n\n" +
      "--- END COMPETITOR NODE MAP ---\n"
    : "";

  // Combined constraint summary used in both prompt paths
  const finalConstraintSummary =
    "FINAL CONSTRAINT SUMMARY:\n" +
    "— Material GAP entity names: Section 10 governed list only — exact match required.\n" +
    "— Service authority GAP entity names: Service Authority Entity Table only — exact match required.\n" +
    "— GAP reasoning: draw from Sections 1–9 of the Technical Specification Matrix — be specific, not generic.\n" +
    "— GAP section numbers: required for material GAP lines only — use the primary application column in Section 10 as your guide.\n" +
    "— SOURCE field: required on every GAP line — must be either \"TSM\" or \"Service Authority\".\n" +
    (supportingEntitiesCore
      ? "— Do NOT produce GAP lines for entities already present in the governed entity profile above.\n"
      : "") +
    (secondaryIntentDecisions
      ? "— Do NOT produce GAP lines for entities belonging exclusively to intents marked N above.\n"
      : "") +
    "— Minimum 8 GAP lines required in total across both sources where sufficient absent entities exist.\n" +
    "— Do not invent, paraphrase, generalise, or substitute placeholder category labels for any entity name under any circumstances.";


  // ── NO GSC DATA PATH ──
  if (!hasGSCData) {
    const govCluster = governanceClusterMap[articleType] || "B";

    var primarySearchTerm = primaryEntity.toLowerCase();
    if (articleType === "Geo Service Page") {
      primarySearchTerm = primaryEntity.toLowerCase() + (location ? " " + location : "");
    } else if (articleType === "Buyer Guide") {
      primarySearchTerm = primaryEntity.toLowerCase() + " cost guide";
    } else if (articleType === "Educational Guide") {
      primarySearchTerm = primaryEntity.toLowerCase() + " explained";
    } else if (articleType === "Method Guide") {
      primarySearchTerm = "how to " + primaryEntity.toLowerCase();
    } else if (articleType === "Diagnostic Guide") {
      primarySearchTerm = primaryEntity.toLowerCase() + " problems";
    } else if (articleType === "Case Study") {
      primarySearchTerm = primaryEntity.toLowerCase() + (location ? " " + location : "");
    }

    var noGscParts = [
      "SECTION: " + (isNewPost
        ? "STAGE 0 — NEW ENTITY SEEDING (GROWTH MODE)"
        : "STAGE 0 — FORENSIC ENTITY ALIGNMENT (RECOVERY MODE)"),
      "ROLE: Forensic SEO Data Analyst",
      "STATUS: " + (isNewPost ? "NEW POST (Post-Silo Change)" : "EXISTING POST (Historical Recovery)"),
      "MATERIAL: " + material,
      requiresLocality ? "TARGET LOCALITY: " + location + (parentArea ? " (Parent Area: " + parentArea + ")" : "") : "",
      requiresLocality ? "LOCALITY LOCK (Hard Lock): The Primary Search Term for this " + articleType + " MUST include \"" + location + "\" exactly. Do not substitute a broader area, drop the locality, or generalise to national/regional intent." : "",
      "",
      "--- GSC DATA MAP ---",
      "MONITOR: No query data available — page has insufficient impressions for GSC to surface queries.",
      "",
      "--- GOVERNANCE-DERIVED CLASSIFICATION ---",
      "GSC query data is absent. Cluster classification and primary search term are derived from governance data:",
      "ARTICLE TYPE: " + articleType,
      "PRIMARY ENTITY: " + primaryEntity,
      "CLUSTER: " + govCluster + " (derived from article type — " + articleType +
        " pages use Cluster " + govCluster + " by governance rule)",
      "PRIMARY SEARCH TERM: " + primarySearchTerm,
      "",
      tsmPromptBlock,
      "",
      serviceAuthorityPromptBlock,
      ""
    ];

    if (governedEntityProfileBlock) {
      noGscParts.push(governedEntityProfileBlock);
      noGscParts.push("");
    }

    if (competitorNodeMapBlock) {
      noGscParts.push(competitorNodeMapBlock);
      noGscParts.push("");
    }

    noGscParts = noGscParts.concat([
      finalConstraintSummary,
      "",
      "--- TASK ---",
      "INTERNAL ANALYSIS (do not output — process silently):",
      "1. No GSC query neighbourhood available. Classify as new territory — no existing ranking signals.",
      "2. COMPREHENSIVE ENTITY AUDIT: Using both the Technical Specification Matrix and the " +
        "Service Authority Entity Table provided above, work through all sections and identify " +
        "which entities from both governed lists are likely absent from a thin " + articleType +
        " about " + primaryEntity + ". Use only entity names from the governed lists. " +
        "Include a SOURCE field on every GAP line identifying whether the entity comes from " +
        "the Technical Specification Matrix or the Service Authority Entity Table. " +
        "Do NOT produce GAP lines for entities already listed in the governed entity profile above.",
      "3. Cluster is governance-derived — do not override with query analysis.",
      "",
      "EXECUTE NOW — Output the following format:",
      "",
      "CLUSTER: " + govCluster,
      "GAP: [Exact Entity Name from governed list] | SOURCE: [TSM or Service Authority] | " +
        "TSM SECTION: [N — for TSM entities only, omit for Service Authority] | " +
        "REASON: [one detailed sentence explaining why this entity is likely absent from a thin " +
        articleType + "]",
      "[continue only for specifically named governed entities that are absent and valid for this material/article type — never generic categories or placeholder concepts]",
      "Primary Search Term: " + primarySearchTerm,
      "Strategic Reasoning: [pipe-separated: entity targeting strategy | specific gaps | ranking outcome | positioning]",
      "Topical Intent: [specific user intent phrase]",
      "Matrix Role: [silo role]",
      "Key Decisions Explained: [pipe-separated: cluster rationale | priority entities | depth strategy | linking strategy]",
      "Stage 0 complete. Waiting for Stage 1A."
    ]);

    return noGscParts.join("\n").trim();
  }


  // ── HAS GSC DATA PATH ──
  var taskList;

  if (isNewPost) {
    taskList =
      "INTERNAL ANALYSIS (do not output — process silently):\n" +
      "1. Identify the initial neighbourhood: what intent/category is Google testing " +
        "for this URL based on Monitor data.\n" +
      "2. COMPREHENSIVE ENTITY SEEDING AUDIT: Using both the Technical Specification " +
        "Matrix and the Service Authority Entity Table provided above, work through all " +
        "sections and identify which entities from both governed lists are missing from " +
        "the GSC footprint. Use only entity names from the governed lists. Include a " +
        "SOURCE field on every GAP line. Do NOT produce GAP lines for entities already " +
        "listed in the governed entity profile above.\n" +
      "3. CLUSTER CLASSIFICATION: Using the decision tree below, determine the dominant " +
        "cluster internally.\n" +
      "   Decision tree — impression-weighted, apply in order:\n" +
      "   STEP 1 — IMPRESSION THRESHOLD: Only queries with 20 or more impressions are " +
        "eligible to trigger cluster classification.\n" +
      "   If fewer than 3 queries meet the 20-impression threshold, lower the threshold " +
        "to 10 impressions.\n" +
      "   If no queries meet either threshold, use all queries.\n" +
      "   STEP 2 — Apply to eligible queries only, stop at first match:\n" +
      "   - Any eligible query contains: professional, service, near me, cost, price, " +
        "quote, contractor → CLUSTER C\n" +
      "   - Any eligible query contains: repair, restore, fix, scratched, etched, dull, " +
        "stained, damaged, worn, marks, faded, cracked, mould → CLUSTER B\n" +
      "   - Any eligible query contains: care, clean, how to, best way, safe, avoid, " +
        "maintain, tips, guide, prevent → CLUSTER A\n" +
      "   - Question word present with no damage signal in eligible queries → CLUSTER A\n" +
      "   - All other eligible queries → CLUSTER B\n" +
      "   NOTE: Low-impression queries (under threshold) may inform GAP analysis but " +
        "must not override cluster classification.\n\n" +
      "OUTPUT (output these lines only — nothing before, nothing after):\n" +
      "CLUSTER: [A/B/C/D]\n" +
      "GAP: [Exact Entity Name from governed list] | SOURCE: [TSM or Service Authority] | " +
        "TSM SECTION: [N — for TSM entities only] | " +
        "REASON: [one detailed sentence explaining why this entity must be seeded]\n" +
      "[continue with one GAP line per specifically named governed entity that is absent and valid — " +
        "minimum 8 GAP lines total across both sources where sufficient governed entities exist; never use generic categories or placeholder concepts]\n" +
      "Primary Search Term: [single highest-authority query this page should target — SCOPE ALIGNMENT CONSTRAINT: the term must match this page's specific governed scope, not just the highest-impression query available. A high-impression query belonging to a broader intent than this page covers (e.g. general repair when this page covers only hole-filling) must be rejected in favour of a lower-impression query that accurately reflects what this specific page delivers. CASE STUDY CONSTRAINT: if Article Type is Case Study, the term must reflect project/location-anchored intent (e.g. material + service + location, or material + restoration), never generic how-to or instructional phrasing such as 'how to [verb]...' — a Case Study documents one completed project, not a method]\n" +
      "Strategic Reasoning: [minimum 4 pipe-separated segments — seeding strategy | " +
        "specific authority gaps | expected ranking outcome | knowledge graph positioning]\n" +
      "Topical Intent: [single specific phrase — not generic]\n" +
      "Matrix Role: [this page's precise role in the silo architecture]\n" +
      "Key Decisions Explained: [minimum 4 pipe-separated decisions with specific reasoning]\n" +
      "Stage 0 complete. Waiting for Stage 1A.";
  } else {
    taskList =
      "INTERNAL ANALYSIS (do not output — process silently):\n" +
      "1. Identify entity drift: compare Benchmark vs Monitor — what technical " +
        "terminology has Google forgotten about this page.\n" +
      "2. COMPREHENSIVE ENTITY RECOVERY AUDIT: Using both the Technical Specification " +
        "Matrix and the Service Authority Entity Table provided above, work through all " +
        "sections and identify which entities from both governed lists are absent from " +
        "the Monitor GSC footprint but should be present. Use only entity names from " +
        "the governed lists. Include a SOURCE field on every GAP line. Do NOT produce " +
        "GAP lines for entities already listed in the governed entity profile above.\n" +
      "3. Write a 1-sentence recovery summary. Retain internally.\n" +
      "4. CLUSTER CLASSIFICATION: Using the decision tree below, determine the dominant " +
        "cluster internally.\n" +
      "   Decision tree — impression-weighted, apply in order:\n" +
      "   STEP 1 — IMPRESSION THRESHOLD: Only queries with 20 or more impressions are " +
        "eligible to trigger cluster classification.\n" +
      "   If fewer than 3 queries meet the 20-impression threshold, lower the threshold " +
        "to 10 impressions.\n" +
      "   If no queries meet either threshold, use all queries.\n" +
      "   STEP 2 — Apply to eligible queries only, stop at first match:\n" +
      "   - Any eligible query contains: professional, service, near me, cost, price, " +
        "quote, contractor → CLUSTER C\n" +
      "   - Any eligible query contains: repair, restore, fix, scratched, etched, dull, " +
        "stained, damaged, worn, marks, faded, cracked, mould → CLUSTER B\n" +
      "   - Any eligible query contains: care, clean, how to, best way, safe, avoid, " +
        "maintain, tips, guide, prevent → CLUSTER A\n" +
      "   - Question word present with no damage signal in eligible queries → CLUSTER A\n" +
      "   - All other eligible queries → CLUSTER B\n" +
      "   NOTE: Low-impression queries (under threshold) may inform GAP analysis but " +
        "must not override cluster classification.\n\n" +
      "OUTPUT (output these lines only — nothing before, nothing after):\n" +
      "CLUSTER: [A/B/C/D]\n" +
      "GAP: [Exact Entity Name from governed list] | SOURCE: [TSM or Service Authority] | " +
        "TSM SECTION: [N — for TSM entities only] | " +
        "REASON: [one detailed sentence grounded in the matrix above explaining drift consequence]\n" +
      "[continue with one GAP line per specifically named governed entity that is absent and valid — " +
        "minimum 8 GAP lines total across both sources where sufficient governed entities exist; never use generic categories or placeholder concepts]\n" +
      "Primary Search Term: [single highest-authority query this page should target — SCOPE ALIGNMENT CONSTRAINT: the term must match this page's specific governed scope, not just the highest-impression query available. A high-impression query belonging to a broader intent than this page covers (e.g. general repair when this page covers only hole-filling) must be rejected in favour of a lower-impression query that accurately reflects what this specific page delivers. CASE STUDY CONSTRAINT: if Article Type is Case Study, the term must reflect project/location-anchored intent (e.g. material + service + location, or material + restoration), never generic how-to or instructional phrasing such as 'how to [verb]...' — a Case Study documents one completed project, not a method]\n" +
      "Strategic Reasoning: [minimum 4 pipe-separated segments — specific entity drift pattern | " +
        "specific recovery actions | expected ranking outcome | authority positioning]\n" +
      "Topical Intent: [single specific phrase — not generic]\n" +
      "Matrix Role: [this page's precise role in the silo architecture]\n" +
      "Key Decisions Explained: [minimum 4 pipe-separated decisions with specific reasoning]\n" +
      "Stage 0 complete. Waiting for Stage 1A.";
  }

  var gscParts = [
    "SECTION: " + sectionTitle,
    "ROLE: Forensic SEO Data Analyst",
    "STATUS: " + (isNewPost ? "NEW POST (Post-Silo Change)" : "EXISTING POST (Historical Recovery)"),
    "MATERIAL: " + material,
    requiresLocality ? "TARGET LOCALITY: " + location + (parentArea ? " (Parent Area: " + parentArea + ")" : "") : "",
    requiresLocality ? "LOCALITY LOCK (Hard Lock): The Primary Search Term for this " + articleType + " MUST include \"" + location + "\" exactly. Do not substitute a broader area, drop the locality, or generalise to national/regional intent." : "",
    "",
    "--- GSC DATA MAP ---",
    isNewPost ? "" : "BENCHMARK (Pre-Change): " + d["Benmrk Impr"] + " | Queries: " + d["Benmrk Queries"],
    "MONITOR (Current State): " + d["Montr Impr"] + " | Queries: " + monitorQueries,
    "",
    tsmPromptBlock,
    "",
    serviceAuthorityPromptBlock,
    "",
    competitorNodeMapBlock
  ];

  if (governedEntityProfileBlock) {
    gscParts.push(governedEntityProfileBlock);
    gscParts.push("");
  }

  gscParts = gscParts.concat([
    finalConstraintSummary,
    "",
    "--- TASK ---",
    taskList,
    "",
    "--- UK LOCALISATION RULE ---",
    "This is a UK service area. Ignore all international traffic. " +
      "Focus on British English and UK-specific intent."
  ]);

  return gscParts.join("\n").trim();
}


function pushStage0FieldsToActiveRow(rawOutput) {
  try {
    if (!rawOutput || rawOutput.trim() === "") {
      return { success: false, message: "ERROR: No output provided." };
    }

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheet     = ss.getSheetByName("posts");
    const activeRow = sheet.getActiveRange().getRow();
    if (activeRow < 2) {
      return { success: false, message: "ERROR: Select a valid data row first." };
    }

    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(h) { return String(h).trim(); });

    const FIELDS = [
      { label: "Primary Search Term",     col: "Primary Search Term"     },
      { label: "Strategic Reasoning",     col: "Strategic Reasoning"     },
      { label: "Topical Intent",          col: "Topical Intent"          },
      { label: "Matrix Role",             col: "Matrix Role"             },
      { label: "Key Decisions Explained", col: "Key Decisions Explained" }
    ];

    const lines  = rawOutput.split(/\r?\n/);
    const parsed = {};
    let currentField = null;

    const gapLines = lines
      .map(function(l) { return l.trim(); })
      .filter(function(l) { return /^GAP:/i.test(l); });

    var stopPatterns = [
      /^CLUSTER:/i,
      /^GAP:/i,
      /^Stage 0 complete/i,
      /^PUSHED/i,
      /^---/,
      /^===/
    ];

    lines.forEach(function(line) {
      var trimmed = line.trim();
      var matched = false;

      FIELDS.forEach(function(f) {
        var prefix = f.label + ":";
        if (trimmed.indexOf(prefix) === 0) {
          currentField    = f.label;
          parsed[f.label] = trimmed.substring(prefix.length).trim();
          matched         = true;
        }
      });

      if (!matched) {
        var isStop = stopPatterns.some(function(re) { return re.test(trimmed); });
        if (isStop) {
          currentField = null;
          return;
        }
        if (currentField && trimmed.length > 0) {
          parsed[currentField] = (parsed[currentField] || "") + " " + trimmed;
        }
      }
    });

    const written  = [];
    const missing  = [];
    const notFound = [];

    FIELDS.forEach(function(f) {
      const colIndex = headers.indexOf(f.col);
      if (colIndex === -1) { notFound.push(f.col); return; }
      const value = parsed[f.label];
      if (!value || value.trim() === "") { missing.push(f.label); return; }
      const cleaned = stripMarkdown(value);
      const cell    = sheet.getRange(activeRow, colIndex + 1);
      try {
        cell.setPlainTextValue(cleaned);
      } catch (plainErr) {
        cell.setNumberFormat("@STRING@");
        cell.setValue(cleaned);
      }
      written.push(f.label);
    });

    if (gapLines.length > 0) {
      const gapColIndex = headers.indexOf("W0 Entity Gap Analysis");
      if (gapColIndex !== -1) {
        const gapText = stripMarkdown(gapLines.join("\n"));
        const gapCell = sheet.getRange(activeRow, gapColIndex + 1);
        try {
          gapCell.setPlainTextValue(gapText);
        } catch (plainErr) {
          gapCell.setNumberFormat("@STRING@");
          gapCell.setValue(gapText);
        }
        written.push("W0 Entity Gap Analysis");
      } else {
        notFound.push("W0 Entity Gap Analysis");
      }
    }

    var message = "PUSHED " + written.length + "/6 fields to Row " + activeRow + ".";
    if (missing.length  > 0) message += "\nNot found in output: "       + missing.join(", ");
    if (notFound.length > 0) message += "\nColumn missing from sheet: " + notFound.join(", ");

    if (written.length > 0) logPipelineResume("W0 — Stage 0 Push", "");
    return {
      success:  written.length > 0,
      message:  message,
      written:  written,
      missing:  missing,
      notFound: notFound
    };

  } catch (e) {
    return { success: false, message: "PUSH ERROR: " + e.toString() };
  }
}

/**
 * ================================================================================
 * STAGE 0 GOVERNANCE BYPASS
 * ================================================================================
 *
 * Purpose: Populates the 5 Stage 0 output fields directly from existing
 *          governance data when GSC monitor data is unavailable.
 *
 * When to use: checkGSCDataStatus() returns success: false
 *
 * Fields populated:
 * - Primary Search Term       → Derived from Article Type + Material Entity / Primary Entity
 * - Strategic Reasoning       → Constructed bypass rationale with governance source record
 * - Topical Intent            → From Primary Query Cluster Owned
 * - Matrix Role               → From Entity Role + Entity Type
 * - Key Decisions Explained   → Constructed from governance fields
 *
 * ================================================================================
 */


/**
 * Derives the Primary Search Term from Article Type and material/entity fields.
 * Extracted as a standalone function so both generateForensicPrompt() and
 * populateStage0FromGovernance() can use the same logic.
 *
 * @param {string} articleType  - e.g. "Hub Page", "Geo Service Page", "Case Study"
 * @param {string} primaryEntity - e.g. "Quarry Tile Hub"
 * @param {string} material      - e.g. "Quarry Tile" (from Material Entity or Stone Type)
 * @param {string} location      - e.g. "Oakham" (from Locality column)
 * @returns {string} The derived primary search term
 */
function derivePrimarySearchTerm(articleType, primaryEntity, material, location) {
  var type = String(articleType || "").trim();
  var mat  = String(material || primaryEntity || "").trim().toLowerCase();
  var pe   = String(primaryEntity || "").trim().toLowerCase();
  var loc  = String(location || "").trim();

  switch (type) {
    case "Hub Page":
      return "restoring " + mat;

    case "Geo Service Page":
      return pe + (loc ? " " + loc : "");

    case "Buyer Guide":
      return pe + " cost guide";

    case "Educational Guide":
      return pe + " explained";

    case "Method Guide":
      return "how to " + pe;

    case "Diagnostic Guide":
      return pe + " problems";

    case "Case Study":
      return loc ? pe + " " + loc : pe;

    default:
      return pe;
  }
}


/**
 * Populates Stage 0 fields from governance data when monitor data is missing.
 * Called from the sidebar when checkGSCDataStatus() returns success: false.
 *
 * @returns {Object} Result object with success, message, written, missing, notFound arrays
 */
function populateStage0FromGovernance() {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var sheet     = ss.getSheetByName("posts");
    var activeRow = sheet.getActiveRange().getRow();

    if (activeRow < 2) {
      return { success: false, message: "ERROR: Select a valid data row first." };
    }

    var d = getActiveRowDataMap();

    // ── Source fields ──
    var material       = String(d["Material Entity"] || d["Stone Type"] || "").trim();
    var articleType    = String(d["Article Type"] || "").trim();
    var primaryEntity  = String(d["Primary Entity"] || "").trim();
    var location       = String(d["Locality"] || d["Location"] || "").trim();
    var requiresLocality = (articleType === "Case Study" || articleType === "Geo Service Page");
    if (requiresLocality && !location) {
      return {
        success: false,
        message: "STOP — Locality is empty for this " + articleType + ". " +
          "Run W0B first to generate Locality, Parent Area, and Location Context, " +
          "then return to W0. The Primary Search Term for this article type must include the Locality."
      };
    }
    var entityRole     = String(d["Entity Role"] || "").trim();
    var entityType     = String(d["Entity Type"] || "").trim();
    var queryCluster   = String(d["Primary Query Cluster Owned"] || "").trim();
    var driftStatus    = String(d["Drift Status"] || "").trim();
    var rewriteStatus  = String(d["Rewrite Status"] || "").trim();
    var cannGuardrail  = String(d["Cannibalisation Guardrail"] || "").trim();
    var confirmedIntent = String(d["Confirmed Primary Intent"] || "").trim();
    var secondaryIntents = String(d["Secondary Intent Decisions"] || "").trim();
    var govStatus      = String(d["Entity Governance Status"] || "").trim();
    var govDate        = String(d["Entity Governance Date"] || d["Governance Date"] || "").trim();
    var benchmarkImpr  = Number(d["Benmrk Impr"]) || 0;

    // ── Validate minimum governance data exists ──
    if (!articleType) {
      return { success: false, message: "ERROR: Article Type is empty. Run Entity Governance first." };
    }
    if (!primaryEntity) {
      return { success: false, message: "ERROR: Primary Entity is empty. Run Entity Governance first." };
    }

    // ── Derive the 5 fields ──

    // 1. Primary Search Term
    var primarySearchTerm = derivePrimarySearchTerm(articleType, primaryEntity, material, location);

    // 2. Strategic Reasoning
    var isNewPost = (benchmarkImpr === 0);
    var statusLabel = isNewPost ? "NEW POST (Post-Silo Change)" : "EXISTING POST (Historical Recovery)";

    var strategicParts = [
      "W0 bypassed — no GSC monitor data available",
      "Classification derived from governance pipeline (" + govStatus + " " + govDate + ")",
      "Article Type: " + articleType + " | Status: " + statusLabel,
      "Drift Status: " + driftStatus + " | Rewrite Status: " + rewriteStatus,
      "No drift baseline established — monitor on first GSC appearance"
    ];
    var strategicReasoning = strategicParts.join(" | ");

    // 3. Topical Intent
    var topicalIntent = (queryCluster && queryCluster.toLowerCase().indexOf("not yet evidenced") === -1)
      ? queryCluster
      : material.toLowerCase() + " behaviour and care — " + articleType.toLowerCase();

    // 4. Matrix Role
    var matrixRole = entityRole;
    if (entityType) {
      matrixRole += " — " + entityType;
    }
    if (!matrixRole || matrixRole === " — ") {
      matrixRole = articleType + " role for " + material;
    }

    // 5. Key Decisions Explained
    var keyDecisionParts = [
      "Cluster derived from article type governance — no GSC query classification possible",
      "Primary intent: " + (confirmedIntent || articleType),
      "Cannibalisation boundary: " + (cannGuardrail || "not set")
    ];

    if (secondaryIntents) {
      // Extract a concise summary of retained intents
      var retainedMatches = secondaryIntents.match(/\d+\.\s+[^\n—]+/g);
      if (retainedMatches && retainedMatches.length > 0) {
        keyDecisionParts.push("Retained supporting intents: " + retainedMatches.join("; "));
      }
    }

    keyDecisionParts.push("GAP analysis deferred to W1 stages — no monitor entity footprint to audit against");
    var keyDecisionsExplained = keyDecisionParts.join(" | ");

    // ── Write to sheet ──
    var headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(h) { return String(h).trim(); });

    var FIELDS = [
      { label: "Primary Search Term",     value: primarySearchTerm     },
      { label: "Strategic Reasoning",     value: strategicReasoning    },
      { label: "Topical Intent",          value: topicalIntent          },
      { label: "Matrix Role",             value: matrixRole             },
      { label: "Key Decisions Explained", value: keyDecisionsExplained  }
    ];

    var written  = [];
    var notFound = [];

    FIELDS.forEach(function(f) {
      var colIndex = headers.indexOf(f.label);
      if (colIndex === -1) {
        notFound.push(f.label);
        return;
      }
      var cell = sheet.getRange(activeRow, colIndex + 1);
      try {
        cell.setPlainTextValue(f.value);
      } catch (plainErr) {
        cell.setNumberFormat("@STRING@");
        cell.setValue(f.value);
      }
      written.push(f.label);
    });

    var message = "GOVERNANCE BYPASS — PUSHED " + written.length + "/6 fields to Row " + activeRow + ".";
    if (notFound.length > 0) {
      message += "\nColumn missing from sheet: " + notFound.join(", ");
    }

    return {
      success:  written.length > 0,
      message:  message,
      written:  written,
      notFound: notFound
    };

  } catch (e) {
    return { success: false, message: "GOVERNANCE BYPASS ERROR: " + e.toString() };
  }
}

/**
 * Checks the posts sheet for other rows sharing the same Locality + Material
 * whose Primary Search Term significantly overlaps with the active row's PST.
 * Returns an array of {postId, title, pst} for any matches found.
 */
function checkPstDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("posts");
  var activeRow = sheet.getActiveRange().getRow();
  if (activeRow < 2) return { success: false, duplicates: [] };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  var pstIdx = headers.indexOf("Primary Search Term");
  var localityIdx = headers.indexOf("Locality");
  var materialIdx = headers.indexOf("Stone Type");
  var postIdIdx = headers.indexOf("Post ID");
  var titleIdx = headers.indexOf("Title");

  if (pstIdx === -1 || localityIdx === -1 || materialIdx === -1) {
    return { success: false, duplicates: [], message: "Required column missing." };
  }

  var data = sheet.getDataRange().getValues();
  var activeData = data[activeRow - 1];

  var myPst = String(activeData[pstIdx] || "").trim().toLowerCase();
  var myLocality = String(activeData[localityIdx] || "").trim().toLowerCase();
  var myMaterial = String(activeData[materialIdx] || "").trim().toLowerCase();
  var myPostId = String(activeData[postIdIdx] || "").trim();

  if (!myPst || !myLocality) return { success: true, duplicates: [] };

  var myWords = myPst.split(/\s+/).filter(function(w) { return w.length > 2; });

  var duplicates = [];
  for (var i = 1; i < data.length; i++) {
    if (i + 1 === activeRow) continue;

    var rowLocality = String(data[i][localityIdx] || "").trim().toLowerCase();
    var rowMaterial = String(data[i][materialIdx] || "").trim().toLowerCase();
    if (rowLocality !== myLocality || rowMaterial !== myMaterial) continue;

    var rowPst = String(data[i][pstIdx] || "").trim().toLowerCase();
    if (!rowPst) continue;

    var rowWords = rowPst.split(/\s+/).filter(function(w) { return w.length > 2; });
    var overlap = myWords.filter(function(w) { return rowWords.indexOf(w) > -1; });
    var overlapRatio = overlap.length / Math.max(myWords.length, rowWords.length);

    if (overlapRatio >= 0.6) {
      duplicates.push({
        postId: String(data[i][postIdIdx] || ""),
        title: String(data[i][titleIdx] || "(no title)"),
        pst: String(data[i][pstIdx] || "")
      });
    }
  }

  return { success: true, duplicates: duplicates, myPostId: myPostId, myPst: activeData[pstIdx] };
}

/**
 * Saves a new, manually-differentiated Primary Search Term to the active row.
 */
function saveNewPrimarySearchTerm(newPst) {
  if (!newPst || !newPst.trim()) {
    return { success: false, message: "New PST cannot be empty." };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("posts");
  var activeRow = sheet.getActiveRange().getRow();
  if (activeRow < 2) return { success: false, message: "Select a valid data row first." };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).trim(); });
  var pstIdx = headers.indexOf("Primary Search Term");
  if (pstIdx === -1) return { success: false, message: "Primary Search Term column not found." };

  sheet.getRange(activeRow, pstIdx + 1).setPlainTextValue(newPst.trim());
  return { success: true, message: "New Primary Search Term saved." };
}
