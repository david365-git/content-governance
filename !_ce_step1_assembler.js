/**
 * ce_Step1_Assembler.gs
 * Abbey Floor Care — Content Generation Pipeline
 * Step 1: Structure Audit & Section Mapping
 */

var STEP1_PROMPT_TEMPLATE = 
"STEP 1 — STRUCTURE AUDIT & SECTION MAPPING\n" +
"ROLE: Senior UK SEO & Stone Restoration Strategist\n" +
"\n" +
"--- SYSTEM INSTRUCTION ---\n" +
"This is Step 1 of the Content Generation Pipeline.\n" +
"Your only job is to run 12 audit checks and produce a section mapping.\n" +
"Do not write any HTML. Do not write any content. Do not summarise or analyse.\n" +
"Your output must follow the exact format specified at the end of this prompt.\n" +
"\n" +
"--- INPUT 1: AUTHORITY BRIEF ---\n" +
"[AUTHORITY_BRIEF]\n" +
"\n" +
"--- INPUT 2: TECHNICAL SPECIFICATION MATRIX ---\n" +
"[TSM]\n" +
"\n" +
"--- INPUT 3: TECHNICAL ENTITIES ---\n" +
"[TECHNICAL_ENTITIES]\n" +
"\n" +
"--- INPUT 4: TIER STRUCTURAL REQUIREMENTS ---\n" +
"[TIER_REQUIREMENTS]\n" +
"\n" +
"--- FAILURE RULE ---\n" +
"If any check fails, you MUST output a one-sentence reason immediately after the FAIL result.\n" +
"Format: AUDIT CHECK [N]: FAIL — [one sentence reason]\n" +
"A FAIL result with no reason is not permitted.\n" +
"\n" +
"--- AUDIT CHECKS ---\n" +
"Run the following 12 checks using the inputs above:\n" +
"\n" +
"AUDIT CHECK 1 — ARTICLE TYPE CLASSIFICATION\n" +
"Verify the article type from the Authority Brief matches the structural expectations for that type.\n" +
"AUDIT CHECK 1: PASS / FAIL\n" +
"\n" +
"AUDIT CHECK 2 — TIER WORD BUDGET RANGE\n" +
"Minimum word budget for Tier 1 is 2600 words.\n" +
"AUDIT CHECK 2: PASS / FAIL\n" +
"\n" +
"AUDIT CHECK 3 — TSM REQUIREMENT COUNT\n" +
"Count the number of TSM sections available for this material and article type.\n" +
"Do NOT count the header block. Do NOT count HUB-INTRO if present.\n" +
"AUDIT CHECK 3: [N] sections required\n" +
"\n" +
"AUDIT CHECK 4 — WORD BUDGET DISTRIBUTION\n" +
"Each section must have a governed word range from the TSM.\n" +
"AUDIT CHECK 4: PASS / FAIL\n" +
"\n" +
"AUDIT CHECK 5 — COMM FLAG VALIDATION\n" +
"Read COMM flag from Authority Brief.\n" +
"AUDIT CHECK 5: PASS / FAIL / NOT APPLICABLE\n" +
"\n" +
"AUDIT CHECK 6 — GEOCTA FLAG VALIDATION\n" +
"Read GEOCTA flag from Authority Brief.\n" +
"AUDIT CHECK 6: PASS / FAIL / NOT APPLICABLE\n" +
"\n" +
"AUDIT CHECK 7 — PRIMARY ENTITY ALIGNMENT\n" +
"Confirm the Primary Entity from the Authority Brief is present in the TSM.\n" +
"AUDIT CHECK 7: PASS / FAIL\n" +
"\n" +
"AUDIT CHECK 8 — ARTICLE TYPE EXPECTATIONS\n" +
"Validate content structure only against the article type expectations from the Authority Brief.\n" +
"AUDIT CHECK 8: PASS / FAIL\n" +
"\n" +
"AUDIT CHECK 9 — TSM ORDER SEQUENCE\n" +
"Verify TSM Orders are sequential (1, 2, 3...) with no gaps.\n" +
"AUDIT CHECK 9: PASS / FAIL\n" +
"\n" +
"AUDIT CHECK 10 — SECTION COUNT MATCH\n" +
"Confirm section count matches TSM requirement count from Audit Check 3.\n" +
"Do NOT count the header block. Do NOT count HUB-INTRO if present.\n" +
"AUDIT CHECK 10: PASS / FAIL\n" +
"\n" +
"AUDIT CHECK 11 — WORD BUDGET TOTAL\n" +
"Total word budget must fall within the Tier minimum and 1.5x that minimum.\n" +
"AUDIT CHECK 11: PASS / FAIL\n" +
"\n" +
"AUDIT CHECK 12 — ORIGINAL IMAGE COUNT\n" +
"Count the number of section-content images in the original HTML from the Authority Brief.\n" +
"Do NOT count: author portrait images, logos, icons, avatars, sprites, tracking pixels, Amazon product images.\n" +
"PASS if one or more qualifying images found and count is reported.\n" +
"NOT APPLICABLE if no qualifying images found.\n" +
"FAIL if count is missing or excluded types are counted.\n" +
"AUDIT CHECK 12: PASS / FAIL / NOT APPLICABLE\n" +
"\n" +
"--- ARTICLE TYPE EXPECTATIONS ---\n" +
"Hub Page: High-level summary of full system, visual router menu, and clear internal linking to all governed entities.\n" +
"Educational Guide: Deep-dive definitions, conceptual explanations, and material behaviour coverage.\n" +
"Method Guide: Numbered steps, required tools, total time, and safety warnings for a specific process.\n" +
"Diagnostic Guide: Defect definition, cause categories, visual indicators, risk evaluation, correction pathway.\n" +
"Case Study: Project location, before and after evidence, specific metrics, job narrative.\n" +
"Service Page: What we do description, trust signals, service availability.\n" +
"Geo Service Page: Localised keywords, service area statement, conversion pathway.\n" +
"\n" +
"--- TSM-TO-SECTION MAPPING ---\n" +
"Using the tier requirements from INPUT 4 (TIER STRUCTURAL REQUIREMENTS), map each Order to a section number.\n" +
"Do NOT use Core Entities from the Authority Brief as section names.\n" +
"Do NOT invent section names. Use only Order names as written in INPUT 4.\n" +
"\n" +
"HUB-INTRO HARD LOCK: If Article Type is Hub Page, you MUST output HUB-INTRO as the FIRST line before SECTION 1.\n" +
"HUB-INTRO is mandatory for Hub Pages, is not a TSM section, and does not count toward the required TSM section count.\n" +
"Use this exact format:\n" +
"HUB-INTRO: id=\"hub-intro\" | H2=\"None\" | CONTENT=\"One orienting sentence and a quick-links navigation list only\"\n" +
"\n" +
"Then output each section using this format:\n" +
"SECTION [N]: Order [N] — [TSM Requirement Name] — [min]–[max] words\n" +
"\n" +
"--- OUTPUT INSTRUCTION ---\n" +
"Output ONLY:\n" +
"1. The 12 audit check results\n" +
"2. The HUB-INTRO line (if Hub Page)\n" +
"3. The section mapping\n" +
"\n" +
"Last line must be exactly: \"Step 1 complete. Waiting for Step 2.\"\n" +
"No other text. No analysis. No commentary.\n";


function getActivePostData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("posts");
  if (!sheet) throw new Error("posts sheet not found.");

  var activeRow = sheet.getActiveRange().getRow();
  if (activeRow < 2) throw new Error("Select a data row in the posts sheet.");

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  function col(name) {
    var idx = headers.map(function(h) { return String(h).trim().toLowerCase(); })
                     .indexOf(name.toLowerCase().trim());
    return idx > -1 ? String(sheet.getRange(activeRow, idx + 1).getValue() || "") : "";
  }

  var postId     = col("Post ID");
  var title      = col("Title");
  var material   = col("Stone Type");
  var articleType = col("Article Type");
  var authorityBrief = col("Authority Brief");

  var tierMap = {
    "Hub Page":          "Tier 1",
    "Educational Guide": "Tier 1",
    "Method Guide":      "Tier 2",
    "Service Page":      "Tier 2",
    "Geo Service Page":  "Tier 2",
    "Diagnostic Guide":  "Tier 3",
    "Buyer Guide":       "Tier 3",
    "FAQ Spoke":         "Tier 3",
    "Case Study":        "Tier 4"
  };
  var tierLabel = tierMap[articleType] || "Tier 2";

  // Get TSM from technical sheet
  var tsm = "";
  var techSheet = ss.getSheetByName("technical");
  if (techSheet) {
    var techData = techSheet.getDataRange().getValues();
    var matLower = material.trim().toLowerCase();
    for (var i = 0; i < techData.length; i++) {
      if (String(techData[i][1] || "").trim().toLowerCase() === matLower) {
        tsm = String(techData[i][2] || "").trim();
        break;
      }
    }
  }

  // Get technical entities from technical_entities sheet
  var technicalEntities = "";
  var entSheet = ss.getSheetByName("technical_entities");
  if (entSheet) {
    var entData = entSheet.getDataRange().getValues();
    var matLower2 = material.trim().toLowerCase();
    var entLines = [];
    entLines.push("Entity | Type | Primary Application | Semantic Fingerprint Vocabulary");
    for (var j = 1; j < entData.length; j++) {
      if (String(entData[j][0] || "").trim().toLowerCase() === matLower2) {
        entLines.push([
          String(entData[j][1] || "").trim(),
          String(entData[j][2] || "").trim(),
          String(entData[j][3] || "").trim(),
          String(entData[j][4] || "").trim()
        ].join(" | "));
      }
    }
    if (entLines.length > 1) technicalEntities = entLines.join("\n");
  }

  // Get tier requirements from Tier Structural Coverage Matrix
  var tierRequirements = getTierStructuralRequirements(material, tierLabel, articleType);
  var tierRequirementsText = "";
  tierRequirements.forEach(function(req, index) {
    var minWords = req.wordCount;
    var maxWords = Math.round(minWords * 1.5);
    tierRequirementsText += "Order " + (index + 1) + ": " + req.name + " (" + minWords + "–" + maxWords + " words)\n";
  });

  return {
    postId:           postId,
    title:            title,
    material:         material,
    articleType:      articleType,
    authorityBrief:   authorityBrief,
    tsm:              tsm,
    technicalEntities: technicalEntities,
    tierRequirements:  tierRequirementsText
  };
}


function assembleStep1Prompt(postData) {
  var prompt = STEP1_PROMPT_TEMPLATE;
  prompt = prompt.replace("[AUTHORITY_BRIEF]",    postData.authorityBrief    || "(Authority Brief not found)");
  prompt = prompt.replace("[TSM]",                postData.tsm               || "(TSM not found for material: " + postData.material + ")");
  prompt = prompt.replace("[TECHNICAL_ENTITIES]", postData.technicalEntities || "(Technical Entities not found for material: " + postData.material + ")");
  prompt = prompt.replace("[TIER_REQUIREMENTS]",  postData.tierRequirements  || "(Tier requirements not found for: " + postData.material + " / " + postData.articleType + ")");
  return prompt;
}


function saveStep1Output(postId, output) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("posts");
  if (!sheet) throw new Error("posts sheet not found.");

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var postIdIdx = headers.map(function(h) { return String(h).trim().toLowerCase(); })
                         .indexOf("post id");
  var step1Idx  = headers.map(function(h) { return String(h).trim().toLowerCase(); })
                         .indexOf("step 1 output");

  if (postIdIdx === -1) throw new Error("Post ID column not found in posts sheet.");
  if (step1Idx  === -1) throw new Error("Step 1 Output column not found in posts sheet.");

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][postIdIdx]) === String(postId)) {
      sheet.getRange(i + 1, step1Idx + 1).setValue(output);
      return;
    }
  }

  throw new Error("Post ID " + postId + " not found in posts sheet.");
}