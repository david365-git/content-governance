/**
 * ============================================================
 * bc_DataFetcher.gs
 * Abbey Floor Care — Governance Prompt Assembler
 * Data Retrieval and Row Reading
 *
 * Version: 1.0 — Split from bc_GovernancePromptAssembler_DataFetcher.gs v3.4
 *
 * Contains:
 * - bc_getTodayDDMMYYYY()
 * - bc_getRowSpecificData()
 * - bc_getTargetMaterialEntity()
 * - bc_readActiveRowFields()
 *
 * Depends on: bc_GovernancePromptAssembler_Config.gs
 * ============================================================
 */


function bc_getTodayDDMMYYYY() {
  const d    = new Date();
  const dd   = String(d.getDate()).padStart(2, '0');
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}


function bc_getRowSpecificData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const postsSheet = ss.getSheets().find(
    s => s.getSheetId() == BC_SHEET_CONFIG.posts
  );
  if (!postsSheet) throw new Error("Posts sheet not found.");

  const activeRowIndex = ss.getActiveRange().getRow();
  if (activeRowIndex < 2) {
    throw new Error("Select a data row in the Posts sheet.");
  }

  const postsHeaders = postsSheet
    .getRange(1, 1, 1, postsSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).toLowerCase().trim());

  const activePostsRow = postsSheet
    .getRange(activeRowIndex, 1, 1, postsSheet.getLastColumn())
    .getValues()[0];

  const findPostVal = (name) => {
    const idx = postsHeaders.indexOf(String(name).toLowerCase().trim());
    return idx > -1 ? activePostsRow[idx] : "";
  };

  const postId = findPostVal("Post ID") ||
                 postsSheet.getRange(activeRowIndex, 4).getValue();

  const articleType = String(findPostVal("Article Type") || findPostVal("Confirmed Article Type") || "").trim();
  Logger.log("postsHeaders = " + JSON.stringify(postsHeaders));
  Logger.log("raw Article Type from posts = [" + findPostVal("Article Type") + "]");
  Logger.log("articleType final = [" + articleType + "]");  

  const exportSheet = ss.getSheetByName(BC_SHEET_CONFIG.export);
  if (!exportSheet) {
    throw new Error("Sheet '" + BC_SHEET_CONFIG.export + "' not found.");
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'export_sheet_data';
  let exportData;
  const cached = cache.get(cacheKey);
  if (cached) {
    exportData = JSON.parse(cached);
  } else {
    exportData = exportSheet.getDataRange().getValues();
    try {
      cache.put(cacheKey, JSON.stringify(exportData), 600);
    } catch(e) {
      // Cache too large — skip caching
    }
  }
  const exportHeaders = exportData[0].map(
    h => String(h).toLowerCase().trim()
  );

  const matchIndex = exportData.findIndex(
    row => String(row[0]) === String(postId)
  );

  if (matchIndex === -1) {
    throw new Error(
      "ID " + postId + " not found in '" + BC_SHEET_CONFIG.export + "'."
    );
  }

  const matchRow = exportData[matchIndex];

  const findExportVal = (name) => {
    const idx = exportHeaders.indexOf(String(name).toLowerCase().trim());
    return idx > -1 ? String(matchRow[idx] || "") : "";
  };

  const fullHtml = findExportVal("Full Post HTML");
  const hasAmazon =
    fullHtml.includes("amzn.to") || fullHtml.includes("amazon.co.uk");

  const artSheet = ss.getSheets().find(
    s => s.getSheetId() == BC_SHEET_CONFIG.articleTypeControl
  );
  if (!artSheet) throw new Error("Article Type Control sheet not found.");

  const artData = artSheet.getDataRange().getValues();
  if (artData.length < 2) {
    throw new Error("Article Type Control sheet is empty.");
  }

  const artHeaders = artData[0].map(h => String(h || "").trim().toLowerCase());

  const findArtIdx = (name) => artHeaders.indexOf(String(name).trim().toLowerCase());

  const artTypeIdx                  = findArtIdx("Canonical Article Type Label");
  const onPageExpectationsIdx       = findArtIdx("On-Page Content Expectations");
  const commercialStrategyLockIdx   = findArtIdx("Commercial Strategy Lock");
  const geoStrategyLockIdx          = findArtIdx("Geo Strategy Lock");
  const reasoningIdx                = findArtIdx("Reasoning");
  const intentDecisionTreeIdx       = findArtIdx("Intent Ranking Rules");
  const serviceAuthorityLayerRuleIdx = findArtIdx("Service Authority Layer Rule");

  let commFlag                  = "[COMM:No]";
  let geoFlag                   = "[GEOCTA:No]";
  let onPageExpectations        = "";
  let reasoning                 = "";
  let intentDecisionTree        = "";
  let serviceAuthorityLayerRule = "Use weighting table";

  for (let i = 1; i < artData.length; i++) {
    const rowArticleType = String(
      artTypeIdx > -1 ? artData[i][artTypeIdx] : artData[i][0]
    ).trim();

  Logger.log("rowArticleType = [" + rowArticleType + "] | articleType = [" + articleType + "]");

    if (rowArticleType === articleType) {
      commFlag = commercialStrategyLockIdx > -1
        ? String(artData[i][commercialStrategyLockIdx] || "[COMM:No]").trim()
        : "[COMM:No]";

      geoFlag = geoStrategyLockIdx > -1
        ? String(artData[i][geoStrategyLockIdx] || "[GEOCTA:No]").trim()
        : "[GEOCTA:No]";

      onPageExpectations = onPageExpectationsIdx > -1
        ? String(artData[i][onPageExpectationsIdx] || "").trim()
        : "";

      reasoning = reasoningIdx > -1
        ? String(artData[i][reasoningIdx] || "").trim()
        : "";

      intentDecisionTree = intentDecisionTreeIdx > -1
        ? String(artData[i][intentDecisionTreeIdx] || "").trim()
        : "";

      serviceAuthorityLayerRule = serviceAuthorityLayerRuleIdx > -1
      ? String(artData[i][serviceAuthorityLayerRuleIdx] || "Use weighting table").trim()
      : "Use weighting table";

      break;
    }
  }

  if (commFlag === "[COMM:BelowFold]") {
    commFlag = hasAmazon ? "[COMM:Yes]" : "[COMM:BelowFold]";
  }

  const searchEntity = String(findPostVal("Search Entity") || "").trim();
  const stoneType = String(findPostVal("Stone Type") || "Unknown").trim();
  const location = String(
    findPostVal("Locality") || findPostVal("Location") || ""
  ).trim();
  const materialEntity = bc_getTargetMaterialEntity(stoneType, searchEntity);

  Logger.log("serviceAuthorityLayerRule = " + serviceAuthorityLayerRule);

  Logger.log("serviceAuthorityLayerRuleIdx = " + serviceAuthorityLayerRuleIdx);

  return {
    title:                     findExportVal("Title"),
    postId:                    postId,
    articleType:               articleType,
    govFlags:                  commFlag + geoFlag,
    stoneType:                 stoneType,
    searchEntity:              searchEntity,
    materialEntity:            materialEntity,
    location:                  location,
    fullHtml:                  fullHtml,
    clicks:                    String(findPostVal("Clicks") || "0"),
    impressions:               String(findPostVal("Impressions") || "0"),
    position:                  String(findPostVal("Position") || "0"),
    queries:                   bc_getQueriesFromSheet(String(postId), 'GSC Queries'),
    canonicalUrl:              findExportVal("Canonical URL"),
    publishedDate:             findExportVal("Published Date"),
    updatedDate:               findExportVal("Last Updated Date"),
    metaTitle:                 findExportVal("Meta Title") || "(none)",
    metaDesc:                  findExportVal("Meta Description") || "(none)",
    schema:                    findExportVal("WPCode Header Schema") || "(none)",
    onPageExpectations:        onPageExpectations,
    reasoning:                 reasoning,
    intentDecisionTree:        intentDecisionTree,
    serviceAuthorityLayerRule: serviceAuthorityLayerRule
  };
}


function bc_getTargetMaterialEntity(hub, variant) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(
      s => s.getSheetId() == BC_SHEET_CONFIG.materialMapping
    );
    if (!sheet) return hub;
    const data  = sheet.getDataRange().getValues();
    const clean = str =>
      String(str || "").trim().toLowerCase().replace(/s$/, "");
    const match = data.find(
      row =>
        clean(row[0]) === clean(hub) &&
        clean(row[1]) === clean(variant)
    );
    return (match && match[2]) ? String(match[2]).trim() : hub;
  } catch (e) {
    return hub;
  }
}


function bc_readActiveRowFields(fieldNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheets().find(
    s => s.getSheetId() == BC_SHEET_CONFIG.posts
  );
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
  for (let i = 0; i < fieldNames.length; i++) {
    const name   = fieldNames[i];
    const idx    = headers.indexOf(name.toLowerCase().trim());
    result[name] = idx > -1 ? String(row[idx] || "") : "";
  }
  return result;
}

function bc_writePrompt1ALockToRow(confirmedArticleType, confirmedPrimaryIntent, classificationContext, secondaryDecisions, p1aResponse, lockedRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheets().find(
    s => s.getSheetId() == BC_SHEET_CONFIG.posts
  );
  if (!postsSheet) throw new Error("Posts sheet not found.");

  const activeRowIndex = lockedRow || ss.getActiveRange().getRow();
  if (activeRowIndex < 2) throw new Error("No valid row to write to.");

  const headers = postsSheet
    .getRange(1, 1, 1, postsSheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h || "").trim().toLowerCase());

  const articleTypeIdx = headers.indexOf("article type");
  const confirmedPrimaryIntentIdx = headers.indexOf("confirmed primary intent");
  const classificationContextIdx = headers.indexOf("1a classification context");
  const secondaryIntentIdx = headers.indexOf("secondary intent decisions");

  if (articleTypeIdx === -1) throw new Error("Column 'Article Type' not found.");
  if (confirmedPrimaryIntentIdx === -1) throw new Error("Column 'Confirmed Primary Intent' not found.");

  postsSheet.getRange(activeRowIndex, articleTypeIdx + 1).setValue(
    String(confirmedArticleType || "").trim()
  );
  postsSheet.getRange(activeRowIndex, confirmedPrimaryIntentIdx + 1).setValue(
    String(confirmedPrimaryIntent || "").trim()
  );

  if (classificationContextIdx > -1 && classificationContext) {
    postsSheet.getRange(activeRowIndex, classificationContextIdx + 1).setValue(
      String(classificationContext || "").trim()
    );
  }

  if (secondaryIntentIdx > -1 && secondaryDecisions) {
    postsSheet.getRange(activeRowIndex, secondaryIntentIdx + 1).setValue(
      String(secondaryDecisions || "").trim()
    );
  }

  const p1aResponseIdx = headers.indexOf("p1a response");
  if (p1aResponseIdx > -1 && p1aResponse) {
    postsSheet.getRange(activeRowIndex, p1aResponseIdx + 1).setValue(
      String(p1aResponse || "").trim()
    );
  }
}

function bc_getQueriesFromSheet(postId, sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return '';
    if (sheet.getLastRow() < 2) return '';

    var data = sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
    var lines = [];
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() !== String(postId).trim()) continue;
      var query = String(data[r][1] || '').trim();
      var clicks = String(data[r][2] || '0').trim();
      var impressions = String(data[r][3] || '0').trim();
      if (!query) continue;
      lines.push(query + '  ' + clicks + '  ' + impressions);
    }

    return lines.join('\n');
  } catch(e) {
    Logger.log('bc_getQueriesFromSheet error: ' + e.message);
    return '';
  }
}

function testQuerySheetLookup() {
  var result = bc_getQueriesFromSheet('43693', 'GSC Queries');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  sheet.getRange(1, 1).setComment('Length: ' + result.length + ' | First 200: ' + result.substring(0, 200));
}