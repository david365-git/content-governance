/* ============================================================
   ce_Stage6_PublicationLog.gs
   W6 — WordPress Export + Publication Log
============================================================ */

function ce_getW6Data() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const posts     = ss.getSheetByName("posts");
    const exportSh  = ss.getSheetByName("site-export");
    if (!posts)    throw new Error("posts sheet not found.");
    if (!exportSh) throw new Error("site-export sheet not found.");

    const row = posts.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const postHeaders = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0];
    function getPostVal(name) {
      const idx = postHeaders.indexOf(name);
      return idx > -1 ? String(posts.getRange(row, idx + 1).getValue() || "").trim() : "";
    }

    const postId      = getPostVal("Post ID");
    const title       = getPostVal("Title");
    const stoneType   = getPostVal("Stone Type");
    const articleType = getPostVal("Article Type");
    const newH1       = getPostVal("New H1");
    const newHtml     = getPostVal("New HTML");
    const newSchema   = getPostVal("Schema (JSON-LD)");
    const locality    = getPostVal("Locality");

    // Parse New Meta Description and Yoast Keyphrase
    const rawMetaDesc     = getPostVal("New Meta Description");
    const keyphraseMarker = "Yoast Keyphrase:";
    let newMetaDesc  = rawMetaDesc;
    let yoastKeyphrase = "";
    const kpIdx = rawMetaDesc.indexOf(keyphraseMarker);
    if (kpIdx > -1) {
      newMetaDesc    = rawMetaDesc.substring(0, kpIdx).trim();
      yoastKeyphrase = rawMetaDesc.substring(kpIdx + keyphraseMarker.length).trim();
    }

    const newMetaTitle = getPostVal("New Meta Title");

    // Read old values from site-export
    const exportData    = exportSh.getDataRange().getValues();
    const exportHeaders = exportData[0].map(function(h) { return String(h).trim(); });
    const expIdIdx      = exportHeaders.indexOf("ID");
    const expTitleIdx   = exportHeaders.indexOf("Title");
    const expDateIdx    = exportHeaders.indexOf("Published Date");
    const expMetaTitleIdx = exportHeaders.indexOf("Meta Title");
    const expMetaDescIdx  = exportHeaders.indexOf("Meta Description");
    const expHtmlIdx    = exportHeaders.indexOf("Full Post HTML");
    const expKpIdx      = exportHeaders.indexOf("Yoast Keyphrase");

    let oldTitle = "", oldDate = "", oldMetaTitle = "", oldMetaDesc = "", oldHtml = "", oldKeyphrase = "";

    for (let i = 1; i < exportData.length; i++) {
      if (String(exportData[i][expIdIdx] || "").trim() === String(postId)) {
        oldTitle      = expTitleIdx      > -1 ? String(exportData[i][expTitleIdx]      || "").trim() : "";
        oldDate       = expDateIdx       > -1 ? String(exportData[i][expDateIdx]       || "").trim() : "";
        oldMetaTitle  = expMetaTitleIdx  > -1 ? String(exportData[i][expMetaTitleIdx]  || "").trim() : "";
        oldMetaDesc   = expMetaDescIdx   > -1 ? String(exportData[i][expMetaDescIdx]   || "").trim() : "";
        oldHtml       = expHtmlIdx       > -1 ? String(exportData[i][expHtmlIdx]       || "").trim() : "";
        oldKeyphrase  = expKpIdx         > -1 ? String(exportData[i][expKpIdx]         || "").trim() : "";
        break;
      }
    }

    // Additional fields for publication log
    const pqc          = getPostVal("Primary Query Cluster Owned");
    const intent       = getPostVal("Confirmed Primary Intent");
    const primaryTerm  = getPostVal("Primary Search Term");
    const entityStatus = getPostVal("Entity Governance Status");
    const driftStatus  = getPostVal("Drift Status");
    const rewriteStatus = getPostVal("Rewrite Status");
    const clicks       = getPostVal("Clicks");
    const impressions  = getPostVal("Impressions");
    const position     = getPostVal("Position");
    const recallScore  = getPostVal("recall_score");
    const recallNotes  = getPostVal("recall_notes");
    const recallScorePrev = getPostVal("recall_score_previous");
    const recallNotesPrev = getPostVal("recall_notes_previous");
    const competitorWordCount = getPostVal("Competitor Word Count");
    const tfidfTerms   = getPostVal("tfidf_terms");
    const tfidfSummary = getPostVal("tfidf_summary");
    const whatChanged  = getPostVal("What Changed");
    const why          = getPostVal("Why");
    const url          = getPostVal("URL");

    return {
      success: true,
      postId, title, stoneType, articleType, locality, url,
      newH1, newHtml, newMetaTitle, newMetaDesc, newSchema, yoastKeyphrase,
      oldTitle, oldDate, oldMetaTitle, oldMetaDesc, oldHtml, oldKeyphrase,
      pqc, intent, primaryTerm, entityStatus, driftStatus, rewriteStatus,
      clicks, impressions, position, recallScore, recallNotes, recallScorePrev, recallNotesPrev,
      competitorWordCount, tfidfTerms, tfidfSummary,
      whatChanged, why
    };

  } catch(e) {
    return { success: false, message: "ce_getW6Data error: " + e.message };
  }
}


function ce_updatePublicationLog(data) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("publication_log");
    if (!logSheet) throw new Error("publication_log sheet not found.");

    const lastCol = logSheet.getLastColumn();
    const headers = lastCol > 0 ? logSheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

    // If no headers exist yet, write them
    if (headers.length === 0 || !headers[0]) {
      const cols = [
        "Log Date","New Published Date","Post ID","Title","URL","Stone Type",
        "Article Type","Locality","Primary Query Cluster Owned",
        "Confirmed Primary Intent","Primary Search Term",
        "Entity Governance Status","Drift Status","Rewrite Status",
        "Clicks","Impressions","Position","recall_score","recall_notes","recall_score_previous","recall_notes_previous","Competitor Word Count","tfidf_terms","tfidf_summary",
        "New H1","New Meta Title","New Meta Description","Yoast Keyphrase",
        "What Changed","Why",
        "Old Title","Old Published Date","Old Meta Title","Old Meta Description",
        "Old Keyphrase","Old HTML","New HTML","New Schema"
      ];
      logSheet.getRange(1, 1, 1, cols.length).setValues([cols]);
    }

    const now = new Date();
    const logDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    const newRow = [
      logDate,
      data.newPublishedDate,
      data.postId,
      data.title,
      data.url,
      data.stoneType,
      data.articleType,
      data.locality,
      data.pqc,
      data.intent,
      data.primaryTerm,
      data.entityStatus,
      data.driftStatus,
      data.rewriteStatus,
      data.clicks,
      data.impressions,
      data.position,
      data.recallScore,
      data.recallNotes,
      data.recallScorePrev,
      data.recallNotesPrev,
      data.competitorWordCount,
      data.tfidfTerms,
      data.tfidfSummary,
      data.newH1,
      data.newMetaTitle,
      data.newMetaDesc,
      data.yoastKeyphrase,
      data.whatChanged,
      data.why,
      data.oldTitle,
      data.oldDate,
      data.oldMetaTitle,
      data.oldMetaDesc,
      data.oldKeyphrase,
      data.oldHtml,
      data.newHtml,
      data.newSchema
    ];

    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);

    // Sort by Log Date descending
    if (lastRow > 1) {
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn())
        .sort({ column: 1, ascending: false });
    }

    logPipelineResume("W6 — Publication Log Updated", "");
    return { success: true, message: "Publication log updated for Post ID " + data.postId };

  } catch(e) {
    return { success: false, message: "ce_updatePublicationLog error: " + e.message };
  }
}

function ce_updateSiteExport() {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const posts    = ss.getSheetByName("posts");
    const exportSh = ss.getSheetByName("site-export");
    if (!posts)    throw new Error("posts sheet not found.");
    if (!exportSh) throw new Error("site-export sheet not found.");

    const row = posts.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const postHeaders = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0];
    function getPostVal(name) {
      const idx = postHeaders.indexOf(name);
      return idx > -1 ? String(posts.getRange(row, idx + 1).getValue() || "").trim() : "";
    }

    const postId    = getPostVal("Post ID");
    const newH1     = getPostVal("New H1");
    const newMetaTitle = getPostVal("New Meta Title");
    const newHtml   = getPostVal("New HTML");
    const schema    = getPostVal("Schema (JSON-LD)");
    const stoneType = getPostVal("Stone Type");

    // Parse New Meta Description and Yoast Keyphrase
    const rawMetaDesc     = getPostVal("New Meta Description");
    const keyphraseMarker = "Yoast Keyphrase:";
    let newMetaDesc    = rawMetaDesc;
    let yoastKeyphrase = "";
    const kpIdx = rawMetaDesc.indexOf(keyphraseMarker);
    if (kpIdx > -1) {
      newMetaDesc    = rawMetaDesc.substring(0, kpIdx).trim();
      yoastKeyphrase = rawMetaDesc.substring(kpIdx + keyphraseMarker.length).trim();
    }

    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    // Find matching row in site-export
    const exportData    = exportSh.getDataRange().getValues();
    const exportHeaders = exportData[0].map(function(h) { return String(h).trim(); });

    const expIdIdx        = exportHeaders.indexOf("ID");
    const expTitleIdx     = exportHeaders.indexOf("Title");
    const expPubDateIdx   = exportHeaders.indexOf("Published Date");
    const expUpdDateIdx   = exportHeaders.indexOf("Updated Date");
    const expKpIdx        = exportHeaders.indexOf("Yoast Keyphrase");
    const expMetaTitleIdx = exportHeaders.indexOf("Meta Title");
    const expMetaDescIdx  = exportHeaders.indexOf("Meta Description");
    const expStoneIdx     = exportHeaders.indexOf("Stone Type");
    const expSchemaIdx    = exportHeaders.indexOf("WPCode Header Schema");
    const expHtmlIdx      = exportHeaders.indexOf("Full Post HTML");

    if (expIdIdx === -1) return { success: false, message: "ID column not found in site-export." };

    let matchRow = -1;
    for (let i = 1; i < exportData.length; i++) {
      if (String(exportData[i][expIdIdx] || "").trim() === String(postId)) {
        matchRow = i + 1; // sheet row is 1-indexed
        break;
      }
    }

    if (matchRow === -1) {
      return { success: false, message: "Post ID " + postId + " not found in site-export." };
    }

    // Write updated values
    function writeIfFound(colIdx, value) {
      if (colIdx > -1) {
        const cell = exportSh.getRange(matchRow, colIdx + 1);
        cell.setNumberFormat("@");
        cell.setValue(value);
      }
    }

    writeIfFound(expTitleIdx,     newH1);
    writeIfFound(expPubDateIdx,   today);
    writeIfFound(expUpdDateIdx,   today);
    writeIfFound(expKpIdx,        yoastKeyphrase);
    writeIfFound(expMetaTitleIdx, newMetaTitle);
    writeIfFound(expMetaDescIdx,  newMetaDesc);
    writeIfFound(expStoneIdx,     stoneType);
    writeIfFound(expSchemaIdx,    schema);
    writeIfFound(expHtmlIdx,      newHtml);

    SpreadsheetApp.flush();

    logPipelineResume("W6 — Site Export Updated", "");
    return { success: true, message: "site-export updated for Post ID " + postId + "." };

  } catch(e) {
    return { success: false, message: "ce_updateSiteExport error: " + e.message };
  }
}
