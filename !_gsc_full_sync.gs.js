/* ============================================================
   gsc_full_sync.gs
   Full GSC sync for the active row — all four destination sheets
   in one automated sequence.

   Sequence:
   Phase 1 — Last 16 months  → GSC Queries + GSC Baseline
   Phase 2 — Months 12 to 6  → GSC Benchmark
   Phase 3 — Last 6 months   → GSC Monitor
============================================================ */

function openGSCFullSyncModal() {
  var html = HtmlService.createHtmlOutputFromFile('gsc_full_sync_modal')
    .setWidth(480)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, 'GSC Full Sync — Active Row');
}

function runGSCFullSyncActiveRow() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName('posts');
  var liveSheet  = ss.getSheetByName('GSC_Live');

  if (!postsSheet) throw new Error("'posts' sheet not found.");
  if (!liveSheet)  throw new Error("'GSC_Live' sheet not found.");

  // Read active row
  var activeRow = postsSheet.getActiveRange().getRow();
  if (activeRow < 2) throw new Error("Select a data row in the posts sheet first.");

  var headers   = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];
  var urlIdx    = headers.indexOf("URL");
  var postIdIdx = headers.indexOf("Post ID");
  if (urlIdx === -1)    throw new Error("URL column not found in posts sheet.");
  if (postIdIdx === -1) throw new Error("Post ID column not found in posts sheet.");

  var siteUrl  = "https://www.abbeyfloorcare.co.uk/";
  var url      = String(postsSheet.getRange(activeRow, urlIdx + 1).getValue()    || "").trim();
  var postId   = String(postsSheet.getRange(activeRow, postIdIdx + 1).getValue() || "").trim();

  if (!url)    throw new Error("No URL found in the active row.");
  if (!postId) throw new Error("No Post ID found in the active row.");

  var result = {
    postId:  postId,
    phases:  [],
    success: true
  };

  /* ── PHASE 1 — Last 16 months → GSC Queries + GSC Baseline ── */
  var phase1Rows = _gscFetchSingleUrl(siteUrl, url, postId, 16, null);
  _gscClearLive(liveSheet);
  _gscWriteLive(liveSheet, phase1Rows);
  var q1 = _gscMergeToSheet(ss, 'GSC Queries', postId, phase1Rows);
  var q2 = _gscMergeToSheet(ss, 'GSC Baseline', postId, phase1Rows);
  result.phases.push({
    label:   'Phase 1 — Last 16 months',
    sheets:  [
      { name: 'GSC Queries',  added: q1.added, total: q1.total },
      { name: 'GSC Baseline', added: q2.added, total: q2.total }
    ],
    fetched: phase1Rows.length
  });

  /* ── PHASE 2 — Months 12 to 6 → GSC Benchmark ── */
  var phase2Rows = _gscFetchSingleUrl(siteUrl, url, postId, 12, 6);
  _gscClearLive(liveSheet);
  _gscWriteLive(liveSheet, phase2Rows);
  var q3 = _gscMergeToSheet(ss, 'GSC Benchmark', postId, phase2Rows);
  result.phases.push({
    label:   'Phase 2 — Months 12 to 6',
    sheets:  [
      { name: 'GSC Benchmark', added: q3.added, total: q3.total }
    ],
    fetched: phase2Rows.length
  });

  /* ── PHASE 3 — Last 6 months → GSC Monitor ── */
  var phase3Rows = _gscFetchSingleUrl(siteUrl, url, postId, 6, null);
  _gscClearLive(liveSheet);
  _gscWriteLive(liveSheet, phase3Rows);
  var q4 = _gscMergeToSheet(ss, 'GSC Monitor', postId, phase3Rows);
  result.phases.push({
    label:   'Phase 3 — Last 6 months',
    sheets:  [
      { name: 'GSC Monitor', added: q4.added, total: q4.total }
    ],
    fetched: phase3Rows.length
  });

  /* ── FINAL CLEAR ── */
  _gscClearLive(liveSheet);

  return result;
}

/* ============================================================
   HELPER — Fetch GSC data for a single URL
   monthsBack: how many months back to start
   endMonthsBack: how many months back to end (null = today)
============================================================ */
function _gscFetchSingleUrl(siteUrl, url, postId, monthsBack, endMonthsBack) {
  var today     = new Date();
  var startDate = new Date();
  startDate.setMonth(today.getMonth() - monthsBack);
  var endDate = endMonthsBack ? new Date() : new Date();
  if (endMonthsBack) endDate.setMonth(today.getMonth() - endMonthsBack);

  var startDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var endDateStr   = Utilities.formatDate(endDate,   Session.getScriptTimeZone(), "yyyy-MM-dd");

  var apiUrl  = "https://www.googleapis.com/webmasters/v3/sites/" + encodeURIComponent(siteUrl) + "/searchAnalytics/query";
  var payload = {
    startDate: startDateStr,
    endDate:   endDateStr,
    dimensions: ["query", "page"],
    dimensionFilterGroups: [{
      filters: [
        { dimension: "page", operator: "equals", expression: url }
      ]
    }],
    rowLimit: 25000
  };
  var options = {
    method:      "post",
    contentType: "application/json",
    headers:     { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error("GSC API Error: " + response.getContentText());
  }

  var rows = JSON.parse(response.getContentText()).rows || [];
  return rows.map(function(r) {
    return [postId, r.keys[0], r.clicks, r.impressions, url];
  });
}

/* ============================================================
   HELPER — Clear GSC_Live from row 2 down
============================================================ */
function _gscClearLive(liveSheet) {
  var last = liveSheet.getLastRow();
  if (last > 1) {
    liveSheet.getRange(2, 1, last - 1, 5).clearContent();
  }
}

/* ============================================================
   HELPER — Write rows to GSC_Live
============================================================ */
function _gscWriteLive(liveSheet, rows) {
  if (rows.length === 0) return;
  liveSheet.getRange(2, 1, rows.length, 5).setValues(rows);
}

/* ============================================================
   HELPER — Merge rows into a destination sheet
   Retains all existing rows for other Post IDs.
   For the current Post ID, only adds rows where the
   Query does not already exist (matched on Post ID + Query).
============================================================ */
function _gscMergeToSheet(ss, sheetName, postId, newRows) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet '" + sheetName + "' not found.");

  var last     = sheet.getLastRow();
  var existing = last > 1 ? sheet.getRange(2, 1, last - 1, 5).getValues() : [];

  // Build a set of existing Post ID + Query combinations for this postId
  var existingKeys = {};
  existing.forEach(function(r) {
    if (String(r[0]).trim() === postId) {
      existingKeys[String(r[1]).trim().toLowerCase()] = true;
    }
  });

  // Only add rows whose query isn't already present
  var toAdd = newRows.filter(function(r) {
    return !existingKeys[String(r[1]).trim().toLowerCase()];
  });

  if (toAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 5).setValues(toAdd);
  }

  return {
    added: toAdd.length,
    total: existing.filter(function(r) { return String(r[0]).trim() === postId; }).length + toAdd.length
  };
}