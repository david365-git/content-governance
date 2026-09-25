/* ============================================================
   gsc_row_stats.gs
   Pulls page-level aggregated GSC stats for the active row
   using a slug-contains filter, and writes results directly
   into the posts sheet columns — only where the new value
   is higher than the existing value (or cell is empty).

   Phase 1 — Last 16 months  → AH(34), AI(35), AJ(36), AK(37)
                                AX(50), AY(51), AZ(52), BA(53),
                                BB(54), BC(55)
   Phase 2 — Months 12 to 6  → BE(57), BF(58), BG(59), BH(60),
                                BI(61), BJ(62)
   Phase 3 — Last 6 months   → BL(64), BM(65), BN(66), BO(67),
                                BP(68), BQ(69)
============================================================ */

function openGSCRowStatsModal() {
  var html = HtmlService.createHtmlOutputFromFile('gsc_row_stats_modal')
    .setWidth(480)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'GSC Row Stats — Active Row');
}

function runGSCRowStatsActiveRow() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName('posts');
  if (!postsSheet) throw new Error("'posts' sheet not found.");

  var activeRow = postsSheet.getActiveRange().getRow();
  if (activeRow < 2) throw new Error("Select a data row in the posts sheet first.");

  // Read URL from column C (3)
  var url = String(postsSheet.getRange(activeRow, 3).getValue() || "").trim();
  if (!url) throw new Error("No URL found in column C for this row.");

  // Extract slug — last path segment including trailing slash
  var slug = _extractSlug(url);
  if (!slug) throw new Error("Could not extract slug from URL: " + url);

  var siteUrl = "https://www.abbeyfloorcare.co.uk/";
  var result  = { slug: slug, phases: [] };

  var allChangedCols = [];

  /* ── PHASE 1 — Last 16 months ── */
  var p1 = _gscFetchPageStats(siteUrl, slug, 16, null);
  var p1dates = _formatDateRange(16, null);
  var p1Changed = _writeStatsToRow(postsSheet, activeRow, [
    { col: 34, val: p1.clicks      },
    { col: 35, val: p1.impressions },
    { col: 36, val: p1.ctr        },
    { col: 37, val: p1.position   },
    { col: 50, val: p1dates        },
    { col: 51, val: '16 months'    },
    { col: 52, val: p1.clicks      },
    { col: 53, val: p1.impressions },
    { col: 54, val: p1.ctr        },
    { col: 55, val: p1.position   }
  ]);
  allChangedCols = allChangedCols.concat(p1Changed);
  result.phases.push({
    label:   'Phase 1 — Last 16 months',
    clicks:  p1.clicks,
    impr:    p1.impressions,
    ctr:     p1.ctr,
    pos:     p1.position,
    dates:   p1dates
  });

  /* ── PHASE 2 — Months 12 to 6 ── */
  var p2 = _gscFetchPageStats(siteUrl, slug, 12, 6);
  var p2dates = _formatDateRange(12, 6);
  var p2Changed = _writeStatsToRow(postsSheet, activeRow, [
    { col: 57, val: p2dates        },
    { col: 58, val: '6 months (12 to 6 months ago)' },
    { col: 59, val: p2.clicks      },
    { col: 60, val: p2.impressions },
    { col: 61, val: p2.ctr        },
    { col: 62, val: p2.position   }
  ]);
  allChangedCols = allChangedCols.concat(p2Changed);
  result.phases.push({
    label:   'Phase 2 — Months 12 to 6',
    clicks:  p2.clicks,
    impr:    p2.impressions,
    ctr:     p2.ctr,
    pos:     p2.position,
    dates:   p2dates
  });

  /* ── PHASE 3 — Last 6 months ── */
  var p3 = _gscFetchPageStats(siteUrl, slug, 6, null);
  var p3dates = _formatDateRange(6, null);
  var p3Changed = _writeStatsToRow(postsSheet, activeRow, [
    { col: 64, val: p3dates        },
    { col: 65, val: '6 months'     },
    { col: 66, val: p3.clicks      },
    { col: 67, val: p3.impressions },
    { col: 68, val: p3.ctr        },
    { col: 69, val: p3.position   }
  ]);
  allChangedCols = allChangedCols.concat(p3Changed);
  result.phases.push({
    label:   'Phase 3 — Last 6 months',
    clicks:  p3.clicks,
    impr:    p3.impressions,
    ctr:     p3.ctr,
    pos:     p3.position,
    dates:   p3dates
  });

  result.changedColumns = allChangedCols;

  return result;
}

/* ============================================================
   HELPER — Fetch aggregated page stats for a slug
   Uses a CONTAINS filter on the slug string.
============================================================ */
function _gscFetchPageStats(siteUrl, slug, monthsBack, endMonthsBack) {
  var today     = new Date();
  var startDate = new Date();
  startDate.setMonth(today.getMonth() - monthsBack);
  var endDate = new Date();
  if (endMonthsBack) endDate.setMonth(today.getMonth() - endMonthsBack);

  var startStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var endStr   = Utilities.formatDate(endDate,   Session.getScriptTimeZone(), "yyyy-MM-dd");

  var apiUrl  = "https://www.googleapis.com/webmasters/v3/sites/" +
                encodeURIComponent(siteUrl) + "/searchAnalytics/query";
  var payload = {
    startDate:  startStr,
    endDate:    endStr,
    dimensions: ["page"],
    dimensionFilterGroups: [{
      filters: [
        { dimension: "page", operator: "contains", expression: slug }
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

  // Aggregate across all matching pages (catches all URL variants)
  var totalClicks = 0, totalImpr = 0, totalCtr = 0, totalPos = 0, count = 0;
  rows.forEach(function(r) {
    totalClicks += r.clicks      || 0;
    totalImpr   += r.impressions || 0;
    totalCtr    += r.ctr         || 0;
    totalPos    += r.position    || 0;
    count++;
  });

  return {
    clicks:      totalClicks,
    impressions: totalImpr,
    ctr:         count > 0 ? Math.round((totalCtr / count) * 10000) / 10000 : 0,
    position:    count > 0 ? Math.round((totalPos / count) * 100)  / 100   : 0
  };
}

/* ============================================================
   HELPER — Write stats to the active row
   Only writes where the new value is higher than existing,
   or the cell is empty. Skips date/timescale columns —
   those always overwrite since they describe the fetch window.
============================================================ */
function _writeStatsToRow(sheet, row, entries) {
  var DATE_COLS = [50, 51, 57, 58, 64, 65]; // always write dates/timescales

  entries.forEach(function(e) {
    var cell     = sheet.getRange(row, e.col);
    var existing = cell.getValue();
    if (DATE_COLS.indexOf(e.col) > -1) {
      cell.setValue(e.val);
    } else {
      var existingNum = parseFloat(existing) || 0;
      var newNum      = parseFloat(e.val)    || 0;
      if (!existing || newNum > existingNum) {
        cell.setValue(e.val);
      }
    }
  });
}

/* ============================================================
   HELPER — Extract slug from full URL
   e.g. https://www.abbeyfloorcare.co.uk/home-garden/slate/
   returns /slate/
============================================================ */
function _extractSlug(url) {
  var match = url.match(/\/([^\/]+\/?)$/);
  return match ? '/' + match[1].replace(/^\//, '') : null;
}

/* ============================================================
   HELPER — Format date range string
   e.g. "01/01/2025 — 01/05/2026"
============================================================ */
function _formatDateRange(monthsBack, endMonthsBack) {
  var today     = new Date();
  var startDate = new Date();
  startDate.setMonth(today.getMonth() - monthsBack);
  var endDate = new Date();
  if (endMonthsBack) endDate.setMonth(today.getMonth() - endMonthsBack);

  function fmt(d) {
    var dd = String(d.getDate()).padStart(2, '0');
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    var yy = d.getFullYear();
    return dd + '/' + mm + '/' + yy;
  }
  return fmt(startDate) + ' — ' + fmt(endDate);
}

/* ============================================================
   COMBINED — Run GSC Full Sync then GSC Row Stats in sequence
   for the active row, in one menu click.
============================================================ */
function runGSCFullSyncThenRowStats() {
  var ui = SpreadsheetApp.getUi();
  try {
    var syncResult = runGSCFullSyncActiveRow();
  } catch (e) {
    ui.alert('GSC Full Sync failed — Row Stats was not run.\n\n' + e.message);
    throw e;
  }

  try {
    var statsResult = runGSCRowStatsActiveRow();
  } catch (e) {
    ui.alert('GSC Full Sync completed, but GSC Row Stats failed.\n\n' + e.message);
    throw e;
  }

  ui.alert('✅ GSC Full Sync and Row Stats both completed for Post ID ' + syncResult.postId + '.');

  return { sync: syncResult, stats: statsResult };
}