/**
 * Opens the HTML Modal for user timescale input
 */
function openGSCTimescaleModal() {
  var html = HtmlService.createHtmlOutputFromFile('gsc_timescale_modal')
    .setWidth(420)
    .setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sync Search Console Data');
}

/**
 * Core Processing Function triggered by the HTML modal interface
 */
function processGSCFetch(months) {
  // =========================================================================
  // USER PROPERTY URL CONFIGURATION
  var siteUrl = "https://www.abbeyfloorcare.co.uk/"; 
  // =========================================================================
  
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GSC_Live");
  var postsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("posts");
  
  if (!targetSheet) throw new Error("Tab named 'GSC_Live' could not be found.");
  if (!postsSheet) throw new Error("Tab named 'posts' could not be found.");

  // SINGLE URL MODE
  if (months === "single") {
    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return "Error: Select a data row in the posts sheet first.";
    var headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];
    var urlIdx = headers.indexOf("URL");
    var postIdIdx = headers.indexOf("Post ID");
    if (urlIdx === -1) urlIdx = headers.indexOf("Canonical URL");
    if (urlIdx === -1) return "Error: URL column not found in posts sheet.";
    var singleUrl = String(postsSheet.getRange(activeRow, urlIdx + 1).getValue() || "").trim();
    var singlePostId = postIdIdx > -1 ? String(postsSheet.getRange(activeRow, postIdIdx + 1).getValue() || "").trim() : "";
    if (!singleUrl) return "Error: No URL found in the active row.";

    var today = new Date();
    var startDate = new Date();
    startDate.setMonth(today.getMonth() - 6);
    var startDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var endDateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

    var apiUrl = "https://www.googleapis.com/webmasters/v3/sites/" + encodeURIComponent(siteUrl) + "/searchAnalytics/query";
    var payload = {
      "startDate": startDateStr,
      "endDate": endDateStr,
      "dimensions": ["query", "page"],
      "dimensionFilterGroups": [{
        "filters": [
          { "dimension": "page", "operator": "equals", "expression": singleUrl }
        ]
      }],
      "rowLimit": 25000
    };
    var options = {
      "method": "post",
      "contentType": "application/json",
      "headers": { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    var response = UrlFetchApp.fetch(apiUrl, options);
    if (response.getResponseCode() !== 200) throw new Error("GSC API Error: " + response.getContentText());
    var json = JSON.parse(response.getContentText());
    var rows = json.rows || [];

    // Remove existing rows for this post ID from GSC_Live
    var lastTargetRow = targetSheet.getLastRow();
    if (lastTargetRow > 1) {
      var existingData = targetSheet.getRange(2, 1, lastTargetRow - 1, 5).getValues();
      for (var r = existingData.length - 1; r >= 0; r--) {
        if (String(existingData[r][0]).trim() === singlePostId) {
          targetSheet.deleteRow(r + 2);
        }
      }
    }

    // Append new rows for this URL
    var outputData = [];
    for (var i = 0; i < rows.length; i++) {
      outputData.push([singlePostId, rows[i].keys[0], rows[i].clicks, rows[i].impressions, singleUrl]);
    }
    if (outputData.length > 0) {
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, outputData.length, 5).setValues(outputData);
    }
    return "Single URL sync complete. Found " + outputData.length + " keyword rows for this page.";
  }

  // 1. Extract target URLs (Col C) and Post IDs (Col D) from the 'posts' sheet
  var postsLastRow = postsSheet.getLastRow();
  if (postsLastRow < 2) {
    return "Error: No URLs found in your 'posts' sheet to match against.";
  }
  
  var postsData = postsSheet.getRange(2, 3, postsLastRow - 1, 2).getValues();
  var urlToPostIdMap = {};
  var urlToOriginalUrlMap = {}; 
  
  for (var i = 0; i < postsData.length; i++) {
    var urlFromSheet = postsData[i][0].toString().trim();
    var postIdFromSheet = postsData[i][1].toString().trim();
    
    if (urlFromSheet) {
      var normUrl = normalizeUrl(urlFromSheet);
      urlToPostIdMap[normUrl] = postIdFromSheet;
      urlToOriginalUrlMap[normUrl] = urlFromSheet; 
    }
  }
  
  // 2. Set up date constraints dynamically
  var today = new Date();
  var startDate = new Date();
  var endDate = new Date();
  
  if (months === "6_from_12") {
    // 6 months starting 12 months ago (ends 6 months ago)
    startDate.setMonth(today.getMonth() - 12);
    endDate.setMonth(today.getMonth() - 6);
  } else {
    // Handles exact choices: last 6 months or last 16 months
    startDate.setMonth(today.getMonth() - Number(months));
    // endDate remains today
  }
  
  var startDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var endDateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // =========================================================================
  // PAGINATED API FETCH LOOP WITH UK GEOGRAPHIC FILTER
  // =========================================================================
  var apiUrl = "https://www.googleapis.com/webmasters/v3/sites/" + encodeURIComponent(siteUrl) + "/searchAnalytics/query";
  var allGscRows = [];
  var startRow = 0;
  var pageSize = 25000;
  var maxRowsToFetch = 150000; 
  
  while (startRow < maxRowsToFetch) {
    var payload = {
      "startDate": startDateStr,
      "endDate": endDateStr,
      "dimensions": ["query", "page"],
      "startRow": startRow,
      "rowLimit": pageSize
    };
    
    var options = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseCode = response.getResponseCode();
    var resText = response.getContentText();
    
    if (responseCode !== 200) {
      throw new Error("GSC API Error (" + responseCode + "): " + resText);
    }
    
    var json = JSON.parse(resText);
    var rows = json.rows;
    
    if (!rows || rows.length === 0) {
      break; 
    }
    
    allGscRows = allGscRows.concat(rows);
    
    if (rows.length < pageSize) {
      break; 
    }
    
    startRow += pageSize; 
  }
  
  // 4. Intercept the complete dataset and match against your target URLs
  var outputData = [];
  for (var i = 0; i < allGscRows.length; i++) {
    var row = allGscRows[i];
    var query = row.keys[0];
    var page = row.keys[1];
    
    var normGscUrl = normalizeUrl(page);
    
    if (urlToPostIdMap.hasOwnProperty(normGscUrl)) {
      var postId = urlToPostIdMap[normGscUrl];
      var cleanDisplayUrl = urlToOriginalUrlMap[normGscUrl]; 
      var clicks = row.clicks;
      var impressions = row.impressions;
      
      outputData.push([postId, query, clicks, impressions, cleanDisplayUrl]);
    }
  }
  
  // 5. Wipes the landing pad clean from Row 2 down across columns A to E
  var lastTargetRow = targetSheet.getLastRow();
  if (lastTargetRow > 1) {
    targetSheet.getRange(2, 1, lastTargetRow - 1, 5).clearContent();
  }
  
  // 6. Output filtered target rows
  if (outputData.length > 0) {
    targetSheet.getRange(2, 1, outputData.length, 5).setValues(outputData);
    return "Success! Exhaustive UK-only scan complete. Found and pushed " + outputData.length + " keyword rows matching your targeted pages.";
  } else {
    return "The script executed perfectly, but none of the URLs listed in your 'posts' tab received UK-based search impressions during this specific timescale.";
  }
}

/**
 * URL Normalization Helper
 */
function normalizeUrl(url) {
  if (!url) return "";
  var clean = url.toString().toLowerCase().trim();
  clean = clean.replace(/^https?:\/\//, "");
  clean = clean.replace(/^www\./, "");
  if (clean.endsWith("/")) {
    clean = clean.slice(0, -1);
  }
  return clean;
}