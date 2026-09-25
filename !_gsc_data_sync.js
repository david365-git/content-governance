/**
 * Synchronizes Google Search Console data from 'GSC_Live' to a user-specified target sheet.
 * Appends new queries and conditionally updates Clicks/Impressions if live data is higher.
 * File Name: GSC_Data_Sync.gs
 */
function syncGSCData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // 1. Prompt the user ONLY for the target sheet name
  var promptResponse = ui.prompt(
    'Select Target Sheet',
    'Enter the exact name of the sheet you want to update:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (promptResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Sync cancelled.');
    return;
  }
  
  var targetSheetName = promptResponse.getResponseText().trim();
  if (!targetSheetName) {
    ui.alert('Error: Sheet name cannot be empty.');
    return;
  }
  
  // 2. Fetch the sheets
  var mainSheet = ss.getSheetByName(targetSheetName);
  var liveSheet = ss.getSheetByName("GSC_Live");
  
  if (!mainSheet) {
    ui.alert("Error: The target sheet '" + targetSheetName + "' could not be found.");
    return;
  }
  if (!liveSheet) {
    ui.alert("Error: The donor sheet 'GSC_Live' could not be found.");
    return;
  }
  
  // Load data into memory
  var mainRange = mainSheet.getDataRange();
  var mainValues = mainRange.getValues();
  var mainHeaders = mainValues[0];
  
  var liveRange = liveSheet.getDataRange();
  var liveValues = liveRange.getValues();
  var liveHeaders = liveValues[0];
  
  // Normalize headers
  var cleanMainHeaders = mainHeaders.map(function(h) { return String(h).toLowerCase().trim(); });
  var cleanLiveHeaders = liveHeaders.map(function(h) { return String(h).toLowerCase().trim(); });
  
  // Identify key metric columns
  var mainPostIdIdx = cleanMainHeaders.indexOf("post id");
  var mainQueryIdx = cleanMainHeaders.indexOf("query");
  var mainClicksIdx = cleanMainHeaders.indexOf("clicks");
  var mainImpressionsIdx = cleanMainHeaders.indexOf("impressions");
  
  var livePostIdIdx = cleanLiveHeaders.indexOf("post id");
  var liveQueryIdx = cleanLiveHeaders.indexOf("query");
  var liveClicksIdx = cleanLiveHeaders.indexOf("clicks");
  var liveImpressionsIdx = cleanLiveHeaders.indexOf("impressions");
  
  if (mainPostIdIdx === -1 || mainQueryIdx === -1 || mainClicksIdx === -1 || mainImpressionsIdx === -1 ||
      livePostIdIdx === -1 || liveQueryIdx === -1 || liveClicksIdx === -1 || liveImpressionsIdx === -1) {
    ui.alert("Error: Could not find required headers ('Post ID', 'Query', 'Clicks', 'Impressions') in both sheets.");
    return;
  }
  
  // Map main headers to live headers
  var headerMapping = [];
  for (var i = 0; i < cleanMainHeaders.length; i++) {
    headerMapping.push(cleanLiveHeaders.indexOf(cleanMainHeaders[i]));
  }
  
  // Build a map of existing records in the target sheet
  var mainMap = {};
  for (var r = 1; r < mainValues.length; r++) {
    var key = String(mainValues[r][mainPostIdIdx]).trim() + "|" + String(mainValues[r][mainQueryIdx]).trim();
    mainMap[key] = r;
  }
  
  var newRowsCount = 0;
  var updatedClicksCount = 0;
  var updatedImpressionsCount = 0;
  
  // Process rows
  for (var l = 1; l < liveValues.length; l++) {
    var liveRow = liveValues[l];
    var livePostId = String(liveRow[livePostIdIdx]).trim();
    var liveQuery = String(liveRow[liveQueryIdx]).trim();
    
    if (!livePostId && !liveQuery) continue;
    
    var liveKey = livePostId + "|" + liveQuery;
    var liveClicks = Number(liveRow[liveClicksIdx]) || 0;
    var liveImpressions = Number(liveRow[liveImpressionsIdx]) || 0;
    
    if (mainMap.hasOwnProperty(liveKey)) {
      var mainRowIdx = mainMap[liveKey];
      
      // Update Clicks if larger
      var currentClicks = Number(mainValues[mainRowIdx][mainClicksIdx]) || 0;
      if (liveClicks > currentClicks) {
        mainValues[mainRowIdx][mainClicksIdx] = liveClicks;
        updatedClicksCount++;
      }
      
      // Update Impressions if larger
      var currentImpressions = Number(mainValues[mainRowIdx][mainImpressionsIdx]) || 0;
      if (liveImpressions > currentImpressions) {
        mainValues[mainRowIdx][mainImpressionsIdx] = liveImpressions;
        updatedImpressionsCount++;
      }
    } else {
      // Assemble new row
      var newRow = new Array(mainHeaders.length);
      for (var h = 0; h < mainHeaders.length; h++) {
        var liveColIdx = headerMapping[h];
        newRow[h] = (liveColIdx !== -1) ? liveRow[liveColIdx] : "";
      }
      
      mainValues.push(newRow);
      mainMap[liveKey] = mainValues.length - 1;
      newRowsCount++;
    }
  }
  
  // Write data back to target sheet
  mainSheet.getRange(1, 1, mainValues.length, mainHeaders.length).setValues(mainValues);
  
  ui.alert(
    "Sync Completed Successfully!\n\n" +
    "Target Sheet: '" + targetSheetName + "'\n" +
    "• New queries added: " + newRowsCount + "\n" +
    "• Clicks updated: " + updatedClicksCount + "\n" +
    "• Impressions updated: " + updatedImpressionsCount
  );
}