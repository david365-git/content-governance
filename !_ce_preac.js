/* ============================================================
   ce_PreAC.gs
   Pre-AC Gap Check functions.
   Can be deleted entirely if the Pre-AC feature is abandoned.
============================================================ */

function preAcBuildGapCheckPrompt() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const postsSheet = ss.getSheetByName('posts');
    const row = ss.getActiveSheet().getActiveCell().getRow();
    const headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];

    const articleTypeCol = headers.indexOf('Article Type');
    const stoneTypeCol = headers.indexOf('Stone Type');
    const rowData = postsSheet.getRange(row, 1, 1, postsSheet.getLastColumn()).getValues()[0];

    const articleType = rowData[articleTypeCol];
    const stoneType = rowData[stoneTypeCol];

    const sheetMap = {
      "Case Study":        "ac-library-case-study",
      "Method Guide":      "ac-library-method-guide",
      "Diagnostic Guide":  "ac-library-diagnostic-guide",
      "Educational Guide": "ac-library-educational-guide",
      "Buyer Guide":       "ac-library-buyer-guide",
      "Service Page":      "ac-library-service-page",
      "Geo Service Page":  "ac-library-geo-service-page",
      "Hub Page":          "ac-library-hub"
    };

    const acSheetName = sheetMap[articleType];
    if (!acSheetName) return { error: 'No AC sheet found for article type: ' + articleType };

    const acSheet = ss.getSheetByName(acSheetName);
    if (!acSheet) return { error: 'AC sheet not found: ' + acSheetName };

    const data = acSheet.getDataRange().getValues();
    const headerRow = data[0];

    const universalCol = headerRow.indexOf('Universal');
    const stoneCol = headerRow.indexOf(stoneType);

    const fields = {};
    const fieldLabels = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const key = row[0];
      const label = row[1];
      const value = row[2];
      const isUniversal = row[universalCol] === 'y';
      const isStone = stoneCol > -1 && row[stoneCol] === 'y';
      if (!key || !value) continue;
      if (isUniversal || isStone) {
        if (!fields[key]) { fields[key] = []; fieldLabels[key] = label; }
        if (!fields[key].includes(value)) fields[key].push(value);
      }
    }

    const fieldOrder = [
      'ac_what_they_found',
      'ac_why_they_called',
      'ac_what_the_floor_was_doing',
      'ac_what_forced_careful_decisions',
      'ac_what_governed_the_approach',
      'ac_what_changed_for_the_homeowner',
      'ac_story_direction'
    ];

    // Fetch article content
    const exportSheet = ss.getSheetByName('site-export');
    let articleContent = '[Article content not found]';
    if (exportSheet) {
      const exportHeaders = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn()).getValues()[0];
      const exportIdCol = exportHeaders.indexOf('ID');
      const exportHtmlCol = exportHeaders.indexOf('Full Post HTML');
      const postIdCol = headers.indexOf('Post ID');
      const postId = String(rowData[postIdCol]).trim();
      if (exportIdCol > -1 && exportHtmlCol > -1 && postId) {
        const exportData = exportSheet.getRange(2, 1, exportSheet.getLastRow() - 1, exportSheet.getLastColumn()).getValues();
        for (let i = 0; i < exportData.length; i++) {
          if (String(exportData[i][exportIdCol]).trim() === postId) {
            articleContent = exportData[i][exportHtmlCol];
            break;
          }
        }
      }
    }

    let prompt = `You are reviewing an article classification system for a specialist floor care content pipeline.\n\n`;
    prompt += `Article Type: ${articleType}\n`;
    prompt += `Stone Type: ${stoneType}\n\n`;
    prompt += `Read the article below. For each field, review the existing options and judge whether any existing option adequately describes what the article is actually about.\n\n`;
    prompt += `If no existing option fits well, flag a gap and suggest a new option.\n\n`;
    prompt += `Output a JSON array only. No preamble, no explanation, no markdown code fences, no backticks. Start your response with [ and end with ].\n\n`;
    prompt += `For each field output one entry in this exact format:\n`;
    prompt += `[\n`;
    prompt += `  {\n`;
    prompt += `    "index": 0,\n`;
    prompt += `    "field": "[field label]",\n`;
    prompt += `    "gap": true or false,\n`;
    prompt += `    "closest_existing": "[the closest existing option even if imperfect]",\n`;
    prompt += `    "suggested_option": "[your suggested new option wording, or empty string if no gap]",\n`;
    prompt += `    "scope": "Universal" or "[Stone Type]",\n`;
    prompt += `    "scope_reason": "[one sentence explaining why this applies universally or only to this stone type]",\n`;
    prompt += `    "gap_reason": "[one sentence explaining why no existing option fits, or empty string if no gap]"\n`;
    prompt += `  }\n`;
    prompt += `]\n\n`;
    prompt += `Only set gap to true if none of the existing options adequately describe the article. If an existing option fits reasonably well, set gap to false.\n\n`;
    prompt += `---\n\nFIELDS AND EXISTING OPTIONS:\n\n`;

    let fieldIndex = 0;
    fieldOrder.forEach(function(key) {
      if (!fields[key]) return;
      prompt += `FIELD ${fieldIndex}: ${fieldLabels[key]}\n`;
      fields[key].forEach(function(opt) { prompt += `- ${opt}\n`; });
      prompt += `\n`;
      fieldIndex++;
    });

    prompt += `---\n\nARTICLE CONTENT:\n\n${articleContent}`;

    return { prompt: prompt, articleType: articleType, stoneType: stoneType };

  } catch(e) {
    return { error: 'preAcBuildGapCheckPrompt error: ' + e.message };
  }
}


function preAcAddGapToSheet(entry) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const postsSheet = ss.getSheetByName('posts');
    const row = ss.getActiveSheet().getActiveCell().getRow();
    const headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];

    const articleTypeCol = headers.indexOf('Article Type');
    const stoneTypeCol = headers.indexOf('Stone Type');
    const rowData = postsSheet.getRange(row, 1, 1, postsSheet.getLastColumn()).getValues()[0];

    const articleType = rowData[articleTypeCol];
    const stoneType = rowData[stoneTypeCol];

    const sheetMap = {
      "Case Study":        "ac-library-case-study",
      "Method Guide":      "ac-library-method-guide",
      "Diagnostic Guide":  "ac-library-diagnostic-guide",
      "Educational Guide": "ac-library-educational-guide",
      "Buyer Guide":       "ac-library-buyer-guide",
      "Service Page":      "ac-library-service-page",
      "Geo Service Page":  "ac-library-geo-service-page",
      "Hub Page":          "ac-library-hub"
    };

    const acSheetName = sheetMap[articleType];
    if (!acSheetName) return { success: false, message: 'No AC sheet found for article type: ' + articleType };

    const acSheet = ss.getSheetByName(acSheetName);
    if (!acSheet) return { success: false, message: 'AC sheet not found: ' + acSheetName };

    const data = acSheet.getDataRange().getValues();
    const headerRow = data[0];

    const universalCol = headerRow.indexOf('Universal');
    const stoneCol = headerRow.indexOf(stoneType);
    const dropdownCol = headerRow.indexOf('dropdown');
    const labelCol = headerRow.indexOf('Label');
    const valueCol = headerRow.indexOf('value') > -1 ? headerRow.indexOf('value') : headerRow.indexOf('Value');

    // Find the dropdown key for this field by matching the label
    let dropdownKey = '';
    let fieldLabel = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][labelCol] || '').trim() === entry.field) {
        dropdownKey = String(data[i][dropdownCol] || '').trim();
        fieldLabel = entry.field;
        break;
      }
    }

    if (!dropdownKey) return { success: false, message: 'Could not find dropdown key for field: ' + entry.field };

    // Build new row matching the sheet structure
    const newRow = new Array(headerRow.length).fill('');
    newRow[dropdownCol] = dropdownKey;
    newRow[labelCol] = fieldLabel;
    newRow[valueCol] = entry.suggested_option;

    if (entry.scope === 'Universal') {
      newRow[universalCol] = 'y';
    } else if (stoneCol > -1) {
      newRow[stoneCol] = 'y';
    }

    acSheet.appendRow(newRow);

    logPipelineResume("Pre-AC — Gap Added to Sheet", "");
    return { success: true, message: 'Added to ' + acSheetName };

  } catch(e) {
    return { success: false, message: 'preAcAddGapToSheet error: ' + e.message };
  }
}
