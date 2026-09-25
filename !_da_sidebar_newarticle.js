// da_Sidebar_NewArticle.gs

function da_openNewArticleSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('da_Sidebar_New_Article')
    .setTitle('New Article AC')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function da_buildNewArticleClassifyPrompt() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('new_article_briefs');
    if (!sheet) return { error: 'new_article_briefs sheet not found.' };

    const activeRow = sheet.getActiveCell().getRow();
    if (activeRow < 2) return { error: 'Please select a data row in new_article_briefs.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(activeRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    const get = (col) => {
      const idx = headers.indexOf(col);
      return idx > -1 ? String(rowData[idx] || '').trim() : '';
    };

    const articleType = get('Article Type');
    const material    = get('Material / Stone Type');

    if (!articleType) return { error: 'Article Type is empty for this row.' };
    if (!material)    return { error: 'Material / Stone Type is empty for this row.' };

    const sheetMap = {
      'Case Study':        'ac-library-case-study',
      'Method Guide':      'ac-library-method-guide',
      'Diagnostic Guide':  'ac-library-diagnostic-guide',
      'Educational Guide': 'ac-library-educational-guide',
      'Buyer Guide':       'ac-library-buyer-guide',
      'Service Page':      'ac-library-service-page',
      'Geo Service Page':  'ac-library-geo-service-page',
      'Hub Page':          'ac-library-hub'
    };

    const acSheetName = sheetMap[articleType];
    if (!acSheetName) return { error: 'No AC library found for article type: ' + articleType };

    const acSheet = ss.getSheetByName(acSheetName);
    if (!acSheet) return { error: 'AC library sheet not found: ' + acSheetName };

    const acData = acSheet.getDataRange().getValues();
    const acHeaders = acData[0];

    const universalCol = acHeaders.indexOf('Universal');
    const materialCol  = acHeaders.indexOf(material);

    const fields = {};
    const fieldLabels = {};

    for (let i = 1; i < acData.length; i++) {
      const row       = acData[i];
      const key       = row[0];
      const label     = row[1];
      const value     = row[2];
      const universal = row[universalCol] === 'y';
      const matMatch  = materialCol > -1 && row[materialCol] === 'y';
      if (!key || !value) continue;
      if (universal || matMatch) {
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

    // Assemble project brief from row
    const brief = [
      'Post ID: '            + get('Post ID'),
      'Article Type: '       + articleType,
      'Location: '           + get('Location'),
      'Material: '           + material,
      'Setting: '            + get('Setting'),
      'Floor Size: '         + get('Floor Size (sqm)') + ' square metres',
      'Client Brief: '       + get('Client Brief'),
      'Condition on Arrival: ' + get('Condition on Arrival'),
      'Process: '            + get('Process Steps'),
      'Products Used: '      + get('Products Used'),
      'Outcome: '            + get('Outcome')
    ].join('\n');

    // Build prompt
    let prompt = `You are an article classifier for a specialist floor care content pipeline.\n\n`;
    prompt += `Article Type: ${articleType}\n`;
    prompt += `Stone Type: ${material}\n\n`;
    prompt += `Read the project brief below and for each field, select the single closest matching option from the list provided.\n\n`;
    prompt += `Output your answer as a JSON array only. No preamble, no explanation, no markdown code fences.\n\n`;
    prompt += `Format:\n`;
    prompt += `[\n`;
    prompt += `  {"index": 0, "field": "[field name]", "selected": "[exact option text]", "reason": "[one sentence]"},\n`;
    prompt += `  ...\n`;
    prompt += `]\n\n`;
    prompt += `The index must match the field number as listed below. Use the exact option text as listed. Do not output anything else.\n\n`;
    prompt += `---\n\nFIELDS AND OPTIONS:\n\n`;

    let fieldIndex = 0;
    fieldOrder.forEach(function(key) {
      if (!fields[key]) return;
      prompt += `FIELD ${fieldIndex}: ${fieldLabels[key]}\n`;
      fields[key].forEach(function(opt) { prompt += `- ${opt}\n`; });
      prompt += `\n`;
      fieldIndex++;
    });

    prompt += `---\n\nPROJECT BRIEF:\n\n${brief}`;

    return {
      prompt:      prompt,
      articleType: articleType,
      material:    material,
      rowNum:      activeRow
    };

  } catch(e) {
    return { error: 'da_buildNewArticleClassifyPrompt error: ' + e.message };
  }
}


function da_saveNewArticleClassification(jsonResponse, rowNum) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('new_article_briefs');
    if (!sheet) return { success: false, message: 'new_article_briefs sheet not found.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Parse JSON
    let parsed;
    try {
      parsed = JSON.parse(jsonResponse);
    } catch(e) {
      return { success: false, message: 'JSON parse error: ' + e.message };
    }

    if (!Array.isArray(parsed)) return { success: false, message: 'Response is not a JSON array.' };

    const fieldOrder = [
      'ac_what_they_found',
      'ac_why_they_called',
      'ac_what_the_floor_was_doing',
      'ac_what_forced_careful_decisions',
      'ac_what_governed_the_approach',
      'ac_what_changed_for_the_homeowner',
      'ac_story_direction'
    ];

    parsed.forEach(function(item) {
      const fieldKey = fieldOrder[item.index];
      if (!fieldKey) return;
      const colIdx = headers.indexOf(fieldKey);
      if (colIdx < 0) return;
      sheet.getRange(rowNum, colIdx + 1).setValue(item.selected);
    });

    // Mark status as Ready if Post ID exists
    const postIdCol = headers.indexOf('Post ID');
    const statusCol = headers.indexOf('Pipeline Status');
    if (postIdCol > -1 && statusCol > -1) {
      const postId = sheet.getRange(rowNum, postIdCol + 1).getValue();
      if (postId) {
        const statusCell = sheet.getRange(rowNum, statusCol + 1);
        statusCell.setValue('Ready');
        statusCell.setBackground('#c6efce').setFontColor('#276221');
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'AC values written to row ' + rowNum + '.' };

  } catch(e) {
    return { success: false, message: 'da_saveNewArticleClassification error: ' + e.message };
  }
}


function da_addNewArticleGapToLibrary(entry, articleType, material) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetMap = {
      'Case Study':        'ac-library-case-study',
      'Method Guide':      'ac-library-method-guide',
      'Diagnostic Guide':  'ac-library-diagnostic-guide',
      'Educational Guide': 'ac-library-educational-guide',
      'Buyer Guide':       'ac-library-buyer-guide',
      'Service Page':      'ac-library-service-page',
      'Geo Service Page':  'ac-library-geo-service-page',
      'Hub Page':          'ac-library-hub'
    };

    const acSheetName = sheetMap[articleType];
    if (!acSheetName) return { success: false, message: 'No AC sheet for: ' + articleType };

    const acSheet = ss.getSheetByName(acSheetName);
    if (!acSheet) return { success: false, message: 'AC sheet not found: ' + acSheetName };

    const acData    = acSheet.getDataRange().getValues();
    const acHeaders = acData[0];

    const dropdownCol  = acHeaders.indexOf('dropdown');
    const labelCol     = acHeaders.indexOf('Label');
    const valueCol     = acHeaders.indexOf('value') > -1 ? acHeaders.indexOf('value') : acHeaders.indexOf('Value');
    const universalCol = acHeaders.indexOf('Universal');
    const materialCol  = acHeaders.indexOf(material);

    // Find dropdown key by matching label
    let dropdownKey = '';
    let fieldLabel  = '';
    for (let i = 1; i < acData.length; i++) {
      if (String(acData[i][labelCol] || '').trim() === entry.field) {
        dropdownKey = String(acData[i][dropdownCol] || '').trim();
        fieldLabel  = entry.field;
        break;
      }
    }

    if (!dropdownKey) return { success: false, message: 'Could not find dropdown key for: ' + entry.field };

    const newRow = new Array(acHeaders.length).fill('');
    newRow[dropdownCol] = dropdownKey;
    newRow[labelCol]    = fieldLabel;
    newRow[valueCol]    = entry.suggested_option;

    if (entry.scope === 'Universal') {
      newRow[universalCol] = 'y';
    } else if (materialCol > -1) {
      newRow[materialCol] = 'y';
    }

    acSheet.appendRow(newRow);
    return { success: true, message: 'Gap added to ' + acSheetName };

  } catch(e) {
    return { success: false, message: 'da_addNewArticleGapToLibrary error: ' + e.message };
  }
}
function da_addPostToSiteExport(rowNum) {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const briefSheet  = ss.getSheetByName('new_article_briefs');
    const exportSheet = ss.getSheetByName('site-export');

    if (!briefSheet)  return { success: false, message: 'new_article_briefs sheet not found.' };
    if (!exportSheet) return { success: false, message: 'site-export sheet not found.' };

    const briefHeaders  = briefSheet.getRange(1, 1, 1, briefSheet.getLastColumn()).getValues()[0];
    const briefRow      = briefSheet.getRange(rowNum, 1, 1, briefSheet.getLastColumn()).getValues()[0];
    const exportHeaders = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn()).getValues()[0];

    const get = (col) => {
      const idx = briefHeaders.indexOf(col);
      return idx > -1 ? String(briefRow[idx] || '').trim() : '';
    };

    const postId = get('Post ID');
    if (!postId) return { success: false, message: 'Post ID is empty. Add the WordPress Post ID before generating the stub.' };

    // Check for existing entry in site-export
    const exportIdCol = exportHeaders.indexOf('ID');
    if (exportIdCol > -1) {
      const exportData = exportSheet.getRange(2, 1, exportSheet.getLastRow() - 1, exportSheet.getLastColumn()).getValues();
      for (let i = 0; i < exportData.length; i++) {
        if (String(exportData[i][exportIdCol]).trim() === String(postId).trim()) {
          return { success: false, message: 'Post ID ' + postId + ' already exists in site-export. No duplicate added.' };
        }
      }
    }

    const articleType = get('Article Type');
    const location    = get('Location');
    const material    = get('Material / Stone Type');
    const setting     = get('Setting');
    const size        = get('Floor Size (sqm)');
    const clientBrief = get('Client Brief');
    const condition   = get('Condition on Arrival');
    const process     = get('Process Steps');
    const products    = get('Products Used');
    const outcome     = get('Outcome');
    const title       = get('Article Title (working)');

    // Build stub HTML
    const stub = `<header>
<p>${material} floor restoration carried out in ${location}. ${size} square metres in ${setting}. ${clientBrief}</p>
</header>

<section id="section-1">
<h2>Condition on arrival</h2>
<p>${condition}</p>
</section>

<section id="section-2">
<h2>Work carried out</h2>
<p>${process}</p>
</section>

<section id="section-3">
<h2>Products used</h2>
<p>${products}</p>
</section>

<section id="section-4">
<h2>Outcome</h2>
<p>${outcome}</p>
</section>`;

    // Build new row for site-export matching its column structure
    const newRow = new Array(exportHeaders.length).fill('');

    const setExportCol = (colName, value) => {
      const idx = exportHeaders.indexOf(colName);
      if (idx > -1) newRow[idx] = value;
    };

    setExportCol('ID',             postId);
    setExportCol('Title',          title);
    setExportCol('Full Post HTML', stub);
    setExportCol('Article Type',   articleType);
    setExportCol('Stone Type',     material);

    exportSheet.appendRow(newRow);

    // Update Pipeline Status in new_article_briefs
    const statusCol = briefHeaders.indexOf('Pipeline Status');
    if (statusCol > -1) {
      const statusCell = briefSheet.getRange(rowNum, statusCol + 1);
      statusCell.setValue('In Pipeline');
      statusCell.setBackground('#cfe2f3').setFontColor('#1c4587');
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Post ID ' + postId + ' added to site-export. Pipeline Status set to In Pipeline.' };

  } catch(e) {
    return { success: false, message: 'da_addPostToSiteExport error: ' + e.message };
  }
}