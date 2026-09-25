// da_new-article-briefs.gs

function onEdit(e) {

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'new_article_briefs') return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 2) return;

  // Trigger on Article Type (col 2) or Material (col 5)
  if (col !== 2 && col !== 5) return;

  const articleType = sheet.getRange(row, 2).getValue();
  const material = sheet.getRange(row, 5).getValue();

  if (!material) return;

  da_applyAcDropdownsToRow(sheet, row, articleType, material);
}


function da_applyAcDropdownsToRow(sheet, rowNum, articleType, material) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const acSheet = ss.getSheetByName('ac-library-case-study');

  if (!acSheet) return;

  const materialColMap = {
    'Victorian Tile':  5,
    'Marble':          6,
    'Travertine':      7,
    'Limestone':       8,
    'Slate':           9,
    'Quarry Tile':    10,
    'Terracotta':     11,
    'Sandstone':      12,
    'Terrazzo':       13,
    'Porcelain Tile': 14,
    'Ceramic Tile':   15,
    'Grout':          16
  };

  const matCol = materialColMap[material];
  if (!matCol) return;

  const acData = acSheet.getDataRange().getValues();

  const acFields = [
    { field: 'ac_what_they_found',                sheetCol: 14 },
    { field: 'ac_why_they_called',                sheetCol: 15 },
    { field: 'ac_what_the_floor_was_doing',       sheetCol: 16 },
    { field: 'ac_what_forced_careful_decisions',  sheetCol: 17 },
    { field: 'ac_what_governed_the_approach',     sheetCol: 18 },
    { field: 'ac_what_changed_for_the_homeowner', sheetCol: 19 },
    { field: 'ac_story_direction',                sheetCol: 20 }
  ];

  acFields.forEach(({ field, sheetCol }) => {

    const values = acData.slice(1)
      .filter(row => row[0] === field && (row[3] === 'y' || row[matCol - 1] === 'y'))
      .map(row => row[2])
      .filter(v => v !== '');

    if (values.length === 0) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(rowNum, sheetCol).setDataValidation(rule);
  });
}

function da_buildValidationPrompt(jsonResponse, rowNum) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('new_article_briefs');
    if (!sheet) return { error: 'new_article_briefs sheet not found.' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

    const get = (col) => {
      const idx = headers.indexOf(col);
      return idx > -1 ? String(rowData[idx] || '').trim() : '';
    };

    const brief = [
      'Post ID: '               + get('Post ID'),
      'Article Type: '          + get('Article Type'),
      'Location: '              + get('Location'),
      'Material: '              + get('Material / Stone Type'),
      'Setting: '               + get('Setting'),
      'Floor Size: '            + get('Floor Size (sqm)') + ' square metres',
      'Client Brief: '          + get('Client Brief'),
      'Condition on Arrival: '  + get('Condition on Arrival'),
      'Process: '               + get('Process Steps'),
      'Products Used: '         + get('Products Used'),
      'Outcome: '               + get('Outcome')
    ].join('\n');

    const fieldLabels = {
      'ac_what_they_found':                'What did the homeowner find when the problem became apparent',
      'ac_why_they_called':                'Why did the homeowner seek professional help',
      'ac_what_the_floor_was_doing':       'What was the floor doing before restoration began',
      'ac_what_forced_careful_decisions':  'What constraint governed every decision on this job',
      'ac_what_governed_the_approach':     'What process emphasis governed this restoration',
      'ac_what_changed_for_the_homeowner': 'What changed for the homeowner after restoration',
      'ac_story_direction':                'What narrative direction does this case study follow'
    };

    // Parse JSON
    let parsed;
    try {
      parsed = JSON.parse(jsonResponse);
    } catch(e) {
      return { error: 'JSON parse error: ' + e.message };
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

    let selectionsBlock = '';
    parsed.forEach(function(item) {
      const key = fieldOrder[item.index];
      const label = fieldLabels[key] || key;
      selectionsBlock += `FIELD ${item.index}: ${label}\nSelected: ${item.selected}\n\n`;
    });

    let prompt = `You are a classification reviewer for a specialist floor care content pipeline.\n\n`;
    prompt += `A project brief has been classified against seven AC fields. Your job is to review each selection and judge whether it is the best fit for the actual job described.\n\n`;
    prompt += `For each field:\n`;
    prompt += `- If the selection is a good fit, confirm it\n`;
    prompt += `- If the selection is a weak or poor fit, flag it and suggest the better option from the original list\n\n`;
    prompt += `Output a JSON array only. No preamble, no explanation, no markdown code fences.\n\n`;
    prompt += `Format:\n`;
    prompt += `[\n`;
    prompt += `  {\n`;
    prompt += `    "index": 0,\n`;
    prompt += `    "field": "[field name]",\n`;
    prompt += `    "selected": "[current selection]",\n`;
    prompt += `    "fit": "good" or "weak" or "poor",\n`;
    prompt += `    "recommended": "[better option if fit is weak or poor, otherwise same as selected]",\n`;
    prompt += `    "reason": "[one sentence explaining your judgement]"\n`;
    prompt += `  }\n`;
    prompt += `]\n\n`;
    prompt += `---\n\nPROJECT BRIEF:\n\n${brief}\n\n`;
    prompt += `---\n\nCURRENT SELECTIONS:\n\n${selectionsBlock}`;

    return { prompt: prompt, rowNum: rowNum };

  } catch(e) {
    return { error: 'da_buildValidationPrompt error: ' + e.message };
  }
}