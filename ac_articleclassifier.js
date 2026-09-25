function classifyDonorArticle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName('posts');
  const exportSheet = ss.getSheetByName('site-export');
  const ui = SpreadsheetApp.getUi();

  // Get active row
  const activeRow = postsSheet.getActiveCell().getRow();
  if (activeRow < 2) {
    ui.alert('Please select a row in the posts sheet first.');
    return;
  }

  // Get headers from posts sheet
  const postsHeaders = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];

  // Get Post ID from active row
  const postIdCol = postsHeaders.indexOf('Post ID') + 1;
  if (postIdCol === 0) {
    ui.alert('Post ID column not found in posts sheet.');
    return;
  }
  const postId = postsSheet.getRange(activeRow, postIdCol).getValue();
  if (!postId) {
    ui.alert('Post ID is empty for this row.');
    return;
  }

  // Get donor article from site-export sheet by matching Post ID to ID
  const exportHeaders = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn()).getValues()[0];
  const exportIdCol = exportHeaders.indexOf('ID') + 1;
  const exportHtmlCol = exportHeaders.indexOf('Full Post HTML') + 1;
  if (exportIdCol === 0 || exportHtmlCol === 0) {
    ui.alert('ID or Full Post HTML column not found in site-export sheet.');
    return;
  }

  const exportData = exportSheet.getRange(2, 1, exportSheet.getLastRow() - 1, exportSheet.getLastColumn()).getValues();
  let donorArticle = '';
  for (let i = 0; i < exportData.length; i++) {
    if (String(exportData[i][exportIdCol - 1]) === String(postId)) {
      donorArticle = exportData[i][exportHtmlCol - 1];
      break;
    }
  }

  if (!donorArticle) {
    ui.alert('No matching donor article found in site-export sheet for Post ID: ' + postId);
    return;
  }

  // Assemble governed brief from posts sheet columns
  const getValue = (header) => {
    const col = postsHeaders.indexOf(header) + 1;
    return col > 0 ? postsSheet.getRange(activeRow, col).getValue() : '';
  };

  const governedBrief = [
    'Article Type: ' + getValue('Article Type'),
    'Primary Entity: ' + getValue('Primary Entity'),
    'Material Entity: ' + getValue('Material Entity'),
    'Confirmed Primary Intent: ' + getValue('Confirmed Primary Intent'),
    'Supporting Entities Core: ' + getValue('Supporting Entities Core'),
    'Rewrite Governance Summary: ' + getValue('Rewrite Governance Summary'),
    'Page Rewrite Brief: ' + getValue('Page Rewrite Brief'),
    'Cannibalisation Guardrail: ' + getValue('Cannibalisation Guardrail')
  ].join('\n');

  // Build classification prompt
  const prompt = buildClassificationPrompt(donorArticle, governedBrief);

  // Call LLM
  const response = callLLMForClassification(prompt);
  if (!response) {
    ui.alert('LLM call failed. Please try again.');
    return;
  }

  // Parse XML tags from response
  const parsed = parseClassificationResponse(response);
  if (!parsed) {
    ui.alert('Classification response could not be parsed. Please try again.');
    return;
  }

  // Write to posts sheet columns
  writeClassificationToSheet(postsSheet, activeRow, postsHeaders, parsed);

  ui.alert('Classification complete. Seven axis values written to posts sheet.');
}

function buildClassificationPromptForSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName('posts');
  const exportSheet = ss.getSheetByName('site-export');

  // Get active row
  const activeRow = postsSheet.getActiveCell().getRow();
  if (activeRow < 2) return null;

  // Get headers from posts sheet
  const postsHeaders = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];

  // Get Post ID from active row
  const postIdCol = postsHeaders.indexOf('Post ID') + 1;
  if (postIdCol === 0) return null;
  const postId = postsSheet.getRange(activeRow, postIdCol).getValue();
  if (!postId) return null;

  // Get donor article from site-export sheet
  const exportHeaders = exportSheet.getRange(1, 1, 1, exportSheet.getLastColumn()).getValues()[0];
  const exportIdCol = exportHeaders.indexOf('ID') + 1;
  const exportHtmlCol = exportHeaders.indexOf('Full Post HTML') + 1;
  if (exportIdCol === 0 || exportHtmlCol === 0) return null;

  const exportData = exportSheet.getRange(2, 1, exportSheet.getLastRow() - 1, exportSheet.getLastColumn()).getValues();
  let donorArticle = '';
  for (let i = 0; i < exportData.length; i++) {
    if (String(exportData[i][exportIdCol - 1]) === String(postId)) {
      donorArticle = exportData[i][exportHtmlCol - 1];
      break;
    }
  }

  if (!donorArticle) return null;

  // Assemble governed brief from posts sheet columns
  const getValue = (header) => {
    const col = postsHeaders.indexOf(header) + 1;
    return col > 0 ? postsSheet.getRange(activeRow, col).getValue() : '';
  };

  const governedBrief = [
    'Article Type: ' + getValue('Article Type'),
    'Primary Entity: ' + getValue('Primary Entity'),
    'Material Entity: ' + getValue('Material Entity'),
    'Confirmed Primary Intent: ' + getValue('Confirmed Primary Intent'),
    'Supporting Entities Core: ' + getValue('Supporting Entities Core'),
    'Rewrite Governance Summary: ' + getValue('Rewrite Governance Summary'),
    'Page Rewrite Brief: ' + getValue('Page Rewrite Brief'),
    'Cannibalisation Guardrail: ' + getValue('Cannibalisation Guardrail')
  ].join('\n');

  return buildClassificationPrompt(donorArticle, governedBrief);
}


function saveClassificationResponseToSheet(response, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName('posts');

  const activeRow = Number(row || postsSheet.getActiveCell().getRow());
  if (activeRow < 2) return { success: false, message: 'No active row selected.' };

  const postsHeaders = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];

  const parsed = parseClassificationResponse(response);
  if (!parsed) return { success: false, message: 'Could not parse classification response.' };

  // Run duplication check before saving
  const dupCheck = checkClassificationDuplication(postsSheet, activeRow, postsHeaders, parsed);

  if (dupCheck.isDuplicate) {
    // Return duplication data to sidebar without saving yet
    return {
      success: true,
      isDuplicate: true,
      message: dupCheck.message,
      nearDuplicates: dupCheck.nearDuplicates,
      axisToNudge: dupCheck.axisToNudge,
      nudgeSuggestion: dupCheck.nudgeSuggestion,
      axisValues: dupCheck.axisValues,
      currentValues: dupCheck.currentValues
    };
  }

  // No duplication — save directly
  writeClassificationToSheet(postsSheet, activeRow, postsHeaders, parsed);
  return {
    success: true,
    isDuplicate: false,
    message: dupCheck.message
  };
}

function buildClassificationPrompt(donorArticle, governedBrief) {
  return `You are a content classification assistant working within a governed floor restoration content pipeline.

You will be given two inputs:
1. A donor article
2. A governed brief

Your task is to classify the donor article against seven governed axes and assign one Problem Angle code, then return all eight values.

CLASSIFICATION RULES:
- Select only from the predefined values listed for each axis
- Do not invent new values
- Do not return free text outside the XML tags
- Do not explain your selections
- Do not add commentary
- Return exactly eight XML tags and nothing else

AXIS 1 — ENTRY CONDITION
Select the single dominant condition the floor presented with when the job came in.
Permitted values:
- Residue and coating build-up
- Ingrained traffic soiling
- Failed or incompatible previous treatment
- Biological growth or salt activity
- Physical damage or instability
- Surface wear and pigment loss

AXIS 2 — HOMEOWNER PERCEPTION
Select the single dominant perception or concern the homeowner brought to the job.
Permitted values:
- Floor looks dirty despite regular cleaning
- Pattern or colour has gradually disappeared
- Previous treatment made it worse
- Floor looks damaged or beyond saving
- Unsure whether cleaning or restoration is needed
- Concerned about damaging it further

AXIS 3 — MATERIAL BEHAVIOUR
Select the single dominant physical mechanism causing the problem on this floor.
Permitted values:
- Porosity and residue lock-in
- Calcium carbonate etch sensitivity
- Salt migration and moisture movement
- Cleavage plane or lamination fragility
- Pigment depth and surface wear
- Binder or aggregate vulnerability

AXIS 4 — CONSTRAINT
Select the single dominant constraint that governed what could safely be done on this job.
Permitted values:
- No aggressive abrasion permitted
- Acid chemistry excluded
- Moisture sensitive installation
- Vapour permeability must be maintained
- Existing damage limits mechanical intervention
- Heritage surface integrity governs all decisions

AXIS 5 — PROCESS EMPHASIS
Select the single intervention stage that carried the most weight on this job.
Permitted values:
- Residue softening and extraction
- Controlled drying and moisture management
- Sealer or coating removal
- Stabilisation before cleaning
- Neutralisation and rinse control
- Breathable protection decision

AXIS 6 — RESULT TYPE
Select the single dominant result that the homeowner could see had changed.
Permitted values:
- Pattern definition became readable again
- Colour depth and tonal contrast returned
- Surface stopped cycling back to dull after cleaning
- Maintenance became noticeably easier
- Physical stability restored alongside appearance
- Original material character preserved without artificial finish

AXIS 7 — NARRATIVE ARCHETYPE
Select the single narrative structure that best fits this article based on where the story naturally begins.
Permitted values:
- Diagnostic-first
- Material-first
- Constraint-first
- Surface-history-first

INPUTS:

DONOR ARTICLE:
${donorArticle}

GOVERNED BRIEF:
${governedBrief}

AXIS 8 — PROBLEM ANGLE
Assign the correct Problem Angle code based on the article type.
Use the donor article and governed brief to confirm the article type then assign accordingly.
Permitted values:
- PROVE (Case Study — has this been fixed for someone like me)
- FIX (Method Guide or Service Page — how is this professionally corrected)
- ID (Diagnostic Guide — which specific condition do I have)
- WHY (Educational Guide or Hub Page — why is this happening to my floor)
- LOCAL (Geo Service Page — can someone fix this near me)
- CHOOSE (Buyer Guide focused on contractor selection)
- DECIDE (Buyer Guide focused on value or comparison)
- PREVENT (Maintenance article — how do I stop this happening again)

For Buyer Guide articles — select CHOOSE if the page focus is finding and evaluating a contractor, select DECIDE if the page focus is whether professional intervention is worth it.

Return your response in this exact format with no additional text:

<entry_condition>value here</entry_condition>
<homeowner_perception>value here</homeowner_perception>
<material_behaviour>value here</material_behaviour>
<constraint>value here</constraint>
<process_emphasis>value here</process_emphasis>
<result_type>value here</result_type>
<narrative_archetype>value here</narrative_archetype>
<problem_angle>value here</problem_angle>`;
}
function callLLMForClassification(prompt) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    const url = 'https://api.anthropic.com/v1/messages';
    
    const payload = {
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages: [
        {
          role: 'user',
          content: prompt
        }
      ]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log('LLM call failed with response code: ' + responseCode);
      Logger.log('Response body: ' + response.getContentText());
      return null;
    }

    const responseData = JSON.parse(response.getContentText());
    return responseData.content[0].text;

  } catch (error) {
    Logger.log('Error calling LLM: ' + error.toString());
    return null;
  }
}
function parseClassificationResponse(response) {
  try {

    const axes = [
      'entry_condition',
      'homeowner_perception',
      'material_behaviour',
      'constraint',
      'process_emphasis',
      'result_type',
      'narrative_archetype'
    ];

    const parsed = {};

    const clean = String(response || '')
      .trim()
      .replace(/^```(?:json)?\s*/i, '')
      .replace(/\s*```$/i, '')
      .trim();

    // New AC JSON-array format
    if (clean.startsWith('[')) {

      const data = JSON.parse(clean);

      data.forEach(function(entry) {

        if (
          entry &&
          Number.isInteger(Number(entry.index)) &&
          Number(entry.index) >= 0 &&
          Number(entry.index) < axes.length
        ) {

          const axis = axes[Number(entry.index)];

          parsed[axis] = String(
            entry.selected ||
            entry.closest_existing ||
            ''
          ).trim();
        }
      });

      axes.forEach(function(axis) {
        if (!parsed[axis]) parsed[axis] = '';
      });

      parsed.problem_angle = '';

      return parsed;
    }

    // Existing XML format
    const xmlAxes = axes.concat(['problem_angle']);

    xmlAxes.forEach(function(axis) {
      const regex =
        new RegExp('<' + axis + '>([^<]+)</' + axis + '>');

      const match = clean.match(regex);

      parsed[axis] =
        match && match[1]
          ? match[1].trim()
          : '';
    });

    return parsed;

  } catch (error) {
    Logger.log(
      'Error parsing classification response: ' +
      error.toString()
    );

    return null;
  }
}

function writeClassificationToSheet(postsSheet, activeRow, headers, parsed) {
  try {
    const columnMap = {
      'entry_condition': 'ac_entry_condition',
      'homeowner_perception': 'ac_homeowner_perception',
      'material_behaviour': 'ac_material_behaviour',
      'constraint': 'ac_constraint',
      'process_emphasis': 'ac_process_emphasis',
      'result_type': 'ac_result_type',
      'narrative_archetype': 'ac_narrative_archetype',
      'problem_angle': 'Problem Angle'
    };

    const angleDefinitions = {
      'PROVE':   { label: 'Case Study / Proof',     readerQ: 'Has this actually been fixed for someone like me?' },
      'FIX':     { label: 'Method / Process',       readerQ: 'How is this professionally corrected?' },
      'ID':      { label: 'Diagnosis',              readerQ: 'Which specific condition do I have?' },
      'WHY':     { label: 'Material Explanation',   readerQ: 'Why is this happening to my floor?' },
      'LOCAL':   { label: 'Geo Service',            readerQ: 'Can someone fix this near me?' },
      'CHOOSE':  { label: 'Contractor Selection',   readerQ: 'Who should I hire and how do I evaluate them?' },
      'DECIDE':  { label: 'Value / Comparison',     readerQ: 'Is professional intervention worth it?' },
      'PREVENT': { label: 'Maintenance',            readerQ: 'How do I stop this happening again?' }
    };

    Object.keys(columnMap).forEach(key => {
      const colName = columnMap[key];
      let colIndex = headers.indexOf(colName) + 1;

      if (colIndex === 0) {
        colIndex = headers.length + 1;
        postsSheet.getRange(1, colIndex).setValue(colName);
        headers.push(colName);
      }

      let value = parsed[key] || '';

      // Format Problem Angle as full block
      if (key === 'problem_angle' && value) {
        const angleCode = value.trim().toUpperCase();
        const angleDef = angleDefinitions[angleCode];
        if (angleDef) {
          const primaryCluster = postsSheet.getRange(activeRow, headers.indexOf('Primary Query Cluster Owned') + 1).getValue() || 'this problem';
          const locality = postsSheet.getRange(activeRow, headers.indexOf('Locality') + 1).getValue() || '';
          const angleTemplates = {
            'PROVE':   'shows how [PROBLEM] was corrected on a real floor, with before and after evidence',
            'FIX':     'explains how [PROBLEM] is professionally corrected, stage by stage',
            'ID':      'helps you identify which specific condition is causing [PROBLEM], so you choose the right solution',
            'WHY':     'explains why [PROBLEM] occurs at material level, so you understand what you are actually dealing with',
            'LOCAL':   'offers professional correction of [PROBLEM] for homeowners in [LOCALITY]',
            'CHOOSE':  'helps you find and evaluate the right specialist to correct [PROBLEM]',
            'DECIDE':  'helps you decide whether professional intervention is the right choice for [PROBLEM]',
            'PREVENT': 'explains how to prevent [PROBLEM] recurring through correct ongoing maintenance'
          };
          const template = angleTemplates[angleCode] || '';
          const statement = 'This article ' + template
            .replace('[PROBLEM]', primaryCluster)
            .replace('[LOCALITY]', locality || 'your area') + '.';
          value = angleCode + ' — ' + angleDef.label + '\n→ ' + angleDef.readerQ + '\n' + statement;
        }
      }

      const cell = postsSheet.getRange(activeRow, colIndex);
      if (value.toString().startsWith('=')) {
        cell.setValue("'" + value);
      } else {
        cell.setValue(value);
      }
    });

    SpreadsheetApp.flush();

  } catch (error) {
    Logger.log('Error writing classification to sheet: ' + error.toString());
  }
}
function checkClassificationDuplication(postsSheet, activeRow, headers, parsed) {
  try {
    const stoneTypeCol = headers.indexOf('Stone Type') + 1;
    const activeStoneType = stoneTypeCol > 0 ? 
      String(postsSheet.getRange(activeRow, stoneTypeCol).getValue() || '').trim() : '';

    const axisKeys = [
      'entry_condition',
      'homeowner_perception',
      'material_behaviour',
      'constraint',
      'process_emphasis',
      'result_type',
      'narrative_archetype'
    ];

    const columnMap = {
      'entry_condition': 'ac_entry_condition',
      'homeowner_perception': 'ac_homeowner_perception',
      'material_behaviour': 'ac_material_behaviour',
      'constraint': 'ac_constraint',
      'process_emphasis': 'ac_process_emphasis',
      'result_type': 'ac_result_type',
      'narrative_archetype': 'ac_narrative_archetype'
    };

    const axisValues = {
      'entry_condition': [
        'Residue and coating build-up',
        'Ingrained traffic soiling',
        'Failed or incompatible previous treatment',
        'Biological growth or salt activity',
        'Physical damage or instability',
        'Surface wear and pigment loss'
      ],
      'homeowner_perception': [
        'Floor looks dirty despite regular cleaning',
        'Pattern or colour has gradually disappeared',
        'Previous treatment made it worse',
        'Floor looks damaged or beyond saving',
        'Unsure whether cleaning or restoration is needed',
        'Concerned about damaging it further'
      ],
      'material_behaviour': [
        'Porosity and residue lock-in',
        'Calcium carbonate etch sensitivity',
        'Salt migration and moisture movement',
        'Cleavage plane or lamination fragility',
        'Pigment depth and surface wear',
        'Binder or aggregate vulnerability'
      ],
      'constraint': [
        'No aggressive abrasion permitted',
        'Acid chemistry excluded',
        'Moisture sensitive installation',
        'Vapour permeability must be maintained',
        'Existing damage limits mechanical intervention',
        'Heritage surface integrity governs all decisions'
      ],
      'process_emphasis': [
        'Residue softening and extraction',
        'Controlled drying and moisture management',
        'Sealer or coating removal',
        'Stabilisation before cleaning',
        'Neutralisation and rinse control',
        'Breathable protection decision'
      ],
      'result_type': [
        'Pattern definition became readable again',
        'Colour depth and tonal contrast returned',
        'Surface stopped cycling back to dull after cleaning',
        'Maintenance became noticeably easier',
        'Physical stability restored alongside appearance',
        'Original material character preserved without artificial finish'
      ],
      'narrative_archetype': [
        'Diagnostic-first',
        'Material-first',
        'Constraint-first',
        'Surface-history-first'
      ]
    };

    const nudgePriority = [
      'narrative_archetype',
      'process_emphasis',
      'result_type',
      'homeowner_perception',
      'entry_condition',
      'material_behaviour',
      'constraint'
    ];

    // Get all classified rows for same stone type
    const lastRow = postsSheet.getLastRow();
    const allData = postsSheet.getRange(2, 1, lastRow - 1, postsSheet.getLastColumn()).getValues();

    const stoneTypeIdx = headers.indexOf('Stone Type');
    const classifiedRows = [];

    allData.forEach(function(row, idx) {
      const rowNum = idx + 2;
      if (rowNum === activeRow) return;
      if (stoneTypeIdx > -1 && String(row[stoneTypeIdx] || '').trim() !== activeStoneType) return;

      const rowValues = {};
      let hasClassification = false;
      axisKeys.forEach(function(key) {
        const colName = columnMap[key];
        const colIdx = headers.indexOf(colName);
        if (colIdx > -1 && row[colIdx]) {
          rowValues[key] = String(row[colIdx]).trim();
          hasClassification = true;
        }
      });

      if (hasClassification) {
        classifiedRows.push({ rowNum: rowNum, values: rowValues });
      }
    });

    if (classifiedRows.length < 2) {
      return { isDuplicate: false, message: 'Only ' + classifiedRows.length + ' classified row(s) found for ' + activeStoneType + '. Continue classifying to build a valid sample.' };
    }

    // Check for near-duplicates
    let nearDuplicates = [];
    classifiedRows.forEach(function(existingRow) {
      let matchCount = 0;
      let matchedAxes = [];
      axisKeys.forEach(function(key) {
        if (parsed[key] && existingRow.values[key] && 
            parsed[key].trim() === existingRow.values[key].trim()) {
          matchCount++;
          matchedAxes.push(key);
        }
      });
      if (matchCount >= 4) {
        nearDuplicates.push({
          rowNum: existingRow.rowNum,
          matchCount: matchCount,
          matchedAxes: matchedAxes,
          values: existingRow.values
        });
      }
    });

    if (nearDuplicates.length === 0) {
      return { isDuplicate: false, message: 'No near-duplicates detected. Classification is sufficiently unique.' };
    }

    // Find best axis to nudge
    let axisToNudge = null;
    let nudgeSuggestion = null;

    for (let p = 0; p < nudgePriority.length; p++) {
      const axis = nudgePriority[p];
      const usedValues = classifiedRows.map(function(r) { return r.values[axis] || ''; });
      const availableValues = axisValues[axis].filter(function(v) {
        return usedValues.indexOf(v) === -1;
      });

      if (availableValues.length > 0) {
        axisToNudge = axis;
        nudgeSuggestion = availableValues[0];
        break;
      }
    }

    return {
      isDuplicate: true,
      nearDuplicates: nearDuplicates,
      axisToNudge: axisToNudge,
      nudgeSuggestion: nudgeSuggestion,
      axisValues: axisValues,
      currentValues: parsed,
      message: nearDuplicates.length + ' near-duplicate(s) detected.'
    };

  } catch(error) {
    Logger.log('Duplication check error: ' + error.toString());
    return { isDuplicate: false, message: 'Duplication check failed: ' + error.toString() };
  }
}
function saveClassificationWithNudge(nudgedValues) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const postsSheet = ss.getSheetByName('posts');
    const activeRow = postsSheet.getActiveCell().getRow();
    if (activeRow < 2) return { success: false, message: 'No active row selected.' };

    const headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];

    writeClassificationToSheet(postsSheet, activeRow, headers, nudgedValues);

    return { success: true };
  } catch(error) {
    Logger.log('Save with nudge error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}