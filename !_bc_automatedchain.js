function bc_runP1AAutomated() {
  var promptData = bc_getPrompt1A();
  var apiResult = bc_sendPromptViaOpenAI(
      promptData.fullPrompt,
      null,
      MODEL_CHEAP
    );
  if (!apiResult.success) throw new Error(apiResult.message);

  var text = apiResult.text;
  var articleTypeMatch = text.match(/CONFIRMED ARTICLE TYPE:\s*(.+)/);
  var primaryIntentMatch = text.match(/CONFIRMED PRIMARY INTENT:\s*(.+)/);
  if (!articleTypeMatch || !primaryIntentMatch) {
    throw new Error("Could not find CONFIRMED ARTICLE TYPE or CONFIRMED PRIMARY INTENT in API response.");
  }
  var confirmedArticleType = articleTypeMatch[1].trim();
  var confirmedPrimaryIntent = primaryIntentMatch[1].trim();

  var rankedMatch = text.match(/BLOCK 4[^\n]*\n([\s\S]*?)(?:BLOCK 5|$)/);
  var rankedBlockText = rankedMatch ? rankedMatch[1].trim() : '';
  var rankedLines = rankedBlockText.split('\n').filter(function(l) { return l.trim(); });

  var secondaryLines = [];
  var retainedCount = 0;
  for (var i = 0; i < rankedLines.length; i++) {
    var lineMatch = rankedLines[i].match(/-\s*(?:Primary Intent|Secondary Intent):\s*(.+)/i);
    if (!lineMatch) continue;
    var intentName = lineMatch[1].replace(/\(.*\)/, '').trim();
    if (!intentName || intentName.toLowerCase() === confirmedPrimaryIntent.toLowerCase()) continue;
    retainedCount++;
    secondaryLines.push(retainedCount + ". " + intentName + " — Confirmed in automated parse.");
  }
  var secondaryBlock = secondaryLines.length ? ("RETAINED SUPPORTING INTENTS:\n" + secondaryLines.join("\n")) : "";

  bc_writePrompt1ALockToRow(
    confirmedArticleType,
    confirmedPrimaryIntent,
    text,
    secondaryBlock,
    text,
    promptData.lockedRow
  );

  var message = "Article Type: " + confirmedArticleType + " | Primary Intent: " + confirmedPrimaryIntent;
  return { success: true, message: message, cost: apiResult.cost };
}

function bc_runP1BAutomated() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName("posts");
  var row = ss.getActiveRange().getRow();
  var headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).trim().toLowerCase(); });
  var atIdx = headers.indexOf("article type");
  var piIdx = headers.indexOf("confirmed primary intent");
  if (atIdx === -1 || piIdx === -1) throw new Error("Article Type or Confirmed Primary Intent column not found.");
  var confirmedArticleType = String(postsSheet.getRange(row, atIdx + 1).getValue() || "").trim();
  var confirmedPrimaryIntent = String(postsSheet.getRange(row, piIdx + 1).getValue() || "").trim();
  if (!confirmedArticleType || !confirmedPrimaryIntent) {
    throw new Error("Run P1A first — Article Type or Confirmed Primary Intent is empty on this row.");
  }

  var lockBlock =
    "===== CONFIRMED PRIMARY INTENT LOCK FROM PROMPT 1A =====\n" +
    "CONFIRMED ARTICLE TYPE: " + confirmedArticleType + "\n" +
    "CONFIRMED PRIMARY INTENT: " + confirmedPrimaryIntent + "\n" +
    "===== END LOCK =====\n\n" +
    "HARD RULE: The values above are locked inputs from Prompt 1A.\n" +
    "Do NOT redetect, reinterpret, validate, or override Article Type or Intent\n" +
    "from HTML, title, slug, meta title, schema, or any other source.\n" +
    "Use them exactly as supplied for all downstream field population.";

  var promptData = bc_getPrompt1B(lockBlock);
  var apiResult = bc_sendPromptViaOpenAI(
  promptData.fullPrompt,
  null,
  MODEL_CHEAP
);
  if (!apiResult.success) throw new Error(apiResult.message);

  var pairs = bc_parsePairsServer(apiResult.text);
  if (!Object.keys(pairs).length) throw new Error("Could not parse Label: Value pairs from P1B response.");

  var writeResult = bc_writeLLMOutputToRow(pairs, apiResult.text, 'P1B Response', promptData.lockedRow);
  return { success: true, message: writeResult, cost: apiResult.cost };
}
function bc_runP2Automated() {
  var promptData = bc_getPrompt2();

  var prewrite = { "Recovery Blueprint": promptData.recoveryBlueprint };
  prewrite["GSC Query List"] = promptData.gscQueryList || "no valid GSC queries";
  bc_writeLLMOutputToRow(prewrite, null, null, promptData.lockedRow);

  var apiResult = bc_sendPromptViaOpenAI(
  promptData.fullPrompt,
  null,
  MODEL_CHEAP
);
  if (!apiResult.success) throw new Error(apiResult.message);

  var pairs = bc_parsePairsServer(apiResult.text);
  if (!Object.keys(pairs).length) throw new Error("Could not parse Label: Value pairs from P2 response.");

  var writeResult = bc_writeLLMOutputToRow(pairs, apiResult.text, 'P2 Response', promptData.lockedRow);
  return { success: true, message: writeResult, cost: apiResult.cost };
}

function bc_runP3Automated() {
  var promptData = bc_getPrompt3();
  var apiResult = bc_sendPromptViaOpenAI(
  promptData.fullPrompt,
  null,
  MODEL_CHEAP
);
  if (!apiResult.success) throw new Error(apiResult.message);

  var pairs = bc_parsePairsServer(apiResult.text);
  if (!Object.keys(pairs).length) throw new Error("Could not parse Label: Value pairs from P3 response.");
  if (!pairs['Rewrite Brief Date']) {
    pairs['Rewrite Brief Date'] = bc_getTodayDDMMYYYY();
  }

  var writeResult = bc_writeLLMOutputToRow(pairs, apiResult.text, 'P3 Response', promptData.lockedRow);
  return { success: true, message: writeResult, cost: apiResult.cost };
}

function bc_runW0Automated() {
  var prompt = generateForensicPrompt();
  if (typeof prompt !== 'string' || prompt.indexOf('STOP') === 0) {
    throw new Error(prompt);
  }
  var apiResult = bc_sendPromptViaOpenAI(
      prompt,
      null,
      MODEL_CHEAP
    );
  if (!apiResult.success) throw new Error(apiResult.message);

  var cleanedText = apiResult.text;
  var w0FieldLabels = ["Strategic Reasoning", "Topical Intent", "Matrix Role", "Key Decisions Explained"];
  for (var wl = 0; wl < w0FieldLabels.length; wl++) {
    var labelRe = new RegExp("\\s*\\|?\\s*(" + w0FieldLabels[wl] + ":)", "g");
    cleanedText = cleanedText.replace(labelRe, "\n$1");
  }

  var pushResult = pushStage0FieldsToActiveRow(cleanedText);
  if (!pushResult.success) throw new Error(pushResult.message);

  return { success: true, message: pushResult.message, cost: apiResult.cost };
}

function bc_sendPromptViaGemini(promptText) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY not set in Script Properties.');
  var payload = {
    contents: [{ parts: [{ text: promptText }] }],
    tools: [{ googleSearch: {} }]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + apiKey;
  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200) {
    return { success: false, message: (data.error && data.error.message) || 'Gemini API error' };
  }
  var text = data.candidates[0].content.parts.map(function(p) { return p.text || ''; }).join('');
  var promptTokens = (data.usageMetadata && data.usageMetadata.promptTokenCount) || 0;
  var completionTokens = (data.usageMetadata && data.usageMetadata.candidatesTokenCount) || 0;
  var cost = (promptTokens / 1000000 * 1.50) + (completionTokens / 1000000 * 7.50);
  return {
    success: true,
    text: text,
    promptTokens: promptTokens,
    completionTokens: completionTokens,
    cost: cost
  };
}
function bc_runSerpBridgeAutomated() {
  var promptData = bc_buildSerpBridgePrompt();
  if (!promptData.success) throw new Error(promptData.message);

  var apiResult = bc_sendPromptViaGemini(promptData.prompt);
  if (!apiResult.success) throw new Error(apiResult.message);

  var pushResult = bc_pushSerpBridgeToSheet(apiResult.text, true);
  if (!pushResult.success) throw new Error(pushResult.message);

  var message = pushResult.message;

  var lines = apiResult.text.split('\n');
  var tsmRows = lines.filter(function(line) {
    var trimmed = line.trim();
    return trimmed.indexOf('|') > -1 &&
           trimmed.indexOf('Open') !== 0 &&
           trimmed.indexOf('Find') !== 0 &&
           trimmed.indexOf('Add') !== 0 &&
           trimmed.indexOf('After') !== 0 &&
           trimmed.indexOf('Then') !== 0 &&
           trimmed.indexOf('[') !== 0 &&
           trimmed.indexOf('---') !== 0 &&
           trimmed.indexOf('Exact') !== 0 &&
           trimmed.length > 10;
  });

  if (tsmRows.length) {
    var tsmResult = bc_pushEntitiesToTSM(tsmRows.join('\n'));
    message += ' | ' + tsmResult.message;
  }

  var blueprintResult = bc_addSerpEntitiesToBlueprint();
  message += ' | ' + blueprintResult.message;

  return { success: true, message: message, cost: apiResult.cost };
}

function bc_appendGovernancePipelineException(stageName, details) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('posts');
  var row = sh.getActiveCell().getRow();

  if (row < 2) {
    throw new Error('Select a data row first.');
  }

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  var colIndex = headers.indexOf('Governence Pipeline Exceptions');

  if (colIndex === -1) {
    throw new Error('Governence Pipeline Exceptions column not found.');
  }

  var cell = sh.getRange(row, colIndex + 1);
  var existing = String(cell.getValue() || '').trim();

  var entry =
    stageName +
    ' — ' +
    String(details || '').trim();

  var updated = existing
    ? existing + '\n' + entry
    : entry;

  cell.setNumberFormat('@');
  cell.setValue(updated);

  return {
    success: true,
    message: 'Pipeline exception recorded in GJ.'
  };
}

function bc_clearGovernancePipelineExceptions() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('posts');
  var row = sh.getActiveCell().getRow();

  if (row < 2) {
    throw new Error('Select a data row first.');
  }

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  var colIndex = headers.indexOf('Governence Pipeline Exceptions');

  if (colIndex === -1) {
    throw new Error('Governence Pipeline Exceptions column not found.');
  }

  sh.getRange(row, colIndex + 1).clearContent();

  return {
    success: true,
    message: 'Previous pipeline exceptions cleared from GJ.'
  };
}

function bc_runFullGovernanceChainAutomated() {

  // STOP BEFORE ANY COST IF NO HUB PAGE EXISTS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postsSheet = ss.getSheetByName("posts");
  var activeRow = ss.getActiveRange().getRow();

  var headers = postsSheet
    .getRange(1, 1, 1, postsSheet.getLastColumn())
    .getValues()[0];

  var stoneTypeIdx = headers.indexOf("Stone Type");

  if (stoneTypeIdx === -1) {
    throw new Error('Column "Stone Type" not found.');
  }

  var stoneType = String(
    postsSheet.getRange(activeRow, stoneTypeIdx + 1).getValue() || ""
  ).trim();

  var feedsHubUrl = bc_getFeedsHubUrlForStoneType_(stoneType);

  // Clear previous Governance Pipeline Exceptions for this row
  var exceptionIdx = headers.indexOf("Governence Pipeline Exceptions");

  if (exceptionIdx !== -1) {
    postsSheet
      .getRange(activeRow, exceptionIdx + 1)
      .clearContent();
  }

  if (!feedsHubUrl || feedsHubUrl === "none") {
    return {
      success: false,
      message:
        'STOPPED — no Hub Page exists for "' +
        stoneType +
        '". Give this Stone Type a Hub Page before running governance.',
      log: "",
      cost: 0
    };
  }

  var startTime = Date.now();
  var totalCost = 0;

  var modelCosts = {
    "gpt-5.6-luna": 0,
    "gpt-5.6-terra": 0,
    "gpt-5.6-sol": 0,
    "gemini": 0
  };

  var log = [];

  function stageSeconds(stageStart) {
    return ((Date.now() - stageStart) / 1000).toFixed(1) + 's';
  }


  // =========================================================
  // 1. GSC SYNC
  // =========================================================
  try {
    var syncResult = runGSCFullSyncActiveRow();
    runGSCRowStatsActiveRow();

    log.push(
      'GSC Sync: completed for Post ID ' +
      syncResult.postId
    );

  } catch (e) {
    return {
      success: false,
      message: 'STOPPED at GSC Sync — ' + e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }


  // =========================================================
  // 2. COMBINED GOVERNANCE
  // =========================================================
  try {
    var promptData = bc_getPromptCombined();

    var apiResult = bc_sendPromptViaOpenAI(
      promptData.fullPrompt,
      8000,
      MODEL_CHEAP
    );

    if (!apiResult.success) {
      throw new Error(apiResult.message);
    }

    totalCost += apiResult.cost;

    modelCosts[MODEL_CHEAP] =
      (modelCosts[MODEL_CHEAP] || 0) +
      apiResult.cost;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lockedRow = ss.getActiveRange().getRow();

    var writeResult = bc_writeCombinedOutputToRow(
      apiResult.text,
      lockedRow
    );

    log.push(
      'Combined Governance: ' +
      writeResult.message
    );

    if (
      writeResult.governance &&
      writeResult.governance.success === false
    ) {
      return {
        success: false,
        message:
          'STOPPED at Governance Validation — ' +
          writeResult.governance.message,
        log: log.join(' | '),
        cost: totalCost
      };
    }

  } catch (e) {
    return {
      success: false,
      message:
        'STOPPED at Combined Governance — ' +
        e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }



  // =========================================================
  // 3. W0
  // =========================================================
  try {
    var r5 = bc_runW0Automated();

    totalCost += r5.cost;

    modelCosts[MODEL_CHEAP] =
      (modelCosts[MODEL_CHEAP] || 0) +
      r5.cost;

    log.push(
      'W0: ' +
      r5.message
    );

  } catch (e) {
    return {
      success: false,
      message:
        'STOPPED at W0 — ' +
        e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }


  // =========================================================
  // 4. W0B — LOCATION CONTEXT
  // =========================================================
  try {
    var r5b = bc_runW0BAutomated();

    totalCost += r5b.cost;

    log.push(
      'W0B: ' +
      (
        r5b.skipped
          ? 'Skipped — ' + r5b.message
          : r5b.message
      )
    );

  } catch (e) {
    return {
      success: false,
      message:
        'STOPPED at W0B — ' +
        e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }


  // =========================================================
  // 5. SERP ENTITY BRIDGE
  // =========================================================
  try {
    var r6 = bc_runSerpBridgeAutomated();

    totalCost += r6.cost;
    modelCosts["gemini"] += r6.cost;

    log.push(
      'SERP: ' +
      r6.message
    );

  } catch (e) {
    return {
      success: false,
      message:
        'STOPPED at SERP Entity Bridge — ' +
        e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }



  // =========================================================
  // 17. SAVE COST + TIME
  // =========================================================

  var elapsedSeconds =
    ((Date.now() - startTime) / 1000).toFixed(1);

  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  var postsSheet =
    ss.getSheetByName("posts");

  var row =
    ss.getActiveRange().getRow();

  bc_writeLLMOutputToRow(
    {
      "API Cost": totalCost,
      "API Time": elapsedSeconds + "s"
    },
    null,
    null,
    row
  );


  var elapsedMinutes =
    Math.floor(
      Number(elapsedSeconds) / 60
    );

  var remainingSeconds =
    Math.round(
      Number(elapsedSeconds) % 60
    );

  var elapsedTime =
    elapsedMinutes > 0
      ? elapsedMinutes +
        "m " +
        remainingSeconds +
        "s"
      : elapsedSeconds + "s";


  return {
    success: true,

    message:
      "✔ Content Intelligence complete. Time: " +
      elapsedTime +
      " | Total cost: $" +
      totalCost.toFixed(4) +
      ' | Luna: $' +
      modelCosts["gpt-5.6-luna"].toFixed(4) +
      ' | Terra: $' +
      modelCosts["gpt-5.6-terra"].toFixed(4) +
      ' | Sol: $' +
      modelCosts["gpt-5.6-sol"].toFixed(4) +
      ' | Gemini: $' +
      modelCosts["gemini"].toFixed(4),

    log: log.join(' | '),
    cost: totalCost
  };
}

function bc_runW0BAutomated() {
  var promptData = buildLocationContextPrompt();
  if (!promptData.success) {
    return { success: true, skipped: true, message: promptData.message, cost: 0 };
  }

  var apiResult = bc_sendPromptViaOpenAI(promptData.prompt);
  if (!apiResult.success) throw new Error(apiResult.message);

  var pushResult = pushLocationContextToSheet(promptData.locality, promptData.parentArea, apiResult.text);
  if (!pushResult.success) throw new Error(pushResult.message);

  return { success: true, skipped: false, message: pushResult.message, cost: apiResult.cost };
}

function bc_runPreACPromptAutomated() {

  var startTime = Date.now();

  var promptData = preAcBuildGapCheckPrompt();

  if (!promptData || promptData.error) {
    throw new Error(
      promptData && promptData.error
        ? promptData.error
        : 'Pre-AC prompt could not be built.'
    );
  }

  var apiResult = bc_sendPromptViaOpenAI(
    promptData.prompt,
    null,
    MODEL_CHEAP
  );

  if (!apiResult.success) {
    throw new Error(apiResult.message);
  }

  var raw = String(apiResult.text || '').trim();

  raw = raw
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();

  var entries;

  try {
    entries = JSON.parse(raw);
  } catch (e) {
    throw new Error(
      'Pre-AC returned invalid JSON — ' + e.message
    );
  }

  if (!Array.isArray(entries)) {
    throw new Error(
      'Pre-AC response was not a JSON array.'
    );
  }

  var gapsFound = 0;
  var gapsAdded = 0;
  var added = [];
  var failed = [];

  entries.forEach(function(entry) {

    if (!entry || entry.gap !== true) {
      return;
    }

    gapsFound++;

    var suggestion =
      String(entry.suggested_option || '').trim();

    if (!suggestion) {
      failed.push(
        String(entry.field || 'Unknown field') +
        ' — gap was TRUE but suggested_option was empty'
      );
      return;
    }

    var addResult =
      preAcAddGapToSheet(entry);

    if (addResult && addResult.success) {

      gapsAdded++;

      added.push(
        String(entry.field || 'Unknown field') +
        ': ' +
        suggestion
      );

    } else {

      failed.push(
        String(entry.field || 'Unknown field') +
        ' — ' +
        (
          addResult && addResult.message
            ? addResult.message
            : 'could not be added'
        )
      );
    }
  });

  bc_addToApiCostAndTime(
    apiResult.cost,
    (Date.now() - startTime) / 1000
  );

  if (failed.length > 0) {
    throw new Error(
      'Pre-AC found ' +
      gapsFound +
      ' gap(s), but not all could be saved. ' +
      failed.join(' | ')
    );
  }

  var message;

  if (gapsFound === 0) {
    message =
      'Pre-AC PASS — existing AC library options adequately cover the article.';
  } else {
    message =
      'Pre-AC complete — ' +
      gapsAdded +
      ' new AC option(s) added';

    if (added.length > 0) {
      message += ': ' + added.join(' | ');
    }
  }

  return {
    success: true,
    message: message,
    text: apiResult.text,
    gapsFound: gapsFound,
    gapsAdded: gapsAdded,
    cost: apiResult.cost
  };
}

function bc_runACPromptAutomated(row) {

  var startTime = Date.now();

  var promptData = acBuildClassifyPrompt();

  if (!promptData || !promptData.prompt) {
    throw new Error(
      promptData && promptData.error
        ? promptData.error
        : 'AC classification prompt could not be built.'
    );
  }

  var apiResult = bc_sendPromptViaOpenAI(
    promptData.prompt,
    null,
    MODEL_CHEAP
  );

  if (!apiResult.success) {
    throw new Error(apiResult.message);
  }

  var response = String(apiResult.text || '').trim();

  if (!response) {
    throw new Error(
      'AC returned an empty classification response.'
    );
  }

  // Run the normal AC save + duplicate check.
  var saveResult =
  saveClassificationResponseToSheet(response, row);

  if (!saveResult || !saveResult.success) {
    throw new Error(
      saveResult && saveResult.message
        ? saveResult.message
        : 'AC classification could not be processed.'
    );
  }

  var message = '';

  // =========================================================
  // DUPLICATE FOUND — REPLICATE MANUAL "ACCEPT NUDGE"
  // =========================================================
  if (saveResult.isDuplicate) {

    var nudgedValues =
      Object.assign({}, saveResult.currentValues || {});

    var axisToNudge =
      String(saveResult.axisToNudge || '').trim();

    var suggestedValue =
      String(saveResult.nudgeSuggestion || '').trim();

    if (!axisToNudge) {
      throw new Error(
        'AC found a duplicate but supplied no axis to nudge.'
      );
    }

    if (!suggestedValue) {
      throw new Error(
        'AC found a duplicate but supplied no suggested nudge value.'
      );
    }

    // Same action as pressing the suggested option
    // in the manual duplicate-warning panel.
    nudgedValues[axisToNudge] = suggestedValue;

    var nudgeResult =
      saveClassificationWithNudge(nudgedValues);

    if (
      nudgeResult &&
      nudgeResult.success === false
    ) {
      throw new Error(
        nudgeResult.message ||
        'AC suggested nudge could not be saved.'
      );
    }

    message =
      'AC duplicate detected — automatic suggested nudge applied: ' +
      axisToNudge +
      ' → ' +
      suggestedValue;

  } else {

    message =
      saveResult.message ||
      'AC classification saved.';
  }

  bc_addToApiCostAndTime(
    apiResult.cost,
    (Date.now() - startTime) / 1000
  );

  return {
    success: true,
    message: message,
    text: apiResult.text,
    duplicateDetected: !!saveResult.isDuplicate,
    cost: apiResult.cost
  };
}

function bc_runArticleAngleAutomated(row) {

  var startTime = Date.now();

  var promptData = buildArticleAngleSuggestionPrompt(row);

  if (!promptData || !promptData.success) {
    throw new Error(
      promptData && promptData.message
        ? promptData.message
        : 'Article Angle prompt could not be built.'
    );
  }

  var apiResult = bc_sendPromptViaOpenAI(
    promptData.prompt,
    null,
    MODEL_CHEAP
  );

  if (!apiResult.success) {
    throw new Error(apiResult.message);
  }

  var raw = String(apiResult.text || '').trim();

  if (!raw) {
    throw new Error(
      'Article Angle returned an empty response.'
    );
  }

  var allowedShapes =
    promptData.allowedShapes || [];

  if (!allowedShapes.length) {
    throw new Error(
      'No permitted Article Angles were supplied.'
    );
  }

  // =========================================================
  // PARSE FIT RESULTS
  // =========================================================

  var results = [];

  var blockRegex =
    /SHAPE:\s*([^\r\n]+)[\r\n]+FIT:\s*(YES|NO|PARTIAL)[\r\n]+REASON:\s*([^\r\n]+)/gi;

  var match;

  while ((match = blockRegex.exec(raw)) !== null) {

    var shape =
      String(match[1] || '').trim();

    var fit =
      String(match[2] || '').trim().toUpperCase();

    if (
      allowedShapes.indexOf(shape) === -1
    ) {
      continue;
    }

    results.push({
      shape: shape,
      fit: fit,
      reason: String(match[3] || '').trim()
    });
  }

  if (!results.length) {
    throw new Error(
      'Article Angle response could not be parsed.'
    );
  }

  // =========================================================
  // READ MODEL RECOMMENDATION
  // =========================================================

  var recommended = '';

  var recommendedMatch =
    raw.match(/RECOMMENDED:\s*([^\r\n]+)/i);

  if (recommendedMatch) {
    recommended =
      String(recommendedMatch[1] || '').trim();
  }

  if (
    recommended &&
    allowedShapes.indexOf(recommended) === -1
  ) {
    recommended = '';
  }

  // =========================================================
  // BUILD FIT CANDIDATE POOL
  // YES FIRST — PARTIAL ONLY AS FALLBACK
  // =========================================================

  var yesShapes =
    results
      .filter(function(r) {
        return r.fit === 'YES';
      })
      .map(function(r) {
        return r.shape;
      });

  var partialShapes =
    results
      .filter(function(r) {
        return r.fit === 'PARTIAL';
      })
      .map(function(r) {
        return r.shape;
      });

  var candidates =
    yesShapes.length
      ? yesShapes
      : partialShapes;

  if (!candidates.length) {
    throw new Error(
      'Article Angle found no YES or PARTIAL fit.'
    );
  }

  // =========================================================
  // GET SAME ARTICLE TYPE + SAME STONE TYPE USAGE
  // =========================================================

  var usageResult =
    getArticleAngleUsageForStoneType(row);

  if (
    !usageResult ||
    !usageResult.success
  ) {
    throw new Error(
      usageResult && usageResult.message
        ? usageResult.message
        : 'Article Angle usage data could not be read.'
    );
  }

  var usageCounts =
    usageResult.usageCounts || {};

  var totalExisting = 0;

  Object.keys(usageCounts).forEach(
    function(shape) {
      totalExisting +=
        Number(usageCounts[shape] || 0);
    }
  );

  var chosen = '';

  // =========================================================
  // FIRST ARTICLE IN THIS GROUP
  // USE MODEL'S BEST FIT
  // =========================================================

  if (totalExisting === 0) {

    if (
      recommended &&
      candidates.indexOf(recommended) !== -1
    ) {
      chosen = recommended;
    } else {
      chosen = candidates[0];
    }

  } else {

    // =======================================================
    // EXISTING ARTICLES
    // CHOOSE LEAST-USED SUITABLE ANGLE
    // =======================================================

    var lowestCount = Infinity;

    candidates.forEach(function(shape) {

      var count =
        Number(usageCounts[shape] || 0);

      if (count < lowestCount) {
        lowestCount = count;
      }
    });

    var leastUsedCandidates =
      candidates.filter(function(shape) {

        return (
          Number(usageCounts[shape] || 0) ===
          lowestCount
        );
      });

    // If model recommendation is among the
    // least-used suitable angles, use it.
    if (
      recommended &&
      leastUsedCandidates.indexOf(recommended) !== -1
    ) {

      chosen = recommended;

    } else {

      // Otherwise preserve model fit order.
      chosen = leastUsedCandidates[0];
    }
  }

  if (!chosen) {
    throw new Error(
      'Article Angle could not choose a structural shape.'
    );
  }

  // =========================================================
  // SAVE ARTICLE ANGLE
  // =========================================================

  var saveResult =
    saveArticleAngleForActiveRow(chosen, row);

  if (
    !saveResult ||
    !saveResult.success
  ) {
    throw new Error(
      saveResult && saveResult.message
        ? saveResult.message
        : 'Article Angle could not be saved.'
    );
  }

  bc_addToApiCostAndTime(
    apiResult.cost,
    (Date.now() - startTime) / 1000
  );

  return {
    success: true,
    articleType: usageResult.articleType,
    stoneType: usageResult.stoneType,
    chosenAngle: chosen,
    fitLevel:
      yesShapes.length ? 'YES' : 'PARTIAL',
    usageCounts: usageCounts,
    message:
      'Article Angle saved: ' +
      chosen,
    cost: apiResult.cost
  };
}

function bc_runStage15AAutomated() {

  var startTime = Date.now();
  var totalCost = 0;
  var maxFixAttempts = 2;
  var fixAttempt = 0;

  var originalPrompt = buildStage15APrompt();

  if (
    typeof originalPrompt !== 'string' ||
    originalPrompt.indexOf('ERROR') === 0
  ) {
    throw new Error(originalPrompt);
  }

  var currentPrompt = originalPrompt;
  var lastOutput = '';
  var lastFailure = '';

  while (fixAttempt <= maxFixAttempts) {

    var apiResult = bc_sendPromptViaOpenAI(
      currentPrompt,
      null,
      MODEL_CHEAP
    );

    if (!apiResult.success) {
      throw new Error(apiResult.message);
    }

    totalCost += apiResult.cost;
    lastOutput = apiResult.text;

    var saveResult =
      saveStage15AStructure(lastOutput);

    // PASS — only a clean structure is saved to CU.
    if (saveResult.success) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

      return {
        success: true,
        message:
          fixAttempt === 0
            ? saveResult.message
            : 'W1.5A PASS after ' +
              fixAttempt +
              ' automated correction attempt(s). ' +
              saveResult.message,
        cost: totalCost
      };
    }

    lastFailure = saveResult.message;

    // Stop after the allowed repair attempts.
    if (fixAttempt === maxFixAttempts) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

    // Coherence never reached PASS.
    // Restore GI from the last governance-approved W2B.2 HTML in GG
    // so a failed coherence repair cannot leave damaged HTML in GI.
    var safeGG = String(
      sh.getRange(row, 189).getValue() || ''
    ).trim();

    if (safeGG) {
      sh.getRange(row, 191).setValue(safeGG);
    }

      bc_appendGovernancePipelineException(
        'W1.5A — Structure Audit',
        lastFailure +
        '\n\nFINAL W1.5A OUTPUT:\n' +
        lastOutput
      );

      return {
        success: false,
        message:
          'W1.5A still contains failed audit checks after ' +
          maxFixAttempts +
          ' automated correction attempts — recorded in GJ.',
        cost: totalCost
      };
    }

    fixAttempt++;

    currentPrompt = `
    STAGE 1.5A — STRUCTURE AUDIT CORRECTION

    The previous Stage 1.5A output failed its own structural audit.

    FAILED VALIDATION:
    ${lastFailure}

    PREVIOUS OUTPUT:
    ${lastOutput}

    ORIGINAL STAGE 1.5A INSTRUCTIONS:
    ${originalPrompt}

    TASK:
    Correct ONLY the structural decisions necessary to resolve the failed AUDIT CHECK items.

    RULES:
    1. Treat the failed AUDIT CHECK items above as locked correction instructions.
    2. Do not redesign parts of the structure that already passed.
    3. Preserve every passing audit decision unless changing it is strictly necessary to resolve a failed check.
    4. Preserve the article type, tier, material, primary entity and required structural intent.
    5. Preserve valid section roles and word budgets wherever possible.
    6. Do not introduce new sections merely for stylistic variety.
    7. Every AUDIT CHECK must be internally re-evaluated before returning the corrected structure.
    8. Do not return any AUDIT CHECK as PASS unless the corrected structure actually satisfies it.
    9. Return the COMPLETE corrected Stage 1.5A structure in exactly the same output format required by the original instructions.
    10. No markdown fences.
    11. No explanation before or after the Stage 1.5A output.

    Return the corrected Stage 1.5A structure only.
`.trim();
  }
}

function bc_runStage15BAutomated() {
  var startTime = Date.now();
  var prompt = buildStage15BPrompt();
  if (typeof prompt !== 'string' || prompt.indexOf('ERROR') === 0) {
    throw new Error(prompt);
  }
  var apiResult = bc_sendPromptViaOpenAI(
  prompt,
  null,
  MODEL_CHEAP
);
  if (!apiResult.success) throw new Error(apiResult.message);

  var saveResult = saveStage15BH2s(apiResult.text);
  if (!saveResult.success) throw new Error(saveResult.message);

  bc_addToApiCostAndTime(apiResult.cost, (Date.now() - startTime) / 1000);
  return { success: true, message: saveResult.message, cost: apiResult.cost };
}

function bc_runEntityCoverageCheckAutomated() {

  var startTime = Date.now();
  var totalCost = 0;

  var prompt = buildEntityCoverageCheckPrompt();

  if (typeof prompt !== 'string' || prompt.indexOf('ERROR') === 0) {
    throw new Error(prompt);
  }

  var apiResult = bc_sendPromptViaOpenAI(
      prompt,
      null,
      MODEL_CHEAP
    );

  if (!apiResult.success) throw new Error(apiResult.message);

  totalCost += apiResult.cost;

  var saveResult = saveEntityCoverageCheckResult(apiResult.text);

  if (!saveResult.success) throw new Error(saveResult.message);

  if (saveResult.passed) {

    bc_addToApiCostAndTime(
      totalCost,
      (Date.now() - startTime) / 1000
    );

    return {
      success: true,
      message: saveResult.message,
      cost: totalCost
    };
  }

  return {
    success: true,
    passed: false,
    message: saveResult.message,
    cost: totalCost
  };
}

function bc_runRewriteBriefComplianceW15BAutomated() {

  var startTime = Date.now();
  var totalCost = 0;
  var attempt = 0;

  while (attempt <= 3) {

    var prompt = buildRewriteBriefComplianceCheckPromptW15B();

    if (typeof prompt !== 'string' || prompt.indexOf('ERROR') === 0) {
      throw new Error(prompt);
    }

    var apiResult = bc_sendPromptViaOpenAI(
        prompt,
        null,
        MODEL_CHEAP
      );

    if (!apiResult.success) {
      throw new Error(apiResult.message);
    }

    totalCost += apiResult.cost;

    var saveResult = saveRewriteBriefComplianceResultW15B(apiResult.text);

    if (!saveResult.success) {
      throw new Error(saveResult.message);
    }

    if (saveResult.passed) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

      return {
        success: true,
        message:
          attempt === 0
            ? saveResult.message
            : 'Passed after ' + attempt + ' correction attempt(s): ' + saveResult.message,
        cost: totalCost
      };
    }

    if (attempt === 3) {
      break;
    }

    attempt++;

    var fixPrompt = buildRewriteBriefFixPromptW15B();

    if (
      typeof fixPrompt !== 'string' ||
      fixPrompt.indexOf('ERROR') === 0
    ) {
      throw new Error(
        'FAIL saved, but fix prompt could not be built — ' + fixPrompt
      );
    }

    var fixResult = bc_sendPromptViaOpenAI(
      fixPrompt,
      null,
      MODEL_CHEAP
    );

    if (!fixResult.success) {
      throw new Error(fixResult.message);
    }

    totalCost += fixResult.cost;

    var saveH2sResult = saveStage15BH2s(fixResult.text);

    if (!saveH2sResult.success) {
      throw new Error(saveH2sResult.message);
    }
  }

   bc_addToApiCostAndTime(
    totalCost,
    (Date.now() - startTime) / 1000
  );

  bc_appendGovernancePipelineException(
    'W1.5B — Rewrite Brief Compliance',
    apiResult && apiResult.text
      ? apiResult.text
      : 'Still failing after 3 correction attempts.'
  );

  return {
    success: false,
    message:
      'Rewrite Brief Compliance still failing after 3 correction attempts — recorded in GJ.',
    cost: totalCost
  };
}

    function bc_runFullW15BAutomated() {

  var totalCost = 0;
  var log = [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();
  var cvCell = sheet.getRange(row, 100);


  // --------------------------------------------------
  // 1. H2 GENERATION
  // --------------------------------------------------

  try {

    var r1 = bc_runStage15BAutomated();

    totalCost += r1.cost;

    log.push(
      'H2 Gen: ' +
      r1.message
    );

  } catch (e) {

    bc_appendGovernancePipelineException(
      'W1.5B — H2 Generation',
      'Execution error — ' + e.message
    );

    log.push(
      'H2 Generation: ERROR recorded in GJ — ' +
      e.message
    );
  }


  // --------------------------------------------------
  // 2. ENTITY COVERAGE
  // --------------------------------------------------

  try {

    var r2 =
      bc_runEntityCoverageCheckAutomated();

    totalCost += r2.cost;

    log.push(
      'Coverage Check: ' +
      r2.message
    );

    var coverageAttempt = 0;

    while (
      r2.passed === false &&
      coverageAttempt < 3
    ) {

      coverageAttempt++;

      var fixPrompt =
        buildEntityCoverageFixPromptW15B();

      if (
        typeof fixPrompt !== 'string' ||
        fixPrompt.indexOf('ERROR') === 0
      ) {
        throw new Error(
          'Coverage FAIL, but fix prompt could not be built — ' +
          fixPrompt
        );
      }

      var fixResult =
        bc_sendPromptViaOpenAI(
          fixPrompt,
          null,
          MODEL_CHEAP
        );

      if (!fixResult.success) {
        throw new Error(
          fixResult.message
        );
      }

      totalCost +=
        fixResult.cost;

      var saveH2sResult =
        saveStage15BH2s(
          fixResult.text
        );

      if (!saveH2sResult.success) {
        throw new Error(
          saveH2sResult.message
        );
      }

      log.push(
        'Coverage Fix ' +
        coverageAttempt +
        ': corrected H2s saved to CV'
      );

      r2 =
        bc_runEntityCoverageCheckAutomated();

      totalCost += r2.cost;

      log.push(
        'Coverage Recheck ' +
        coverageAttempt +
        ': ' +
        r2.message
      );
    }

    if (r2.passed === false) {

      bc_appendGovernancePipelineException(
        'W1.5B — Entity Coverage',
        r2.rawResult ||
        r2.message ||
        'Still failing after 3 correction attempts.'
      );

      log.push(
        'Entity Coverage: final FAIL recorded in GJ.'
      );
    }

  } catch (e) {

    bc_appendGovernancePipelineException(
      'W1.5B — Entity Coverage',
      'Execution error — ' +
      e.message
    );

    log.push(
      'Entity Coverage: ERROR recorded in GJ — ' +
      e.message
    );
  }


  // --------------------------------------------------
  // 3. REWRITE BRIEF COMPLIANCE
  // --------------------------------------------------

  try {

    var r3 =
      bc_runRewriteBriefComplianceW15BAutomated();

    totalCost += r3.cost;

    log.push(
      'Compliance Check: ' +
      r3.message
    );

  } catch (e) {

    bc_appendGovernancePipelineException(
      'W1.5B — Rewrite Brief Compliance',
      'Execution error — ' +
      e.message
    );

    log.push(
      'Rewrite Brief Compliance: ERROR recorded in GJ — ' +
      e.message
    );
  }


  // --------------------------------------------------
  // REMEMBER H2s BEFORE RULE 17
  // --------------------------------------------------

  var h2sBeforeRule17 =
    String(
      cvCell.getValue() || ''
    ).trim();


  // --------------------------------------------------
  // 4. RULE 17
  // --------------------------------------------------

  try {

    var r4 =
      bc_runH2GovernanceW15BAutomated();

    totalCost += r4.cost;

    log.push(
      'H2 Governance: ' +
      r4.message
    );

  } catch (e) {

    bc_appendGovernancePipelineException(
      'W1.5B — H2 Heading Governance / Rule 17',
      'Execution error — ' +
      e.message
    );

    log.push(
      'H2 Governance / Rule 17: ERROR recorded in GJ — ' +
      e.message
    );
  }


  // --------------------------------------------------
  // CHECK WHETHER RULE 17 CHANGED THE H2s
  // --------------------------------------------------

  var h2sAfterRule17 =
    String(
      cvCell.getValue() || ''
    ).trim();

  var rule17ChangedH2s =
    h2sBeforeRule17 !==
    h2sAfterRule17;


  // --------------------------------------------------
  // 5. TECHNICAL SAFETY
  // --------------------------------------------------

  try {

    var rTech =
      bc_runH2TechnicalSafetyW15BAutomated();

    totalCost += rTech.cost;

    log.push(
      'H2 Technical Safety: ' +
      rTech.message
    );

  } catch (e) {

    bc_appendGovernancePipelineException(
      'W1.5B — H2 Technical & Scope Safety',
      'Execution error — ' +
      e.message
    );

    log.push(
      'H2 Technical & Scope Safety: ERROR recorded in GJ — ' +
      e.message
    );
  }


  var h2sAfterTechnicalSafety =
    String(
      cvCell.getValue() || ''
    ).trim();

  var technicalSafetyChangedH2s =
    h2sAfterRule17 !==
    h2sAfterTechnicalSafety;


  // --------------------------------------------------
  // 6. CROSS-CHECK AFTER HEADING FIXES
  // --------------------------------------------------

  var headingFixChangedH2s =
    rule17ChangedH2s ||
    technicalSafetyChangedH2s;


  if (headingFixChangedH2s) {


    // --------------------------------------------------
    // 6A. RULE 17 CHECK AFTER TECHNICAL SAFETY
    // --------------------------------------------------

    if (technicalSafetyChangedH2s) {

      try {

        var postTechRule17 =
          bc_runH2GovernanceCheckOnlyW15B();

        totalCost +=
          postTechRule17.cost;

        if (!postTechRule17.passed) {

          bc_appendGovernancePipelineException(
            'W1.5B — Post-Technical-Safety Rule 17',
            postTechRule17.rawResult ||
            'Technical Safety changed the H2s and the revised set no longer passes Rule 17.'
          );

          log.push(
            'Post-Technical-Safety Rule 17: FAIL recorded in GJ.'
          );

        } else {

          log.push(
            'Post-Technical-Safety Rule 17: ' +
            postTechRule17.message
          );
        }

      } catch (e) {

        bc_appendGovernancePipelineException(
          'W1.5B — Post-Technical-Safety Rule 17',
          'Execution error — ' +
          e.message
        );

        log.push(
          'Post-Technical-Safety Rule 17: ERROR recorded in GJ — ' +
          e.message
        );
      }
    }


    // --------------------------------------------------
    // 6B. ENTITY COVERAGE
    //     RECHECK + AUTO-REPAIR
    // --------------------------------------------------

    try {

      var postCoverage =
        bc_runEntityCoverageCheckAutomated();

      totalCost +=
        postCoverage.cost;

      log.push(
        'Post-Heading-Fix Coverage Check: ' +
        postCoverage.message
      );

      var postCoverageAttempt = 0;

      while (
        postCoverage.passed === false &&
        postCoverageAttempt < 3
      ) {

        postCoverageAttempt++;

        var postCoverageFixPrompt =
          buildEntityCoverageFixPromptW15B();

        if (
          typeof postCoverageFixPrompt !== 'string' ||
          postCoverageFixPrompt.indexOf('ERROR') === 0
        ) {
          throw new Error(
            'Post-heading-fix Coverage FAIL, but fix prompt could not be built — ' +
            postCoverageFixPrompt
          );
        }

        var postCoverageFix =
          bc_sendPromptViaOpenAI(
            postCoverageFixPrompt,
            null,
            MODEL_CHEAP
          );

        if (!postCoverageFix.success) {
          throw new Error(
            postCoverageFix.message
          );
        }

        totalCost +=
          postCoverageFix.cost;

        var postCoverageSave =
          saveStage15BH2s(
            postCoverageFix.text
          );

        if (!postCoverageSave.success) {
          throw new Error(
            postCoverageSave.message
          );
        }

        log.push(
          'Post-Heading-Fix Coverage Repair ' +
          postCoverageAttempt +
          ': corrected H2s saved to CV'
        );

        postCoverage =
          bc_runEntityCoverageCheckAutomated();

        totalCost +=
          postCoverage.cost;

        log.push(
          'Post-Heading-Fix Coverage Recheck ' +
          postCoverageAttempt +
          ': ' +
          postCoverage.message
        );
      }


      if (
        postCoverage.passed === false
      ) {

        bc_appendGovernancePipelineException(
          'W1.5B — Post-Heading-Fix Entity Coverage',
          postCoverage.rawResult ||
          postCoverage.message ||
          'Still failing after 3 post-heading-fix correction attempts.'
        );

        log.push(
          'Post-Heading-Fix Entity Coverage: final FAIL recorded in GJ.'
        );

      } else {

        log.push(
          'Post-Heading-Fix Entity Coverage: PASS.'
        );
      }

    } catch (e) {

      bc_appendGovernancePipelineException(
        'W1.5B — Post-Heading-Fix Entity Coverage',
        'Execution error — ' +
        e.message
      );

      log.push(
        'Post-Heading-Fix Entity Coverage: ERROR recorded in GJ — ' +
        e.message
      );
    }


    // --------------------------------------------------
    // 6C. REWRITE BRIEF COMPLIANCE
    //     RECHECK + AUTO-REPAIR
    // --------------------------------------------------

    try {

      var postBrief =
        bc_runRewriteBriefComplianceW15BAutomated();

      totalCost +=
        postBrief.cost;

      log.push(
        'Post-Heading-Fix Rewrite Brief Compliance: ' +
        postBrief.message
      );

      if (
        postBrief.success === false
      ) {

        log.push(
          'Post-Heading-Fix Rewrite Brief Compliance: final FAIL already recorded in GJ.'
        );

      } else {

        log.push(
          'Post-Heading-Fix Rewrite Brief Compliance: PASS.'
        );
      }

    } catch (e) {

      bc_appendGovernancePipelineException(
        'W1.5B — Post-Heading-Fix Rewrite Brief Compliance',
        'Execution error — ' +
        e.message
      );

      log.push(
        'Post-Heading-Fix Rewrite Brief Compliance: ERROR recorded in GJ — ' +
        e.message
      );
    }


    // --------------------------------------------------
    // 6D. FINAL RULE 17 CHECK
    //
    // Coverage / Brief repairs can alter H2 wording.
    // Recheck Rule 17 once more, but do NOT start an
    // endless correction loop.
    // --------------------------------------------------

    try {

      var finalRule17 =
        bc_runH2GovernanceCheckOnlyW15B();

      totalCost +=
        finalRule17.cost;

      if (!finalRule17.passed) {

        bc_appendGovernancePipelineException(
          'W1.5B — Final Rule 17 Recheck',
          finalRule17.rawResult ||
          'A post-heading-fix governance repair caused the final H2 set to fail Rule 17.'
        );

        log.push(
          'Final Rule 17 Recheck: FAIL recorded in GJ.'
        );

      } else {

        log.push(
          'Final Rule 17 Recheck: ' +
          finalRule17.message
        );
      }

    } catch (e) {

      bc_appendGovernancePipelineException(
        'W1.5B — Final Rule 17 Recheck',
        'Execution error — ' +
        e.message
      );

      log.push(
        'Final Rule 17 Recheck: ERROR recorded in GJ — ' +
        e.message
      );
    }
  }


  // --------------------------------------------------
  // SUCCESS
  // --------------------------------------------------

  var recheckMessage =
    headingFixChangedH2s
      ? 'Post-heading-fix checks and correction attempts completed.'
      : 'Post-heading-fix cross-checks were not required.';


  return {

    success: true,

    message:
      '✔ W1.5B sequence complete. ' +
      'Any unresolved governance failures have been recorded in GJ. ' +
      recheckMessage,

    log:
      log.join(' | '),

    cost:
      totalCost
  };
}

function bc_runH2GovernanceW15BAutomated() {
  var startTime = Date.now();
  var prompt = buildH2GovernanceCheckPromptW15B();
  if (typeof prompt !== 'string' || prompt.indexOf('ERROR') === 0) {
    throw new Error(prompt);
  }
  var apiResult = bc_sendPromptViaOpenAI(
      prompt,
      null,
      MODEL_CHEAP
    );
  if (!apiResult.success) throw new Error(apiResult.message);

  var saveResult = saveH2GovernanceCheckResultW15B(apiResult.text);
  if (!saveResult.success) throw new Error(saveResult.message);

  var totalCost = apiResult.cost;

  if (saveResult.passed) {
    bc_addToApiCostAndTime(totalCost, (Date.now() - startTime) / 1000);
    return { success: true, message: saveResult.message, cost: totalCost };
  }

  var fixPrompt = buildH2GovernanceFixPromptW15B();
  if (typeof fixPrompt !== 'string' || fixPrompt.indexOf('ERROR') === 0) {
    throw new Error('FAIL saved, but fix prompt could not be built — ' + fixPrompt);
  }
  var fixResult = bc_sendPromptViaOpenAI(
      fixPrompt,
      null,
      MODEL_CHEAP
    );
  if (!fixResult.success) throw new Error(fixResult.message);
  totalCost += fixResult.cost;

  var saveH2sResult = saveStage15BH2s(fixResult.text);
  if (!saveH2sResult.success) throw new Error(saveH2sResult.message);

  var recheckPrompt = buildH2GovernanceCheckPromptW15B();
  var recheckResult = bc_sendPromptViaOpenAI(
      recheckPrompt,
      null,
      MODEL_CHEAP
    );
  if (!recheckResult.success) throw new Error(recheckResult.message);
  totalCost += recheckResult.cost;

  var finalSave = saveH2GovernanceCheckResultW15B(recheckResult.text);

  if (!finalSave.success) throw new Error(finalSave.message);

  bc_addToApiCostAndTime(totalCost, (Date.now() - startTime) / 1000);

  if (!finalSave.passed) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();

  var cvCell = sheet.getRange(row, 100);
  var currentH2s = String(cvCell.getValue() || '').trim();

  var rawReport = String(finalSave.rawResult || '').trim();

  var blocks = rawReport.split(/\n(?=H2:)/);

  var failureBlocks = blocks.filter(function(block) {
  var testFail =
    /TEST\s*[1-4][^\n]*:\s*FAIL/i.test(block);

  var wholeSetFail =
    /PROBLEM:/i.test(block) ||
    /WHAT TO CHANGE:/i.test(block);

  return testFail || wholeSetFail;
  });

  var reportParts = [];

  failureBlocks.forEach(function(block) {

    var h2Match = block.match(/H2:\s*(.+)/i);
    var heading = h2Match ? h2Match[1].trim() : 'Unknown heading';

    var sectionLine = heading;

    currentH2s.split('\n').forEach(function(line) {
      if (
        line.indexOf('SECTION ') === 0 &&
        line.indexOf(heading) !== -1
      ) {
        sectionLine = line.trim();
      }
    });

    var problemMatch = block.match(
      /PROBLEM:\s*([\s\S]*?)(?=\nWHY THIS IS A PROBLEM:|\nWHAT TO CHANGE:|$)/i
    );

    var whyMatch = block.match(
      /WHY THIS IS A PROBLEM:\s*([\s\S]*?)(?=\nWHAT TO CHANGE:|$)/i
    );

    var changeMatch = block.match(
      /WHAT TO CHANGE:\s*([\s\S]*?)(?=\nSUGGESTED ALTERNATIVES:|$)/i
    );

    var alternativesMatch = block.match(
      /SUGGESTED ALTERNATIVES:\s*([\s\S]*?)$/i
    );

    var problem = problemMatch
      ? problemMatch[1].trim()
      : 'This heading still fails the H2 Heading Governance check.';

    var why = whyMatch
      ? whyMatch[1].trim()
      : 'The heading does not fully meet the Rule 17 requirements.';

    var change = changeMatch
      ? changeMatch[1].trim()
      : 'Review and rewrite this heading before continuing.';

    var alternatives = alternativesMatch
      ? alternativesMatch[1].trim()
      : 'No suggested alternatives were returned.';

    reportParts.push(
        sectionLine +
        '\n\nProblem:\n' + problem +
        '\n\nWhy this is a problem:\n' + why +
        '\n\nWhat to change:\n' + change +
        '\n\nSuggested alternatives:\n' + alternatives
      );
    });

    var failureReport =
      currentH2s +
      '\n\n-----------\n' +
      'FINAL FAILURE REPORT — H2 HEADING GOVERNANCE CHECK / RULE 17\n\n' +
      reportParts.join('\n\n-----------\n\n');

        bc_appendGovernancePipelineException(
      'W1.5B — H2 Heading Governance / Rule 17',
      failureReport
    );

    return {
      success: false,
      message:
        'Final H2 validation failed — Rule 17 failure report recorded in GJ.',
      cost: totalCost
    };
  }

  return {
    success: true,
    message: 'Initial FAIL — auto-fixed and rechecked: ' + finalSave.message,
    cost: totalCost
  };
}

function bc_runH2GovernanceCheckOnlyW15B() {
  var startTime = Date.now();

  var prompt = buildH2GovernanceCheckPromptW15B();

  if (
    typeof prompt !== 'string' ||
    prompt.indexOf('ERROR') === 0
  ) {
    throw new Error(prompt);
  }

  var apiResult = bc_sendPromptViaOpenAI(
    prompt,
    null,
    MODEL_CHEAP
  );

  if (!apiResult.success) {
    throw new Error(apiResult.message);
  }

  var saveResult =
    saveH2GovernanceCheckResultW15B(apiResult.text);

  if (!saveResult.success) {
    throw new Error(saveResult.message);
  }

  var totalCost = apiResult.cost;

  bc_addToApiCostAndTime(
    totalCost,
    (Date.now() - startTime) / 1000
  );

  return {
    success: true,
    passed: saveResult.passed,
    message: saveResult.message,
    rawResult: saveResult.rawResult,
    cost: totalCost
  };
}

function bc_appendH2GovernanceFailureReportW15B(rawReport) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select a data row first.');
  }

  var cvCell = sheet.getRange(row, 100);
  var currentH2s = String(cvCell.getValue() || '').trim();

  if (!currentH2s) {
    throw new Error('No H2 headings found in column CV.');
  }

  rawReport = String(rawReport || '').trim();

  if (!rawReport) {
    throw new Error('No Rule 17 failure result was supplied.');
  }

  var blocks = rawReport.split(/\n(?=H2:)/);

  var failureBlocks = blocks.filter(function(block) {
    var testFail =
      /TEST\s*[1-4][^\n]*:\s*FAIL/i.test(block);

    var wholeSetFail =
      /PROBLEM:/i.test(block) ||
      /WHAT TO CHANGE:/i.test(block);

    return testFail || wholeSetFail;
  });

  var reportParts = [];

  failureBlocks.forEach(function(block) {

    var h2Match = block.match(/H2:\s*(.+)/i);
    var heading = h2Match
      ? h2Match[1].trim()
      : 'Whole-set governance issue';

    var sectionLine = heading;

    currentH2s.split('\n').forEach(function(line) {
      if (
        line.indexOf('SECTION ') === 0 &&
        line.indexOf(heading) !== -1
      ) {
        sectionLine = line.trim();
      }
    });

    var problemMatch = block.match(
      /PROBLEM:\s*([\s\S]*?)(?=\nWHY THIS IS A PROBLEM:|\nWHAT TO CHANGE:|$)/i
    );

    var whyMatch = block.match(
      /WHY THIS IS A PROBLEM:\s*([\s\S]*?)(?=\nWHAT TO CHANGE:|$)/i
    );

    var changeMatch = block.match(
      /WHAT TO CHANGE:\s*([\s\S]*?)(?=\nSUGGESTED ALTERNATIVES:|$)/i
    );

    var alternativesMatch = block.match(
      /SUGGESTED ALTERNATIVES:\s*([\s\S]*?)$/i
    );

    var problem = problemMatch
      ? problemMatch[1].trim()
      : 'This heading still fails the H2 Heading Governance check.';

    var why = whyMatch
      ? whyMatch[1].trim()
      : 'The heading does not fully meet the Rule 17 requirements.';

    var change = changeMatch
      ? changeMatch[1].trim()
      : 'Review and rewrite this heading before continuing.';

    var alternatives = alternativesMatch
      ? alternativesMatch[1].trim()
      : 'No suggested alternatives were returned.';

    reportParts.push(
      sectionLine +
      '\n\nProblem:\n' + problem +
      '\n\nWhy this is a problem:\n' + why +
      '\n\nWhat to change:\n' + change +
      '\n\nSuggested alternatives:\n' + alternatives
    );
  });

  var failureReport =
    currentH2s +
    '\n\n-----------\n' +
    'FINAL FAILURE REPORT — H2 HEADING GOVERNANCE CHECK / RULE 17\n\n' +
    reportParts.join('\n\n-----------\n\n');

  cvCell.setNumberFormat('@');
  cvCell.setValue(failureReport);

  return {
    success: true,
    message: 'Rule 17 failure report appended to CV.'
  };
}

function bc_resumeW15BValidationAutomated() {
  var totalCost = 0;
  var log = [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();

  if (row < 2) {
    return {
      success: false,
      message: 'Select a data row first.',
      log: '',
      cost: 0
    };
  }

  var cvCell = sheet.getRange(row, 100);
  var currentH2s = String(cvCell.getValue() || '').trim();

  if (!currentH2s) {
    return {
      success: false,
      message: 'No H2 headings found in column CV.',
      log: '',
      cost: 0
    };
  }

  if (currentH2s.indexOf('FINAL FAILURE REPORT') !== -1) {
    return {
      success: false,
      message:
        'CV still contains a FINAL FAILURE REPORT. ' +
        'Edit the failed H2 headings and delete the failure report before resuming validation.',
      log: '',
      cost: 0
    };
  }

  try {
    var r1 = bc_runH2GovernanceCheckOnlyW15B();
    totalCost += r1.cost;
    log.push('Rule 17 Recheck: ' + r1.message);

    if (!r1.passed) {
      bc_appendH2GovernanceFailureReportW15B(r1.rawResult);

      return {
        success: false,
        message:
          'STOPPED at Rule 17 Recheck — manually corrected H2s still fail. ' +
          'See the failure report appended to CV.',
        log: log.join(' | '),
        cost: totalCost
      };
    }

  } catch (e) {
    return {
      success: false,
      message: 'STOPPED at Rule 17 Recheck — ' + e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }

  try {
    var r2 = bc_runH2TechnicalSafetyW15BAutomated();
    totalCost += r2.cost;
    log.push('Technical Safety: ' + r2.message);

  } catch (e) {
    return {
      success: false,
      message:
        'STOPPED at H2 Technical & Scope Safety Check — ' + e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }

  try {
    var r3 = bc_runEntityCoverageCheckAutomated();
    totalCost += r3.cost;
    log.push('Final Coverage Recheck: ' + r3.message);

    if (r3.passed === false) {

      var repairPrompt =
        buildEntityCoverageRepairRecommendationsW15B();

      if (
        typeof repairPrompt !== 'string' ||
        repairPrompt.indexOf('ERROR') === 0
      ) {
        throw new Error(
          'Manually corrected H2s failed Entity Coverage, but repair recommendations could not be built — ' +
          repairPrompt
        );
      }

      var repairResult = bc_sendPromptViaOpenAI(
        repairPrompt,
        null,
        MODEL_CHEAP
      );

      if (!repairResult.success) {
        throw new Error(repairResult.message);
      }

      totalCost += repairResult.cost;

      var cvCell = sheet.getRange(row, 100);
      var currentH2s = String(cvCell.getValue() || '').trim();

      var repairReport =
        currentH2s +
        '\n\n-----------\n' +
        'ENTITY COVERAGE REPAIR REPORT\n\n' +
        repairResult.text.trim();

      cvCell.setNumberFormat('@');
      cvCell.setValue(repairReport);

      return {
        success: false,
        message:
          'STOPPED at Final Entity Coverage Recheck — repair recommendations have been appended to CV.',
        log: log.join(' | '),
        cost: totalCost
      };
    }

  } catch (e) {
    return {
      success: false,
      message:
        'STOPPED at Final Entity Coverage Recheck — ' + e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }

  try {
    var finalBriefPrompt =
      buildRewriteBriefComplianceCheckPromptW15B();

    if (
      typeof finalBriefPrompt !== 'string' ||
      finalBriefPrompt.indexOf('ERROR') === 0
    ) {
      throw new Error(finalBriefPrompt);
    }

    var finalBriefResult = bc_sendPromptViaOpenAI(
      finalBriefPrompt,
      null,
      MODEL_CHEAP
    );

    if (!finalBriefResult.success) {
      throw new Error(finalBriefResult.message);
    }

    totalCost += finalBriefResult.cost;

    var finalBriefSave =
      saveRewriteBriefComplianceResultW15B(
        finalBriefResult.text
      );

    if (!finalBriefSave.success) {
      throw new Error(finalBriefSave.message);
    }

    if (!finalBriefSave.passed) {
      throw new Error(
        'Manually corrected H2s failed Rewrite Brief Compliance.'
      );
    }

    log.push(
      'Final Rewrite Brief Recheck: ' +
      finalBriefSave.message
    );

  } catch (e) {
    return {
      success: false,
      message:
        'STOPPED at Final Rewrite Brief Recheck — ' + e.message,
      log: log.join(' | '),
      cost: totalCost
    };
  }

  return {
    success: true,
    message:
      '✔ W1.5B resume validation complete.\n' +
      'Rule 17 Recheck: PASS\n' +
      'H2 Technical & Scope Safety: PASS\n' +
      'Final Coverage Recheck: PASS\n' +
      'Final Rewrite Brief Recheck: PASS',
    log: log.join(' | '),
    cost: totalCost
  };
}

function bc_runH2TechnicalSafetyW15BAutomated() {
  var startTime = Date.now();
  var totalCost = 0;
  var maxFixAttempts = 2;
  var attempt = 0;
  var lastFailureText = '';

  while (attempt <= maxFixAttempts) {

    var prompt = buildH2TechnicalSafetyCheckPromptW15B();

    if (
      typeof prompt !== 'string' ||
      prompt.indexOf('ERROR') === 0
    ) {
      throw new Error(prompt);
    }

    var apiResult = bc_sendPromptViaOpenAI(
      prompt,
      null,
      MODEL_CHEAP
    );

    if (!apiResult.success) {
      throw new Error(apiResult.message);
    }

    totalCost += apiResult.cost;

    var saveResult =
      saveH2TechnicalSafetyCheckResultW15B(
        apiResult.text
      );

    if (!saveResult.success) {
      throw new Error(saveResult.message);
    }

    if (saveResult.passed) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

      return {
        success: true,
        message:
          attempt === 0
            ? saveResult.message
            : 'H2 Technical & Scope Safety PASS after ' +
              attempt +
              ' correction attempt(s).',
        cost: totalCost
      };
    }

    lastFailureText = apiResult.text;

    if (attempt === maxFixAttempts) {
      break;
    }

    attempt++;

    var fixPrompt =
      buildH2TechnicalSafetyFixPromptW15B(
        lastFailureText
      );

    if (
      typeof fixPrompt !== 'string' ||
      fixPrompt.indexOf('ERROR') === 0
    ) {
      throw new Error(
        'Technical Safety FAIL, but fix prompt could not be built — ' +
        fixPrompt
      );
    }

    var fixResult = bc_sendPromptViaOpenAI(
      fixPrompt,
      null,
      MODEL_CHEAP
    );

    if (!fixResult.success) {
      throw new Error(fixResult.message);
    }

    totalCost += fixResult.cost;

    var saveH2sResult =
      saveStage15BH2s(fixResult.text);

    if (!saveH2sResult.success) {
      throw new Error(saveH2sResult.message);
    }
  }

  bc_addToApiCostAndTime(
    totalCost,
    (Date.now() - startTime) / 1000
  );

  bc_appendGovernancePipelineException(
    'W1.5B — H2 Technical & Scope Safety',
    lastFailureText ||
      (
        'Still failing after ' +
        maxFixAttempts +
        ' automated correction attempts.'
      )
  );

  return {
    success: false,
    message:
      'H2 Technical & Scope Safety still FAILING after ' +
      maxFixAttempts +
      ' automated correction attempts — recorded in GJ.',
    cost: totalCost
  };
}

function bc_runStage15CAutomated() {
  var startTime = Date.now();
  var totalCost = 0;
  var maxFixAttempts = 2;
  var attempt = 0;
  var lastOutput = '';
  var lastFailure = '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select a data row first.');
  }

  // Clear stale W1.5C output before rebuilding from current CU/CV.
  sheet.getRange(row, 101).clearContent();

  var prompt = buildStage15CPrompt();

  if (
    typeof prompt !== 'string' ||
    prompt.indexOf('ERROR') === 0
  ) {
    throw new Error(prompt);
  }

  var apiResult = bc_sendPromptViaOpenAI(
    prompt,
    null,
    MODEL_CHEAP
  );

  if (!apiResult.success) {
    throw new Error(apiResult.message);
  }

  totalCost += apiResult.cost;
  lastOutput = apiResult.text;

  while (attempt <= maxFixAttempts) {

    var saveResult =
      saveStage15CEnrichedPlan(
        lastOutput
      );

    if (saveResult.success) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

      return {
        success: true,
        message:
          attempt === 0
            ? saveResult.message
            : 'W1.5C PASS after ' +
              attempt +
              ' correction attempt(s). ' +
              saveResult.message,
        cost: totalCost
      };
    }

    lastFailure =
      saveResult.message;

    if (
      attempt ===
      maxFixAttempts
    ) {
      break;
    }

    attempt++;

    var repairPrompt =
      buildStage15CRepairPrompt(
        lastOutput,
        lastFailure
      );

    if (
      typeof repairPrompt !== 'string' ||
      repairPrompt.indexOf('ERROR') === 0
    ) {
      throw new Error(
        'W1.5C validation failed, but repair prompt could not be built — ' +
        repairPrompt
      );
    }

    var repairResult =
      bc_sendPromptViaOpenAI(
        repairPrompt,
        null,
        MODEL_CHEAP
      );

    if (
      !repairResult.success
    ) {
      throw new Error(
        repairResult.message
      );
    }

    totalCost +=
      repairResult.cost;

    lastOutput =
      repairResult.text;
  }

  bc_addToApiCostAndTime(
    totalCost,
    (Date.now() - startTime) / 1000
  );

  bc_appendGovernancePipelineException(
    'W1.5C — Enriched Plan',
    lastFailure ||
      (
        'Still failing after ' +
        maxFixAttempts +
        ' automated correction attempts.'
      )
  );

  return {
    success: false,
    message:
      'W1.5C still FAILING after ' +
      maxFixAttempts +
      ' automated correction attempts — recorded in GJ.',
    cost: totalCost
  };
}

function bc_runStage15DAutomated() {
  var startTime = Date.now();
  var totalCost = 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select a data row first.');
  }

  // Clear stale W1.5D output before rebuilding.
  sheet.getRange(row, 102).clearContent();

  var data = getW15DData();

  if (data.error) {
    throw new Error(data.error);
  }

  var failureMessage = '';

  var prompt = buildW15DPrompt();

  if (
    typeof prompt !== 'string' ||
    prompt.indexOf('ERROR') === 0
  ) {

    failureMessage =
      'W1.5D selection prompt could not be built — ' +
      prompt;

  } else {

    var apiResult =
      bc_sendPromptViaOpenAI(
        prompt,
        null,
        MODEL_CHEAP
      );

    if (apiResult.success) {

      totalCost += Number(
        apiResult.cost || 0
      );

      var saveResult =
        saveW15DUpdatedPlan(
          apiResult.text
        );

      if (saveResult.success) {

        bc_addToApiCostAndTime(
          totalCost,
          (Date.now() - startTime) / 1000
        );

        return {
          success: true,
          cost: totalCost,
          message: saveResult.message
        };
      }

      failureMessage =
        saveResult.message;

    } else {

      failureMessage =
        'W1.5D API call failed — ' +
        apiResult.message;
    }
  }

  /*
   * The AI selection failed.
   * Preserve the governed W1.5C plan in CX
   * so the pipeline can safely continue.
   */
  var fallbackResult =
    saveW15DFallbackPlan();

  bc_addToApiCostAndTime(
    totalCost,
    (Date.now() - startTime) / 1000
  );

  bc_appendGovernancePipelineException(
    'W1.5D — Lateral Links',
    failureMessage +
    ' | Safe fallback used; pipeline continued.'
  );

  if (!fallbackResult.success) {
    return {
      success: false,
      cost: totalCost,
      message:
        'W1.5D fallback also failed — ' +
        fallbackResult.message
    };
  }

  bc_appendPipelineRunLog(
    row,
    '⚠ W1.5D lateral-link selection failed — governed W1.5C plan copied safely to CX and pipeline continued.'
  );

  return {
    success: true,
    fallback: true,
    cost: totalCost,
    message:
      'W1.5D fallback used — governed W1.5C plan preserved and pipeline continued.'
  };
}

function bc_addToApiCostAndTime(cost, seconds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("posts");
  var row = ss.getActiveRange().getRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).trim(); });
  var costIdx = headers.indexOf("API Cost");
  var timeIdx = headers.indexOf("API Time");
  if (costIdx > -1) {
    var existingCost = parseFloat(sheet.getRange(row, costIdx + 1).getValue()) || 0;
    sheet.getRange(row, costIdx + 1).setValue(existingCost + cost);
  }
  if (timeIdx > -1 && seconds) {
    var existingTimeRaw = String(sheet.getRange(row, timeIdx + 1).getValue() || '0').replace('s', '');
    var existingTime = parseFloat(existingTimeRaw) || 0;
    sheet.getRange(row, timeIdx + 1).setValue((existingTime + seconds).toFixed(1) + 's');
  }
}

function bc_runStage15EAutomated() {
  var startTime = Date.now();
  var totalCost = 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select a data row first.');
  }

  // Clear stale W1.5E output before rebuilding from current CX.
  sheet.getRange(row, 103).clearContent();

  var prompt = buildStage15EPrompt();

  if (
    typeof prompt !== 'string' ||
    prompt.indexOf('ERROR') === 0
  ) {
    throw new Error(prompt);
  }

  var apiResult = bc_sendPromptViaOpenAI(
    prompt,
    null,
    MODEL_CHEAP
  );

  if (!apiResult.success) {
    throw new Error(apiResult.message);
  }

  totalCost += apiResult.cost;

  var saveResult =
    saveStage15EOptimizations(
      apiResult.text
    );

  bc_addToApiCostAndTime(
    totalCost,
    (Date.now() - startTime) / 1000
  );

  if (!saveResult.success) {

    bc_appendGovernancePipelineException(
      'W1.5E — Google Optimizations',
      saveResult.message
    );

    return {
      success: false,
      message:
        'W1.5E validation failed — CY left unchanged and failure recorded in GJ.',
      cost: totalCost
    };
  }

  return {
    success: true,
    message: saveResult.message,
    cost: totalCost
  };
}

function bc_runFactCheckAutomated() {
  var startTime = Date.now();
  var promptData = buildFactCheckPrompt('EU');
  if (!promptData.success) throw new Error(promptData.message);

  var apiResult = bc_sendPromptViaOpenAI(
  promptData.prompt,
  2000,
  MODEL_CHEAP
);
  if (!apiResult.success) throw new Error(apiResult.message);

  bc_addToApiCostAndTime(apiResult.cost, (Date.now() - startTime) / 1000);
  return { success: true, text: apiResult.text, cost: apiResult.cost };
}

function bc_runFactCheckFixAutomated(findingsText) {
  var startTime = Date.now();
  if (!findingsText || !findingsText.trim()) {
    throw new Error('No fact-check findings provided — run the fact-check first.');
  }

  var promptData = buildFactCheckFixPrompt(findingsText, 'EU');
  if (!promptData.success) throw new Error(promptData.message);

  var apiResult = bc_sendPromptViaOpenAI(
    promptData.prompt,
    4000,
    MODEL_CHEAP
  );
  if (!apiResult.success) throw new Error(apiResult.message);

  var applyResult = applyFactCheckFix(apiResult.text, 'EU');

  bc_addToApiCostAndTime(apiResult.cost, (Date.now() - startTime) / 1000);
  return { success: applyResult.success, message: applyResult.message, cost: apiResult.cost };
}

function ce_findInternalFactConsistencyFinding_(html) {

  html = String(html || '');

  var matches = [];
  var regex =
    /\b(nearly|almost|over|more than|about|around|approximately)\s+(\d+)\s+years\b/gi;

  var match;

  while ((match = regex.exec(html)) !== null) {

    var phrase = match[0];

    if (matches.indexOf(phrase) === -1) {
      matches.push(phrase);
    }
  }

  if (matches.length <= 1) {
    return '';
  }

  return (
    'FACTUAL CONTRADICTION\n' +
    'The article contains conflicting versions of the same experience-duration claim: ' +
    matches.map(function(item) {
      return '"' + item + '"';
    }).join(', ') +
    '. Standardise these references to one internally consistent formulation. ' +
    'Make only the minimum exact-fragment replacements required. ' +
    'Do not rewrite surrounding paragraphs or alter unrelated facts.'
  );
}

function bc_runFactCheckFullAutomated() {

  var totalCost = 0;
  var maxFixAttempts = 3;
  var attempt = 0;
  var lastFixMessage = '';

  // Start every automated W2B.05 session from the CURRENT EU.
  // This prevents an old GE from a previous test/run being re-used.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('posts');
  var row = sh.getActiveCell().getRow();

  var currentEU = ce_stripHtmlCodeFences_(
    sh.getRange(row, 151).getValue()
  );

  if (!currentEU) {
    throw new Error(
      'Column EU is empty — W2B must complete before W2B.05 Fact Check.'
    );
  }

  sh.getRange(row, 187).setValue(currentEU); // GE — fresh working copy

  while (attempt <= maxFixAttempts) {

    var checkResult = bc_runFactCheckAutomated();

    if (!checkResult.success) {
      throw new Error(checkResult.message);
    }

    totalCost += checkResult.cost;

  // HARD MATERIAL-CATEGORY CHECK
  // Run this independently of the AI fact-check result.
  var currentFactCheckHtml = String(
    sh.getRange(row, 187).getValue() || ''
  ).trim();

  var materialIntegrity =
  ce_checkMaterialCategoryIntegrity_(currentFactCheckHtml);

  var consistencyFinding =
    ce_findInternalFactConsistencyFinding_(
      currentFactCheckHtml
    );

  var consistencyOnlyFail = false;

  if (!materialIntegrity.passed) {

    checkResult.text =
      materialIntegrity.finding +
      '\n\nOVERALL: FAIL (1 issue found)';

  } else if (consistencyFinding) {

    consistencyOnlyFail = true;

    checkResult.text =
      consistencyFinding +
      '\n\nOVERALL: FAIL (1 issue found)';
  }

  var isFail = /OVERALL:\s*FAIL/i.test(checkResult.text);

    // PASS
    if (!isFail) {

      // Only copy EU to GE when it passed first time.
      // After a fix, GE already contains the corrected HTML.
      if (attempt === 0) {

        var passResult = passThroughFactCheckToGE();

        if (!passResult.success) {
          throw new Error(passResult.message);
        }
      }

      return {
        success: true,
        message:
          attempt === 0
            ? 'Fact-check PASS — no changes needed. Saved in column GE.'
            : 'Fact-check PASS after ' +
              attempt +
              ' automated fix attempt(s). ' +
              lastFixMessage +
              ' Final corrected HTML saved in column GE.',
        text: checkResult.text,
        cost: totalCost
      };
    }

        // FAIL after maximum attempts — record exception and allow pipeline to continue
    if (attempt === maxFixAttempts) {

    bc_appendGovernancePipelineException(
      'W2B.05 — Fact Check',
      checkResult.text
    );

    if (consistencyOnlyFail) {

      return {
        success: true,
        warning: true,
        message:
          'Internal fact consistency issue remains after ' +
          maxFixAttempts +
          ' automated fix attempts — recorded in GJ and pipeline continuing.',
        text: checkResult.text,
        cost: totalCost
      };
    }

    return {
      success: false,
      message:
        'Fact-check still FAILING after ' +
        maxFixAttempts +
        ' automated fix attempts — recorded in GJ.',
      text: checkResult.text,
      cost: totalCost
    };
  }

    attempt++;

    var fixResult =
      bc_runFactCheckFixAutomated(checkResult.text);

    totalCost += fixResult.cost;

    lastFixMessage = fixResult.message;

        if (!fixResult.success) {

  bc_appendGovernancePipelineException(
    'W2B.05 — Fact Check',
    'Automated fix could not be applied — ' + fixResult.message
      );

      if (consistencyOnlyFail) {

        return {
          success: true,
          warning: true,
          message:
            'Internal fact consistency repair could not be applied safely — recorded in GJ and pipeline continuing.',
          text: checkResult.text,
          cost: totalCost
        };
      }

      return {
        success: false,
        message:
          'Fact-check FAIL detected, but automated fix could not be applied — recorded in GJ.',
        text: checkResult.text,
        cost: totalCost
      };
    }
  }
}


function bc_runSimilarityCheckAutomated() {

  var startTime = Date.now();

  var promptData = buildCaseStudySimilarityPrompt('GE');

  if (!promptData.success) throw new Error(promptData.message);

  var apiResult = bc_sendPromptViaOpenAI(
    promptData.prompt,
    2000,
    MODEL_CHEAP
  );

  if (!apiResult.success) throw new Error(apiResult.message);

  bc_addToApiCostAndTime(
    apiResult.cost,
    (Date.now() - startTime) / 1000
  );

  return {
    success: true,
    text: apiResult.text,
    cost: apiResult.cost
  };
}

function bc_runSimilarityFixAutomated(findingsText) {

  var startTime = Date.now();

  if (!findingsText || !findingsText.trim()) {
    throw new Error(
      'No similarity check findings provided — run the check first.'
    );
  }

  var promptData =
    buildSimilarityFixPrompt(findingsText, 'GE');

  if (!promptData.success) throw new Error(promptData.message);

  var apiResult = bc_sendPromptViaOpenAI(
    promptData.prompt,
    4000,
    MODEL_CHEAP
  );

  if (!apiResult.success) throw new Error(apiResult.message);

  var applyResult =
    applySimilarityFix(apiResult.text, 'GE');

  bc_addToApiCostAndTime(
    apiResult.cost,
    (Date.now() - startTime) / 1000
  );

  return {
    success: applyResult.success,
    message: applyResult.message,
    cost: apiResult.cost
  };
}

function bc_runSimilarityFullAutomated() {

  // FIRST: check article type before making any API call
  var d = getActiveRowDataMap();
  var articleType = String(d["Article Type"] || "").trim();

  // Similarity Check only applies to Case Studies
  if (articleType !== 'Case Study') {

    var skipResult = passThroughSimilarityCheckToGF();

    if (!skipResult.success) {
      throw new Error(skipResult.message);
    }

    return {
      success: true,
      message:
        'Similarity check skipped — Article Type is "' +
        articleType +
        '". GE copied unchanged to column GF. No API cost.',
      text: '',
      cost: 0
    };
  }


  // CASE STUDY — run similarity check normally
  var totalCost = 0;
  var maxFixAttempts = 3;
  var attempt = 0;
  var lastFixMessage = '';

  while (attempt <= maxFixAttempts) {

    var checkResult =
      bc_runSimilarityCheckAutomated();

    if (!checkResult.success) {
      throw new Error(checkResult.message);
    }

    totalCost += checkResult.cost;

    var isFail =
      /SEVERITY:\s*GENUINE RISK/i.test(checkResult.text);

    // PASS
    if (!isFail) {

      // If it passed first time, copy GE unchanged to GF.
      // After a fix, GF already contains the corrected HTML.
      if (attempt === 0) {

        var passResult =
          passThroughSimilarityCheckToGF();

        if (!passResult.success) {
          throw new Error(passResult.message);
        }
      }

      return {
        success: true,
        message:
          attempt === 0
            ? 'Similarity check PASS — no changes needed. GE copied unchanged to column GF.'
            : 'Similarity check PASS after ' +
              attempt +
              ' automated fix attempt(s). ' +
              lastFixMessage +
              ' Final corrected HTML saved in column GF.',
        text: checkResult.text,
        cost: totalCost
      };
    }

        // Still failing after maximum attempts
      if (attempt === maxFixAttempts) {

        bc_appendGovernancePipelineException(
          'W2B.1 — Similarity',
          checkResult.text ||
            (
              'Still showing GENUINE RISK after ' +
              maxFixAttempts +
              ' automated fix attempts.'
            )
        );

        return {
          success: false,
          message:
            'Similarity check still showing GENUINE RISK after ' +
            maxFixAttempts +
            ' automated fix attempts — recorded in GJ.',
          text: checkResult.text,
          cost: totalCost
        };
      }

    attempt++;

    var fixResult =
      bc_runSimilarityFixAutomated(checkResult.text);

    totalCost += fixResult.cost;

    lastFixMessage = fixResult.message;

        if (!fixResult.success) {

        bc_appendGovernancePipelineException(
          'W2B.1 — Similarity',
          'Similarity risk detected, but automated fix could not be applied — ' +
          fixResult.message +
          '\n\nFINAL CHECK REPORT:\n' +
          (checkResult.text || '')
        );

        return {
          success: false,
          message:
            'Similarity risk detected, but automated fix could not be applied — recorded in GJ.',
          text: checkResult.text,
          cost: totalCost
        };
      }
  }
}

function bc_runRewriteBriefComplianceW2BAutomated() {

  var startTime = Date.now();
  var totalCost = 0;
  var maxFixAttempts = 3;
  var fixAttempt = 0;
  var lastMessage = '';

  // Start W2B.2 from the CURRENT upstream HTML, not an old GG.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('posts');
  var row = sh.getActiveCell().getRow();

  var currentHtml = String(sh.getRange(row, 188).getValue() || '').trim(); // GF
  if (!currentHtml) {
    currentHtml = String(sh.getRange(row, 187).getValue() || '').trim(); // GE
  }

  if (!currentHtml) {
    throw new Error('No current HTML found in GF or GE — previous W2B stage must complete first.');
  }

  sh.getRange(row, 189).setValue(currentHtml); // GG — fresh W2B.2 working copy

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function(h) {
      return String(h).trim();
    });

  var complianceCol = headers.indexOf('Rewrite Brief Compliance Check W2B');

  if (complianceCol > -1) {
    sh.getRange(row, complianceCol + 1).clearContent();
  }

  while (fixAttempt <= maxFixAttempts) {

    var prompt = buildRewriteBriefComplianceCheckPromptW2B();

    if (typeof prompt !== 'string' || prompt.indexOf('ERROR') === 0) {
      throw new Error(prompt);
    }

    var isRecheck =
      prompt.indexOf('REWRITE BRIEF COMPLIANCE RECHECK (SCOPED') > -1;

    var apiResult = bc_sendPromptViaOpenAI(
      prompt,
      3000,
      MODEL_CHEAP
    );

    if (!apiResult.success) {
      throw new Error(apiResult.message);
    }

    totalCost += apiResult.cost;

    var saveResult = isRecheck
      ? saveRewriteBriefComplianceRecheckResultW2B(apiResult.text)
      : saveRewriteBriefComplianceResultW2B(apiResult.text);

    if (!saveResult.success) {
      throw new Error(saveResult.message);
    }

    if (saveResult.passed) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

      return {
        success: true,
        message:
          fixAttempt === 0
            ? saveResult.message
            : 'Compliance PASS after ' +
              fixAttempt +
              ' automated fix attempt(s). ' +
              lastMessage,
        cost: totalCost
      };
    }

    if (fixAttempt === maxFixAttempts) {

    bc_addToApiCostAndTime(
      totalCost,
      (Date.now() - startTime) / 1000
    );

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('posts');
    var row = sh.getActiveCell().getRow();

    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });

    var reportCol =
      headers.indexOf('Rewrite Brief Compliance Check W2B');

    var finalComplianceReport = '';

    if (reportCol > -1) {
      finalComplianceReport =
        String(sh.getRange(row, reportCol + 1).getValue() || '').trim();
    }

    bc_appendGovernancePipelineException(
      'W2B.2 — Rewrite Brief Compliance',
      finalComplianceReport ||
        (
          'Still failing after ' +
          maxFixAttempts +
          ' automated fix attempts.'
        )
    );

    return {
      success: false,
      message:
        'Compliance still FAILING after ' +
        maxFixAttempts +
        ' automated fix attempts — recorded in GJ.',
      cost: totalCost
    };
  }

    fixAttempt++;

    var fixPrompt = buildRewriteBriefFixPromptW2B();

    if (
      typeof fixPrompt !== 'string' ||
      fixPrompt.indexOf('ERROR') === 0
    ) {
      throw new Error(
        'FAIL saved, but fix prompt could not be built — ' +
        fixPrompt
      );
    }

    var fixResult = bc_sendPromptViaOpenAI(
      fixPrompt,
      6000,
      MODEL_CHEAP
    );

    if (!fixResult.success) {
      throw new Error(fixResult.message);
    }

    totalCost += fixResult.cost;

    var applyResult =
  applyRewriteBriefFixW2B(fixResult.text);

    if (!applyResult.success) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

      bc_appendGovernancePipelineException(
        'W2B.2 — Rewrite Brief Compliance Fix',
        'Fix could not be applied safely — ' + applyResult.message
      );

      return {
        success: false,
        message:
          'Rewrite Brief Compliance fix could not be applied safely — recorded in GJ.',
        cost: totalCost
      };
    }

    lastMessage = applyResult.message;
  }
}

function bc_runCoherenceFullAutomated() {

  var startTime = Date.now();
  var totalCost = 0;

  /*
   * W2B.3 SINGLE-PASS DESIGN
   *
   * GG = completed W2B.2 HTML
   * GH = single coherence audit JSON
   * GI = W2B.3 working/final HTML
   *
   * Flow:
   *
   * 1. Copy GG to GI.
   * 2. Run ONE whole-article coherence audit.
   * 3. If PASS, finish.
   * 4. If FAIL, generate ALL required repairs in ONE repair call.
   * 5. Apply those repairs.
   * 6. Finish.
   *
   * There is NO coherence recheck.
   * There is NO repair loop.
   */

  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  var sh =
    ss.getSheetByName('posts');

  var row =
    sh.getActiveCell().getRow();

  if (row < 2) {
    throw new Error(
      'Select a data row before running W2B.3.'
    );
  }

  var currentGG =
    ce_stripHtmlCodeFences_(
      sh.getRange(row, 189).getValue()
    );

  if (!currentGG) {
    throw new Error(
      'Column GG is empty — W2B.2 must complete before W2B.3.'
    );
  }

  /*
   * Always start from the current W2B.2 output.
   */
  sh
    .getRange(row, 190)
    .clearContent(); // GH

  sh
    .getRange(row, 191)
    .setValue(currentGG); // GI

  /*
   * =========================================================
   * STEP 1 — ONE WHOLE-ARTICLE COHERENCE AUDIT
   * =========================================================
   */

  var promptData =
    buildCoherenceCheckPrompt();

  if (
    !promptData ||
    !promptData.success
  ) {
    throw new Error(
      promptData &&
      promptData.message
        ? promptData.message
        : 'Could not build coherence check prompt.'
    );
  }

  var checkResult =
    bc_sendPromptViaOpenAI(
      promptData.prompt,
      5000,
      MODEL_CHEAP
    );

  if (!checkResult.success) {
    throw new Error(
      checkResult.message
    );
  }

  totalCost +=
    Number(
      checkResult.cost || 0
    );

  /*
   * =========================================================
   * STEP 2 — NORMALISE AND PARSE THE AUDIT JSON
   * =========================================================
   */

  var checkJsonText =
    String(
      checkResult.text || ''
    )
      .trim()
      .replace(
        /^```json\s*/i,
        ''
      )
      .replace(
        /^```\s*/i,
        ''
      )
      .replace(
        /\s*```$/i,
        ''
      )
      .trim();

  /*
   * If the model accidentally adds text outside the JSON,
   * recover the outer JSON object where possible.
   */
  var firstBrace =
    checkJsonText.indexOf('{');

  var lastBrace =
    checkJsonText.lastIndexOf('}');

  if (
    firstBrace > -1 &&
    lastBrace > firstBrace
  ) {
    checkJsonText =
      checkJsonText.substring(
        firstBrace,
        lastBrace + 1
      );
  }

  var checkJson;

  try {

    checkJson =
      JSON.parse(
        checkJsonText
      );

  } catch (e) {

    /*
     * Coherence is advisory.
     *
     * A malformed model response must not destroy
     * the completed W2B.2 article or trap the
     * governance pipeline.
     *
     * Keep GG already copied into GI and record
     * the exception for later review.
     */

    sh
      .getRange(row, 190)
      .setValue(
        'W2B.3 single-pass audit returned unusable JSON:\n\n' +
        String(
          checkResult.text || ''
        )
      );

    bc_appendGovernancePipelineException(
      'W2B.3 — Coherence',
      'Single-pass coherence audit returned invalid JSON. ' +
      'Article was passed through unchanged to GI. ' +
      e.message
    );

    bc_addToApiCostAndTime(
      totalCost,
      (
        Date.now() -
        startTime
      ) / 1000
    );

    return {
      success: true,
      message:
        'W2B.3 completed with advisory exception — audit JSON was unusable, so W2B.2 HTML was retained unchanged in GI.',
      cost: totalCost
    };
  }

  /*
   * =========================================================
   * STEP 3 — VALIDATE BASIC AUDIT STRUCTURE
   * =========================================================
   */

  var status =
    String(
      checkJson &&
      checkJson.status
        ? checkJson.status
        : ''
    )
      .trim()
      .toUpperCase();

  var issues =
    checkJson &&
    Array.isArray(
      checkJson.issues
    )
      ? checkJson.issues
      : null;

  if (
    (
      status !== 'PASS' &&
      status !== 'FAIL'
    ) ||
    issues === null
  ) {

    sh
      .getRange(row, 190)
      .setValue(
        JSON.stringify(
          checkJson,
          null,
          2
        )
      );

    bc_appendGovernancePipelineException(
      'W2B.3 — Coherence',
      'Single-pass coherence audit returned an unexpected JSON structure. ' +
      'Article was passed through unchanged to GI.'
    );

    bc_addToApiCostAndTime(
      totalCost,
      (
        Date.now() -
        startTime
      ) / 1000
    );

    return {
      success: true,
      message:
        'W2B.3 completed with advisory exception — unexpected audit structure, so W2B.2 HTML was retained unchanged in GI.',
      cost: totalCost
    };
  }

  /*
   * Save the ONE audit result in GH.
   */
  sh
    .getRange(row, 190)
    .setNumberFormat('@');

  sh
    .getRange(row, 190)
    .setValue(
      JSON.stringify(
        checkJson,
        null,
        2
      )
    );

  /*
   * =========================================================
   * STEP 4 — PASS
   * =========================================================
   */

  if (
    status === 'PASS' ||
    issues.length === 0
  ) {

    bc_addToApiCostAndTime(
      totalCost,
      (
        Date.now() -
        startTime
      ) / 1000
    );

    return {
      success: true,
      message:
        'Coherence single-pass audit found no issues — W2B.2 HTML retained in GI.',
      cost: totalCost
    };
  }

  /*
   * =========================================================
   * STEP 5 — BUILD ONE REPAIR REQUEST FOR ALL ISSUES
   * =========================================================
   */

  var fixPromptData =
    buildCoherenceFixPrompt(
      JSON.stringify(
        checkJson
      )
    );

  if (
    !fixPromptData ||
    !fixPromptData.success
  ) {

    bc_appendGovernancePipelineException(
      'W2B.3 — Coherence Repair',
      fixPromptData &&
      fixPromptData.message
        ? fixPromptData.message
        : 'Could not build single-pass coherence repair prompt.'
    );

    bc_addToApiCostAndTime(
      totalCost,
      (
        Date.now() -
        startTime
      ) / 1000
    );

    return {
      success: true,
      message:
        'W2B.3 audit completed, but the repair prompt could not be built — article retained unchanged in GI.',
      cost: totalCost
    };
  }

  /*
   * =========================================================
   * STEP 6 — ONE AI REPAIR CALL
   * =========================================================
   */

  var repairResult =
    bc_sendPromptViaOpenAI(
      fixPromptData.prompt,
      8000,
      MODEL_CHEAP
    );

  if (!repairResult.success) {

    /*
     * Actual API failure remains a genuine execution error.
     */
    throw new Error(
      repairResult.message
    );
  }

  totalCost +=
    Number(
      repairResult.cost || 0
    );

  /*
   * =========================================================
   * STEP 7 — APPLY ALL SECTION REPLACEMENTS
   * =========================================================
   */

  var applyResult =
    applyCoherenceFix(
      repairResult.text
    );

  if (
    !applyResult ||
    !applyResult.success
  ) {

    bc_appendGovernancePipelineException(
      'W2B.3 — Coherence Repair',
      applyResult &&
      applyResult.message
        ? applyResult.message
        : 'Single-pass coherence repairs could not be applied.'
    );

    bc_addToApiCostAndTime(
      totalCost,
      (
        Date.now() -
        startTime
      ) / 1000
    );

    return {
      success: true,
      message:
        'W2B.3 audit completed, but its repairs could not be applied safely — original W2B.2 HTML retained in GI.',
      cost: totalCost
    };
  }

  /*
   * =========================================================
   * STEP 8 — FINISH
   *
   * NO SECOND COHERENCE CHECK.
   * NO LOOP.
   * NO "STILL FAILING" STATE.
   * =========================================================
   */

  bc_addToApiCostAndTime(
    totalCost,
    (
      Date.now() -
      startTime
    ) / 1000
  );

  return {
    success: true,
    message:
      'W2B.3 single-pass coherence completed — ' +
      issues.length +
      ' issue(s) identified and repair pass completed. ' +
      applyResult.message,
    cost: totalCost
  };
}

function bc_runStage3Automated() {
  var startTime = Date.now();
  var htmlResult = getW2BHtmlFromSheet();
  if (!htmlResult.success) throw new Error(htmlResult.message);

  var promptData = buildHumanisationPrompt(htmlResult.html);
  if (!promptData.success) throw new Error(promptData.message);

  var apiResult = bc_sendPromptViaOpenAI(promptData.prompt, 8000);
  if (!apiResult.success) throw new Error(apiResult.message);

  var saveResult = saveHumanisedHtml(apiResult.text);
  if (!saveResult.success) throw new Error(saveResult.message);

  bc_addToApiCostAndTime(apiResult.cost, (Date.now() - startTime) / 1000);
  return { success: true, message: saveResult.message, cost: apiResult.cost };
}

function bc_runW4BAutomated() {
  var startTime = Date.now();
  var htmlResult = getHumanisedHtmlForW4B();
  if (!htmlResult.success) throw new Error(htmlResult.message);
  var html = htmlResult.html;

  var figResult = storeFigureInventory();

  var imgCount = (html.match(/<img[^>]+>/gi) || []).length;
  if (imgCount === 0) throw new Error('No <img> tags found in the humanised HTML.');

  var meta = getW4BMetadata();

  var imgFilenames = [];
  var imgRegex = /<img[^>]+>/gi;
  var imgMatch;
  while ((imgMatch = imgRegex.exec(html)) !== null) {
    if (/data-w4b-skip\s*=\s*["']true["']/i.test(imgMatch[0])) continue;
    var srcMatch = imgMatch[0].match(/src=["']([^"']+)["']/i);
    if (srcMatch) imgFilenames.push(srcMatch[1].split('/').pop());
  }

  var prompt = [
    'You are an SEO image specialist working on a UK natural stone floor restoration website.',
    '',
    'TASK: For each image filename below, return descriptive alt text and a caption appropriate to its category.',
    '',
    'CONTEXT:',
    '  Stone type: ' + (meta.stoneType || 'unknown'),
    '  Article type: ' + (meta.articleType || 'unknown'),
    '  Primary search term: ' + (meta.primaryTerm || 'unknown'),
    '',
    'IMAGE CATEGORIES (classify each filename before writing alt/caption):',
    '1. FLOOR CONDITION IMAGES — photos of the actual stone floor showing a condition, stage, or result. Write a descriptive alt AND a diagnostic caption that helps the reader identify their problem (e.g. "If your floor looks like this...").',
    '2. PRODUCT IMAGES — any image from m.media-amazon.com, or any filename that is clearly a commercial product (cleaner bottle, vacuum, mop, brush, sealer, tool). These are NOT decorative. Write a plain descriptive alt stating what the product is (e.g. "Fila Pro Floor Cleaner for impregnated stone surfaces"). Caption should be empty string "" — product captions are handled separately by the page template, not by this diagnostic caption system.',
    '3. TRUE DECORATIVE IMAGES — logos, icons, dividers, spacers with no informative content. Only these get alt="" and caption="".',
    '',
    'CRITICAL: Do not default to category 3 for product images. A product photo always gets a real descriptive alt text in category 2 — only alt text is ever blank for icons/logos/dividers with zero informative value.',
    '',
    'RULES:',
    '- Alt text: descriptive, specific, under 125 characters. No keyword stuffing.',
    '- Captions (floor condition images only): plain text only, max 15 words, diagnostic — help the reader identify their problem.',
    '- Return ONLY a JSON array. No preamble. No markdown fences. No explanation.',
    '',
    'OUTPUT FORMAT:',
    '[{"filename":"[filename]","alt":"[alt text]","caption":"[caption]"}]',
    '',
    'IMAGE FILENAMES:',
    imgFilenames.join('\n')
  ].join('\n');

  var apiResult = bc_sendPromptViaOpenAI(prompt, 3000);
  if (!apiResult.success) throw new Error(apiResult.message);

  var pushResult = pushAltCaptionUpdates(apiResult.text);
  if (!pushResult.success) throw new Error(pushResult.message);

  bc_addToApiCostAndTime(apiResult.cost, (Date.now() - startTime) / 1000);
  return { success: true, message: pushResult.message, cost: apiResult.cost };
}

function bc_buildDeferredFixPromptServer_(html, result) {
  var failureLines = [];
  if (result.check1 === 'FAIL') failureLines.push('CHECK 1 (Cluster tone): ' + result.obs1);
  if (result.check4 === 'FAIL') failureLines.push('CHECK 4 (Section opening tone): ' + result.obs4);
  if (result.check5 === 'FAIL') failureLines.push('CHECK 5 (Header cluster tone): ' + result.obs5);
  if (result.check7 === 'FAIL') failureLines.push('CHECK 7 (Named defect completion): ' + result.obs7);

  return 'Read the following and respond as instructed:\n\n' +
    'DEFERRED CHECK FIX — TARGETED SECTION CORRECTIONS\n\n' +
    'ROLE: Senior UK SEO content editor for a natural stone floor restoration website.\n\n' +
    'TASK: Fix only the sections identified in the failures below. Return the complete HTML with only those sections corrected. Do not reword, restructure or improve any other section.\n\n' +
    'CRITICAL RULES:\n' +
    '1. Return complete HTML only — no markdown, no explanation, no preamble.\n' +
    '2. Do not change any section not named in the failures.\n' +
    '3. Do not change images, figures, captions, links, schema, header, footer or bio box.\n' +
    '4. Each section opening sentence must orient the reader to their visible problem before introducing mechanism or technical terms.\n' +
    '5. Named defects must be complete — three mandatory elements in the same paragraph before the next paragraph begins:\n' +
    '   ELEMENT 1 (definition): One sentence stating what the defect is as a physical mechanism.\n' +
    '   ELEMENT 2 (symptom): One sentence stating what the homeowner sees or notices.\n' +
    '   ELEMENT 3 (correction): One sentence stating what professional intervention does about it.\n' +
    '   A defect named without all three elements present is incomplete.\n\n' +
    'FAILURES TO FIX:\n' + failureLines.join('\n') + '\n\n' +
    'ARTICLE HTML:\n' + html;
}

function bc_runW4Automated() {
  var startTime = Date.now();
  var totalCost = 0;
  var log = [];

  var html = getHtmlFromActiveRow();
  if (!html || html.indexOf('ERROR') === 0) throw new Error(html || 'No HTML found.');

  var audit = runAtomicAuditWithContext(html);
  if (audit.error) throw new Error(audit.error);

  if (!audit.allPass) {
    var remaining = [];
    audit.failures.forEach(function(f) {
      if (f.canAutoFix) {
        var fixRes = applyAutoFix(html, f.checkId);
        if (fixRes.success) {
          html = fixRes.patchedHtml;
          log.push('Auto-fixed CHECK ' + f.checkId);
        } else {
          remaining.push(f);
        }
      } else {
        remaining.push(f);
      }
    });

    if (remaining.length > 0) {
      var batchPrompt = buildBatchAtomicFixPrompt(remaining);
      if (!batchPrompt.success) throw new Error(batchPrompt.message);

      var batchApiResult = bc_sendPromptViaOpenAI(batchPrompt.prompt, 6000);
      if (!batchApiResult.success) throw new Error(batchApiResult.message);
      totalCost += batchApiResult.cost;

      var batchApply = applyBatchAtomicFix(html, batchApiResult.text);
      if (!batchApply.success) throw new Error('Batch fix failed: ' + batchApply.message);
      html = batchApply.patchedHtml;
      log.push('Batch fix: ' + batchApply.message);
    }

    var recheck = runAtomicAuditWithContext(html);
    if (recheck.error) throw new Error(recheck.error);
        if (!recheck.allPass) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

      var reviewMsg = recheck.failures.length > 0
        ? recheck.failures.length +
          ' mechanical failure(s) remain after auto-fix and batch fix.\n\n' +
          (recheck.report || '')
        : 'Audit still reports failure(s) that could not be extracted into fixable items.\n\n' +
          (recheck.report || '');

      bc_appendGovernancePipelineException(
        'W2B — Mechanical Audit',
        reviewMsg
      );

      return {
        success: false,
        message:
          'W2B mechanical audit still has unresolved failures — recorded in GJ.',
        cost: totalCost
      };
    }
    log.push('Mechanical audit now passes after fixes.');
  }

  var deferredPrompt = buildDeferredCheckPrompt(html);
  var deferredApiResult = bc_sendPromptViaOpenAI(deferredPrompt, 2000);
  if (!deferredApiResult.success) throw new Error(deferredApiResult.message);
  totalCost += deferredApiResult.cost;

  var deferredResult = parseDeferredCheckResponse(deferredApiResult.text);

  if (deferredResult.hasFailures) {
    var deferredFixPrompt = bc_buildDeferredFixPromptServer_(html, deferredResult);
    var deferredFixApiResult = bc_sendPromptViaOpenAI(deferredFixPrompt, 8000);
    if (!deferredFixApiResult.success) throw new Error(deferredFixApiResult.message);
    totalCost += deferredFixApiResult.cost;

    html = deferredFixApiResult.text;
    log.push('Deferred check fix applied.');

    var recheckPrompt = buildDeferredCheckPrompt(html);
    var recheckApiResult = bc_sendPromptViaOpenAI(recheckPrompt, 2000);
    if (!recheckApiResult.success) throw new Error(recheckApiResult.message);
    totalCost += recheckApiResult.cost;

    var recheckDeferred = parseDeferredCheckResponse(recheckApiResult.text);
        if (recheckDeferred.hasFailures) {

      bc_addToApiCostAndTime(
        totalCost,
        (Date.now() - startTime) / 1000
      );

      bc_appendGovernancePipelineException(
        'W2B — Deferred Checks',
        recheckApiResult.text ||
          'Deferred checks still failing after one automated fix pass.'
      );

      return {
        success: false,
        message:
          'W2B deferred checks still failing after one fix pass — recorded in GJ.',
        cost: totalCost
      };
    }
    log.push('Deferred checks now pass after fix.');
  } else {
    log.push('Deferred checks passed first time.');
  }

  var pushMessage = pushHtmlToActiveRow(html);

  bc_addToApiCostAndTime(totalCost, (Date.now() - startTime) / 1000);
  return {
    success: true,
    message: 'Pushed — ' + pushMessage + ' | ' + log.join(' | '),
    cost: totalCost
  };
}
function bc_getActivePostsRow() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");

    if (!sheet) {
      return {
        success: false,
        message: "Posts sheet not found."
      };
    }

    var activeRow = sheet.getActiveRange().getRow();

    if (activeRow < 2) {
      return {
        success: false,
        message: "Select a data row first."
      };
    }

    return {
      success: true,
      row: activeRow
    };

  } catch (e) {
    return {
      success: false,
      message: "Active row error: " + e.toString()
    };
  }
}

function bc_appendPipelineRunLog(row, line) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");

    if (!sheet) {
      return {
        success: false,
        message: "Posts sheet not found."
      };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });

    var logIdx = headers.indexOf("Pipeline Run Log");

    if (logIdx === -1) {
      return {
        success: false,
        message: "Pipeline Run Log column not found."
      };
    }

    var cell = sheet.getRange(row, logIdx + 1);
    var existing = String(cell.getValue() || "").trim();

    var updated = existing
      ? existing + "\n" + line
      : line;

    cell.setValue(updated);

    return {
      success: true
    };

  } catch (e) {
    return {
      success: false,
      message: "Pipeline log error: " + e.toString()
    };
  }
}

function bc_clearPipelineRunLog(row) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");

    if (!sheet) {
      return {
        success: false,
        message: "Posts sheet not found."
      };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });

    var logIdx = headers.indexOf("Pipeline Run Log");

    if (logIdx === -1) {
      return {
        success: false,
        message: "Pipeline Run Log column not found."
      };
    }

    sheet.getRange(row, logIdx + 1).clearContent();

    return {
      success: true
    };

  } catch (e) {
    return {
      success: false,
      message: "Pipeline log clear error: " + e.toString()
    };
  }
}

function bc_recordPipelineStage(row, stageName, cost) {
  var stageCost = Number(cost || 0);

  var line =
    "✓ " +
    stageName +
    " complete — $" +
    stageCost.toFixed(4);

  return bc_appendPipelineRunLog(row, line);
}

function bc_recordPipelineFailure(row, stageName, message) {
  var line =
    "✗ " +
    stageName +
    " failed — " +
    String(message || "Unknown error");

  return bc_appendPipelineRunLog(row, line);
}

function bc_startPipelineRun() {
  var rowResult = bc_getActivePostsRow();

  if (!rowResult.success) {
    return rowResult;
  }

  var activeRow = rowResult.row;
  PropertiesService
  .getDocumentProperties()
  .deleteProperty("PIPELINE_STOP_ROW_" + activeRow);

  var clearResult = bc_clearPipelineRunLog(activeRow);

  if (!clearResult.success) {
    return clearResult;
  }

  bc_clearGovernancePipelineExceptions();

  bc_appendPipelineRunLog(
    activeRow,
    "Pipeline started — row " + activeRow
  );

  return {
    success: true,
    row: activeRow,
    message: "Pipeline started for row " + activeRow
  };
}

function bc_pipelineRunPreAC(row) {
  try {
    var result = bc_runPreACPromptAutomated();

    if (!result || result.success === false) {
      var message = result && result.message
        ? result.message
        : "Unknown Pre-AC error";

      bc_recordPipelineFailure(row, "Pre-AC", message);

      return {
        success: false,
        message: message
      };
    }

    bc_recordPipelineStage(
      row,
      "Pre-AC",
      result.cost || 0
    );

    return {
      success: true,
      cost: Number(result.cost || 0),
      message: result.message || "Pre-AC complete"
    };

  } catch (e) {
    bc_recordPipelineFailure(
      row,
      "Pre-AC",
      e.toString()
    );

    return {
      success: false,
      message: e.toString()
    };
  }
}

function pipelineRunAC(row) {
  try {
    var result = bc_runACPromptAutomated(row);

    if (!result || result.success === false) {
      var message = result && result.message
        ? result.message
        : "Unknown AC error";

      bc_recordPipelineFailure(row, "AC", message);

      return {
        success: false,
        message: message
      };
    }

    bc_recordPipelineStage(
      row,
      "AC",
      result.cost || 0
    );

    return {
      success: true,
      cost: Number(result.cost || 0),
      message: result.message || "AC complete"
    };

  } catch (e) {
    bc_recordPipelineFailure(
      row,
      "AC",
      e.toString()
    );

    return {
      success: false,
      message: e.toString()
    };
  }
}

function pipelineRunAuthorityBrief(row) {
  try {
    var result = autoRunAuthorityBrief(row);

    if (!result || result.success === false) {
      var message = result && result.message
        ? result.message
        : "Unknown W-1 error";

      bc_recordPipelineFailure(row, "W-1", message);

      return {
        success: false,
        message: message
      };
    }

    bc_recordPipelineStage(
      row,
      "W-1",
      0
    );

    return {
      success: true,
      cost: 0,
      message: result.message || "W-1 complete"
    };

  } catch (e) {
    bc_recordPipelineFailure(
      row,
      "W-1",
      e.toString()
    );

    return {
      success: false,
      message: e.toString()
    };
  }
}

function pipelineRequestStop(row) {
  try {
    row = Number(row);

    if (!row || row < 2) {
      return {
        success: false,
        message: "No valid pipeline row supplied."
      };
    }

    PropertiesService
      .getDocumentProperties()
      .setProperty(
        "PIPELINE_STOP_ROW_" + row,
        "TRUE"
      );

    bc_appendPipelineRunLog(
      row,
      "■ Stop requested by user."
    );

    return {
      success: true,
      row: row,
      message: "Stop requested."
    };

  } catch (e) {
    return {
      success: false,
      message: "Could not request stop: " + e.toString()
    };
  }
}

function pipelineStopRequested(row) {
  var value = PropertiesService
    .getDocumentProperties()
    .getProperty("PIPELINE_STOP_ROW_" + row);

  return value === "TRUE";
}

function pipelineRunPreACToW2B3(row) {

  var totalCost = 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("posts");

  if (!sheet) {
    return {
      success: false,
      cost: 0,
      message: "Posts sheet not found."
    };
  }

  sheet.setActiveRange(
    sheet.getRange(row, 1)
  );

  // ============================================================
  // PRE-AC
  // ============================================================

  var preAcResult =
    bc_pipelineRunPreAC(row);

  if (
    !preAcResult ||
    preAcResult.success === false
  ) {
    return {
      success: false,
      cost: totalCost,
      message:
        preAcResult && preAcResult.message
          ? preAcResult.message
          : "Unknown Pre-AC error"
    };
  }

  totalCost +=
    Number(preAcResult.cost || 0);

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after Pre-AC."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message:
        "Pipeline stopped by user after Pre-AC."
    };
  }

  // ============================================================
  // AC
  // ============================================================

  var acResult =
    pipelineRunAC(row);

  if (
    !acResult ||
    acResult.success === false
  ) {
    return {
      success: false,
      cost: totalCost,
      message:
        acResult && acResult.message
          ? acResult.message
          : "Unknown AC error"
    };
  }

  totalCost +=
    Number(acResult.cost || 0);

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after AC."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message:
        "Pipeline stopped by user after AC."
    };
  }

  // ============================================================
  // AUTHORITY BRIEF / W-1
  // ============================================================

  var authorityResult =
    pipelineRunAuthorityBrief(row);

  if (
    !authorityResult ||
    authorityResult.success === false
  ) {
    return {
      success: false,
      cost: totalCost,
      message:
        authorityResult &&
        authorityResult.message
          ? authorityResult.message
          : "Unknown Authority Brief error"
    };
  }

  totalCost +=
    Number(authorityResult.cost || 0);

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W-1."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message:
        "Pipeline stopped by user after W-1."
    };
  }

  // ============================================================
  // EXISTING W1A → W2B.3 PIPELINE
  // ============================================================

  var downstreamResult =
    pipelineRunW1AToW2B3(row);

  if (
    !downstreamResult ||
    downstreamResult.success === false
  ) {
    return {
      success: false,
      stopped:
        downstreamResult &&
        downstreamResult.stopped === true,
      cost:
        totalCost +
        Number(
          downstreamResult &&
          downstreamResult.cost
            ? downstreamResult.cost
            : 0
        ),
      message:
        downstreamResult &&
        downstreamResult.message
          ? downstreamResult.message
          : "W1A → W2B.3 pipeline failed."
    };
  }

  totalCost +=
    Number(downstreamResult.cost || 0);

  bc_appendPipelineRunLog(
    row,
    "✓ Pre-AC → W2B.3 complete"
  );

  bc_appendPipelineRunLog(
    row,
    "TOTAL AI COST — $" +
    totalCost.toFixed(4)
  );

  return {
    success: true,
    cost: totalCost,
    message:
      "Pre-AC → W2B.3 complete."
  };
}

function pipelineRunW1AToW2B3(row) {

  var totalCost = 0;

  // ============================================================
  // That locks the existing W1A→W2B.3 functions onto the row passed in by the sidebar,
  // ============================================================

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("posts");

  if (!sheet) {
    return {
      success: false,
      cost: 0,
      message: "Posts sheet not found."
    };
  }

  sheet.setActiveRange(
    sheet.getRange(row, 1)
  );


  // ============================================================
  // STOP CHECK BEFORE START
  // ============================================================

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user."
    };
  }


  // ============================================================
  // W1A-W1D
  // ============================================================

  var w1Result = saveAllW1Outputs();

  if (!w1Result || w1Result.success === false) {
    var message = w1Result && w1Result.message
      ? w1Result.message
      : "Unknown W1A-W1D error";

    bc_recordPipelineFailure(
      row,
      "W1A-W1D",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  bc_recordPipelineStage(
    row,
    "W1A-W1D",
    0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W1A-W1D."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W1A-W1D."
    };
  }


  // ============================================================
  // W1.5A
  // ============================================================

  var w15a = bc_runStage15AAutomated();

  if (!w15a || w15a.success === false) {
    var message = w15a && w15a.message
      ? w15a.message
      : "Unknown W1.5A error";

    bc_recordPipelineFailure(
      row,
      "W1.5A",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w15a.cost || 0);

  bc_recordPipelineStage(
    row,
    "W1.5A",
    w15a.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W1.5A."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W1.5A."
    };
  }


  // ============================================================
  // W1.5B
  // ============================================================

  var w15b = bc_runFullW15BAutomated();

  if (!w15b || w15b.success === false) {
    var message = w15b && w15b.message
      ? w15b.message
      : "Unknown W1.5B error";

    bc_recordPipelineFailure(
      row,
      "W1.5B",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w15b.cost || 0);

  bc_recordPipelineStage(
    row,
    "W1.5B",
    w15b.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W1.5B."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W1.5B."
    };
  }


  // ============================================================
  // W1.5C
  // ============================================================

  var w15c = bc_runStage15CAutomated();

  if (!w15c || w15c.success === false) {
    var message = w15c && w15c.message
      ? w15c.message
      : "Unknown W1.5C error";

    bc_recordPipelineFailure(
      row,
      "W1.5C",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w15c.cost || 0);

  bc_recordPipelineStage(
    row,
    "W1.5C",
    w15c.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W1.5C."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W1.5C."
    };
  }


  // ============================================================
  // W1.5D
  // ============================================================

  var w15d = bc_runStage15DAutomated();

  if (!w15d || w15d.success === false) {
    var message = w15d && w15d.message
      ? w15d.message
      : "Unknown W1.5D error";

    bc_recordPipelineFailure(
      row,
      "W1.5D",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w15d.cost || 0);

  bc_recordPipelineStage(
    row,
    "W1.5D",
    w15d.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W1.5D."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W1.5D."
    };
  }


  // ============================================================
  // W1.5E
  // ============================================================

  var w15e = bc_runStage15EAutomated();

  if (!w15e || w15e.success === false) {
    var message = w15e && w15e.message
      ? w15e.message
      : "Unknown W1.5E error";

    bc_recordPipelineFailure(
      row,
      "W1.5E",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w15e.cost || 0);

  bc_recordPipelineStage(
    row,
    "W1.5E",
    w15e.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W1.5E."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W1.5E."
    };
  }


  // ============================================================
  // W2B
  // ============================================================

  var w2b = bc_runStage2BAutomated();

  if (!w2b || w2b.success === false) {
    var message = w2b && w2b.message
      ? w2b.message
      : "Unknown W2B error";

    bc_recordPipelineFailure(
      row,
      "W2B",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w2b.cost || 0);

  bc_recordPipelineStage(
    row,
    "W2B",
    w2b.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W2B."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W2B."
    };
  }


  // ============================================================
  // W2B.05
  // ============================================================

  var w2b05 = bc_runFactCheckFullAutomated();

  if (!w2b05 || w2b05.success === false) {
    var message = w2b05 && w2b05.message
      ? w2b05.message
      : "Unknown W2B.05 error";

    bc_recordPipelineFailure(
      row,
      "W2B.05",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w2b05.cost || 0);

  bc_recordPipelineStage(
    row,
    "W2B.05",
    w2b05.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W2B.05."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W2B.05."
    };
  }


  // ============================================================
  // W2B.1
  // ============================================================

  var w2b1 = bc_runSimilarityFullAutomated();

  if (!w2b1 || w2b1.success === false) {
    var message = w2b1 && w2b1.message
      ? w2b1.message
      : "Unknown W2B.1 error";

    bc_recordPipelineFailure(
      row,
      "W2B.1",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w2b1.cost || 0);

  bc_recordPipelineStage(
    row,
    "W2B.1",
    w2b1.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W2B.1."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W2B.1."
    };
  }


  // ============================================================
  // W2B.2
  // ============================================================

  var w2b2 =
    bc_runRewriteBriefComplianceW2BAutomated();

  if (!w2b2 || w2b2.success === false) {
    var message = w2b2 && w2b2.message
      ? w2b2.message
      : "Unknown W2B.2 error";

    bc_recordPipelineFailure(
      row,
      "W2B.2",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w2b2.cost || 0);

  bc_recordPipelineStage(
    row,
    "W2B.2",
    w2b2.cost || 0
  );

  if (pipelineStopRequested(row)) {
    bc_appendPipelineRunLog(
      row,
      "■ Pipeline stopped by user after W2B.2."
    );

    return {
      success: false,
      stopped: true,
      cost: totalCost,
      message: "Pipeline stopped by user after W2B.2."
    };
  }


  // ============================================================
  // W2B.3
  // ============================================================

  var w2b3 = bc_runCoherenceFullAutomated();

  if (!w2b3 || w2b3.success === false) {
    var message = w2b3 && w2b3.message
      ? w2b3.message
      : "Unknown W2B.3 error";

    bc_recordPipelineFailure(
      row,
      "W2B.3",
      message
    );

    return {
      success: false,
      cost: totalCost,
      message: message
    };
  }

  totalCost += Number(w2b3.cost || 0);

  bc_recordPipelineStage(
    row,
    "W2B.3",
    w2b3.cost || 0
  );


  // ============================================================
  // COMPLETE
  // ============================================================

  bc_appendPipelineRunLog(
    row,
    "✓ W1A → W2B.3 complete — $" +
    totalCost.toFixed(4)
  );

  return {
    success: true,
    cost: totalCost,
    message: "W1A → W2B.3 complete."
  };
}

function getPipelineRunLog(row) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");

    if (!sheet) {
      return {
        success: false,
        message: "Posts sheet not found."
      };
    }

    var headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(h) {
        return String(h).trim();
      });

    var logIdx = headers.indexOf("Pipeline Run Log");

    if (logIdx === -1) {
      return {
        success: false,
        message: "Pipeline Run Log column not found."
      };
    }

    var logText = String(
      sheet.getRange(row, logIdx + 1).getValue() || ""
    );

    return {
      success: true,
      log: logText
    };

  } catch (e) {
    return {
      success: false,
      message: "Pipeline log read error: " + e.toString()
    };
  }
}