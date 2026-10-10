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

  try {

    var apiKey =
      PropertiesService
        .getScriptProperties()
        .getProperty('GEMINI_API_KEY');

    if (!apiKey) {
      throw new Error(
        'GEMINI_API_KEY not set in Script Properties.'
      );
    }


    var payload = {
      contents: [
        {
          parts: [
            {
              text: promptText
            }
          ]
        }
      ],
      tools: [
        {
          googleSearch: {}
        }
      ]
    };


    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };


    var url =
      'https://generativelanguage.googleapis.com/v1beta/models/' +
      'gemini-3.6-flash:generateContent?key=' +
      apiKey;


    var response =
      UrlFetchApp.fetch(
        url,
        options
      );


    var responseText =
      response.getContentText();


    var data;

    try {

      data =
        JSON.parse(
          responseText
        );

    } catch (parseError) {

      return {
        success: false,
        message:
          'Gemini returned invalid JSON: ' +
          responseText.substring(0, 500)
      };
    }


    if (
      response.getResponseCode() !== 200
    ) {

      return {
        success: false,
        message:
          (
            data.error &&
            data.error.message
          )
            ? data.error.message
            : 'Gemini API error — HTTP ' +
              response.getResponseCode()
      };
    }


    if (
      !data.candidates ||
      !data.candidates.length
    ) {

      return {
        success: false,
        message:
          'Gemini returned no candidates. Response: ' +
          responseText.substring(0, 500)
      };
    }


    var candidate =
      data.candidates[0];


    if (
      !candidate.content ||
      !candidate.content.parts ||
      !candidate.content.parts.length
    ) {

      return {
        success: false,
        message:
          'Gemini returned no text content. Finish reason: ' +
          String(
            candidate.finishReason || 'unknown'
          )
      };
    }


    var text =
      candidate.content.parts
        .map(
          function(part) {
            return part.text || '';
          }
        )
        .join('')
        .trim();


    if (!text) {

      return {
        success: false,
        message:
          'Gemini returned an empty text response.'
      };
    }


    var promptTokens =
      (
        data.usageMetadata &&
        data.usageMetadata.promptTokenCount
      ) || 0;


    var completionTokens =
      (
        data.usageMetadata &&
        data.usageMetadata.candidatesTokenCount
      ) || 0;


    var cost =
      (
        promptTokens /
        1000000 *
        1.50
      ) +
      (
        completionTokens /
        1000000 *
        7.50
      );


    return {
      success: true,
      text: text,
      promptTokens: promptTokens,
      completionTokens: completionTokens,
      cost: cost
    };


  } catch (e) {

    return {
      success: false,
      message:
        'Gemini request error: ' +
        e.toString()
    };
  }
}
function bc_runSerpBridgeAutomated() {
  var promptData = bc_buildSerpBridgePrompt();
  if (promptData.skipped) {
  return {
    success: true,
    skipped: true,
    text: '',
    message: promptData.message,
    cost: 0
  };
}

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

function bc_contentIntelligencePreflight_() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName("posts");
  var exportSheet = ss.getSheetByName("site-export");
  var row = posts.getActiveRange().getRow();

  if (row < 2) {
    return {
      success: false,
      message: "Select a valid article row first."
    };
  }

  var headers = posts
    .getRange(1, 1, 1, posts.getLastColumn())
    .getValues()[0]
    .map(function(h) {
      return String(h || "").trim();
    });

  function getPostValue(name) {
    var idx = headers.indexOf(name);
    return idx > -1
      ? String(posts.getRange(row, idx + 1).getValue() || "").trim()
      : "";
  }

  var postId = getPostValue("Post ID");
  var url = getPostValue("URL");
  var title = getPostValue("Title");

  var monitorQueries = postId
    ? String(bc_getQueriesFromSheet(postId, "GSC Monitor") || "").trim()
    : String(getPostValue("Montr Queries") || "").trim();

  // Existing article with usable GSC data — continue normally.
  if (monitorQueries.length > 10) {
    return {
      success: true,
      isNewArticle: false
    };
  }

  if (!exportSheet) {
    return {
      success: false,
      message:
        "No usable GSC data found. This looks like a new article, but the site-export sheet was not found."
    };
  }

  var data = exportSheet.getDataRange().getValues();

  if (data.length < 2) {
    return {
      success: false,
      message:
        "No usable GSC data found. This looks like a new article, but site-export is empty."
    };
  }

  var exportHeaders = data[0].map(function(h) {
    return String(h || "").trim();
  });

  var idIdx = exportHeaders.indexOf("ID");
  var titleIdx = exportHeaders.indexOf("Title");
  var urlIdx = exportHeaders.indexOf("Canonical URL");
  var keyphraseIdx = exportHeaders.indexOf("Yoast Keyphrase");
  var metaTitleIdx = exportHeaders.indexOf("Meta Title");
  var metaDescIdx = exportHeaders.indexOf("Meta Description");

  var matchedRow = null;

  for (var i = 1; i < data.length; i++) {

    var exportId = idIdx > -1
      ? String(data[i][idIdx] || "").trim()
      : "";

    var exportUrl = urlIdx > -1
      ? String(data[i][urlIdx] || "").trim()
      : "";

    var exportTitle = titleIdx > -1
      ? String(data[i][titleIdx] || "").trim()
      : "";

    if (
      (postId && exportId === postId) ||
      (url && exportUrl === url) ||
      (title && exportTitle === title)
    ) {
      matchedRow = data[i];
      break;
    }
  }

  if (!matchedRow) {
    return {
      success: false,
      message:
        "No usable GSC data found. This looks like a new article. Add the article to site-export with its Yoast keyphrase and meta data before running Content Intelligence."
    };
  }

  var keyphrase = keyphraseIdx > -1
    ? String(matchedRow[keyphraseIdx] || "").trim()
    : "";

  var metaTitle = metaTitleIdx > -1
    ? String(matchedRow[metaTitleIdx] || "").trim()
    : "";

  var metaDescription = metaDescIdx > -1
    ? String(matchedRow[metaDescIdx] || "").trim()
    : "";

  var missing = [];

  if (!keyphrase) missing.push("Yoast Keyphrase");
  if (!metaTitle) missing.push("Meta Title");
  if (!metaDescription) missing.push("Meta Description");

  if (missing.length > 0) {
    return {
      success: false,
      message:
        "No usable GSC data found. This looks like a new article. Add: " +
        missing.join(", ") +
        " before running Content Intelligence."
    };
  }

  var pstIdx = headers.indexOf("Primary Search Term");

  if (pstIdx === -1) {
    return {
      success: false,
      message: 'Column "Primary Search Term" was not found.'
    };
  }

  posts
    .getRange(row, pstIdx + 1)
    .setValue(keyphrase);

  return {
    success: true,
    isNewArticle: true,
    primarySearchTerm: keyphrase
  };
}


function getMasterCIRowsForPostIds(postIdsText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");

    if (!sheet) {
      return { success: false, message: "Posts sheet not found." };
    }

    var rawLines = String(postIdsText || "")
      .split(/\r?\n/)
      .map(function(line) { return String(line || "").trim(); })
      .filter(function(line) { return line !== ""; });

    if (!rawLines.length) {
      return { success: false, message: "Enter at least one Post ID, one per line." };
    }

    var requestedIds = [];
    var seen = {};
    var duplicates = [];

    rawLines.forEach(function(id) {
      if (seen[id]) {
        duplicates.push(id);
        return;
      }
      seen[id] = true;
      requestedIds.push(id);
    });

    var data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return { success: false, message: "Posts sheet has no data rows." };
    }

    var headers = data[0].map(function(h) {
      return String(h || "").trim();
    });

    var postIdIdx = headers.indexOf("Post ID");
    var titleIdx = headers.indexOf("Title");
    var stoneIdx = headers.indexOf("Stone Type");

    if (postIdIdx === -1) {
      return { success: false, message: 'Column "Post ID" not found.' };
    }

    if (stoneIdx === -1) {
      return { success: false, message: 'Column "Stone Type" not found.' };
    }

    var byId = {};

    for (var r = 1; r < data.length; r++) {
      var id = String(data[r][postIdIdx] || "").trim();
      if (!id) continue;

      byId[id] = {
        row: r + 1,
        postId: id,
        title: titleIdx > -1 ? String(data[r][titleIdx] || "").trim() : "",
        stoneType: String(data[r][stoneIdx] || "").trim()
      };
    }

    var rows = [];
    var missing = [];

    requestedIds.forEach(function(id) {
      if (byId[id]) {
        rows.push(byId[id]);
      } else {
        missing.push(id);
      }
    });

    if (missing.length) {
      return {
        success: false,
        message:
          "These Post IDs were not found in Column D: " +
          missing.join(", "),
        missing: missing,
        duplicates: duplicates
      };
    }

    return {
      success: true,
      rows: rows,
      count: rows.length,
      duplicates: duplicates,
      message:
        rows.length +
        " Post ID" +
        (rows.length === 1 ? "" : "s") +
        " loaded."
    };

  } catch (e) {
    return {
      success: false,
      message: "getMasterCIRowsForPostIds error: " + e.message
    };
  }
}

function runMasterContentIntelligenceBatchRow(row, expectedStoneType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("posts");

  if (!sheet) {
    return { success: false, message: "Posts sheet not found." };
  }

  row = Number(row);

  if (!row || row < 2 || row > sheet.getLastRow()) {
    return { success: false, message: "Invalid posts row: " + row };
  }

  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(function(h) { return String(h || "").trim(); });

  var stoneIdx = headers.indexOf("Stone Type");
  var titleIdx = headers.indexOf("Title");

  if (stoneIdx === -1) {
    return { success: false, message: 'Column "Stone Type" not found.' };
  }

  var liveStone = String(
    sheet.getRange(row, stoneIdx + 1).getValue() || ""
  ).trim();

  if (
    expectedStoneType &&
    liveStone.toLowerCase() !==
      String(expectedStoneType).trim().toLowerCase()
  ) {
    return {
      success: false,
      message:
        "Row " + row +
        ' is now "' + liveStone +
        '" rather than "' + expectedStoneType +
        '". Batch stopped to avoid processing the wrong article.'
    };
  }

  var title = titleIdx > -1
    ? String(sheet.getRange(row, titleIdx + 1).getValue() || "").trim()
    : "";

  sheet.setActiveRange(sheet.getRange(row, 1));

  try {
    var result = bc_runFullGovernanceChainAutomated();

    if (!result || !result.success) {
      return {
        success: false,
        row: row,
        title: title,
        stoneType: liveStone,
        message:
          result && result.message
            ? result.message
            : "Content Intelligence failed.",
        log:
          result && result.log
            ? result.log
            : "",
        cost:
          result && result.cost
            ? Number(result.cost)
            : 0
      };
    }

    return {
      success: true,
      row: row,
      title: title,
      stoneType: liveStone,
      message: result.message || "Content Intelligence complete.",
      log: result.log || "",
      cost: Number(result.cost || 0)
    };

  } catch (e) {
    return {
      success: false,
      row: row,
      title: title,
      stoneType: liveStone,
      message: e && e.message ? e.message : String(e),
      log: "",
      cost: 0
    };
  }
}



function runMasterContentGenerationBatchRow(row, expectedPostId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");

    if (!sheet) {
      return {
        success: false,
        message: "Posts sheet not found.",
        cost: 0
      };
    }

    row = Number(row);

    if (!row || row < 2 || row > sheet.getLastRow()) {
      return {
        success: false,
        message: "Invalid posts row: " + row,
        cost: 0
      };
    }

    var headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(h) {
        return String(h || "").trim();
      });

    var postIdIdx = headers.indexOf("Post ID");
    var titleIdx = headers.indexOf("Title");

    if (postIdIdx === -1) {
      return {
        success: false,
        message: 'Column "Post ID" not found.',
        cost: 0
      };
    }

    var livePostId = String(
      sheet.getRange(row, postIdIdx + 1).getValue() || ""
    ).trim();

    if (
      expectedPostId &&
      livePostId !== String(expectedPostId).trim()
    ) {
      return {
        success: false,
        message:
          "Row " + row +
          " now contains Post ID " + livePostId +
          " rather than " + expectedPostId +
          ". Batch stopped to avoid processing the wrong article.",
        cost: 0
      };
    }

    var title = titleIdx > -1
      ? String(sheet.getRange(row, titleIdx + 1).getValue() || "").trim()
      : "";

    sheet.setActiveRange(
      sheet.getRange(row, 1)
    );

    var startResult =
      bc_startPipelineRun();

    if (
      !startResult ||
      startResult.success === false
    ) {
      return {
        success: false,
        row: row,
        postId: livePostId,
        title: title,
        message:
          startResult && startResult.message
            ? startResult.message
            : "Could not initialise pipeline.",
        cost: 0
      };
    }

    var result =
      pipelineResumeToW8E(
        row,
        "Pre-AC"
      );

    var logResult =
      getPipelineRunLog(row);

    if (
      !result ||
      result.success === false
    ) {
      return {
        success: false,
        stopped:
          result &&
          result.stopped === true,
        row: row,
        postId: livePostId,
        title: title,
        message:
          result && result.message
            ? result.message
            : "Content Generation pipeline failed.",
        log:
          logResult && logResult.success
            ? String(logResult.log || "")
            : "",
        cost:
          Number(
            result && result.cost
              ? result.cost
              : 0
          )
      };
    }

    return {
      success: true,
      row: row,
      postId: livePostId,
      title: title,
      message: "Pre-AC → W8E complete.",
      log:
        logResult && logResult.success
          ? String(logResult.log || "")
          : "",
      cost: Number(result.cost || 0),
      w8e: result.w8e || null
    };

  } catch (e) {
    return {
      success: false,
      row: Number(row) || 0,
      postId: String(expectedPostId || ""),
      title: "",
      message:
        e && e.message
          ? e.message
          : String(e),
      log: "",
      cost: 0
    };
  }
}

function bc_runFullGovernanceChainAutomated() {
  var preflight = bc_contentIntelligencePreflight_();

if (!preflight.success) {
  throw new Error(preflight.message);
}

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
    var skipGSC = false;
    var monitorDateText = String(
      postsSheet.getRange(activeRow, 64).getDisplayValue() || ""
    ).trim();

    if (monitorDateText) {
      var monitorDateParts = monitorDateText.split("—");

      if (monitorDateParts.length === 2) {
        var latestDateText = monitorDateParts[1].trim();
        var latestDateParts = latestDateText.split("/");

        if (latestDateParts.length === 3) {
          var latestMonitorDate = new Date(
            Number(latestDateParts[2]),
            Number(latestDateParts[1]) - 1,
            Number(latestDateParts[0])
          );

          latestMonitorDate.setHours(0, 0, 0, 0);

          var today = new Date();
          today.setHours(0, 0, 0, 0);

          var monitorAgeDays = Math.floor(
            (today.getTime() - latestMonitorDate.getTime()) / 86400000
          );

          if (monitorAgeDays >= 0 && monitorAgeDays < 30) {
            skipGSC = true;

            log.push(
              'GSC Sync: skipped — Monitor data is ' +
              monitorAgeDays +
              ' days old (' +
              latestDateText +
              ')'
            );
          }
        }
      }
    }

    if (!skipGSC) {
      var syncResult = runGSCFullSyncActiveRow();
      runGSCRowStatsActiveRow();

      log.push(
        'GSC Sync: completed for Post ID ' +
        syncResult.postId
      );
    }

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

  // Content Intelligence complete — mark Column D blue with black text.
  postsSheet
    .getRange(row, 4)
    .setBackground("#2196F3")
    .setFontColor("#000000");


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

        log.push(
          'Final Rule 17 Recheck: FAIL — running one final targeted W1.5B.1 correction.'
        );

        var finalRule17FixPrompt =
          buildH2GovernanceFixPromptW15B();

        if (
          typeof finalRule17FixPrompt !== 'string' ||
          finalRule17FixPrompt.indexOf('ERROR') === 0
        ) {
          throw new Error(
            'Final Rule 17 FAIL, but targeted fix prompt could not be built — ' +
            finalRule17FixPrompt
          );
        }

        var finalRule17BeforeFix =
          String(
            cvCell.getValue() || ''
          ).trim();

        var finalRule17Fix =
          bc_sendPromptViaOpenAI(
            finalRule17FixPrompt,
            null,
            MODEL_CHEAP
          );

        if (!finalRule17Fix.success) {
          throw new Error(
            finalRule17Fix.message
          );
        }

        totalCost +=
          finalRule17Fix.cost;

        var finalRule17Save =
          saveStage15BH2s(
            finalRule17Fix.text
          );

        if (!finalRule17Save.success) {
          throw new Error(
            finalRule17Save.message
          );
        }

        var finalRule17AfterFix =
          String(
            cvCell.getValue() || ''
          ).trim();

        if (
          finalRule17AfterFix ===
          finalRule17BeforeFix
        ) {
          throw new Error(
            'Final Rule 17 repair produced no change in CV.'
          );
        }

        log.push(
          'Final Rule 17 Repair: corrected headings saved to CV.'
        );

        var finalRule17Confirmation =
          bc_runH2GovernanceCheckOnlyW15B();

        totalCost +=
          finalRule17Confirmation.cost;

        if (!finalRule17Confirmation.passed) {

          bc_appendGovernancePipelineException(
            'W1.5B — Final Rule 17 Recheck',
            finalRule17Confirmation.rawResult ||
            'Final targeted W1.5B.1 correction did not resolve all Rule 17 failures.'
          );

          log.push(
            'Final Rule 17 Confirmation: FAIL recorded in GJ after one final correction attempt.'
          );

        } else {

          log.push(
            'Final Rule 17 Confirmation: PASS after targeted W1.5B.1 correction.'
          );
        }

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

  var h2sBeforeGovernanceFix =
    String(
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('posts')
        .getRange(
          SpreadsheetApp
            .getActiveSpreadsheet()
            .getSheetByName('posts')
            .getActiveRange()
            .getRow(),
          100
        )
        .getValue() || ''
    ).trim();

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

  var h2sAfterGovernanceFix =
    String(
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName('posts')
        .getRange(
          SpreadsheetApp
            .getActiveSpreadsheet()
            .getSheetByName('posts')
            .getActiveRange()
            .getRow(),
          100
        )
        .getValue() || ''
    ).trim();

  if (
    h2sAfterGovernanceFix ===
    h2sBeforeGovernanceFix
  ) {
    throw new Error(
      'Rule 17 repair produced no change in CV.'
    );
  }

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
  6000,
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
    8000,
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

function bc_applyVisibleTextFactCheckFixes_(fixes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('posts');
  var row = sh.getActiveCell().getRow();
  var html = String(sh.getRange(row, 187).getValue() || '');

  if (!html) {
    return { success: false, message: 'GE is empty.' };
  }

  if (!Array.isArray(fixes) || fixes.length === 0) {
    return { success: false, message: 'No fixes supplied.' };
  }

  function decodeEntity_(entity) {
    var map = {
      '&amp;': '&',
      '&lt;': '<',
      '&gt;': '>',
      '&quot;': '"',
      '&#39;': "'",
      '&apos;': "'",
      '&nbsp;': ' '
    };
    return map[entity] !== undefined ? map[entity] : entity;
  }

  function buildVisibleMap_(source) {
    var text = '';
    var starts = [];
    var ends = [];
    var i = 0;

    while (i < source.length) {
      if (source[i] === '<') {
        var close = source.indexOf('>', i);
        if (close === -1) {
          i++;
        } else {
          i = close + 1;
        }
        continue;
      }

      if (source[i] === '&') {
        var semi = source.indexOf(';', i);
        if (semi > i && semi - i <= 10) {
          var entity = source.slice(i, semi + 1);
          var decoded = decodeEntity_(entity);
          if (decoded !== entity && decoded.length === 1) {
            text += decoded;
            starts.push(i);
            ends.push(semi + 1);
            i = semi + 1;
            continue;
          }
        }
      }

      text += source[i];
      starts.push(i);
      ends.push(i + 1);
      i++;
    }

    return { text: text, starts: starts, ends: ends };
  }

  var validationErrors = [];
  var prepared = [];

  fixes.forEach(function(fix, index) {
    var label = String(fix.fixLabel || ('Fix ' + (index + 1)));
    var target = String(fix.targetText || '');
    var contextText = String(fix.contextText || '');
    var replacement =
      (typeof fix.newText === 'string') ? fix.newText : null;

    if (!target) {
      validationErrors.push(label + ' — targetText is missing.');
      return;
    }

    if (replacement === null) {
      validationErrors.push(label + ' — newText must be a string.');
      return;
    }

    var mapped = buildVisibleMap_(html);

    // The AI occasionally includes literal HTML closing tags in targetText,
    // although the fact-check contract requests visible text only.
    // Match against the same visible-text representation used for the article.
    // Never use the stripped string as replacement HTML.
    var visibleTarget = buildVisibleMap_(target).text;
    var visibleContext = buildVisibleMap_(contextText).text;

    if (!visibleTarget.trim()) {
      validationErrors.push(label + ' — targetText contains no visible text.');
      return;
    }

    var occurrences = [];
    var searchFrom = 0;
    var foundAt = -1;

    while ((foundAt = mapped.text.indexOf(visibleTarget, searchFrom)) !== -1) {
      occurrences.push(foundAt);
      searchFrom = foundAt + Math.max(1, visibleTarget.length);
    }

    if (occurrences.length === 0) {
      validationErrors.push(label + ' — targetText not found in visible article text.');
      return;
    }

    var first = -1;

    if (occurrences.length === 1) {
      first = occurrences[0];
    } else {
      if (!visibleContext) {
        validationErrors.push(
          label +
          ' — targetText is not unique in visible article text and contextText was not supplied.'
        );
        return;
      }

      var contextHits = [];
      var contextFrom = 0;
      var contextAt = -1;

      while ((contextAt = mapped.text.indexOf(visibleContext, contextFrom)) !== -1) {
        contextHits.push(contextAt);
        contextFrom = contextAt + Math.max(1, visibleContext.length);
      }

      if (contextHits.length === 0) {
        validationErrors.push(label + ' — contextText not found in visible article text.');
        return;
      }

      var candidates = occurrences.filter(function(targetPos) {
        return contextHits.some(function(contextPos) {
          var contextEnd = contextPos + visibleContext.length;
          var targetEnd = targetPos + visibleTarget.length;

          // Same local passage: overlap, containment, or within 350 visible chars.
          return (
            (targetPos >= contextPos && targetPos <= contextEnd) ||
            (contextPos >= targetPos && contextPos <= targetEnd) ||
            Math.abs(targetPos - contextPos) <= 350
          );
        });
      });

      if (candidates.length !== 1) {
        validationErrors.push(
          label +
          ' — contextText did not identify exactly one target occurrence.'
        );
        return;
      }

      first = candidates[0];
    }

    var htmlStart = mapped.starts[first];
    var htmlEnd = mapped.ends[first + visibleTarget.length - 1];
    var htmlSpan = html.slice(htmlStart, htmlEnd);
    // Include an adjacent closing anchor tag when the target ends inside a link.
    var opens = (htmlSpan.match(/<a\b[^>]*>/gi) || []).length;
    var closes = (htmlSpan.match(/<\/a\s*>/gi) || []).length;
    if (opens === closes + 1) {
      var closeTag = html.slice(htmlEnd).match(/^<\/a\s*>/i);
      if (closeTag) {
        htmlEnd += closeTag[0].length;
        htmlSpan = html.slice(htmlStart, htmlEnd);
      }
    }

    // If the visible target crosses governed links, preserve the ORIGINAL
    // exact anchor HTML using [[LINK_1]], [[LINK_2]], etc. placeholders.
    var anchors = htmlSpan.match(/<a\b[^>]*>[\s\S]*?<\/a>/gi) || [];
    var replacementHtml = replacement;

    if (anchors.length > 0) {
      for (var a = 0; a < anchors.length; a++) {
        var placeholder = '[[LINK_' + (a + 1) + ']]';

        if (replacementHtml.indexOf(placeholder) === -1) {
          if (anchors.length !== 1 || !replacementHtml.trim()) {
            validationErrors.push(label + ' — link preservation is ambiguous.');
            return;
          }
          replacementHtml = replacementHtml.trim() + ' (see ' + placeholder + ')';
        }

        replacementHtml =
          replacementHtml.replace(
            placeholder,
            anchors[a]
          );
      }

      if (/\[\[LINK_\d+\]\]/.test(replacementHtml)) {
        validationErrors.push(
          label + ' — newText contains an unmatched link placeholder.'
        );
        return;
      }
    } else if (/\[\[LINK_\d+\]\]/.test(replacementHtml)) {
      validationErrors.push(
        label + ' — newText contains a link placeholder but targetText crosses no link.'
      );
      return;
    }

    prepared.push({
      label: label,
      start: htmlStart,
      end: htmlEnd,
      newText: replacementHtml
    });
  });

  if (validationErrors.length > 0) {
    return {
      success: false,
      message:
        'Visible-text fact-check fix batch rejected — no changes applied. ' +
        validationErrors.join(' | ')
    };
  }

  // Reject overlapping replacements before writing anything to GE.
  var ordered = prepared.slice().sort(function(a, b) { return a.start - b.start; });
  for (var p = 1; p < ordered.length; p++) {
    if (ordered[p].start < ordered[p - 1].end) {
      return {
        success: false,
        message: 'Visible-text fact-check fix batch rejected — no changes applied. ' +
          ordered[p - 1].label + ' overlaps ' + ordered[p].label + '.'
      };
    }
  }

  // Apply from end to start so earlier character positions stay valid.
  prepared.sort(function(a, b) {
    return b.start - a.start;
  });

  var updated = html;

  prepared.forEach(function(item) {
    updated =
      updated.slice(0, item.start) +
      item.newText +
      updated.slice(item.end);
  });

  sh.getRange(row, 187).setValue(updated);

  if (typeof logPipelineResume === 'function') {
    logPipelineResume(
      'W2B.05 — Visible-Text Fact-Check Fixes Applied Atomically (GE)',
      ''
    );
  }

  return {
    success: true,
    message:
      prepared.length +
      '/' +
      prepared.length +
      ' visible-text fact-check fix(es) applied atomically to GE.'
  };
}


function bc_runFactCheckFullAutomated() {

  var totalCost = 0;
  var maxFixAttempts = 1; // One correction batch, then one independent verification.
  var attempt = 0;
  var lastFixMessage = '';
  var diagnosticRunId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') +
    '-' + Utilities.getUuid().slice(0, 8);
  function recordW2B05_(step, status, summary, details) {
    try {
      var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('W2B05 Audit Log');
      if (!log) {
        log = SpreadsheetApp.getActiveSpreadsheet().insertSheet('W2B05 Audit Log');
        log.appendRow(['Timestamp', 'Run ID', 'Post Row', 'Check/Attempt', 'Step', 'Status', 'Summary', 'Full Details', 'GE Length', 'Cost So Far']);
        log.setFrozenRows(1);
      }
      // A Sheet cell has a 50,000-character limit; split long responses
      // across successive audit rows so nothing is silently discarded.
      var raw = String(details == null ? '' : details);
      var pieces = raw.match(/[\s\S]{1,45000}/g) || [''];
      var geLength = String(sh.getRange(row, 187).getValue() || '').length;
      for (var z = 0; z < pieces.length; z++) {
        log.appendRow([new Date(), diagnosticRunId, row, attempt,
          step + (pieces.length > 1 ? ' (part ' + (z + 1) + '/' + pieces.length + ')' : ''),
          status, String(summary || ''), pieces[z], geLength, totalCost]);
      }
    } catch (logError) {
      console.error('W2B.05 diagnostic logging failed: ' + logError);
    }
  }

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

  // W2B.05 always starts clean: discard any previous GE result,
  // then seed GE from the current W2B Raw HTML in EU.
  var geCell = sh.getRange(row, 187);
  geCell.clearContent();
  geCell.setValue(currentEU);
  recordW2B05_('START', 'START', 'GE reset from current EU', 'EU characters: ' + currentEU.length);

  while (attempt <= maxFixAttempts) {

    var checkResult = bc_runFactCheckAutomated();

    if (!checkResult.success) {
      recordW2B05_('FACT CHECK API', 'ERROR', checkResult.message, JSON.stringify(checkResult));
      throw new Error(checkResult.message);
    }

    totalCost += Number(checkResult.cost || 0);
    recordW2B05_('AI RESPONSE', 'RETURNED', 'Full raw fact-check response', checkResult.text);

    var currentFactCheckHtml = String(
      sh.getRange(row, 187).getValue() || ''
    ).trim();

    // Keep the existing deterministic hard checks independent of the AI.
    var materialIntegrity =
      ce_checkMaterialCategoryIntegrity_(currentFactCheckHtml);

    var consistencyFinding =
      ce_findInternalFactConsistencyFinding_(
        currentFactCheckHtml
      );

    var consistencyOnlyFail = false;
    var parsed = null;

    // New W2B.05 contract: the FIRST fact-check call returns the atomic
    // old/new fixes itself. No second AI fix call is needed for normal
    // fact-check findings.
    try {
      var cleanedFactCheck = String(checkResult.text || '')
        .trim()
        .replace(/^\`\`\`json/i, '')
        .replace(/^\`\`\`/, '')
        .replace(/\`\`\`$/, '')
        .trim();

      parsed = JSON.parse(cleanedFactCheck);
    } catch (parseError) {
      recordW2B05_('PARSE', 'ERROR', parseError.message, checkResult.text);
      throw new Error('W2B.05 returned invalid JSON: ' + parseError.message);
    }

    if (
      !parsed ||
      (parsed.status !== 'PASS' && parsed.status !== 'PASS_WITH_NOTES' && parsed.status !== 'FAIL') ||
      !Array.isArray(parsed.fixes)
    ) {
      throw new Error(
        'W2B.05 returned JSON in an unexpected format.'
      );
    }

    recordW2B05_('PARSED RESULT', parsed.status, parsed.summary, JSON.stringify(parsed.fixes));

    // Deterministic hard checks still take priority.
    if (!materialIntegrity.passed) {

      parsed = {
        status: 'FAIL',
        summary: materialIntegrity.finding,
        fixes: []
      };

    } else if (consistencyFinding) {

      consistencyOnlyFail = true;

      parsed = {
        status: 'FAIL',
        summary: consistencyFinding,
        fixes: []
      };
    }

    if (!materialIntegrity.passed || consistencyFinding) {
      recordW2B05_('HARD CHECK', 'FAIL', parsed.summary, JSON.stringify({materialIntegrity: materialIntegrity, consistencyFinding: consistencyFinding}));
    }

    // PASS WITH NOTES is non-blocking: preserve GE and record the observation.
    if (parsed.status === 'PASS_WITH_NOTES') {
      if (attempt === 0) {
        var notesPass = passThroughFactCheckToGE();
        if (!notesPass.success) throw new Error(notesPass.message);
      }
      recordW2B05_('FINAL', 'PASS_WITH_NOTES', parsed.summary, checkResult.text);
      bc_appendGovernancePipelineException('W2B.05 — Non-blocking Notes',
        parsed.summary || 'Minor factual review observations.');
      return {
        success: true,
        warning: true,
        message: 'Fact-check PASS WITH NOTES — observation recorded in GJ. GE preserved.',
        text: checkResult.text,
        cost: totalCost
      };
    }

    // PASS
    if (parsed.status === 'PASS') {

      if (attempt === 0) {
        var passResult = passThroughFactCheckToGE();

        if (!passResult.success) {
          throw new Error(passResult.message);
        }
      }

      recordW2B05_('FINAL', 'PASS', 'Fact-check passed', lastFixMessage);
      return {
        success: true,
        message:
          attempt === 0
            ? 'Fact-check PASS — no changes needed. Saved in column GE.'
            : 'Fact-check PASS after ' +
              attempt +
              ' atomic fix attempt(s). ' +
              lastFixMessage +
              ' Final corrected HTML saved in column GE.',
        text: checkResult.text,
        cost: totalCost
      };
    }

    // Deterministic checks currently do not generate exact old/new pairs.
    // Preserve the existing repair fallback only for those rare hard-check
    // failures; ordinary AI fact-check failures never use a second AI call.
    if (parsed.fixes.length === 0) {
      recordW2B05_('FINAL', 'FAIL', 'Material error without safe exact correction', checkResult.text);
      bc_appendGovernancePipelineException('W2B.05 — Fact Check',
        'Material check failed without exact safe corrections. ' +
        (parsed.summary || '') + '\n\nFACT-CHECK RESPONSE:\n' + checkResult.text);
      return {
        success: false,
        message: 'Fact-check FAIL — no safe correction batch supplied; findings in GJ.',
        text: checkResult.text,
        cost: totalCost
      };
    }

    // The final verification may still FAIL after all permitted repairs.
    // Record that outcome explicitly rather than falling out of the loop
    // and returning undefined to the Master Workflow.
    if (attempt >= maxFixAttempts) {
      recordW2B05_('FINAL', 'FAIL', 'Correction limit reached', checkResult.text);
      bc_appendGovernancePipelineException(
        'W2B.05 — Fact Check',
        'Final verification still FAIL after ' + maxFixAttempts +
        ' correction attempt. ' + (parsed.summary || '') +
        '\n\nFACT-CHECK RESPONSE:\n' + checkResult.text
      );
      return {
        success: false,
        message: 'Fact-check still FAIL after one correction batch and verification — findings recorded in GJ.',
        text: checkResult.text,
        cost: totalCost
      };
    }

    // Normal path: apply the fixes proposed by the SAME fact-check call.
    var atomicFixResult =
      bc_applyVisibleTextFactCheckFixes_(
        parsed.fixes
      );

    lastFixMessage = atomicFixResult.message;
    recordW2B05_('ATOMIC FIX', atomicFixResult.success ? 'APPLIED' : 'REJECTED', atomicFixResult.message, JSON.stringify(parsed.fixes));

    if (!atomicFixResult.success) {

      bc_appendGovernancePipelineException(
        'W2B.05 — Fact Check',
        'Visible-text atomic fact-check fixes were rejected — ' +
        atomicFixResult.message +
        '\n\nFACT-CHECK RESPONSE:\n' +
        checkResult.text
      );

      return {
        success: false,
        message:
          'Fact-check found issues, but its atomic fixes could not be applied safely — recorded in GJ.',
        text: checkResult.text,
        cost: totalCost
      };
    }

    attempt++;

    // Loop now performs only a fresh fact-check of the corrected GE.
    // There is no separate AI fix-generation call.
  }

  recordW2B05_('FINAL', 'FAIL', 'Loop exited without PASS', lastFixMessage);
  return {
    success: false,
    message: 'W2B.05 stopped without a PASS — check GJ for details.',
    cost: totalCost
  };
}

function bc_runSimilarityCheckAutomated() {

  var startTime = Date.now();

  var promptData = buildCaseStudySimilarityPrompt('GE');

  if (!promptData.success) throw new Error(promptData.message);

  if (promptData.skipped) {
    return {
      success: true,
      skipped: true,
      text: '',
      message: promptData.message,
      cost: 0
    };
  }

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

function bc_runW2CPlanningAutomated() {
  var startTime = Date.now();

  var promptData = buildImageSuggestionPrompt('new_html');
  if (!promptData.success) throw new Error(promptData.message);

  var apiResult = bc_sendPromptViaOpenAI(promptData.prompt);
  if (!apiResult.success) throw new Error(apiResult.message);

  var saveResult = saveW2CSuggestions(apiResult.text);
  if (!saveResult.success) throw new Error(saveResult.message);

  var reconcileResult = reconcileImageSuggestions('new_html');
  if (!reconcileResult.success) throw new Error(reconcileResult.message);

  bc_addToApiCostAndTime(
    apiResult.cost,
    (Date.now() - startTime) / 1000
  );

  return {
    success: true,
    message:
      'W2C planning complete — visual requirements assessed and FL prepared for review. ' +
      reconcileResult.message,
    cost: apiResult.cost
  };
}

function bc_runStage3Automated() {
  var startTime = Date.now();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
    var posts = ss.getSheetByName('posts');
    var row = posts.getActiveRange().getRow();

    var finalW2B3Html =
      String(posts.getRange(row, 191).getValue() || '').trim(); // GI

    if (!finalW2B3Html) {
      throw new Error(
        'Column GI (W2B.3 Final HTML) is empty — W2B.3 must complete before W3.'
      );
    }

    var promptData = buildHumanisationPrompt(finalW2B3Html);
  if (!promptData.success) throw new Error(promptData.message);

  var apiResult = bc_sendPromptViaOpenAI(promptData.prompt, 8000);
  if (!apiResult.success) throw new Error(apiResult.message);

  var saveResult = saveHumanisedHtml(apiResult.text);
  if (!saveResult.success) throw new Error(saveResult.message);

  bc_addToApiCostAndTime(
    apiResult.cost,
    (Date.now() - startTime) / 1000
  );

  return {
    success: true,
    message: 'W3 humanisation complete — result saved to Column CZ.',
    cost: apiResult.cost
  };
}

function bc_runW4BAutomated() {
  var startTime = Date.now();
  var htmlResult = getHumanisedHtmlForW4B();
  if (!htmlResult.success) throw new Error(htmlResult.message);
  var html = htmlResult.html;

  var figResult = storeFigureInventory();

  var imgCount = (html.match(/<img[^>]+>/gi) || []).length;

    if (imgCount === 0) {
      return {
        success: true,
        message: 'W4B skipped — no images are present in the humanised HTML.',
        cost: 0
      };
    }

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
    '1. FLOOR CONDITION IMAGES — photos of the actual stone floor showing a condition, stage, or result. Write a descriptive alt AND a diagnostic caption that helps the reader identify their problem. Use varied natural framing and do not default to a fixed opening such as "If your floor...".',
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

  if (imgFilenames.length === 0) {
  return {
    success: true,
    message: 'W4B skipped — no eligible original images require alt/caption updates.',
    cost: 0
  };
  }

  var apiResult = bc_sendPromptViaOpenAI(prompt, 3000);
  if (!apiResult.success) throw new Error(apiResult.message);

  var pushResult = pushAltCaptionUpdates(apiResult.text);
  if (!pushResult.success) throw new Error(pushResult.message);

  bc_addToApiCostAndTime(apiResult.cost, (Date.now() - startTime) / 1000);
  return { success: true, message: pushResult.message, cost: apiResult.cost };
}

function bc_buildDeferredFixPromptServer_(html, result) {

  var failures = [];
  var fragments = [];


  /*
   * ---------------------------------------------------------
   * HEADER
   * Needed only for cluster/header failures.
   * ---------------------------------------------------------
   */

  if (
    result.check1 === 'FAIL' ||
    result.check5 === 'FAIL'
  ) {

    var headerMatch =
      html.match(
        /<header[^>]*>[\s\S]*?<\/header>/i
      );

    if (headerMatch) {

      fragments.push(
        'FRAGMENT TYPE: HEADER\n' +
        headerMatch[0]
      );
    }
  }


  if (result.check1 === 'FAIL') {
    failures.push(
      'CHECK 1 — Cluster tone: ' +
      result.obs1
    );
  }

  if (result.check5 === 'FAIL') {
    failures.push(
      'CHECK 5 — Header cluster tone: ' +
      result.obs5
    );
  }


  /*
   * ---------------------------------------------------------
   * SECTION OPENERS
   * Only send the first paragraph of affected sections.
   * ---------------------------------------------------------
   */

  if (result.check4 === 'FAIL') {

    failures.push(
      'CHECK 4 — Section opening tone: ' +
      result.obs4
    );

    var sectionIds =
      String(result.obs4 || '')
        .match(/section-\d+/gi) || [];

    var seenSections = {};

    sectionIds.forEach(function(sectionId) {

      sectionId =
        sectionId.toLowerCase();

      if (seenSections[sectionId]) {
        return;
      }

      seenSections[sectionId] = true;

      var sectionRegex =
        new RegExp(
          '<section[^>]*id=["\']' +
          sectionId +
          '["\'][^>]*>[\\s\\S]*?<\\/section>',
          'i'
        );

      var sectionMatch =
        html.match(sectionRegex);

      if (!sectionMatch) {
        return;
      }

      var h2Match =
        sectionMatch[0].match(
          /<h2[^>]*>[\s\S]*?<\/h2>/i
        );

      var pMatch =
        sectionMatch[0].match(
          /<p[^>]*>[\s\S]*?<\/p>/i
        );

      if (pMatch) {

        fragments.push(
          'FRAGMENT TYPE: SECTION OPENER — ' +
          sectionId +
          '\n' +
          (h2Match ? h2Match[0] + '\n' : '') +
          pMatch[0]
        );
      }
    });
  }


  /*
   * ---------------------------------------------------------
   * NAMED DEFECT PARAGRAPHS
   * Only send paragraphs containing relevant defect terms.
   * ---------------------------------------------------------
   */

  if (result.check7 === 'FAIL') {

    failures.push(
      'CHECK 7 — Named defect completion: ' +
      result.obs7
    );

    var defectTerms = [
    'delamination',
    'sealer failure',
    'efflorescence',
    'filler collapse',
    'colour loss',
    'grout haze',
    'lippage',
    'spalling',
    'micro-scratching',
    'residue lock-in',
    'adhesive failure',
    'debonded substrate',
    'glaze crazing',
    'water ingress'
  ];

    var paragraphs =
      html.match(
        /<p[^>]*>[\s\S]*?<\/p>/gi
      ) || [];

    paragraphs.forEach(function(p) {

      var plain =
        p.replace(/<[^>]+>/g, ' ')
         .toLowerCase();

      var relevant =
        defectTerms.some(function(term) {
          return plain.indexOf(term) !== -1;
        });

      if (relevant) {

        fragments.push(
          'FRAGMENT TYPE: DEFECT PARAGRAPH\n' +
          p
        );
      }
    });
  }


  /*
   * ---------------------------------------------------------
   * ENTITY-DUMP CHECK
   * Only send suspicious comma-heavy paragraphs.
   * ---------------------------------------------------------
   */

  if (result.check8 === 'FAIL') {

    failures.push(
      'CHECK 8 — Entity dump/comma-list prose: ' +
      result.obs8
    );

    var dumpParagraphs =
      html.match(
        /<p[^>]*>[\s\S]*?<\/p>/gi
      ) || [];

    dumpParagraphs.forEach(function(p) {

      var plain =
        p.replace(/<[^>]+>/g, ' ');

      var commaCount =
        (plain.match(/,/g) || []).length;

      if (commaCount >= 4) {

        fragments.push(
          'FRAGMENT TYPE: COMMA-HEAVY PARAGRAPH\n' +
          p
        );
      }
    });
  }


  /*
   * ---------------------------------------------------------
   * DEDUPLICATE
   * ---------------------------------------------------------
   */

  fragments =
    fragments.filter(
      function(value, index, array) {
        return array.indexOf(value) === index;
      }
    );


  return [
    'DEFERRED CHECK FIX — TARGETED FRAGMENT CORRECTION',
    '',
    'ROLE: Senior UK SEO content editor for a natural stone floor restoration website.',
    '',
    'TASK:',
    'Correct ONLY the supplied HTML fragments needed to resolve the failures.',
    'Do NOT return the complete article.',
    '',
    'RETURN ONLY A JSON ARRAY.',
    '',
    'FORMAT:',
    '[{"checkId":"4","old":"EXACT ORIGINAL HTML","new":"CORRECTED HTML"}]',
    '',
    'RULES:',
    '- "old" MUST be copied exactly from the supplied fragment.',
    '- "new" must contain only the corrected replacement for that exact fragment.',
    '- Preserve HTML tags.',
    '- Preserve links.',
    '- Preserve figures and images.',
    '- Do not introduce new facts.',
    '- Do not rewrite unaffected material.',
    '- Do not return markdown.',
    '- Do not return explanations.',
    '',
    'FAILURES:',
    failures.join('\n'),
    '',
    'TARGET FRAGMENTS:',
    fragments.join('\n\n---\n\n')
  ].join('\n');
}

function bc_runW4Automated() {

  var startTime = Date.now();
  var totalCost = 0;
  var log = [];
  var exceptionCount = 0;

  var w4RunStamp =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );


  function recordW4Exception(label, message) {

    exceptionCount++;

    bc_appendGovernancePipelineException(
      label + ' [' + w4RunStamp + ']',
      message || 'Unresolved W4 issue.'
    );

    log.push(
      label + ' recorded in GJ.'
    );
  }


  function isGovernedLateralLinkFailure(failure) {

    if (!failure) return false;

    var checkId =
      String(
        failure.checkId || ''
      ).toLowerCase();

    var description =
      String(
        failure.description || ''
      ).toLowerCase();

    return (
      checkId === '2-lateral' ||
      (
        checkId === '2' &&
        description.indexOf(
          'lateral'
        ) !== -1
      )
    );
  }


  function buildCheck4OnlyRecheckPrompt(html) {

    var sectionOpeners = [];
    var secRe =
      /<section[^>]*id="([^"]+)"[^>]*>([\s\S]*?)<\/section>/gi;

    var sm;

    while (
      (sm = secRe.exec(html)) !== null
    ) {

      var h2m =
        sm[2].match(
          /<h2[^>]*>([\s\S]*?)<\/h2>/i
        );

      var pm =
        sm[2].match(
          /<p[^>]*>([\s\S]*?)<\/p>/i
        );

      if (h2m && pm) {

        sectionOpeners.push(
          'Section "' +
          sm[1] +
          '" — H2: "' +
          h2m[1]
            .replace(/<[^>]+>/g, '')
            .trim() +
          '"\nOpening: "' +
          pm[1]
            .replace(/<[^>]+>/g, '')
            .trim() +
          '"'
        );
      }
    }


    return [
      'CHECK 4 ONLY — SECTION OPENING TONE',
      '',
      'ROLE: Senior UK SEO & Editorial Quality Auditor',
      '',
      'TASK:',
      'Check only whether each section opening sentence orients the reader to their visible problem before introducing mechanism, company, entity or process.',
      '',
      'Do not assess cluster tone.',
      'Do not assess the header.',
      'Do not assess named defects.',
      'Do not assess any other editorial rule.',
      '',
      'SECTION OPENERS:',
      sectionOpeners.join('\n\n'),
      '',
      'RETURN EXACTLY:',
      'CHECK 4: PASS|FAIL',
      'OBSERVATION 4: [one sentence — name failing section ids if any]'
    ].join('\n');
  }


  function buildCheck7OnlyRecheckPrompt(html) {

    return [
      'CHECK 7 ONLY — NAMED DEFECT COMPLETION',
      '',
      'ROLE: Senior UK SEO & Editorial Quality Auditor',
      '',
      'TASK:',
      'Check only whether every named technical defect is explained to completion before the next paragraph.',
      '',
      'Each named defect must contain:',
      'ELEMENT 1 — what the defect is physically.',
      'ELEMENT 2 — what the homeowner sees or notices.',
      'ELEMENT 3 — what professional correction does about it.',
      '',
      'Named defects include:',
      'delamination, sealer failure, efflorescence, filler collapse, colour loss, grout haze, lippage, spalling, micro-scratching, residue lock-in.',
      '',
      'Do not assess section opening tone.',
      'Do not assess cluster tone.',
      'Do not assess links.',
      'Do not assess any other editorial rule.',
      '',
      'ARTICLE HTML:',
      html,
      '',
      'RETURN EXACTLY:',
      'CHECK 7: PASS|FAIL',
      'OBSERVATION 7: [one sentence — name any incomplete defect]'
    ].join('\n');
  }


  /*
   * =========================================================
   * STEP 1 — LOAD CZ
   * =========================================================
   */

  var html =
    getHtmlFromActiveRow();

  if (
    !html ||
    html.indexOf('ERROR') === 0
  ) {

    throw new Error(
      html ||
      'No HTML found in CZ.'
    );
  }


  /*
   * =========================================================
   * STEP 2 — MECHANICAL AUDIT
   * =========================================================
   */

  var audit =
    runAtomicAuditWithContext(
      html
    );

  if (audit.error) {

    recordW4Exception(
      'W4 — Mechanical Audit',
      audit.error
    );

  } else if (!audit.allPass) {

    var remaining = [];


    audit.failures.forEach(
      function(f) {

        if (
          isGovernedLateralLinkFailure(f)
        ) {

          log.push(
            'Governed lateral-link issue deferred to Final Exception Check.'
          );

          return;
        }


        if (f.canAutoFix) {

          var fixRes =
            applyAutoFix(
              html,
              f.checkId
            );

          if (
            fixRes &&
            fixRes.success
          ) {

            html =
              fixRes.patchedHtml;

            log.push(
              'Auto-fixed CHECK ' +
              f.checkId
            );

          } else {

            remaining.push(f);
          }

        } else {

          remaining.push(f);
        }
      }
    );


    if (remaining.length > 0) {

      var batchPrompt =
        buildBatchAtomicFixPrompt(
          remaining
        );

      if (
        batchPrompt &&
        batchPrompt.success
      ) {

        var batchApiResult =
          bc_sendPromptViaOpenAI(
            batchPrompt.prompt,
            3500
          );

        if (
          batchApiResult &&
          batchApiResult.success
        ) {

          totalCost +=
            batchApiResult.cost || 0;

          var batchApply =
            applyBatchAtomicFix(
              html,
              batchApiResult.text
            );

          if (
            batchApply &&
            batchApply.success
          ) {

            html =
              batchApply.patchedHtml;

            log.push(
              'Mechanical JSON patches applied.'
            );

          } else {

            log.push(
              'Mechanical JSON patch could not be applied safely.'
            );
          }

        } else {

          log.push(
            'Mechanical correction API call did not complete.'
          );
        }
      }
    }


    var mechanicalRecheck =
      runAtomicAuditWithContext(
        html
      );

    if (mechanicalRecheck.error) {

      recordW4Exception(
        'W4 — Mechanical Audit',
        mechanicalRecheck.error
      );

    } else if (
      !mechanicalRecheck.allPass
    ) {

      var lateralFailures = [];
      var otherFailures = [];


      (
        mechanicalRecheck.failures ||
        []
      ).forEach(
        function(f) {

          if (
            isGovernedLateralLinkFailure(f)
          ) {

            lateralFailures.push(f);

          } else {

            otherFailures.push(f);
          }
        }
      );


      if (
        lateralFailures.length > 0
      ) {

        recordW4Exception(
          'W4 — Governed Lateral Link',
          'No governed lateral internal link is present in the article. ' +
          'W4 did not create or invent a link because lateral-link selection belongs to the governed plan. ' +
          'Review this item in the Final Exception Check.'
        );
      }


      if (
        otherFailures.length > 0
      ) {

        var failureSummary =
          otherFailures.map(
            function(f) {

              return (
                'CHECK ' +
                (f.checkId || '?') +
                ': ' +
                (
                  f.description ||
                  'Unresolved mechanical issue'
                )
              );
            }
          ).join('\n');

        recordW4Exception(
          'W4 — Mechanical Audit',
          failureSummary
        );

      } else {

        log.push(
          'No other mechanical failures remain.'
        );
      }

    } else {

      log.push(
        'Mechanical audit passed.'
      );
    }

  } else {

    log.push(
      'Mechanical audit passed first time.'
    );
  }


  /*
   * =========================================================
   * STEP 3 — ONE FULL DEFERRED AUDIT
   * =========================================================
   */

  var deferredPrompt =
    buildDeferredCheckPrompt(
      html
    );

  var deferredApiResult =
    bc_sendPromptViaOpenAI(
      deferredPrompt,
      1800
    );

  if (
    deferredApiResult &&
    deferredApiResult.success
  ) {

    totalCost +=
      deferredApiResult.cost || 0;

    var deferredResult =
      parseDeferredCheckResponse(
        deferredApiResult.text
      );


    /*
     * Checks 1, 5 and 8 remain locked.
     * Check 7 now receives its own targeted repair below.
     */

    var lockedFailures = [];

    if (
      deferredResult.check1 === 'FAIL'
    ) {
      lockedFailures.push(
        'CHECK 1: ' +
        deferredResult.obs1
      );
    }

    if (
      deferredResult.check5 === 'FAIL'
    ) {
      lockedFailures.push(
        'CHECK 5: ' +
        deferredResult.obs5
      );
    }

    if (
      deferredResult.check8 === 'FAIL'
    ) {
      lockedFailures.push(
        'CHECK 8: ' +
        deferredResult.obs8
      );
    }


    if (
      lockedFailures.length > 0
    ) {

      recordW4Exception(
        'W4 — Deferred Checks',
        lockedFailures.join('\n')
      );
    }


    /*
     * =======================================================
     * CHECK 7 ONLY — TARGETED DEFECT COMPLETION REPAIR
     * =======================================================
     */

    if (
      deferredResult.check7 === 'FAIL'
    ) {

      var check7OnlyResult = {
        check1: 'PASS',
        obs1: '',
        check4: 'PASS',
        obs4: '',
        check5: 'PASS',
        obs5: '',
        check7: 'FAIL',
        obs7: deferredResult.obs7,
        check8: 'PASS',
        obs8: ''
      };


      var check7FixPrompt =
        bc_buildDeferredFixPromptServer_(
          html,
          check7OnlyResult
        );


      var check7FixApi =
        bc_sendPromptViaOpenAI(
          check7FixPrompt,
          2200
        );


      if (
        check7FixApi &&
        check7FixApi.success
      ) {

        totalCost +=
          check7FixApi.cost || 0;


        var check7Apply =
          applyBatchAtomicFix(
            html,
            check7FixApi.text
          );


        if (
          check7Apply &&
          check7Apply.success
        ) {

          html =
            check7Apply.patchedHtml;

          log.push(
            'Check 7 JSON patch applied.'
          );


          var check7RecheckPrompt =
            buildCheck7OnlyRecheckPrompt(
              html
            );


          var check7RecheckApi =
            bc_sendPromptViaOpenAI(
              check7RecheckPrompt,
              700
            );


          if (
            check7RecheckApi &&
            check7RecheckApi.success
          ) {

            totalCost +=
              check7RecheckApi.cost || 0;


            var check7Text =
              String(
                check7RecheckApi.text || ''
              );


            var check7Match =
              check7Text.match(
                /CHECK 7:\s*(PASS|FAIL)/i
              );


            if (
              !check7Match ||
              check7Match[1]
                .toUpperCase() !==
                'PASS'
            ) {

              recordW4Exception(
                'W4 — Deferred Check 7',
                check7Text ||
                'Named-defect check remains unresolved.'
              );

            } else {

              log.push(
                'Check 7 passed after targeted correction.'
              );
            }

          } else {

            recordW4Exception(
              'W4 — Deferred Check 7',
              'Check 7 recheck could not be completed.'
            );
          }

        } else {

          recordW4Exception(
            'W4 — Deferred Check 7',
            deferredResult.obs7 ||
            'Check 7 correction could not be applied safely.'
          );
        }

      } else {

        recordW4Exception(
          'W4 — Deferred Check 7',
          deferredResult.obs7 ||
          'Check 7 correction API call failed.'
        );
      }

    } else {

      log.push(
        'Check 7 passed first time.'
      );
    }


    /*
     * =======================================================
     * CHECK 4 ONLY — SAFE JSON REPAIR
     * =======================================================
     */

    if (
      deferredResult.check4 === 'FAIL'
    ) {

      var check4OnlyResult = {
        check1: 'PASS',
        obs1: '',
        check4: 'FAIL',
        obs4: deferredResult.obs4,
        check5: 'PASS',
        obs5: '',
        check7: 'PASS',
        obs7: '',
        check8: 'PASS',
        obs8: ''
      };


      var check4FixPrompt =
        bc_buildDeferredFixPromptServer_(
          html,
          check4OnlyResult
        );


      var check4FixApi =
        bc_sendPromptViaOpenAI(
          check4FixPrompt,
          2200
        );


      if (
        check4FixApi &&
        check4FixApi.success
      ) {

        totalCost +=
          check4FixApi.cost || 0;


        var check4Apply =
          applyBatchAtomicFix(
            html,
            check4FixApi.text
          );


        if (
          check4Apply &&
          check4Apply.success
        ) {

          html =
            check4Apply.patchedHtml;

          log.push(
            'Check 4 JSON patches applied.'
          );


          var check4RecheckPrompt =
            buildCheck4OnlyRecheckPrompt(
              html
            );


          var check4RecheckApi =
            bc_sendPromptViaOpenAI(
              check4RecheckPrompt,
              700
            );


          if (
            check4RecheckApi &&
            check4RecheckApi.success
          ) {

            totalCost +=
              check4RecheckApi.cost || 0;


            var check4Text =
              String(
                check4RecheckApi.text || ''
              );


            var check4Match =
              check4Text.match(
                /CHECK 4:\s*(PASS|FAIL)/i
              );


            if (
              !check4Match ||
              check4Match[1]
                .toUpperCase() !==
                'PASS'
            ) {

              recordW4Exception(
                'W4 — Deferred Check 4',
                check4Text ||
                'Section-opening check remains unresolved.'
              );

            } else {

              log.push(
                'Check 4 passed after targeted correction.'
              );
            }

          } else {

            recordW4Exception(
              'W4 — Deferred Check 4',
              'Check 4 recheck could not be completed.'
            );
          }

        } else {

          recordW4Exception(
            'W4 — Deferred Check 4',
            deferredResult.obs4 ||
            'Check 4 correction could not be applied safely.'
          );
        }

      } else {

        recordW4Exception(
          'W4 — Deferred Check 4',
          deferredResult.obs4 ||
          'Check 4 correction API call failed.'
        );
      }

    } else {

      log.push(
        'Check 4 passed first time.'
      );
    }

  } else {

    recordW4Exception(
      'W4 — Deferred Checks',
      deferredApiResult &&
      deferredApiResult.message
        ? deferredApiResult.message
        : 'Deferred audit API call failed.'
    );
  }


  /*
   * =========================================================
   * STEP 4 — SAVE FINAL HTML TO CT
   * =========================================================
   */

  var pushMessage =
    pushHtmlToActiveRow(
      html
    );

  if (
    !pushMessage ||
    /^ERROR|^PUSH ERROR/i.test(
      pushMessage
    )
  ) {

    throw new Error(
      pushMessage ||
      'Could not save W4 HTML to CT.'
    );
  }


  /*
   * =========================================================
   * STEP 5 — FINISH
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
      exceptionCount === 0
        ? 'W4 complete — all checks passed and final HTML saved to Column CT.'
        : 'W4 complete — final HTML saved to Column CT. ' +
          exceptionCount +
          ' unresolved issue(s) recorded in Column GJ for the Final Exception Check.',

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

function pipelineResumeToW8E(row, startStage) {

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

  row = Number(row);

  if (!row || row < 2) {
    return {
      success: false,
      cost: 0,
      message: "Invalid pipeline row."
    };
  }

  sheet.setActiveRange(
    sheet.getRange(row, 1)
  );


  var stages = [

    {
      name: "Pre-AC",
      selfLogged: true,
      run: function() {
        return bc_pipelineRunPreAC(row);
      }
    },

    {
      name: "AC",
      selfLogged: true,
      run: function() {
        return pipelineRunAC(row);
      }
    },

    {
      name: "W-1",
      selfLogged: true,
      run: function() {
        return pipelineRunAuthorityBrief(row);
      }
    },

    {
      name: "W1A-W1D",
      run: function() {
        return saveAllW1Outputs();
      }
    },

    {
      name: "W1.5A",
      run: function() {
        return bc_runStage15AAutomated();
      }
    },

    {
      name: "W1.5B",
      run: function() {
        return bc_runFullW15BAutomated();
      }
    },

    {
      name: "W1.5C",
      run: function() {
        return bc_runStage15CAutomated();
      }
    },

    {
      name: "W1.5D",
      run: function() {
        return bc_runStage15DAutomated();
      }
    },

    {
      name: "W1.5E",
      run: function() {
        return bc_runStage15EAutomated();
      }
    },

    {
      name: "W2B",
      run: function() {
        return bc_runStage2BAutomated();
      }
    },

    {
      name: "W2B.05",
      run: function() {
        return bc_runFactCheckFullAutomated();
      }
    },

    {
      name: "W2B.1",
      run: function() {
        return bc_runSimilarityFullAutomated();
      }
    },

    {
      name: "W2B.2",
      run: function() {
        return bc_runRewriteBriefComplianceW2BAutomated();
      }
    },

    {
      name: "W2B.3",
      run: function() {
        return bc_runCoherenceFullAutomated();
      }
    },

    {
      name: "W3",
      run: function() {
        return bc_runStage3Automated();
      }
    },

    {
      name: "W4B",
      run: function() {
        return bc_runW4BAutomated();
      }
    },

    {
      name: "W4",
      run: function() {
        return bc_runW4Automated();
      }
    },

    {
      name: "W4.5",
      run: function() {
        return bc_runCompetitorSerpNotesAutomated();
      }
    },

    {
      name: "W5",
      run: function() {
        return bc_runW5Automated();
      }
    },

    {
      name: "W5C",
      run: function() {
        return generateSchemaForActiveRow();
      }
    },

    {
      name: "W8E",
      run: function() {
        return bc_runW8EAutomated();
      }
    }

  ];


  var startIndex = -1;

  for (var i = 0; i < stages.length; i++) {

    if (stages[i].name === startStage) {
      startIndex = i;
      break;
    }
  }


  if (startIndex === -1) {

    return {
      success: false,
      cost: 0,
      message:
        "Unknown resume stage: " +
        startStage
    };
  }


  for (
    var stageIndex = startIndex;
    stageIndex < stages.length;
    stageIndex++
  ) {

    var stage = stages[stageIndex];


    if (pipelineStopRequested(row)) {

      bc_appendPipelineRunLog(
        row,
        "■ Pipeline stopped by user before " +
        stage.name +
        "."
      );

      return {
        success: false,
        stopped: true,
        cost: totalCost,
        message:
          "Pipeline stopped before " +
          stage.name +
          "."
      };
    }


    var result;

    try {

      result = stage.run();

    } catch (e) {

      if (!stage.selfLogged) {
        bc_recordPipelineFailure(
          row,
          stage.name,
          e.toString()
        );
      }

      return {
        success: false,
        cost: totalCost,
        message:
          stage.name +
          " failed: " +
          e.toString()
      };
    }


    if (
      !result ||
      result.success === false
    ) {

      var message =
        result && result.message
          ? result.message
          : "Unknown " +
            stage.name +
            " error";

      if (!stage.selfLogged) {

        bc_recordPipelineFailure(
          row,
          stage.name,
          message
        );
      }

      return {
        success: false,
        stopped:
          result &&
          result.stopped === true,
        cost:
          totalCost +
          Number(
            result && result.cost
              ? result.cost
              : 0
          ),
        message: message
      };
    }


    var stageCost =
      Number(result.cost || 0);

    totalCost += stageCost;


    if (!stage.selfLogged) {

      bc_recordPipelineStage(
        row,
        stage.name,
        stageCost
      );
    }
  }


  bc_appendPipelineRunLog(
    row,
    "✓ Recovery " +
    startStage +
    " → W8E complete — $" +
    totalCost.toFixed(4)
  );


  return {
    success: true,
    cost: totalCost,
    message:
      startStage +
      " → W8E complete."
  };
}
  function runStage(stageName, runner) {

    var result;

    try {
      result = runner();
    } catch (e) {
      bc_recordPipelineFailure(
        row,
        stageName,
        e.toString()
      );

      return {
        success: false,
        message: e.toString()
      };
    }

    if (
      !result ||
      result.success === false
    ) {

      var message =
        result && result.message
          ? result.message
          : "Unknown " +
            stageName +
            " error";

      bc_recordPipelineFailure(
        row,
        stageName,
        message
      );

      return {
        success: false,
        message: message
      };
    }

    var stageCost =
      Number(result.cost || 0);

    totalCost += stageCost;

    bc_recordPipelineStage(
      row,
      stageName,
      stageCost
    );

    return {
      success: true
    };
  }

  
function pipelineRunW3ToW8E(row) {

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


  function runStage(stageName, runner) {

    if (pipelineStopRequested(row)) {
      return {
        success: false,
        stopped: true,
        message:
          "Pipeline stopped by user before " +
          stageName +
          "."
      };
    }

    try {

      var result = runner();

      if (
        !result ||
        result.success === false
      ) {

        var message =
          result && result.message
            ? result.message
            : "Unknown " + stageName + " error";

        bc_recordPipelineFailure(
          row,
          stageName,
          message
        );

        return {
          success: false,
          message: message,
          cost:
            Number(
              result && result.cost
                ? result.cost
                : 0
            )
        };
      }

      var stageCost =
        Number(result.cost || 0);

      totalCost += stageCost;

      bc_recordPipelineStage(
        row,
        stageName,
        stageCost
      );

      return {
        success: true,
        result: result
      };

    } catch (e) {

      bc_recordPipelineFailure(
        row,
        stageName,
        e.toString()
      );

      return {
        success: false,
        message: e.toString(),
        cost: 0
      };
    }
  }


  // ============================================================
  // W3
  // ============================================================

  var stage = runStage(
    "W3",
    function() {
      return bc_runStage3Automated();
    }
  );

  if (!stage.success) {
    return {
      success: false,
      stopped: stage.stopped === true,
      cost: totalCost,
      message: stage.message
    };
  }


  // ============================================================
  // W4B
  // ============================================================

  stage = runStage(
    "W4B",
    function() {
      return bc_runW4BAutomated();
    }
  );

  if (!stage.success) {
    return {
      success: false,
      stopped: stage.stopped === true,
      cost: totalCost,
      message: stage.message
    };
  }


  // ============================================================
  // W4
  // ============================================================

  stage = runStage(
    "W4",
    function() {
      return bc_runW4Automated();
    }
  );

  if (!stage.success) {
    return {
      success: false,
      stopped: stage.stopped === true,
      cost: totalCost,
      message: stage.message
    };
  }


  // ============================================================
  // W4.5
  // ============================================================

  stage = runStage(
    "W4.5",
    function() {
      return bc_runCompetitorSerpNotesAutomated();
    }
  );

  if (!stage.success) {
    return {
      success: false,
      stopped: stage.stopped === true,
      cost: totalCost,
      message: stage.message
    };
  }


  // ============================================================
  // W5
  // ============================================================

  stage = runStage(
    "W5",
    function() {
      return bc_runW5Automated();
    }
  );

  if (!stage.success) {
    return {
      success: false,
      stopped: stage.stopped === true,
      cost: totalCost,
      message: stage.message
    };
  }


  // ============================================================
  // W5C
  // ============================================================

  stage = runStage(
    "W5C",
    function() {
      return generateSchemaForActiveRow();
    }
  );

  if (!stage.success) {
    return {
      success: false,
      stopped: stage.stopped === true,
      cost: totalCost,
      message: stage.message
    };
  }


  // ============================================================
  // W8E
  // ============================================================

  stage = runStage(
    "W8E",
    function() {
      return bc_runW8EAutomated();
    }
  );

  if (!stage.success) {
    return {
      success: false,
      stopped: stage.stopped === true,
      cost: totalCost,
      message: stage.message
    };
  }


  bc_appendPipelineRunLog(
    row,
    "✓ W3 → W8E complete — $" +
    totalCost.toFixed(4)
  );


  return {
    success: true,
    cost: totalCost,
    message: "W3 → W8E complete.",
    w8e: stage.result
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

function lockW5Keyphrase_(text, lockedKeyphrase) {

  if (!lockedKeyphrase) {
    return String(text || "");
  }

  return String(text || "").replace(
    /Yoast\s+Keyphrase:\s*.+/i,
    "Yoast Keyphrase: " + lockedKeyphrase
  );
}


function bc_runW5Automated() {

  var startTime = Date.now();
  var totalCost = 0;

  var w5RowData = getActiveRowDataMap();

  var w5PostId =
    String(w5RowData["Post ID"] || "").trim();


  var w5MonitorQueries = w5PostId
    ? String(bc_getQueriesFromSheet(w5PostId, "GSC Monitor") || "").trim()
    : String(w5RowData["Montr Queries"] || "").trim();


  var w5LockedKeyphrase =
    w5MonitorQueries.length <= 10
      ? String(w5RowData["Primary Search Term"] || "").trim()
      : "";



  /*
   * =========================================================
   * PARSE W5 OUTPUT
   * =========================================================
   */

  function parseW5Output(text) {

    text = String(text || "")
      .replace(
        /\s+(New\s+H1:|New\s+Meta\s+Title:|New\s+Meta\s+Description:|Yoast\s+Keyphrase:)/gi,
        '\n$1'
      )
      .trim();


    function extract(pattern) {

      var match =
        text.match(pattern);

      return match
        ? String(match[1] || "").trim()
        : "";
    }


    return {

      h1:
        extract(
          /(?:^|\n)New\s+H1:\s*(.+)/i
        ),

      title:
        extract(
          /(?:^|\n)New\s+Meta\s+Title:\s*(.+)/i
        ),

      description:
        extract(
          /(?:^|\n)New\s+Meta\s+Description:\s*(.+)/i
        ),

      keyphrase:
        extract(
          /(?:^|\n)Yoast\s+Keyphrase:\s*(.+)/i
        )
    };
  }


  /*
   * =========================================================
   * DETERMINISTIC VALIDATION
   * =========================================================
   */

  function validateW5Output(fields) {

    var issues = [];


    /*
     * REQUIRED FIELDS
     */

    if (!fields.h1) {
      issues.push(
        "New H1 missing."
      );
    }

    if (!fields.title) {
      issues.push(
        "New Meta Title missing."
      );
    }

    if (!fields.description) {
      issues.push(
        "New Meta Description missing."
      );
    }

    if (!fields.keyphrase) {
      issues.push(
        "Yoast Keyphrase missing."
      );
    }


    /*
     * H1
     */

    if (fields.h1) {

      if (
        fields.h1.length < 40 ||
        fields.h1.length > 60
      ) {

        issues.push(
          "H1 is " +
          fields.h1.length +
          " characters; required range is 40-60."
        );
      }
    }


    /*
     * META TITLE
     */

    if (fields.title) {

      if (
        fields.title.length > 60
      ) {

        issues.push(
          "Meta Title is " +
          fields.title.length +
          " characters; maximum is 60."
        );
      }


      var rowData =
        getActiveRowDataMap();


      var primaryTerm =
        String(
          rowData["Primary Search Term"] ||
          rowData["Primary Query Cluster Owned"] ||
          ""
        )
          .toLowerCase()
          .replace(/\bnear me\b/g, "")
          .replace(/[^a-z0-9\s]/g, " ")
          .replace(/\s+/g, " ")
          .trim();


      var location =
        String(
          rowData["Locality"] ||
          rowData["Location"] ||
          ""
        )
          .toLowerCase()
          .trim();


      var titleCore =
        fields.title
          .toLowerCase()
          .replace(
            /abbey floor care/g,
            ""
          )
          .replace(
            /[|–—:-]/g,
            " "
          )
          .replace(
            /[^a-z0-9\s]/g,
            " "
          )
          .replace(
            /\s+/g,
            " "
          )
          .trim();


      if (location) {

        var escapedLocation =
          location.replace(
            /[.*+?^${}()|[\]\\]/g,
            "\\$&"
          );


        titleCore =
          titleCore
            .replace(
              new RegExp(
                "\\b" +
                escapedLocation +
                "\\b",
                "gi"
              ),
              ""
            )
            .replace(
              /\s+/g,
              " "
            )
            .trim();
      }


      /*
       * Basic deterministic protection against
       * service/search-term + location titles.
       *
       * The semantic check below makes the final
       * linguistic judgement.
       */

      if (
        primaryTerm &&
        titleCore === primaryTerm
      ) {

        issues.push(
          "Meta Title is effectively only the search term/location and has no distinct CTR-driving angle."
        );
      }
    }


    /*
     * META DESCRIPTION
     */

    if (fields.description) {

      if (
        fields.description.length < 140 ||
        fields.description.length > 160
      ) {

        issues.push(
          "Meta Description is " +
          fields.description.length +
          " characters; required range is 140-160."
        );
      }


      var sentenceCount =
        (
          fields.description.match(
            /[.!?](?:\s|$)/g
          ) || []
        ).length;


      if (
        sentenceCount !== 2
      ) {

        issues.push(
          "Meta Description must contain exactly two sentences."
        );
      }
    }


    /*
     * YOAST KEYPHRASE
     */

    if (fields.keyphrase) {

      var keyphrase =
        fields.keyphrase.trim();


      var words =
        keyphrase
          .split(/\s+/)
          .filter(Boolean);


      if (
        keyphrase !==
        keyphrase.toLowerCase()
      ) {

        issues.push(
          "Yoast Keyphrase is not lowercase."
        );
      }


      if (
        words.length < 2 ||
        words.length > 6
      ) {

        issues.push(
          "Yoast Keyphrase must contain 2-6 words."
        );
      }
    }


    return {

      success:
        issues.length === 0,

      issues:
        issues
    };
  }


  /*
   * =========================================================
   * SEMANTIC META TITLE QUALITY CHECK
   *
   * This deliberately does NOT use a blacklist of phrases.
   * It judges whether the title contains a concrete,
   * page-specific informational tension.
   * =========================================================
   */

  function validateW5TitleSemantically(fields) {

    try {

      var d =
        getActiveRowDataMap();


      var title =
        String(
          fields.title || ""
        ).trim();


      if (!title) {

        return {
          success: false,
          issue:
            "Meta Title is missing.",
          cost: 0
        };
      }


      var primaryTerm =
        String(
          d["Primary Search Term"] ||
          d["Primary Query Cluster Owned"] ||
          ""
        ).trim();


      var location =
        String(
          d["Locality"] ||
          d["Location"] ||
          ""
        ).trim();


      var problemAngle =
        String(
          d["Problem Angle"] ||
          ""
        ).trim();


      var authorityBrief =
        String(
          d["Authority Brief"] ||
          ""
        ).trim();


      var competitorNotes =
        String(
          d["Competitor SERP Notes"] ||
          ""
        ).trim();


      var semanticPrompt =
        "W5 META TITLE QUALITY CHECK\n\n" +

        "Evaluate the Meta Title linguistically and semantically.\n" +
        "Do NOT judge it by matching against a blacklist of phrases.\n\n" +

        "META TITLE:\n" +
        title +
        "\n\n" +

        "PAGE CONTEXT:\n" +

        "Primary Search Term: " +
        primaryTerm +
        "\n" +

        (
          location
            ? "Location: " +
              location +
              "\n"
            : ""
        ) +

        (
          problemAngle
            ? "Problem Angle: " +
              problemAngle +
              "\n"
            : ""
        ) +

        (
          authorityBrief
            ? "Authority Brief:\n" +
              authorityBrief.substring(
                0,
                1600
              ) +
              "\n"
            : ""
        ) +

        (
          competitorNotes
            ? "Competitor SERP Context:\n" +
              competitorNotes.substring(
                0,
                1200
              ) +
              "\n"
            : ""
        ) +

        "\nQUALITY TEST:\n" +

        "PASS only if the title gives the searcher a specific and concrete reason to click.\n\n" +

        "A strong title should express at least one meaningful form of informational tension supported by this page:\n" +

        "- a real decision the homeowner needs to make\n" +
        "- a recognisable problem\n" +
        "- a consequence that matters\n" +
        "- a limitation or boundary\n" +
        "- a useful contrast between choices or outcomes\n" +
        "- a specific uncertainty the article resolves\n" +
        "- a counterintuitive fact genuinely supported by the page\n\n" +

        "FAIL if the supposed hook is semantically vague.\n" +

        "A vague hook could be attached to many unrelated pages without materially changing its meaning.\n\n" +

        "Do not fail or pass a title merely because it contains words such as 'what', 'why', 'when', 'works', 'lasts' or 'matters'.\n" +

        "Judge what those words MEAN in this specific title.\n\n" +

        "Ask this decisive question:\n" +

        "After reading this title, does the searcher know WHAT specific uncertainty, problem, consequence, distinction or decision the article will resolve?\n\n" +

        "If the answer is no, FAIL.\n\n" +

        "Also FAIL if the title is essentially just a service/search phrase plus location with decorative wording added that contributes no concrete meaning.\n\n" +

        "Do not require Abbey Floor Care in the title.\n" +

        "Do not reward exaggerated clickbait.\n" +

        "Do not require emotional language.\n" +

        "Do not invent a better title.\n\n" +

        "Return JSON only:\n" +

        '{"pass":true,"reason":"brief linguistic reason"}';


      var semanticResult =
        bc_sendPromptViaOpenAI(
          semanticPrompt,
          900
        );


      if (
        !semanticResult ||
        !semanticResult.success
      ) {

        return {
          success: false,
          issue:
            "Meta Title semantic validation could not be completed.",
          cost:
            Number(
              semanticResult &&
              semanticResult.cost
                ? semanticResult.cost
                : 0
            )
        };
      }


      var raw =
        String(
          semanticResult.text || ""
        )
          .replace(
            /^```(?:json)?\s*/i,
            ""
          )
          .replace(
            /\s*```$/i,
            ""
          )
          .trim();


      var parsed;


      try {

        parsed =
          JSON.parse(
            raw
          );

      } catch (e) {

        return {
          success: false,
          issue:
            "Meta Title semantic validator returned invalid JSON.",
          cost:
            Number(
              semanticResult.cost || 0
            )
        };
      }


      if (
        parsed.pass !== true
      ) {

        return {
          success: false,
          issue:
            "Meta Title lacks a specific CTR-driving informational tension: " +
            String(
              parsed.reason ||
              "the title is too vague."
            ),
          cost:
            Number(
              semanticResult.cost || 0
            )
        };
      }


      return {
        success: true,
        cost:
          Number(
            semanticResult.cost || 0
          )
      };


    } catch (e) {

      return {
        success: false,
        issue:
          "Meta Title semantic validation error: " +
          e.toString(),
        cost: 0
      };
    }
  }


  /*
   * =========================================================
   * COMPLETE VALIDATION
   *
   * Run deterministic checks first.
   * Only spend money on semantic title checking if those pass.
   * =========================================================
   */

  function runCompleteValidation(fields) {

    var deterministic =
      validateW5Output(
        fields
      );


    if (
      !deterministic.success
    ) {

      return deterministic;
    }


    var semantic =
      validateW5TitleSemantically(
        fields
      );


    totalCost +=
      Number(
        semantic.cost || 0
      );


    if (
      !semantic.success
    ) {

      return {
        success: false,
        issues: [
          semantic.issue
        ]
      };
    }


    return {
      success: true,
      issues: []
    };
  }


  /*
   * =========================================================
   * BUILD W5 PROMPT
   * =========================================================
   */

  var prompt =
    getMetaPrompt();


  if (
    !prompt ||
    String(prompt).trim() === ""
  ) {

    return {
      success: false,
      message:
        "W5 could not build the H1/meta prompt.",
      cost: 0
    };
  }


  /*
   * =========================================================
   * FIRST GENERATION
   * =========================================================
   */

  var apiResult =
    bc_sendPromptViaOpenAI(
      prompt,
      5000
    );


  if (
    !apiResult ||
    !apiResult.success
  ) {

    return {
      success: false,
      message:
        apiResult &&
        apiResult.message
          ? apiResult.message
          : "W5 OpenAI generation failed.",
      cost: 0
    };
  }


  totalCost +=
    Number(
      apiResult.cost || 0
    );


  var output =
    String(
      apiResult.text || ""
    ).trim();


  var fields =
    parseW5Output(
      output
    );

  if (w5LockedKeyphrase) {
    fields.keyphrase = w5LockedKeyphrase;
  }


  var validation =
    runCompleteValidation(
      fields
    );


  /*
   * =========================================================
   * ONE FULL CORRECTION ATTEMPT
   * =========================================================
   */

  if (
    !validation.success
  ) {

    var correctionPrompt =
      prompt +
      "\n\n" +

      "IMPORTANT — YOUR PREVIOUS OUTPUT FAILED AUTOMATED VALIDATION.\n\n" +

      "PREVIOUS OUTPUT:\n" +
      output +
      "\n\n" +

      "FAILURES:\n- " +
      validation.issues.join(
        "\n- "
      ) +
      "\n\n" +

      "Correct ONLY the four requested fields so every failure above is resolved.\n\n" +

      "For the Meta Title, do not merely add a generic curiosity phrase.\n" +

      "The title must communicate a concrete problem, decision, consequence, limitation, contrast or uncertainty that THIS page actually resolves.\n" +

      "A phrase that could be attached unchanged to many unrelated service pages is not a meaningful CTR hook.\n\n" +

      "Do not include analysis, reasoning, notes, markdown or explanation.\n\n" +

      "Return ONLY these four lines:\n" +

      "New H1: [value]\n" +
      "New Meta Title: [value]\n" +
      "New Meta Description: [value]\n" +
      "Yoast Keyphrase: [value]";


    var retry =
      bc_sendPromptViaOpenAI(
        correctionPrompt,
        5000
      );


    if (
      !retry ||
      !retry.success
    ) {

      return {
        success: false,
        message:
          "W5 first output failed validation and the correction call failed.",
        cost:
          totalCost
      };
    }


    totalCost +=
      Number(
        retry.cost || 0
      );


    output =
      String(
        retry.text || ""
      ).trim();


    fields =
      parseW5Output(
        output
      );

    if (w5LockedKeyphrase) {
      fields.keyphrase = w5LockedKeyphrase;
    }


    validation =
      runCompleteValidation(
        fields
      );
  }


  /*
   * =========================================================
   * FINAL TARGETED MICRO-REPAIR
   *
   * Do not regenerate the enormous W5 prompt for a small
   * remaining defect such as 161 characters or a vague title.
   * =========================================================
   */

  if (
    !validation.success
  ) {

    var rowData =
      getActiveRowDataMap();


    var microPrimaryTerm =
      String(
        rowData["Primary Search Term"] ||
        rowData["Primary Query Cluster Owned"] ||
        ""
      ).trim();


    var microLocation =
      String(
        rowData["Locality"] ||
        rowData["Location"] ||
        ""
      ).trim();


    var microMaterial =
      String(
        rowData["Stone Type"] ||
        ""
      ).trim();


    var microProblemAngle =
      String(
        rowData["Problem Angle"] ||
        ""
      ).trim();


    var microAuthority =
      String(
        rowData["Authority Brief"] ||
        ""
      ).trim();


    var microPrompt =
      "W5 FINAL TARGETED REPAIR\n\n" +

      "Repair ONLY the field or fields that fail the validation below.\n" +

      "Keep fields that already pass unchanged unless a small consistency change is unavoidable.\n\n" +

      "PAGE CONTEXT:\n" +

      "Primary Search Term: " +
      microPrimaryTerm +
      "\n" +

      "Material: " +
      microMaterial +
      "\n" +

      (
        microLocation
          ? "Location: " +
            microLocation +
            "\n"
          : ""
      ) +

      (
        microProblemAngle
          ? "Problem Angle: " +
            microProblemAngle +
            "\n"
          : ""
      ) +

      (
        microAuthority
          ? "Authority Brief:\n" +
            microAuthority.substring(
              0,
              1400
            ) +
            "\n"
          : ""
      ) +

      "\nCURRENT FIELDS:\n" +

      "New H1: " +
      fields.h1 +
      "\n" +

      "New Meta Title: " +
      fields.title +
      "\n" +

      "New Meta Description: " +
      fields.description +
      "\n" +

      "Yoast Keyphrase: " +
      fields.keyphrase +
      "\n\n" +

      "FAILED VALIDATION:\n- " +

      validation.issues.join(
        "\n- "
      ) +

      "\n\nREPAIR RULES:\n" +

      "- H1 must be 40-60 characters.\n" +

      "- Meta Title must be no more than 60 characters.\n" +

      "- Meta Title must preserve the search intent and required location where applicable.\n" +

      "- Abbey Floor Care is optional if branding prevents a useful title within 60 characters.\n" +

      "- Meta Title must contain a concrete informational tension supported by this page.\n" +

      "- The reader should be able to tell what specific problem, decision, consequence, contrast, limitation or uncertainty the page resolves.\n" +

      "- Do not substitute a vague curiosity phrase merely to make the title sound interesting.\n" +

      "- Do not invent facts, drama, urgency, risk or unsupported claims.\n" +

      "- Meta Description must be 140-160 characters INCLUDING spaces and punctuation.\n" +
      "- Before returning the Meta Description, count its characters and rewrite it until it is between 140 and 160 characters inclusive.\n" +
      "- Do not return a Meta Description outside that range.\n" +

      "- Meta Description must contain exactly two sentences.\n" +

      "- Yoast Keyphrase must be lowercase and 2-6 words.\n" +

      "- Do not include analysis, commentary, markdown or character counts.\n\n" +

      "Return ONLY these four lines:\n" +

      "New H1: [value]\n" +

      "New Meta Title: [value]\n" +

      "New Meta Description: [value]\n" +

      "Yoast Keyphrase: [value]";


    var microResult =
      bc_sendPromptViaOpenAI(
        microPrompt,
        1600
      );


    if (
      microResult &&
      microResult.success
    ) {

      totalCost +=
        Number(
          microResult.cost || 0
        );


      output =
        String(
          microResult.text || ""
        ).trim();


      fields =
        parseW5Output(
          output
        );

      if (w5LockedKeyphrase) {
        fields.keyphrase = w5LockedKeyphrase;
      }


      validation =
        runCompleteValidation(
          fields
        );
    }
  }


  /*
   * =========================================================
   * DO NOT SAVE FAILED OUTPUT
   * =========================================================
   */

  if (
    !validation.success
  ) {

    return {
      success: false,

      message:
        "W5 output failed validation after correction: " +
        validation.issues.join(
          " | "
        ) +
        "\n\nRAW W5 OUTPUT:\n" +
        String(
          output || ""
        ).substring(
          0,
          1400
        ),

      cost:
        totalCost
    };
  }


  /*
   * =========================================================
   * SAVE ALL FOUR FIELDS
   * =========================================================
   */

  // Rebuild the saved output from the validated fields.
  // For a new/no-GSC article this guarantees the original
  // Yoast keyphrase stored in Primary Search Term is preserved.

  output =
    "New H1: " + fields.h1 + "\n" +
    "New Meta Title: " + fields.title + "\n" +
    "New Meta Description: " + fields.description + "\n" +
    "Yoast Keyphrase: " + fields.keyphrase;


  var pushResult =
    pushGovernanceFieldsToActiveRow(
      output
    );


  if (
    !pushResult ||
    !pushResult.success
  ) {

    return {
      success: false,

      message:
        pushResult &&
        pushResult.message
          ? pushResult.message
          : "W5 could not save the generated fields.",

      cost:
        totalCost
    };
  }


  /*
   * =========================================================
   * RECORD COST + TIME
   * =========================================================
   */

  if (
    typeof bc_addToApiCostAndTime ===
    "function"
  ) {

    bc_addToApiCostAndTime(
      totalCost,
      (
        Date.now() -
        startTime
      ) / 1000
    );
  }


  /*
   * =========================================================
   * SUCCESS
   * =========================================================
   */

  return {

    success: true,

    message:
      "W5 complete — H1, CTR-focused Meta Title, Meta Description and Yoast Keyphrase saved.",

    h1:
      fields.h1,

    metaTitle:
      fields.title,

    metaDescription:
      fields.description,

    yoastKeyphrase:
      fields.keyphrase,

    cost:
      totalCost
  };
}

function bc_runW8EAutomated() {

  var startTime = Date.now();
  var totalCost = 0;

  try {

    /*
     * =========================================================
     * BUILD SINGLE W8E PROMPT
     * =========================================================
     */

    var promptResult =
      buildW8EAnalysisPrompt();

    if (
      !promptResult ||
      !promptResult.success
    ) {
      return {
        success: false,
        message:
          promptResult &&
          promptResult.message
            ? promptResult.message
            : "W8E could not build the analysis prompt.",
        cost: 0
      };
    }


    var prompt =
      String(
        promptResult.prompt || ""
      ).trim();

    if (!prompt) {
      return {
        success: false,
        message:
          "W8E analysis prompt is empty.",
        cost: 0
      };
    }


    /*
     * =========================================================
     * ONE OPENAI CALL ONLY
     * =========================================================
     */

    var apiResult =
      bc_sendPromptViaOpenAI(
        prompt,
        7000
      );

    if (
      !apiResult ||
      !apiResult.success
    ) {
      return {
        success: false,
        message:
          apiResult &&
          apiResult.message
            ? apiResult.message
            : "W8E OpenAI analysis failed.",
        cost: 0
      };
    }


    totalCost +=
      Number(
        apiResult.cost || 0
      );


    var output =
      String(
        apiResult.text || ""
      ).trim();

    if (!output) {
      return {
        success: false,
        message:
          "W8E returned an empty response.",
        cost:
          totalCost
      };
    }


    /*
     * =========================================================
     * VALIDATE ANALYSIS + ATOMIC FIXES
     * =========================================================
     */

    var saveResult =
      saveW8EAnalysisResult(
        output
      );


    /*
     * IMPORTANT:
     * No second AI correction call.
     *
     * If the JSON is invalid, stop.
     * This keeps W8E to one paid API call.
     * =========================================================
     */

    if (
      !saveResult ||
      !saveResult.success
    ) {
      return {
        success: false,

        message:
          "W8E stopped because the one-call result failed validation: " +
          (
            saveResult &&
            saveResult.message
              ? saveResult.message
              : "Unknown validation failure."
          ),

        cost:
          totalCost
      };
    }


    /*
     * =========================================================
     * PARSE THE ALREADY-VALIDATED JSON
     * =========================================================
     */

    var cleaned =
      output
        .replace(/^```json\s*/i, "")
        .replace(/^```\s*/i, "")
        .replace(/\s*```$/i, "")
        .trim();


    var checks =
  Array.isArray(saveResult.checks)
    ? saveResult.checks
    : [];


    var fixes =
      Array.isArray(saveResult.fixes)
        ? saveResult.fixes
        : [];


    var autoItems =
      checks.filter(
        function(item) {
          return (
            item &&
            item.action_type ===
            "AUTO-FIXABLE"
          );
        }
      );


    var authorItems =
      checks.filter(
        function(item) {
          return (
            item &&
            item.action_type ===
            "NEEDS AUTHOR INPUT"
          );
        }
      );


    /*
     * =========================================================
     * APPLY ALREADY-VALIDATED ATOMIC FIXES
     * =========================================================
     */

    var appliedFixCount = 0;
    var skippedFixCount = 0;
    var fixMessages = [];


    if (
      fixes.length > 0
    ) {

      var ss =
        SpreadsheetApp
          .getActiveSpreadsheet();


      var sh =
        ss.getSheetByName(
          'posts'
        );


      var row =
        sh
          .getActiveCell()
          .getRow();


      var fieldColumns = {
        html: 'CT',
        h1: 'CK',
        metaTitle: 'CL',
        metaDescription: 'CM',
        schema: 'CN'
      };


      /*
       * Work in memory first so multiple fixes in the
       * same field can be applied safely before one write.
       */

      var working = {
        html:
          String(
            sh.getRange(
              'CT' + row
            ).getValue() || ''
          ),

        h1:
          String(
            sh.getRange(
              'CK' + row
            ).getValue() || ''
          ),

        metaTitle:
          String(
            sh.getRange(
              'CL' + row
            ).getValue() || ''
          ),

        metaDescription:
          String(
            sh.getRange(
              'CM' + row
            ).getValue() || ''
          ),

        schema:
          String(
            sh.getRange(
              'CN' + row
            ).getValue() || ''
          )
      };


      var touched = {};


      for (
        var i = 0;
        i < fixes.length;
        i++
      ) {

        var fix =
          fixes[i] || {};


        var label =
          String(
            fix.fixLabel || ""
          ).trim();


        var field =
          String(
            fix.field || ""
          ).trim();


        var oldText =
          String(
            fix.old || ""
          );


        var newText =
          String(
            fix.new || ""
          );


        if (
          !fieldColumns[field]
        ) {

          skippedFixCount++;

          fixMessages.push(
            "Skipped " +
            label +
            " — invalid field."
          );

          continue;
        }


        var current =
          working[field];


        var firstIndex =
          current.indexOf(
            oldText
          );


        if (
          firstIndex === -1
        ) {

          skippedFixCount++;

          fixMessages.push(
            "Skipped " +
            label +
            " — OLD text no longer found."
          );

          continue;
        }


        var secondIndex =
          current.indexOf(
            oldText,
            firstIndex +
            oldText.length
          );


        if (
          secondIndex !== -1
        ) {

          skippedFixCount++;

          fixMessages.push(
            "Skipped " +
            label +
            " — OLD text is not unique."
          );

          continue;
        }


        working[field] =
          current.substring(
            0,
            firstIndex
          ) +
          newText +
          current.substring(
            firstIndex +
            oldText.length
          );


        touched[field] = true;

        appliedFixCount++;

        fixMessages.push(
          "Applied " +
          label
        );
      }


      /*
       * =========================================================
       * WRITE ONLY TOUCHED FIELDS
       * =========================================================
       */

      Object.keys(
        touched
      ).forEach(
        function(field) {

          sh.getRange(
            fieldColumns[field] +
            row
          ).setValue(
            working[field]
          );
        }
      );
    }


    /*
     * =========================================================
     * FINAL STATUS
     *
     * This is based on the audit result from the one call.
     * No second audit is performed.
     * =========================================================
     */

    var finalStatus;


    if (
      authorItems.length > 0
    ) {

      finalStatus =
        "NEEDS_AUTHOR_INPUT";

    } else if (
      skippedFixCount > 0
    ) {

      finalStatus =
        "PASS_WITH_FIXES";

    } else {

      finalStatus =
        "PASS";
    }


    /*
     * =========================================================
     * MARK W8E COMPLETE
     * =========================================================
     */

    var w8Ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    var w8Sheet =
      w8Ss.getSheetByName(
        "posts"
      );

    var w8Row =
      w8Sheet
        .getActiveCell()
        .getRow();

    // W8E complete — mark Column D pink with white text.
    w8Sheet
      .getRange(w8Row, 4)
      .setBackground("#D81B60")
      .setFontColor("#FFFFFF");


    /*
     * =========================================================
     * RECORD COST + TIME
     * =========================================================
     */

    if (
      typeof bc_addToApiCostAndTime ===
      "function"
    ) {

      bc_addToApiCostAndTime(
        totalCost,
        (
          Date.now() -
          startTime
        ) / 1000
      );
    }


    /*
     * =========================================================
     * RETURN
     * =========================================================
     */

    return {

      success: true,

      message:
        "W8E complete — one-call audit finished and validated safe fixes were applied.",

      overallStatus:
        finalStatus,

      autoFixableCount:
        autoItems.length,

      authorInputCount:
        authorItems.length,

      autoFixableItems:
        autoItems,

      authorInputItems:
        authorItems,

      appliedFixCount:
        appliedFixCount,

      skippedFixCount:
        skippedFixCount,

      fixMessage:
        fixMessages.join(
          "\n"
        ),

      cost:
        totalCost
    };


  } catch (e) {

    return {

      success: false,

      message:
        "W8E automation error: " +
        e.toString(),

      cost:
        totalCost
    };
  }
}