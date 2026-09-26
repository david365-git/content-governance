/* ============================================================
   ce_CompetitorSerpNotes
   Competitor SERP Differentiation Notes
   Reads a pasted Gemini grounded-search output, condenses it
   into a short structured note, and stores it in the posts
   sheet for use by getMetaPrompt().

   Column: "Competitor SERP Notes" (add this header to the
   posts sheet if not already present)
============================================================ */

function buildCompetitorSerpNotes(rawGeminiOutput) {
  try {
    if (!rawGeminiOutput || rawGeminiOutput.trim() === "") {
      return { success: false, message: "Paste the Gemini output first." };
    }

    var note = "";
    note += "Generated: " + new Date().toLocaleDateString("en-GB") + "\n";
    note += "Source: Gemini grounded search (manual)\n\n";
    note += "RAW COMPETITOR SNAPSHOT:\n";
    note += rawGeminiOutput.trim();

    return { success: true, note: note };
  } catch (e) {
    return { success: false, message: "Build error: " + e.toString() };
  }
}

function pushCompetitorSerpNotesToSheet(note) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                         .getValues()[0]
                         .map(function(h) { return String(h).trim(); });

    var colIdx = headers.indexOf("Competitor SERP Notes");
    if (colIdx === -1) {
      return { success: false, message: "Competitor SERP Notes column not found. Add this header to the posts sheet first." };
    }

    var cell = sheet.getRange(row, colIdx + 1);
    cell.setNumberFormat("@");
    cell.setValue(note.trim());

    return { success: true, message: "Competitor SERP Notes saved for row " + row };
  } catch (e) {
    return { success: false, message: "Push error: " + e.toString() };
  }
}

function getCompetitorSerpNotes() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return "";

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                         .getValues()[0]
                         .map(function(h) { return String(h).trim(); });

    var colIdx = headers.indexOf("Competitor SERP Notes");
    if (colIdx === -1) return "";
    return String(sheet.getRange(row, colIdx + 1).getValue() || "").trim();
  } catch (e) {
    return "";
  }
}

function buildCompetitorSerpGeminiPrompt() {
  try {
    const d = getActiveRowDataMap();
    const primaryTerm  = String(d["Primary Search Term"]         || "").trim();
    const queryCluster = String(d["Primary Query Cluster Owned"] || "").trim();
    const term = primaryTerm || queryCluster;

    if (!term) {
      return { success: false, message: "No Primary Search Term or Query Cluster found for this row." };
    }

    var prompt =
      "Search Google UK for the term: " + term + "\n\n" +
      "Return the top 8 organic results (not ads, not shopping results) in this exact format:\n\n" +
      "1. [Domain] | [Exact page title as it appears in the SERP] | [Exact meta description as it appears in the SERP]\n" +
      "2. ...\n\n" +
      "Show your sources/citations for this.";

    return { success: true, prompt: prompt, term: term };
  } catch (e) {
    return { success: false, message: "Build error: " + e.toString() };
 
 
 function bc_runCompetitorSerpNotesAutomated() {
  try {

    var startTime = Date.now();

    /*
     * ---------------------------------------------------------
     * BUILD SEARCH PROMPT
     * ---------------------------------------------------------
     */

    var promptData =
      buildCompetitorSerpGeminiPrompt();

    if (
      !promptData ||
      !promptData.success
    ) {
      return {
        success: false,
        message:
          promptData && promptData.message
            ? promptData.message
            : "Could not build W4.5 SERP prompt.",
        cost: 0
      };
    }


    /*
     * ---------------------------------------------------------
     * RUN GROUNDED GOOGLE SEARCH
     * ---------------------------------------------------------
     */

    var apiResult =
      bc_sendPromptViaGemini(
        promptData.prompt
      );

    if (
      !apiResult ||
      !apiResult.success
    ) {
      return {
        success: false,
        message:
          apiResult && apiResult.message
            ? apiResult.message
            : "Gemini grounded search failed.",
        cost: 0
      };
    }


    var rawOutput =
      String(
        apiResult.text || ""
      ).trim();

    if (!rawOutput) {
      return {
        success: false,
        message:
          "Gemini grounded search returned no usable SERP output.",
        cost: Number(apiResult.cost || 0)
      };
    }


    /*
     * ---------------------------------------------------------
     * BASIC SERP RESULT VALIDATION
     * ---------------------------------------------------------
     */

    var resultLines =
      rawOutput.match(
        /^\s*[1-8][.)]\s+.+$/gm
      ) || [];

    if (resultLines.length < 5) {

      return {
        success: false,
        message:
          "Grounded search returned only " +
          resultLines.length +
          " recognisable organic result(s). W4.5 was not saved.",
        cost: Number(apiResult.cost || 0)
      };
    }


    /*
     * ---------------------------------------------------------
     * BUILD STORED NOTE
     * ---------------------------------------------------------
     */

    var generatedDate =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm"
      );

    var note =
      "Generated: " +
      generatedDate +
      "\n" +
      "Source: Gemini grounded Google Search — automated\n" +
      "Search term: " +
      promptData.term +
      "\n\n" +
      "RAW COMPETITOR SNAPSHOT:\n" +
      rawOutput;


    /*
     * ---------------------------------------------------------
     * SAVE TO COMPETITOR SERP NOTES
     * ---------------------------------------------------------
     */

    var pushResult =
      pushCompetitorSerpNotesToSheet(
        note
      );

    if (
      !pushResult ||
      !pushResult.success
    ) {

      return {
        success: false,
        message:
          pushResult && pushResult.message
            ? pushResult.message
            : "Could not save Competitor SERP Notes.",
        cost: Number(apiResult.cost || 0)
      };
    }


    /*
     * ---------------------------------------------------------
     * RECORD COST / TIME
     * ---------------------------------------------------------
     */

    if (
      typeof bc_addToApiCostAndTime ===
      "function"
    ) {

      bc_addToApiCostAndTime(
        Number(apiResult.cost || 0),
        (
          Date.now() -
          startTime
        ) / 1000
      );
    }


    return {
      success: true,
      message:
        "W4.5 complete — grounded competitor SERP notes saved for \"" +
        promptData.term +
        "\".",
      term: promptData.term,
      resultsFound: resultLines.length,
      cost: Number(apiResult.cost || 0)
    };


  } catch (e) {

    return {
      success: false,
      message:
        "W4.5 automation error: " +
        e.toString(),
      cost: 0
    };
  }
} }
}

function bc_runCompetitorSerpNotesAutomated() {
  try {

    var startTime = Date.now();

    /*
     * ---------------------------------------------------------
     * BUILD SEARCH PROMPT
     * ---------------------------------------------------------
     */

    var promptData =
      buildCompetitorSerpGeminiPrompt();

    if (
      !promptData ||
      !promptData.success
    ) {
      return {
        success: false,
        message:
          promptData && promptData.message
            ? promptData.message
            : "Could not build W4.5 SERP prompt.",
        cost: 0
      };
    }


    /*
     * ---------------------------------------------------------
     * RUN GROUNDED GOOGLE SEARCH
     * ---------------------------------------------------------
     */

    var apiResult =
      bc_sendPromptViaGemini(
        promptData.prompt
      );

    if (
      !apiResult ||
      !apiResult.success
    ) {
      return {
        success: false,
        message:
          apiResult && apiResult.message
            ? apiResult.message
            : "Gemini grounded search failed.",
        cost: 0
      };
    }


    var rawOutput =
      String(
        apiResult.text || ""
      ).trim();

    if (!rawOutput) {
      return {
        success: false,
        message:
          "Gemini grounded search returned no usable SERP output.",
        cost: Number(apiResult.cost || 0)
      };
    }


    /*
     * ---------------------------------------------------------
     * BASIC SERP RESULT VALIDATION
     * ---------------------------------------------------------
     */

    var resultLines =
      rawOutput.match(
        /^\s*[1-8][.)]\s+.+$/gm
      ) || [];

    if (resultLines.length < 5) {

      return {
        success: false,
        message:
          "Grounded search returned only " +
          resultLines.length +
          " recognisable organic result(s). W4.5 was not saved.",
        cost: Number(apiResult.cost || 0)
      };
    }


    /*
     * ---------------------------------------------------------
     * BUILD STORED NOTE
     * ---------------------------------------------------------
     */

    var generatedDate =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm"
      );

    var note =
      "Generated: " +
      generatedDate +
      "\n" +
      "Source: Gemini grounded Google Search — automated\n" +
      "Search term: " +
      promptData.term +
      "\n\n" +
      "RAW COMPETITOR SNAPSHOT:\n" +
      rawOutput;


    /*
     * ---------------------------------------------------------
     * SAVE TO COMPETITOR SERP NOTES
     * ---------------------------------------------------------
     */

    var pushResult =
      pushCompetitorSerpNotesToSheet(
        note
      );

    if (
      !pushResult ||
      !pushResult.success
    ) {

      return {
        success: false,
        message:
          pushResult && pushResult.message
            ? pushResult.message
            : "Could not save Competitor SERP Notes.",
        cost: Number(apiResult.cost || 0)
      };
    }


    /*
     * ---------------------------------------------------------
     * RECORD COST / TIME
     * ---------------------------------------------------------
     */

    if (
      typeof bc_addToApiCostAndTime ===
      "function"
    ) {

      bc_addToApiCostAndTime(
        Number(apiResult.cost || 0),
        (
          Date.now() -
          startTime
        ) / 1000
      );
    }


    return {
      success: true,
      message:
        "W4.5 complete — grounded competitor SERP notes saved for \"" +
        promptData.term +
        "\".",
      term: promptData.term,
      resultsFound: resultLines.length,
      cost: Number(apiResult.cost || 0)
    };


  } catch (e) {

    return {
      success: false,
      message:
        "W4.5 automation error: " +
        e.toString(),
      cost: 0
    };
  }
}