/**
 * ================================================================================
 * ce_stage0A.gs - STAGEOA SerpValidation Prompt
 * ================================================================================
 * 
 * Build SerpValidation Prompt
 * 
 * Part of Abbey Floor Care Content Pipeline v77+
 * Reorganised: April 2026
 * ================================================================================
 */

function buildStage0APrompt() {
  const d = getActiveRowDataMap();

  const primaryTerm = d["Primary Search Term"] || "";
  const intent      = d["Topical Intent"] || "";
  const material    = d["Stone Type"] || "";
  const articleType = d["Article Type"] || "";

  if (!primaryTerm) return "ERROR: No Primary Search Term found.";

  return `
STAGE 0A — SERP VALIDATION
ROLE: Senior SEO Analyst

TASK:
Validate and extract the REAL entities Google associates with this query.

PRIMARY SEARCH TERM:
${primaryTerm}

CONTEXT:
Material: ${material}
Intent: ${intent}
Article Type: ${articleType}

INSTRUCTIONS:
1. Analyse the top 10 Google results for this query.
2. Identify recurring entities across:
   - headings
   - topics
   - terminology
3. Extract ONLY entities that appear across multiple results.

OUTPUT FORMAT (STRICT):
Return a comma-separated list only.

No explanation.
No sentences.
No markdown.

EXAMPLE OUTPUT:
etching, limescale, grout haze, sealer residue, surface dullness
`.trim();
}

function pushStage0AToSheet(raw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("posts");
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  const colIndex = headers.indexOf("SERP Validation");
  if (colIndex === -1) {
    return { success:false, message:"Column 'SERP Validation' not found." };
  }

  const row = sheet.getActiveRange().getRow();
  if (row < 2) {
    return { success:false, message:"No active row selected." };
  }

  // Clean input (simple normalisation)
  const cleaned = String(raw || "")
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  sheet.getRange(row, colIndex + 1).setValue(cleaned);

  return { success:true, message:"SERP Validation entities saved to column DP." };
}

function runStage0A() {
  setGenerating('status0a','gen0abtn','0A');
  google.script.run
    .withSuccessHandler(function(res) {
      document.getElementById('part0a').value = res;
      setDone('status0a','gen0abtn','Generate','Stage 0A ready — copy to ChatGPT','conf0abtn');
    })
    .withFailureHandler(function(err) {
      setFail('status0a','gen0abtn','Generate',err.message);
    })
    .buildStage0APrompt();
}

function pushStage0A() {
  var raw = document.getElementById('part0a_output').value.trim();
  if (!raw) {
    setStatus('status0a','fail','❌ Paste output first');
    return;
  }

  google.script.run
    .withSuccessHandler(function(res) {
      setStatus('status0a', res.success ? 'done' : 'fail',
        res.success ? '✔ ' + res.message : '❌ ' + res.message);
    })
    .withFailureHandler(function(err) {
      setStatus('status0a','fail','❌ ' + err.message);
    })
    .pushStage0AToSheet(raw);
}
function testBuildStage0APromptToDQ() {
  const prompt = buildStage0APrompt();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("posts");
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error("Select a data row in the posts sheet first.");
  }

  // DQ = column 121
  sheet.getRange(row, 121).setValue(prompt);
}