/**
 * Abbey SEO: Master Workflow (Validated SEO + Graph Schema + Validator Link)
 * Fixes: hasPart, about, supply types, and adds one-click testing.
 */

function GENERATE_DYNAMIC_SEO_PROMPT() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRow = sheet.getActiveCell().getRow();
  
  // 1. GATHER DATA
  const articleType  = sheet.getRange(activeRow, 8).getValue();   // H
  const pageUrl      = sheet.getRange(activeRow, 3).getValue();   // C
  const pageHtml     = sheet.getRange(activeRow, 98).getValue();  // CT
  const stoneType    = sheet.getRange(activeRow, 7).getValue();   // G
  const primaryTerm  = sheet.getRange(activeRow, 93).getValue();  // CO
  const featuredImg  = sheet.getRange(activeRow, 85).getValue();  // CG
  const logoUrl      = sheet.getRange(activeRow, 86).getValue();  // CH

  if (!primaryTerm || !pageHtml) {
    SpreadsheetApp.getUi().alert("Error: Missing Keyword (CO) or HTML (CT).");
    return;
  }

  // 2. HUB SCANNER
  const hubData = getRelatedHubData(stoneType, activeRow);

  // 3. CONSTRUCT THE VALIDATED PROMPT
  const promptText = `Act as a Stone Restoration SEO Specialist and Schema Architect.
Generate H1, Meta Title, Meta Description, and a VALIDATED JSON-LD Graph.

PAGE DATA:
- URL: ${pageUrl}
- Article Type: ${articleType}
- Stone: ${stoneType}
- Primary Keyword: ${primaryTerm}
- Featured Image: ${featuredImg}
- Publisher Logo: ${logoUrl}

STRICT SCHEMA VALIDATION RULES (Stop Google Errors):
1. USE FLAT GRAPH: Keep 'LocalBusiness', 'Service', and 'HowTo' as separate objects in the "@graph" array.
2. SERVICE ENTITY: 
   - DO NOT use 'about' or 'hasPart' inside 'Service'.
   - Use 'provider' to link to the '@id' of the 'LocalBusiness'.
   - Use 'mainEntityOfPage' to link 'Service' to the current URL.
3. HOWTO ENTITY:
   - Use '@type': 'HowToSupply' for all supplies (NEVER use 'Product' inside 'supply').
   - Link 'HowTo' to the page via 'mainEntityOfPage'.
4. RELATIONSHIPS: Connect everything via '@id' anchors (#service, #localbusiness, #process).

RELATED HUB ENTITIES (Select 5 most relevant):
${hubData}

PAGE CONTENT (HTML):
"""${pageHtml.substring(0, 3200)}"""

Return exactly:
H1: [Result]
Title: [Result]
Description: [Result]
Schema:
[OUTPUT EVERYTHING INSIDE A SINGLE CODE WINDOW]
<script type="application/ld+json">
[VALIDATED JSON-LD GRAPH]
</script>`;

  // 4. THE MODAL INTERFACE
  const htmlOutput = HtmlService.createHtmlOutput(`
    <html>
    <body style="font-family: sans-serif; padding: 15px; background: #f4f7f6;">
      <div style="margin-bottom:10px; padding:12px; background:white; border-radius:8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
        <label style="font-weight:bold; font-size:13px;">1. Copy Prompt</label>
        <textarea id="pBox" style="width:100%; height:80px; font-size:10px; border:1px solid #ddd;">${promptText}</textarea>
        <button onclick="copyPrompt()" style="background:#1a73e8; color:white; width:100%; padding:10px; border:none; border-radius:4px; margin-top:5px; cursor:pointer;">📋 Copy Prompt</button>
      </div>
      <div style="padding:12px; background:white; border-radius:8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
        <label style="font-weight:bold; font-size:13px;">2. Paste LLM Code Block</label>
        <textarea id="pastedText" placeholder="Paste the entire code block here..." style="width:100%; height:280px; font-size:10px; border:1px solid #ddd;"></textarea>
        <button id="pushBtn" onclick="pushToSheet()" style="background:#34a853; color:white; width:100%; padding:12px; border:none; border-radius:4px; margin-top:10px; font-weight:bold; cursor:pointer;">💾 Push to Sheet</button>
        <button id="testBtn" onclick="openTester()" style="display:none; background:#fbbc04; color:black; width:100%; padding:12px; border:none; border-radius:4px; margin-top:10px; font-weight:bold; cursor:pointer;">🔍 Validate in Google Tester</button>
      </div>
      <script>
        function copyPrompt() { var t = document.getElementById("pBox"); t.select(); document.execCommand('copy'); }
        function pushToSheet() {
          var val = document.getElementById("pastedText").value;
          if(!val) { alert("Paste results first."); return; }
          google.script.run.withSuccessHandler(() => {
            document.getElementById("pushBtn").innerText = "✅ Published!";
            document.getElementById("testBtn").style.display = "block";
          }).processSEOInput(val);
        }
        function openTester() {
          window.open('https://search.google.com/test/rich-results?url=' + encodeURIComponent("${pageUrl}"), '_blank');
        }
      </script>
    </body>
    </html>
  `).setWidth(550).setHeight(680);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Abbey SEO & Validated Graph");
}

function getRelatedHubData(currentStone, currentRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  let hubLinks = [];
  for (let i = 1; i < data.length; i++) {
    if (i + 1 === currentRow) continue; 
    if (data[i][6] === currentStone) {
      if (data[i][2] && data[i][79]) hubLinks.push({ url: data[i][2], title: data[i][79] });
    }
  }
  return JSON.stringify(hubLinks.slice(0, 7), null, 2);
}

function processSEOInput(text) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  let h1 = text.match(/H1:\s*(.*)/i)?.[1] || "";
  let title = text.match(/Title:\s*(.*)/i)?.[1] || "";
  let desc = text.match(/Description:\s*(.*)/i)?.[1] || "";
  let schemaRaw = "";
  const tagMatch = text.match(/<script[^>]*>([\s\S]*?)<\/script>/i);
  if (tagMatch) { schemaRaw = tagMatch[1].trim(); }
  else { const jsonMatch = text.match(/\{[\s\S]*\}/); schemaRaw = jsonMatch ? jsonMatch[0].trim() : ""; }
  const fixCasing = (str) => str ? str.trim().charAt(0).toUpperCase() + str.trim().slice(1) : "";
  let finalSchema = schemaRaw ? `<script type="application/ld+json">\n${schemaRaw}\n</script>` : "";
  sheet.getRange(row, 89).setValue(h1);              
  sheet.getRange(row, 90).setValue(fixCasing(title)); 
  sheet.getRange(row, 91).setValue(fixCasing(desc));  
  sheet.getRange(row, 92).setValue(finalSchema);      
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Row Updated Successfully!");
}