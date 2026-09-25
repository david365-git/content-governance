/* =========================================================
    🛡️ STRATEGIC BENCHMARK TOOL (REFRESHED V3.2)
========================================================= */

function openStoneSelector() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const refSheet = ss.getSheetByName("Material Governance Reference Sheet");

  if (!refSheet) throw new Error("Reference Sheet not found.");

  const materials = refSheet.getRange(2, 1, refSheet.getLastRow() - 1, 1).getValues().flat().filter(String);

  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', sans-serif; padding: 15px; background: #f4f7f6; color: #333; }
          .search-box { width: 100%; padding: 12px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-size: 14px; }
          .list-container { height: 220px; overflow-y: auto; background: white; border: 1px solid #ddd; padding: 10px; border-radius: 4px; }
          .item { margin-bottom: 8px; display: flex; align-items: center; cursor: pointer; padding: 5px; }
          .item:hover { background: #eef2f7; }
          input[type="radio"] { margin-right: 12px; cursor: pointer; }
          label { cursor: pointer; flex-grow: 1; font-size: 14px; }
          button { width: 100%; padding: 12px; background: #3498db; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 15px; }
          button:hover { background: #2980b9; }
        </style>
      </head>
      <body>
        <input type="text" id="search" class="search-box" placeholder="Type to search (e.g. Victorian)..." onkeyup="filterList()">
        <form id="stoneForm">
          <div class="list-container" id="list">
            ${materials.map((m, i) => `
              <div class="item">
                <input type="radio" id="m${i}" name="stone" value="${m}" ${i === 0 ? 'checked' : ''}>
                <label for="m${i}">${m}</label>
              </div>
            `).join('')}
          </div>
          <button type="button" onclick="submitForm()">📊 Update Benchmarks</button>
        </form>
        <script>
          function filterList() {
            const val = document.getElementById('search').value.toLowerCase();
            const items = document.querySelectorAll('.item');
            items.forEach(item => {
              const text = item.innerText.toLowerCase();
              item.style.display = text.includes(val) ? 'flex' : 'none';
            });
          }
          function submitForm() {
            const stone = document.querySelector('input[name="stone"]:checked').value;
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .fillBenchmarkData(stone);
          }
        </script>
      </body>
    </html>
  `).setWidth(400).setHeight(460);

  SpreadsheetApp.getUi().showModalDialog(html, "🛡️ Stone Selector");
}

function fillBenchmarkData(selectedStone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getActiveSheet();
  const gscSheet = ss.getSheetByName("last-28-day-gsc");

  if (!gscSheet) {
    SpreadsheetApp.getUi().alert("❌ ERROR: Sheet 'last-28-day-gsc' not found.");
    return;
  }

  const COL = { URL: 3, STONE: 6, BENCH_START: 53 };

  // 1. Build GSC Map with strict cleaning (removes trailing slashes and hashes)
  const gscValues = gscSheet.getRange(2, 1, gscSheet.getLastRow() - 1, 5).getValues();
  const gscMap = new Map();
  gscValues.forEach(r => { 
    if(r[0]) {
      const cleanUrl = r[0].toString().toLowerCase().trim().split('#')[0].replace(/\/$/, "");
      gscMap.set(cleanUrl, [r[1], r[2], r[3], r[4]]); 
    }
  });

  // 2. Normalize Selected Stone
  const targetStone = selectedStone.toString().trim().toLowerCase();

  // 3. Process Main Sheet
  const lastRow = mainSheet.getLastRow();
  const mainData = mainSheet.getRange(2, 1, lastRow - 1, mainSheet.getLastColumn()).getValues();

  let matchCount = 0;
  const outputData = mainData.map((row) => {
    // Column F normalized
    const rowStone = row[COL.STONE - 1].toString().trim().toLowerCase();
    // Column C normalized
    const rowUrl = row[COL.URL - 1].toString().toLowerCase().trim().split('#')[0].replace(/\/$/, "");
    
    const currentBench = [row[52], row[53], row[54], row[55]];

    // Strict comparison: Both must be identical after cleaning
    if (rowStone === targetStone && gscMap.has(rowUrl)) {
      matchCount++;
      return gscMap.get(rowUrl);
    }
    return currentBench;
  });

  // 4. Final Write
  if (matchCount > 0) {
    mainSheet.getRange(2, COL.BENCH_START, outputData.length, 4).setValues(outputData);
    SpreadsheetApp.getUi().alert(`✅ SUCCESS\n\nUpdated ${matchCount} rows for "${selectedStone}".`);
  } else {
    SpreadsheetApp.getUi().alert(`⚠️ 0 ROWS UPDATED\n\nThe script couldn't find a match where Column F is "${selectedStone}" AND the URL exists in your GSC sheet.`);
  }
}