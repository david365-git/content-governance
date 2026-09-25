function bc_batchPopulateGscQueryList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const posts = ss.getSheetByName("posts");
  const data = posts.getDataRange().getValues();
  const headers = data[0];

  const colIndex = function(name) {
    return headers.indexOf(name);
  };

  const queriesIdx = colIndex("Queries");
  const materialIdx = colIndex("Material Entity");
  const drIdx = colIndex("GSC Query List");

  if (queriesIdx === -1 || materialIdx === -1 || drIdx === -1) {
    throw new Error("Required column not found. Check Queries, Material Entity, GSC Query List headers.");
  }

  let written = 0;

  for (let i = 1; i < data.length; i++) {
    const raw = String(data[i][queriesIdx] || "").trim().replace(/\t/g, '  ');
    const material = String(data[i][materialIdx] || "").trim();
    const existing = String(data[i][drIdx] || "").trim();

    if (!raw) continue;
    if (existing) continue;

    const result = bc_extractGscQueryList_(raw, material);
    if (result) {
      posts.getRange(i + 1, drIdx + 1).setValue(result);
      written++;
    }
  }

  Logger.log("Done. " + written + " rows populated.");
}