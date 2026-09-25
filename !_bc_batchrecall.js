function bc_batchCalculateRecall() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const posts = ss.getSheetByName("posts");

  if (!posts) throw new Error("posts sheet not found.");

  const headers = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0];

  const colIndex = function(name) {
    const idx = headers.indexOf(name);
    if (idx === -1) throw new Error("Column not found: " + name);
    return idx;
  };

  const postIdIdx = colIndex("Post ID");
  const drIdx = colIndex("GSC Query List");
  const pqcIdx = colIndex("Primary Query Cluster Owned");
  const tfidfSummaryIdx = colIndex("tfidf_summary");
  const recallScoreIdx = colIndex("recall_score");
  const recallNotesIdx = colIndex("recall_notes");

  const data = posts.getDataRange().getValues();
  let written = 0;

  for (let i = 1; i < data.length; i++) {
    const dr = String(data[i][drIdx] || "").trim();
    const pqc = String(data[i][pqcIdx] || "").trim().toLowerCase();
    const tfidfSummary = String(data[i][tfidfSummaryIdx] || "").trim();

    const articleType = String(data[i][colIndex("Article Type")] || "").trim();
      if (!dr || !tfidfSummary) continue;
      if (articleType === "Case Study") {
        posts.getRange(i + 1, recallScoreIdx + 1).setValue("N/A");
        posts.getRange(i + 1, recallNotesIdx + 1).setValue("Case Study — recall not applicable");
        written++;
        continue;
      }

    // Extract TF-IDF terms (already tokenised words)
    const tfidfTerms = tfidfSummary.split(",").map(function(t) {
      return t.trim().toLowerCase();
    }).filter(Boolean);

    if (tfidfTerms.length === 0) continue;

    // Tokenise the Primary Query Cluster into words
    const pqcWords = pqc.replace(/[^a-z\s]/g, "").split(/\s+/).filter(function(w) {
      return w.length >= 3;
    });

    // Count how many TF-IDF terms appear in the PQC
    let matches = 0;
    const matched = [];
    const unmatched = [];

    tfidfTerms.forEach(function(term) {
      if (pqcWords.indexOf(term) !== -1) {
        matches++;
        matched.push(term);
      } else {
        unmatched.push(term);
      }
    });

    const recallScore = Math.round((matches / tfidfTerms.length) * 100) / 100;

    let notes = "";
    if (matched.length > 0) {
      notes += "Matched: " + matched.join(", ") + ". ";
    }
    if (unmatched.length > 0) {
      notes += "Not in PQC: " + unmatched.slice(0, 5).join(", ");
      if (unmatched.length > 5) notes += " (+" + (unmatched.length - 5) + " more)";
    }

    posts.getRange(i + 1, recallScoreIdx + 1).setValue(recallScore);
    posts.getRange(i + 1, recallNotesIdx + 1).setValue(notes);
    written++;
  }

  Logger.log("Recall complete. " + written + " rows updated.");
}
