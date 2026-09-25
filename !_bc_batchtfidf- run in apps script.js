// bc_batchTfIdf- RUN IN APPS SCRIPT
function bc_batchCalculateTfIdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const posts = ss.getSheetByName("posts");

  if (!posts) throw new Error("posts sheet not found.");

  const postHeaders = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0];
  const postIdIdx = postHeaders.indexOf("Post ID");
  const drIdx = postHeaders.indexOf("GSC Query List");
  const tfidfTermsIdx = postHeaders.indexOf("tfidf_terms");
  const tfidfSummaryIdx = postHeaders.indexOf("tfidf_summary");

  if (postIdIdx === -1) throw new Error("Post ID column not found.");
  if (drIdx === -1) throw new Error("GSC Query List column not found.");
  if (tfidfTermsIdx === -1) throw new Error("tfidf_terms column not found.");
  if (tfidfSummaryIdx === -1) throw new Error("tfidf_summary column not found.");

  const postsData = posts.getDataRange().getValues();

  // Build corpus: postId -> array of words from DR column
  const corpus = {};
  const rowIndex = {};

  for (let i = 1; i < postsData.length; i++) {
    const postId = String(postsData[i][postIdIdx] || "").trim();
    const drValue = String(postsData[i][drIdx] || "").trim();
    if (!postId || !drValue) continue;

    const words = [];
    drValue.split(",").forEach(function(query) {
      query.trim().toLowerCase().split(/\s+/).forEach(function(word) {
        const clean = word.replace(/[^a-z]/g, "");
        if (clean.length >= 3 && !TFIDF_STOPWORDS.has(clean)) {
          words.push(clean);
        }
      });
    });

    corpus[postId] = words;
    rowIndex[postId] = i + 1;
  }

  const docCount = Object.keys(corpus).length;
  Logger.log("Corpus size: " + docCount + " documents.");

  if (docCount === 0) {
    Logger.log("No corpus data found.");
    return;
  }

  // Build document frequency table
  const df = {};
  Object.keys(corpus).forEach(function(postId) {
    const uniqueWords = new Set(corpus[postId]);
    uniqueWords.forEach(function(word) {
      df[word] = (df[word] || 0) + 1;
    });
  });

  // Calculate TF-IDF per document and write to posts sheet
  let written = 0;

  Object.keys(corpus).forEach(function(postId) {
    const words = corpus[postId];
    if (words.length === 0) return;

    const tf = {};
    words.forEach(function(word) {
      tf[word] = (tf[word] || 0) + 1;
    });

    const scores = {};
    Object.keys(tf).forEach(function(word) {
      const termTf = tf[word] / words.length;
      const termIdf = Math.log(docCount / (df[word] || 1));
      scores[word] = termTf * termIdf;
    });

    const ranked = Object.keys(scores).sort(function(a, b) {
      return scores[b] - scores[a];
    }).slice(0, 15);

    const tfidfTerms = ranked.map(function(word) {
      return word + ":" + scores[word].toFixed(3);
    }).join(", ");

    const tfidfSummary = ranked.slice(0, 8).join(", ");

    const row = rowIndex[postId];
    posts.getRange(row, tfidfTermsIdx + 1).setValue(tfidfTerms);
    posts.getRange(row, tfidfSummaryIdx + 1).setValue(tfidfSummary);
    written++;
  });

  Logger.log("TF-IDF complete. " + written + " rows updated.");
}

const TFIDF_STOPWORDS = new Set([
  "the", "and", "for", "with", "how", "can", "you", "your", "from", "that",
  "this", "are", "was", "have", "has", "not", "use", "used", "using",
  "what", "why", "when", "will", "does", "best", "get", "all", "its",
  "more", "been", "into", "our", "out", "their", "they", "who", "which",
  "but", "also", "than", "then", "them", "these", "those", "some", "any",
  "one", "two", "per", "new", "old", "top", "way", "need", "make",
  "take", "much", "many", "long", "very", "well", "good", "able"
]);