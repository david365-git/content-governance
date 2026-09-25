/**
 * ============================================================
 * bc_QueryVectorCalculator.gs
 * Abbey Floor Care — Query Vector Analysis
 * Calculates query vector from Primary Query Cluster Owned
 * and writes to query_vector_analysis sheet
 * ============================================================
 */

function bc_calculateQueryVector(lockedRow) {
  Logger.log("bc_calculateQueryVector called");
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const postsSheet = ss.getSheets().find(s => s.getSheetId() == BC_SHEET_CONFIG.posts);
    if (!postsSheet) throw new Error("Posts sheet not found.");
    
    const activeRowIndex = lockedRow || ss.getActiveRange().getRow();
    if (activeRowIndex < 2) throw new Error("No valid row to process.");
    
    const headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];
    const postIdIdx = headers.indexOf("Post ID");
    const gscQueryListIdx = headers.indexOf("GSC Query List");
    const queriesIdx = headers.indexOf("Queries");

    if (postIdIdx === -1) throw new Error("Post ID column not found.");
    if (gscQueryListIdx === -1) throw new Error("GSC Query List column not found.");

    const postId = postsSheet.getRange(activeRowIndex, postIdIdx + 1).getValue();
    const primaryQueryCluster = String(postsSheet.getRange(activeRowIndex, gscQueryListIdx + 1).getValue() || "").trim();
    const rawQueries = queriesIdx !== -1 ? String(postsSheet.getRange(activeRowIndex, queriesIdx + 1).getValue() || "") : "";
    const rawHtml = bc_getPostHtmlById_(String(postId));

    if (!postId) throw new Error("Post ID is empty.");
    if (!primaryQueryCluster) throw new Error("GSC Query List is empty.");

    // Strip HTML tags to get plain text
    const pageText = rawHtml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").toLowerCase();

    // --- Parse raw queries with CTR and impressions ---
    const stopwords = ["the", "and", "or", "but", "in", "on", "at", "to", "for", "of", "with", "a", "an", "is", "it", "that", "this", "as", "was", "are", "be", "by", "from", "has", "have", "had", "not", "can", "will", "all", "would", "there", "their", "how", "you", "do", "does", "what", "your"];

    // Load UK towns and English words for filtering
    const townSheet = ss.getSheetByName("uk-towns");
    const townSet = new Set();
    if (townSheet) {
      const townData = townSheet.getRange(2, 1, townSheet.getLastRow() - 1, 1).getValues();
      townData.forEach(function(row) {
        if (row[0]) townSet.add(String(row[0]).toLowerCase().trim());
      });
    }

    const wordsSheet = ss.getSheetByName("uk-words");
    const englishWordSet = new Set();
    if (wordsSheet) {
      const wordsData = wordsSheet.getRange(2, 1, wordsSheet.getLastRow() - 1, 1).getValues();
      wordsData.forEach(function(row) {
        if (row[0]) englishWordSet.add(String(row[0]).toLowerCase().trim());
      });
    }
    

    const excludeBrand = /\b(abbey|tile doctor)\b/i;
    const excludeGeo = /\b(near me|near you|in my area|local|locally)\b/i;
    const excludeWrongMaterial = /\b(carpet tile|vinyl|laminate|porcelain|terracotta|wood floor|wooden floor|hardwood)\b/i;
    const excludeUnrelated = /\b(fireplace|countertop|worktop|dry lin|lift install|youtube|definition|cleaning job|surface repair)\b/i;
    const excludeOperators = /inurl:|site:|"\w+.*"\s+\w+/i;

    // Parse raw queries retaining CTR and impressions
    const filteredQueryData = [];
    if (rawQueries) {
      const lines = String(rawQueries).split('\n');
      lines.forEach(function(line) {
        if (!line.trim()) return;
        const parts = line.split(/\s{2,}/);
        if (parts.length < 3) return;

        const query = parts[0].trim().toLowerCase();
        const clicks = parseInt(parts[1]) || 0;
        const impressions = parseInt(parts[2]) || 0;
        const ctr = parts.length >= 4 ? parseFloat(parts[3]) || 0 : 0;

        if (!query) return;
        if (/^(query|top queries|clicks|impressions|ctr|position)$/i.test(query)) return;
        if (excludeBrand.test(query)) return;
        if (excludeGeo.test(query)) return;
        if (excludeUnrelated.test(query)) return;
        if (excludeOperators.test(query)) return;
        if (/[^\u0000-\u007F]/.test(query)) return;

        const queryWords = query.split(/\s+/);
        if (!queryWords.some(function(word) { return englishWordSet.has(word); })) return;
        if (queryWords.some(function(word) { return word.length >= 4 && townSet.has(word); })) return;
        if (!queryWords.some(function(word) { return englishWordSet.has(word); })) return;

        if (excludeWrongMaterial.test(query)) {
          const material = primaryQueryCluster.split(',')[0].trim().toLowerCase().split(' ')[0];
          if (!query.indexOf(material) !== -1) return;
        }

        if (clicks === 0 && impressions < 10) return;

        filteredQueryData.push({ query: query, clicks: clicks, impressions: impressions, ctr: ctr });
      });
    }

    // Determine weighting mode
    const totalClicks = filteredQueryData.reduce(function(sum, q) { return sum + q.clicks; }, 0);
    const totalImpressions = filteredQueryData.reduce(function(sum, q) { return sum + q.impressions; }, 0);

    let weightMode = "raw";
    if (totalClicks > 0) {
      weightMode = "ctr";
    } else if (totalImpressions > 0) {
      weightMode = "impressions";
    }

    // Calculate weighted term frequencies
    const weightedTerms = {};
    filteredQueryData.forEach(function(item) {
      const weight = weightMode === "ctr" ? (item.ctr || 0.01) :
                     weightMode === "impressions" ? item.impressions :
                     1;

      item.query.split(/\s+/).forEach(function(word) {
        const clean = word.replace(/[^a-z]/g, "");
        if (clean.length > 3 && stopwords.indexOf(clean) === -1) {
          weightedTerms[clean] = (weightedTerms[clean] || 0) + weight;
        }
      });
    });

    // Sort weighted terms
    const sortedWeighted = Object.keys(weightedTerms).sort(function(a, b) {
      return weightedTerms[b] - weightedTerms[a];
    });

    const weightedFreqString = sortedWeighted.map(function(term) {
      return term + ":" + Math.round(weightedTerms[term] * 100) / 100;
    }).join(", ");

    // Build CTR query data string for storage
    const ctrQueryDataString = filteredQueryData.map(function(item) {
      return item.query + " [" + (weightMode === "ctr" ? "CTR:" + item.ctr : weightMode === "impressions" ? "Imp:" + item.impressions : "raw") + "]";
    }).join(", ");

    // Get or create query_vector_analysis sheet
    let qvSheet = ss.getSheetByName("query_vector_analysis");
    if (!qvSheet) {
      qvSheet = ss.insertSheet("query_vector_analysis");
      qvSheet.getRange(1, 1, 1, 10).setValues([[
        "post_id",
        "primary_query_cluster",
        "extracted_terms",
        "term_frequencies",
        "high_weight_terms",
        "medium_weight_terms",
        "query_vector_summary",
        "precision_score",
        "weighted_query_data",
        "weighted_term_frequencies"
      ]]);
    } else {
      const existingHeaders = qvSheet.getRange(1, 1, 1, qvSheet.getLastColumn()).getValues()[0];
      if (existingHeaders.indexOf("precision_score") === -1) {
        qvSheet.getRange(1, 8).setValue("precision_score");
      }
      if (existingHeaders.indexOf("weighted_query_data") === -1) {
        qvSheet.getRange(1, 9).setValue("weighted_query_data");
      }
      if (existingHeaders.indexOf("weighted_term_frequencies") === -1) {
        qvSheet.getRange(1, 10).setValue("weighted_term_frequencies");
      }
    }

    // Find or create row for this Post ID
    const qvData = qvSheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < qvData.length; i++) {
      if (String(qvData[i][0]) === String(postId)) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      targetRow = qvSheet.getLastRow() + 1;
    }

    // Write Post ID and GSC Query List
    qvSheet.getRange(targetRow, 1).setValue(postId);
    qvSheet.getRange(targetRow, 2).setValue(primaryQueryCluster);

    // Compute values directly (one-time) instead of live formulas, to stop
    // constant background recalculation across the whole query_vector_analysis sheet.
    function extractTermsPlain(text) {
      if (!text) return "";
      var words = String(text).toLowerCase().replace(/[^a-z ]/g, " ").split(/\s+/);
      var seen = {};
      var out = [];
      words.forEach(function(w) {
        if (w.length > 3 && !seen[w]) { seen[w] = true; out.push(w); }
      });
      return out.join(", ");
    }

    const extractedTermsValue = primaryQueryCluster ? extractTermsPlain(primaryQueryCluster) : "";
    const termFreqValue = primaryQueryCluster ? countTermFrequencies(primaryQueryCluster) : "";
    const highWeightValue = extractedTermsValue ? filterHighWeightTerms(termFreqValue, 3) : "";
    const mediumWeightValue = extractedTermsValue ? filterMediumWeightTerms(termFreqValue, 1, 2) : "";
    const summaryValue = (highWeightValue || mediumWeightValue)
      ? "High-weight: " + highWeightValue + " | Medium-weight: " + mediumWeightValue
      : "";

    qvSheet.getRange(targetRow, 3).setValue(extractedTermsValue);
    qvSheet.getRange(targetRow, 4).setValue(termFreqValue);
    qvSheet.getRange(targetRow, 5).setValue(highWeightValue);
    qvSheet.getRange(targetRow, 6).setValue(mediumWeightValue);
    qvSheet.getRange(targetRow, 7).setValue(summaryValue);

    // Write precision score
    if (pageText) {
      const queryTerms = {};
      primaryQueryCluster.split(",").forEach(function(q) {
        q.trim().toLowerCase().replace(/[^a-z ]/g, "").split(/\s+/).forEach(function(word) {
          if (word.length > 3 && stopwords.indexOf(word) === -1) {
            queryTerms[word] = true;
          }
        });
      });

      const uniqueQueryTerms = Object.keys(queryTerms);
      let matchCount = 0;
      uniqueQueryTerms.forEach(function(term) {
        if (pageText.indexOf(term) !== -1) matchCount++;
      });

      const precisionScore = uniqueQueryTerms.length > 0
        ? Math.round((matchCount / uniqueQueryTerms.length) * 100) + "%"
        : "n/a";

      qvSheet.getRange(targetRow, 8).setValue(precisionScore);
    } else {
      qvSheet.getRange(targetRow, 8).setValue("no HTML");
    }

    // Write weighted query data and weighted term frequencies
    qvSheet.getRange(targetRow, 9).setValue(ctrQueryDataString || "no query data");
    qvSheet.getRange(targetRow, 10).setValue(weightedFreqString || "no weighted terms");

    return "Query vector calculated for Post ID " + postId + " [weight mode: " + weightMode + "] FINISHED";

  } catch (e) {
    throw new Error("bc_calculateQueryVector: " + e.message);
  }
}
/**
 * Count term frequencies in a query cluster
 * @param {string} queryCluster - Comma-separated queries or newline-separated
 * @return {string} Term:frequency pairs (e.g., "quarry:4, tile:4, efflorescence:3")
 */
function countTermFrequencies(queryCluster) {
  if (!queryCluster || queryCluster.trim() === "") return "";
  
  // Common English stopwords to exclude
  var stopwords = ["the", "and", "or", "but", "in", "on", "at", "to", "for", "of", "with", "a", "an", "is", "it", "that", "this", "as", "was", "are", "be", "by", "from", "has", "have", "had", "not", "can", "will", "all", "would", "there", "their"];
  
  var terms = {};
  
  // Split by comma or newline
  var queries = queryCluster.split(/[,\n]+/);
  
  queries.forEach(function(query) {
    // Clean and split into words
    var words = query.toLowerCase()
                     .replace(/[^a-z ]/g, ' ')
                     .trim()
                     .split(/\s+/)
                     .filter(function(w) { 
                       return w.length > 3 && stopwords.indexOf(w) === -1; 
                     });
    
    words.forEach(function(word) {
      terms[word] = (terms[word] || 0) + 1;
    });
  });
  
  // Sort by frequency (descending)
  var sorted = Object.keys(terms).sort(function(a, b) {
    return terms[b] - terms[a];
  });
  
  var result = sorted.map(function(term) {
    return term + ":" + terms[term];
  });
  
  return result.join(", ");
}

/**
 * Filter high-weight terms (frequency >= threshold)
 * @param {string} termFreqs - "term:count, term:count" format
 * @param {number} threshold - Minimum frequency (default 3)
 * @return {string} Comma-separated high-weight terms
 */
function filterHighWeightTerms(termFreqs, threshold) {
  if (!termFreqs || termFreqs.trim() === "") return "";
  
  threshold = threshold || 3;
  var pairs = termFreqs.split(", ");
  var highWeight = [];
  
  pairs.forEach(function(pair) {
    var parts = pair.split(":");
    var term = parts[0];
    var count = parseInt(parts[1]);
    
    if (count >= threshold) {
      highWeight.push(term);
    }
  });
  
  return highWeight.join(", ");
}

/**
 * Filter medium-weight terms (frequency between min and max)
 * @param {string} termFreqs - "term:count, term:count" format
 * @param {number} minThreshold - Minimum frequency (default 1)
 * @param {number} maxThreshold - Maximum frequency (default 2)
 * @return {string} Comma-separated medium-weight terms
 */
function filterMediumWeightTerms(termFreqs, minThreshold, maxThreshold) {
  if (!termFreqs || termFreqs.trim() === "") return "";
  
  minThreshold = minThreshold || 1;
  maxThreshold = maxThreshold || 2;
  var pairs = termFreqs.split(", ");
  var mediumWeight = [];
  
  pairs.forEach(function(pair) {
    var parts = pair.split(":");
    var term = parts[0];
    var count = parseInt(parts[1]);
    
    if (count >= minThreshold && count <= maxThreshold) {
      mediumWeight.push(term);
    }
  });
  
  return mediumWeight.join(", ");
}

function bc_buildQueryVectorPreview_(gscQueryList, fullHtml) {
  try {
    if (!gscQueryList || !fullHtml) {
      return { precisionScore: "n/a", weightedTermsString: "no data available", topTerms: [] };
    }

    const stopwords = ["the", "and", "or", "but", "in", "on", "at", "to", "for", "of", "with", "a", "an", "is", "it", "that", "this", "as", "was", "are", "be", "by", "from", "has", "have", "had", "not", "can", "will", "all", "would", "there", "their", "how", "you", "do", "does", "what", "your"];

    const pageText = String(fullHtml).replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").toLowerCase();

    // gscQueryList is a comma-separated string of queries (from bc_extractGscQueryList_)
    const queries = String(gscQueryList).split(',').map(function(q) { return q.trim().toLowerCase(); }).filter(Boolean);

    const termCounts = {};
    queries.forEach(function(query) {
      query.replace(/[^a-z ]/g, ' ').split(/\s+/).forEach(function(word) {
        if (word.length > 3 && stopwords.indexOf(word) === -1) {
          termCounts[word] = (termCounts[word] || 0) + 1;
        }
      });
    });

    const uniqueTerms = Object.keys(termCounts);
    let matchCount = 0;
    const missingTerms = [];
    const presentTerms = [];

    uniqueTerms.forEach(function(term) {
      if (pageText.indexOf(term) !== -1) {
        matchCount++;
        presentTerms.push(term);
      } else {
        missingTerms.push(term);
      }
    });

    const precisionScore = uniqueTerms.length > 0
      ? Math.round((matchCount / uniqueTerms.length) * 100) + "%"
      : "n/a";

    const sortedByFreq = uniqueTerms.sort(function(a, b) { return termCounts[b] - termCounts[a]; });
    const topTerms = sortedByFreq.slice(0, 15);

    return {
      precisionScore: precisionScore,
      missingTerms: missingTerms.slice(0, 15),
      presentTerms: presentTerms.slice(0, 15),
      topTerms: topTerms
    };

  } catch (e) {
    return { precisionScore: "error", missingTerms: [], presentTerms: [], topTerms: [], error: e.message };
  }
}