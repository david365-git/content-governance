/**
 * ================================================================================
 * ce_Stage7_RunAnalysis.gs - STAGE 7 — RUN ANALYSIS
 * ================================================================================
 *
 * Reads the two most recent publication_log entries for the active post ID
 * and returns a recommendation: Run Again, Borderline, or Publish.
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * ================================================================================
 */

function getW7RunAnalysis() {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const posts    = ss.getSheetByName("posts");
    const logSheet = ss.getSheetByName("publication_log");
    if (!posts)    return { success: false, message: "posts sheet not found." };
    if (!logSheet) return { success: false, message: "publication_log sheet not found." };

    const row = posts.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const postHeaders = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0];
    const postIdIdx   = postHeaders.indexOf("Post ID");
    const postId      = postIdIdx > -1 ? String(posts.getRange(row, postIdIdx + 1).getValue() || "").trim() : "";
    if (!postId) return { success: false, message: "No Post ID found on active row." };

    // Read publication log
    const logData    = logSheet.getDataRange().getValues();
    const logHeaders = logData[0].map(function(h) { return String(h).trim(); });

    const colPostId        = logHeaders.indexOf("Post ID");
    const colRecallScore   = logHeaders.indexOf("recall_score");
    const colRecallNotes   = logHeaders.indexOf("recall_notes");
    const colRecallPrev    = logHeaders.indexOf("recall_score_previous");
    const colRecallNotesPrev = logHeaders.indexOf("recall_notes_previous");
    const colCompWords     = logHeaders.indexOf("Competitor Word Count");
    const colTfidf         = logHeaders.indexOf("tfidf_terms");
    const colTfidfSummary  = logHeaders.indexOf("tfidf_summary");
    const colArticleType   = logHeaders.indexOf("Article Type");
    const colHtml          = logHeaders.indexOf("New HTML");

    // Collect all rows for this post ID — log is sorted descending so first match is most recent
    const matchingRows = [];
    for (let i = 1; i < logData.length; i++) {
      if (String(logData[i][colPostId] || "").trim() === postId) {
        matchingRows.push(logData[i]);
      }
    }

    if (matchingRows.length === 0) {
      return { success: false, message: "No publication log entries found for Post ID " + postId + ". Run W6 first." };
    }

    const latest = matchingRows[0];
    const previous = matchingRows.length > 1 ? matchingRows[1] : null;

    // Extract values
    const recallScore     = parseFloat(latest[colRecallScore]) || 0;
    const recallNotes     = String(latest[colRecallNotes] || "").trim();
    const recallScorePrev = previous ? (parseFloat(previous[colRecallScore]) || 0) : null;
    const recallNotesPrev = previous ? String(previous[colRecallNotes] || "").trim() : null;
    const compWords       = parseInt(latest[colCompWords]) || 0;
    const tfidfTerms      = String(latest[colTfidf] || "").trim();
    const tfidfSummary    = String(latest[colTfidfSummary] || "").trim();
    const articleType     = String(latest[colArticleType] || "").trim();
    const htmlContent     = String(latest[colHtml] || "").trim();

    // Estimate article word count from HTML
    const plainText   = htmlContent.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
    const articleWords = plainText.split(" ").filter(function(w) { return w.length > 0; }).length;

    // Score improvement between runs
    const scoreImproved = recallScorePrev !== null && recallScore > recallScorePrev;
    const scoreFlat     = recallScorePrev !== null && recallScore === recallScorePrev;
    const scoreDrop     = recallScorePrev !== null && recallScore < recallScorePrev;
    const firstRun      = recallScorePrev === null;

    // Word count gap
    const wordCountGap  = compWords > 0 ? compWords - articleWords : 0;
    const wordCountLow  = compWords > 0 && articleWords < (compWords * 0.75);

    // Missing PQC terms from recall notes
    const missingTerms  = extractMissingTerms(recallNotes);
    const prevMissingTerms = recallNotesPrev ? extractMissingTerms(recallNotesPrev) : [];
    const sameMissing   = missingTerms.length > 0 && arraysMatch(missingTerms, prevMissingTerms);

    // Case Study override — recall not applicable
    const isCaseStudy   = articleType.toLowerCase().indexOf("case study") > -1;
    const recallNA      = recallNotes.toLowerCase().indexOf("not applicable") > -1;

    // Build signals array
    const signals = [];
    const runAgainSignals = [];
    const publishSignals  = [];
    const borderlineSignals = [];

    // Signal 1 — Recall score (page-type aware via Article Type Control Sheet)
    if (isCaseStudy || recallNA) {
      signals.push({ label: "Recall Score", value: "N/A — Case Study", verdict: "neutral", detail: "Recall scoring does not apply to Case Study article type." });
    } else {
      var missingForRec = extractMissingTerms(recallNotes);
      var recallRec = buildRecallRecommendation(recallScore, articleType, [], missingForRec);
      var recallVerdict = "borderline";
      var recallMeasure = "";

      // Parse recommendation to determine verdict
      if (recallRec.indexOf("RECALL MEASURE: Structural completeness") > -1) {
        // Hub Page — score above 0.8 is a cannibalisation warning, not a publish signal
        if (recallScore > 0.8) {
          recallVerdict = "borderline";
          recallMeasure = "Hub Page — score above 0.8 suggests hub may be doing spoke page jobs. Review for cannibalisation before publishing.";
        } else {
          recallVerdict = "publish";
          recallMeasure = "Hub Page — score of " + recallScore + " is within expected range for a routing page.";
        }
      } else if (recallScore >= 0.85) {
        recallVerdict = "publish";
        recallMeasure = "Score of " + recallScore + " meets the threshold for this article type.";
        publishSignals.push("recall");
      } else if (recallScore >= 0.65) {
        recallVerdict = "borderline";
        recallMeasure = "Score of " + recallScore + " is borderline. " + recallRec;
        borderlineSignals.push("recall");
      } else {
        recallVerdict = "run_again";
        recallMeasure = "Score of " + recallScore + " is below threshold. " + recallRec;
        runAgainSignals.push("recall");
      }

      signals.push({ label: "Recall Score", value: recallScore, verdict: recallVerdict, detail: recallMeasure });
    }

    // Signal 2 — Score improvement between runs
    if (!firstRun && !isCaseStudy && !recallNA) {
      if (scoreImproved) {
        signals.push({ label: "Score Trend", value: recallScorePrev + " → " + recallScore, verdict: "run_again", detail: "Score improved between runs (" + recallScorePrev + " to " + recallScore + "). A further run may continue this improvement." });
        runAgainSignals.push("trend");
      } else if (scoreFlat && sameMissing) {
        signals.push({ label: "Score Trend", value: "Flat at " + recallScore, verdict: "publish", detail: "Score held flat between runs and the same terms remain missing. Additional rewrites are unlikely to change coverage — the ceiling has been reached." });
        publishSignals.push("trend");
      } else if (scoreFlat && !sameMissing) {
        signals.push({ label: "Score Trend", value: "Flat at " + recallScore, verdict: "borderline", detail: "Score is flat but different terms are missing between runs, suggesting the pipeline is still shifting coverage. One more run may stabilise." });
        borderlineSignals.push("trend");
      } else if (scoreDrop) {
        signals.push({ label: "Score Trend", value: recallScorePrev + " → " + recallScore, verdict: "borderline", detail: "Score dropped between runs. This may indicate over-optimisation or entity dilution. Review before running again." });
        borderlineSignals.push("trend");
      }
    } else if (firstRun && !isCaseStudy && !recallNA) {
      // If score is already at or above threshold on first run, do not push run_again
      var alreadyAtThreshold = false;
      if (articleType === 'Hub Page' && recallScore <= 0.8) alreadyAtThreshold = true;
      else if (articleType === 'Buyer Guide' && recallScore >= 0.7) alreadyAtThreshold = true;
      else if (recallScore >= 0.8) alreadyAtThreshold = true;

      if (alreadyAtThreshold) {
        signals.push({ label: "Score Trend", value: "First run", verdict: "publish", detail: "First pipeline run. Recall score is already at or above threshold — a second run is unlikely to improve coverage further." });
        publishSignals.push("trend");
      } else {
        signals.push({ label: "Score Trend", value: "First run", verdict: "run_again", detail: "This is the first pipeline run. A second run is recommended to check for further entity coverage gains." });
        runAgainSignals.push("trend");
      }
    }

    // Signal 3 — Missing PQC terms
    // Filter out stemming artefacts and recall recommendation text masquerading as missing terms
    var STEMMING_ARTEFACTS = ['surfac', 'choos', 'turn', 'tri', 'use', 'caus', 'provid', 'includ', 'requir'];
    var genuineMissingTerms = missingTerms.filter(function(t) {
      return t.length <= 30 && STEMMING_ARTEFACTS.indexOf(t.toLowerCase()) === -1 && !/RECALL MEASURE|RECOMMENDATION|SCORE|pipeline rerun/i.test(t);
    });

    if (!isCaseStudy && !recallNA && genuineMissingTerms.length > 0) {
      signals.push({ label: "Missing PQC Terms", value: genuineMissingTerms.join(", "), verdict: "run_again", detail: "The following PQC terms are not yet matched: " + genuineMissingTerms.join(", ") + ". A further run may add coverage for these." });
      runAgainSignals.push("pqc");
    } else if (!isCaseStudy && !recallNA) {
      signals.push({ label: "Missing PQC Terms", value: "None", verdict: "publish", detail: "All PQC terms are matched. No coverage gaps identified." });
      publishSignals.push("pqc");
    }

    // Signal 4 — Word count vs competitor
    if (compWords > 0) {
      if (wordCountLow) {
        signals.push({ label: "Word Count", value: articleWords + " words (competitor: " + compWords + ")", verdict: "run_again", detail: "Article is " + wordCountGap + " words below 75% of the competitor benchmark (" + compWords + " words). A further run is likely to produce more complete content." });
        runAgainSignals.push("wordcount");
      } else if (wordCountGap > 0 && wordCountGap <= compWords * 0.25) {
        signals.push({ label: "Word Count", value: articleWords + " words (competitor: " + compWords + ")", verdict: "borderline", detail: "Article is within 25% of the competitor benchmark. Content depth is adequate but a further run could close the gap." });
        borderlineSignals.push("wordcount");
      } else {
        signals.push({ label: "Word Count", value: articleWords + " words (competitor: " + compWords + ")", verdict: "publish", detail: "Article meets or exceeds the competitor word count benchmark." });
        publishSignals.push("wordcount");
      }
    }

    // Signal 5 — TF-IDF terms
    if (tfidfTerms) {
      signals.push({ label: "TF-IDF Terms", value: tfidfTerms, verdict: "neutral", detail: tfidfSummary || "Dominant vocabulary fingerprint for this run." });
    }

    // Signal 6 — Entity dump detection (informational only — the mechanical regex
    // over-flags normal prose with verbs outside its whitelist. Genuine judgment now
    // happens in the Stage 3B deferred LLM audit, CHECK 8.)
    if (htmlContent) {
      var entityDumps = detectEntityDumps(htmlContent);
      if (entityDumps.length > 0) {
        var dumpDetail = entityDumps.length + " paragraph(s) mechanically flagged as possible comma-separated lists. This is informational only — confirm with the Stage 3B CHECK 8 result before treating as a genuine issue. Sections: " +
          entityDumps.map(function(d) { return d.sectionId; }).join(", ");
        signals.push({ label: "Entity Dumps (mechanical scan)", value: entityDumps.length + " flagged", verdict: "neutral", detail: dumpDetail });
      } else {
        signals.push({ label: "Entity Dumps (mechanical scan)", value: "None flagged", verdict: "neutral", detail: "No comma-separated entity list paragraphs mechanically flagged." });
      }
    }

    // Final recommendation
    let recommendation, recommendationDetail;

    // Entity dumps are now judged by Stage 3B CHECK 8 (LLM), not hard-blocked here.
    if (runAgainSignals.length >= 2) {
      recommendation = "RUN AGAIN";
      recommendationDetail = "Multiple signals suggest a further pipeline run will improve this article. " + buildReasonSummary(signals, "run_again");
    } else if (runAgainSignals.length === 1 && borderlineSignals.length === 0) {
      recommendation = "RUN AGAIN";
      recommendationDetail = "One significant signal suggests a further run is worthwhile. " + buildReasonSummary(signals, "run_again");
    } else if (publishSignals.length >= 2 && runAgainSignals.length === 0) {
      recommendation = "PUBLISH";
      recommendationDetail = "Multiple positive signals confirm this article is ready to publish. " + buildReasonSummary(signals, "publish");
    } else if (isCaseStudy || recallNA) {
      recommendation = "PUBLISH";
      recommendationDetail = "Case Study article type — recall scoring does not apply. Review content quality manually before publishing.";
    } else {
      recommendation = "BORDERLINE";
      recommendationDetail = "Signals are mixed. Review the detail below before deciding whether to run again or publish. " + buildReasonSummary(signals, "borderline");
    }

    // Store W7 analysis to posts sheet col 163 (FG — W8_Analysis)
    try {
      var analysisSummary = 'RECOMMENDATION: ' + recommendation + '\n' +
        recommendationDetail + '\n\n' +
        'ARTICLE TYPE: ' + articleType + '\n' +
        'SIGNAL BREAKDOWN\n' +
        signals.map(function(s) {
          return s.label + ': ' + s.value + '\n' + s.detail;
        }).join('\n');
      var analysisCell = posts.getRange(row, 163);
      analysisCell.setNumberFormat('@');
      analysisCell.setValue(analysisSummary);
    } catch(storeErr) {
      // Non-fatal — analysis display still returns normally
    }

    return {
      success: true,
      postId: postId,
      articleType: articleType,
      recommendation: recommendation,
      recommendationDetail: recommendationDetail,
      signals: signals,
      runCount: matchingRows.length
    };

  } catch(e) {
    return { success: false, message: "getW7RunAnalysis error: " + e.message };
  }
}
function getW7BioData() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    function getVal(name) {
      const idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
    }

    const html        = getVal("New HTML");
    const articleType = getVal("Article Type");
    const stoneType   = getVal("Stone Type");
    const locality    = getVal("Locality");
    const parentArea  = getVal("Parent Area");
    const postId      = getVal("Post ID");

    if (!html) return { success: false, message: "No HTML found in New HTML column. Run W4 first." };

    // Extract current bio paragraph text from bio box
    var currentBio = "";
    var bioBoxMatch = html.match(/<div[^>]*abbey-bio-box[^>]*>[\s\S]*?<p[^>]*style[^>]*color:#444[^>]*>([\s\S]*?)<\/p>[\s\S]*?<\/div>/i);
    if (bioBoxMatch) {
      currentBio = bioBoxMatch[1].replace(/<[^>]+>/g, "").trim();
    }

    // Build suggested bio based on article type and context
    var location = locality ? locality + (parentArea ? ", " + parentArea : "") : parentArea;
    var suggested = "";

    if (articleType === "Case Study") {
      suggested = "David Allen has spent more than 30 years restoring " + stoneType.toLowerCase() + " floors across the UK through Abbey Floor Care" +
        (location ? ", including projects in " + location : "") +
        ". This case study documents a real restoration visit where on-site assessment shaped every decision. His work focuses on accurate diagnosis, sympathetic repair and preserving the original character of period flooring.";
    } else if (articleType === "Geo Service Page") {
      suggested = "David Allen brings more than 30 years of practical experience restoring " + stoneType.toLowerCase() + " floors across the UK through Abbey Floor Care" +
        (location ? ", with projects completed across " + location + " and the surrounding area" : "") +
        ". His guidance reflects hands-on knowledge of local building stock, period floor conditions and the restoration decisions that produce lasting results.";
    } else if (articleType === "Method Guide") {
      suggested = "David Allen has restored " + stoneType.toLowerCase() + " floors across the UK for over 30 years through Abbey Floor Care. His approach to this method reflects direct experience of how " + stoneType.toLowerCase() + " surfaces respond under real working conditions, where understanding material behaviour shapes every intervention decision.";
    } else if (articleType === "Diagnostic Guide") {
      suggested = "David Allen has diagnosed and corrected " + stoneType.toLowerCase() + " floor conditions across the UK for over 30 years through Abbey Floor Care. His diagnostic approach is grounded in direct observation of how these conditions develop, progress and respond to professional intervention in period and modern properties alike.";
    } else {
      suggested = "David Allen has worked with " + stoneType.toLowerCase() + " floors across the UK for over 30 years through Abbey Floor Care. His practical experience with material behaviour, restoration sequencing and long-term floor care informs every article published under the Abbey Floor Care name.";
    }

    // Bio quality check
    var qualitySignals = [];
    var qualityScore   = 0;

    if (currentBio.toLowerCase().indexOf("david allen") > -1)   { qualityScore++; } else { qualitySignals.push("Missing: David Allen name reference"); }
    if (currentBio.toLowerCase().indexOf("30") > -1 || currentBio.toLowerCase().indexOf("thirty") > -1) { qualityScore++; } else { qualitySignals.push("Missing: Years of experience signal"); }
    if (currentBio.toLowerCase().indexOf("abbey floor care") > -1) { qualityScore++; } else { qualitySignals.push("Missing: Abbey Floor Care brand reference"); }
    if (location && currentBio.toLowerCase().indexOf(location.toLowerCase().split(",")[0]) > -1) { qualityScore++; } else if (location) { qualitySignals.push("Missing: Location reference (" + location + ")"); }
    if (currentBio.length > 100) { qualityScore++; } else { qualitySignals.push("Bio text too short — needs more depth"); }

    var bioQuality = qualityScore >= 4 ? "good" : qualityScore >= 2 ? "weak" : "poor";

    return {
      success:       true,
      currentBio:    currentBio,
      debugBioInfo:  debugBioInfo,
      suggestedBio:  suggested,
      bioQuality:    bioQuality,
      qualityScore:  qualityScore,
      qualitySignals: qualitySignals,
      articleType:   articleType,
      location:      location
    };

  } catch(e) {
    return { success: false, message: "getW7BioData error: " + e.message };
  }
}

function saveW7BioToHtml(newBioText) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const htmlIdx = headers.indexOf("New HTML");
    if (htmlIdx === -1) return { success: false, message: "New HTML column not found." };

    var html = String(sheet.getRange(row, htmlIdx + 1).getValue() || "").trim();
    if (!html) return { success: false, message: "No HTML found in New HTML column." };

    // Build new bio box HTML with corrected bio text
    var newBioBox =
      '\n<div class="abbey-bio-box" style="border:1px solid #ddd;border-radius:8px;padding:16px;background:#f9f9f9;margin:24px 0;display:flex;gap:16px;align-items:flex-start">\n' +
      '  <div style="flex:0 0 80px">\n' +
      '    <img src="https://www.abbeyfloorcare.co.uk/wp-content/uploads/David_Allen.jpg" alt="David Allen, marble and stone restoration specialist" style="width:80px;height:auto;border-radius:6px" />\n' +
      '  </div>\n' +
      '  <div style="flex:1">\n' +
      '    <p style="margin:0 0 6px 0;font-weight:700">David Allen \u2014 <a href="https://maps.app.goo.gl/W8GSsZUiWoxYPQ1Y6" target="_blank" rel="noopener noreferrer">Abbey Floor Care</a></p>\n' +
      '    <p style="margin:0;font-size:0.95em;color:#444">' + newBioText.trim() + '</p>\n' +
      '  </div>\n' +
      '</div>\n';

    // Remove ALL existing bio box instances — find start and end by marker
    var bioStart = html.indexOf('<div class="abbey-bio-box"');
    if (bioStart === -1) bioStart = html.indexOf('<div class=\\"abbey-bio-box\\"');
    if (bioStart === -1) bioStart = html.search(/<div[^>]*abbey-bio-box/i);

    var updatedHtml = html;
    if (bioStart > -1) {
      // Walk forward counting div opens and closes to find the correct closing tag
      var depth = 0;
      var pos   = bioStart;
      var bioEnd = -1;
      while (pos < html.length) {
        var nextOpen  = html.indexOf('<div', pos);
        var nextClose = html.indexOf('</div>', pos);
        if (nextClose === -1) break;
        if (nextOpen !== -1 && nextOpen < nextClose) {
          depth++;
          pos = nextOpen + 4;
        } else {
          depth--;
          pos = nextClose + 6;
          if (depth === 0) { bioEnd = pos; break; }
        }
      }
      if (bioEnd > -1) {
        // Remove everything from bioStart to bioEnd — handles multiple bio boxes
        var before = html.substring(0, bioStart);
        var after  = html.substring(bioEnd);
        // Remove any additional bio box fragments in the after section
        after = after.replace(/<div[^>]*abbey-bio-box[^>]*>[\s\S]*?<\/div>\s*<\/div>/gi, "");
        after = after.replace(/<div style="flex:1">[\s\S]*?<\/div>\s*<\/div>/gi, "");
        updatedHtml = before + newBioBox + after;
      } else {
        updatedHtml = html.replace(/(<footer)/i, newBioBox + '$1');
      }
    } else {
      // No bio box found — insert before footer
      updatedHtml = html.replace(/(<footer)/i, newBioBox + '$1');
    }

    var cell = sheet.getRange(row, htmlIdx + 1);
    cell.setNumberFormat("@");
    cell.setValue(updatedHtml);

    return { success: true, message: "Bio box updated in New HTML column." };

  } catch(e) {
    return { success: false, message: "saveW7BioToHtml error: " + e.message };
  }
}
function finaliseAndLog() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    // Get all W6 data
    var w6Data = ce_getW6Data();
    if (!w6Data.success) return { success: false, message: "Could not read row data: " + w6Data.message };

    // Set published date to today
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    w6Data.newPublishedDate = today;

    // Update publication log
    var logResult = ce_updatePublicationLog(w6Data);
    if (!logResult.success) return { success: false, message: "Publication log update failed: " + logResult.message };

    // Update site export
    var exportResult = ce_updateSiteExport();
    if (!exportResult.success) return { success: false, message: "Site export update failed: " + exportResult.message };

    return { success: true, message: "Publication log and site export updated for Post ID " + w6Data.postId + "." };

  } catch(e) {
    return { success: false, message: "finaliseAndLog error: " + e.message };
  }
}

function extractMissingTerms(recallNotes) {
  if (!recallNotes) return [];
  const notInIdx = recallNotes.indexOf("Not in PQC:");
  if (notInIdx === -1) return [];
  const raw = recallNotes.substring(notInIdx + "Not in PQC:".length).trim();
  return raw.split(",").map(function(t) { return t.trim(); }).filter(function(t) { return t.length > 0; });
}

function arraysMatch(a, b) {
  if (a.length !== b.length) return false;
  const sortedA = a.slice().sort();
  const sortedB = b.slice().sort();
  for (let i = 0; i < sortedA.length; i++) {
    if (sortedA[i] !== sortedB[i]) return false;
  }
  return true;
}

function buildReasonSummary(signals, verdict) {
  const matching = signals.filter(function(s) { return s.verdict === verdict; });
  if (matching.length === 0) return "";
  return matching.map(function(s) { return s.label + ": " + s.detail; }).join(" | ");
}