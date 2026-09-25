/**
 * ================================================================================
 * ce_W8D_Analysis.gs - W8D — POST-REFINEMENT ANALYSIS
 * ================================================================================
 *
 * Runs article-type aware analysis against W8C HTML (col 166) and compares
 * against the original to determine whether the refinement improved the article.
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * ================================================================================
 */

function getW8DAnalysis() {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet = ss.getSheetByName('posts');
    if (!postsSheet) return { success: false, message: 'posts sheet not found.' };

    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return { success: false, message: 'Select a data row first.' };

    var headers = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).trim(); });

    function col(name) {
      var idx = headers.indexOf(name);
      return idx > -1 ? String(postsSheet.getRange(activeRow, idx + 1).getValue() || '').trim() : '';
    }

    var articleType  = col('Article Type');
    var pqc          = col('Primary Query Cluster Owned').toLowerCase();
    var tfidfSummary = col('tfidf_summary');
    var compWords    = parseInt(col('Competitor Word Count')) || 0;

    var originalAnalysis = String(postsSheet.getRange(activeRow, 163).getValue() || '').trim();
    if (!originalAnalysis) return { success: false, message: 'No W8_Analysis found. Run W7 first.' };

    var origHtml = String(postsSheet.getRange(activeRow, 98).getValue() || '').trim();
    if (!origHtml) return { success: false, message: 'No original HTML found in col 98.' };

    var w8cHtml = String(postsSheet.getRange(activeRow, 166).getValue() || '').trim();
    if (!w8cHtml) return { success: false, message: 'No W8C HTML found. Run W8C first.' };

    var originalScore = 0;
    var origScoreMatch = originalAnalysis.match(/Recall Score[:\s]+([0-9.]+)/i);
    if (origScoreMatch) originalScore = parseFloat(origScoreMatch[1]) || 0;

    var originalRec = '';
    var origRecMatch = originalAnalysis.match(/RECOMMENDATION:\s*([A-Z ]+)/i);
    if (origRecMatch) originalRec = origRecMatch[1].trim();

    var origWordCount = origHtml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter(function(w){ return w.length > 0; }).length;
    var w8cWordCount  = w8cHtml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter(function(w){ return w.length > 0; }).length;
    var wordDelta     = w8cWordCount - origWordCount;

    var primaryMeasure = '';
    var beforeValue    = '';
    var afterValue     = '';
    var beforeLabel    = '';
    var afterLabel     = '';
    var scoreImproved  = false;
    var notes          = '';
    var newRec         = '';

    if (articleType === 'Hub Page') {
      primaryMeasure = 'Spoke-signal detection';
      var origSignals = identifyHubSpokeDrift(origHtml);
      var w8cSignals  = identifyHubSpokeDrift(w8cHtml);
      beforeValue     = origSignals.length + ' spoke signal(s)';
      afterValue      = w8cSignals.length + ' spoke signal(s)';
      beforeLabel     = 'Original';
      afterLabel      = 'W8C';
      scoreImproved   = w8cSignals.length < origSignals.length;
      newRec          = w8cSignals.length === 0 ? 'PUBLISH' : w8cSignals.length < origSignals.length ? 'IMPROVED — signals remain' : 'NO IMPROVEMENT';
      if (w8cSignals.length === 0) {
        notes = 'No spoke-level signals detected in W8C HTML.\n\nRECOMMENDATION: Push W8C to New HTML. All spoke-level content has been removed. The hub is operating at summary level.';
      } else if (w8cSignals.length <= 2 && w8cSignals.length < origSignals.length) {
        notes = 'Remaining spoke signals in W8C:\n' + w8cSignals.map(function(s){
          return '• ' + s.id + ': ' + s.issue;
        }).join('\n');
        notes += '\n\nRECOMMENDATION: Push W8C to New HTML. Significant improvement achieved — ' + (origSignals.length - w8cSignals.length) + ' of ' + origSignals.length + ' signals removed. The ' + w8cSignals.length + ' remaining signal(s) are in routing sections where some vocabulary is expected. Pushing is the right call.';
      } else if (w8cSignals.length < origSignals.length) {
        notes = 'Remaining spoke signals in W8C:\n' + w8cSignals.map(function(s){
          return '• ' + s.id + ': ' + s.issue;
        }).join('\n');
        notes += '\n\nRECOMMENDATION: Run W8 again. Improvement has been made (' + (origSignals.length - w8cSignals.length) + ' signals removed) but ' + w8cSignals.length + ' spoke-level signals remain. A second W8 pass targeting the remaining sections is needed before publishing.';
      } else {
        notes = 'Remaining spoke signals in W8C:\n' + w8cSignals.map(function(s){
          return '• ' + s.id + ': ' + s.issue;
        }).join('\n');
        notes += '\n\nRECOMMENDATION: Reject W8C. No improvement detected — the spoke-level signals are unchanged. Keep the original and review the W8B rewrite prompt before running again.';
      }
      notes += '\n\nRecall score (secondary): ' + originalScore + ' — not recalculated for Hub Page.';
    }

    else if (articleType === 'Case Study') {
      primaryMeasure = 'Evidence completeness';
      var EVIDENCE_SIGNALS = [
        { label: 'Location',     pattern: /\b(in|at|near|from)\s+[A-Z][a-z]+/  },
        { label: 'Condition',    pattern: /before|condition|problem|issue|stain|damage|worn|dirty/i },
        { label: 'Intervention', pattern: /clean|restor|seal|treat|repair|strip|polish|hone/i },
        { label: 'Outcome',      pattern: /after|result|finish|complet|now|transform|restor/i }
      ];
      function scoreEvidence(html) {
        var plain = html.replace(/<[^>]+>/g, ' ');
        return EVIDENCE_SIGNALS.filter(function(s){ return s.pattern.test(plain); }).length;
      }
      var origEvidence = scoreEvidence(origHtml);
      var w8cEvidence  = scoreEvidence(w8cHtml);
      beforeValue      = origEvidence + '/4 signals';
      afterValue       = w8cEvidence + '/4 signals';
      beforeLabel      = 'Original';
      afterLabel       = 'W8C';
      scoreImproved    = w8cEvidence >= origEvidence;
      newRec           = w8cEvidence === 4 ? 'PUBLISH' : w8cEvidence >= 3 ? 'BORDERLINE' : 'REVIEW';
      notes            = 'Evidence signals checked: Location, Condition, Intervention, Outcome.';
    }

    else if (
      articleType === 'Educational Guide' ||
      articleType === 'Diagnostic Guide'  ||
      articleType === 'Method Guide'      ||
      articleType === 'Buyer Guide'
    ) {
      primaryMeasure = 'Recall score';
      if (!tfidfSummary || !pqc) {
        return { success: false, message: 'TF-IDF summary or PQC missing — cannot calculate recall score.' };
      }
      var tfidfTerms = tfidfSummary.split(',').map(function(t){ return t.trim().toLowerCase(); }).filter(Boolean);
      var w8cPlain   = w8cHtml.replace(/<[^>]+>/g, ' ').toLowerCase();
      var matches    = 0;
      var unmatched  = [];
      tfidfTerms.forEach(function(term) {
        if (w8cPlain.indexOf(term) !== -1) { matches++; } else { unmatched.push(term); }
      });
      var newScore  = Math.round((matches / tfidfTerms.length) * 100) / 100;
      var threshold = articleType === 'Buyer Guide' ? 0.7 : 0.8;
      beforeValue   = originalScore;
      afterValue    = newScore;
      beforeLabel   = 'Original score';
      afterLabel    = 'W8C score';
      newRec        = newScore >= threshold ? 'PUBLISH' : newScore >= threshold - 0.2 ? 'BORDERLINE' : 'RUN AGAIN';
      notes         = unmatched.length > 0 ? 'Still missing: ' + unmatched.join(', ') : 'All TF-IDF terms present.';

      // Score already at threshold before W8 ran
      if (originalScore >= threshold && newScore >= threshold) {
        scoreImproved = true;
        notes += '\n\nNOTE: Recall score was already at or above threshold before W8 ran (' + originalScore + '). W8 was not needed for recall improvement. Check whether W8 was triggered for a different reason.';
      } else {
        scoreImproved = newScore > originalScore;
      }
    }

    else if (articleType === 'Geo Service Page' || articleType === 'Service Page') {
      primaryMeasure = 'Conversion element completeness';
      function scoreConversion(html) {
        var score = 0;
        if (/contact|quote|enquir|call us/i.test(html))    score++;
        if (/abbeyfloorcare\.co\.uk/i.test(html))          score++;
        if (/service|clean|restor|seal|treat/i.test(html)) score++;
        if (articleType === 'Geo Service Page' && /\b[A-Z][a-z]+\b/.test(html)) score++;
        return score;
      }
      var maxScore  = articleType === 'Geo Service Page' ? 4 : 3;
      var origConv  = scoreConversion(origHtml);
      var w8cConv   = scoreConversion(w8cHtml);
      beforeValue   = origConv + '/' + maxScore + ' elements';
      afterValue    = w8cConv + '/' + maxScore + ' elements';
      beforeLabel   = 'Original';
      afterLabel    = 'W8C';
      scoreImproved = w8cConv >= origConv;
      newRec        = w8cConv === maxScore ? 'PUBLISH' : w8cConv >= maxScore - 1 ? 'BORDERLINE' : 'REVIEW';
      notes         = 'Conversion elements checked: CTA, domain link, service description' + (articleType === 'Geo Service Page' ? ', location binding.' : '.');
    }

    else {
      primaryMeasure = 'Word count';
      beforeValue    = origWordCount + ' words';
      afterValue     = w8cWordCount + ' words';
      beforeLabel    = 'Original';
      afterLabel     = 'W8C';
      scoreImproved  = true;
      newRec         = 'REVIEW';
      notes          = 'No specific measure defined for article type: ' + articleType + '. Review manually.';
    }

    return {
      success:        true,
      articleType:    articleType,
      primaryMeasure: primaryMeasure,
      beforeValue:    beforeValue,
      afterValue:     afterValue,
      beforeLabel:    beforeLabel,
      afterLabel:     afterLabel,
      originalRec:    originalRec,
      newRec:         newRec,
      scoreImproved:  scoreImproved,
      notes:          notes,
      origWordCount:  origWordCount,
      w8cWordCount:   w8cWordCount,
      wordDelta:      wordDelta,
      compWords:      compWords
    };

  } catch(e) {
    return { success: false, message: 'getW8DAnalysis error: ' + e.message };
  }
}

function pushW8CToNewHtml() {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet = ss.getSheetByName('posts');
    if (!postsSheet) return { success: false, message: 'posts sheet not found.' };

    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return { success: false, message: 'Select a data row first.' };

    var w8cHtml = String(postsSheet.getRange(activeRow, 166).getValue() || '').trim();
    if (!w8cHtml) return { success: false, message: 'No W8C HTML found. Run W8C first.' };

    var cell = postsSheet.getRange(activeRow, 98);
    cell.setNumberFormat('@');
    cell.setValue(w8cHtml);

    return { success: true, message: 'W8C HTML pushed to New HTML column — ready to publish.' };

  } catch(e) {
    return { success: false, message: 'pushW8CToNewHtml error: ' + e.message };
  }
}