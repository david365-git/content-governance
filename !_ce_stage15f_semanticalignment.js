/**
 * ============================================================
 * ce_Stage15F_SemanticAlignment.gs
 * Abbey Floor Care — Semantic Alignment Scoring (W1.5F)
 *
 * Uses OpenAI embeddings to score how closely each H2 section's
 * text aligns with the article's target query cluster. This is
 * a leading indicator — available immediately after W4 — rather
 * than waiting weeks/months for GSC data to show whether a
 * rewrite closed the topical gap.
 *
 * Columns used (posts sheet):
 *   173 — Semantic Alignment Score      (overall, 0–1 decimal)
 *   174 — Semantic Alignment Breakdown  (per-H2, pipe-separated)
 *   175 — Semantic Alignment Date       (dd/mm/yyyy)
 *
 * Requires Script Property: OPENAI_API_KEY
 * ============================================================
 */

var SEMANTIC_LOW_THRESHOLD = 0.75;

var SEMANTIC_THRESHOLDS_BY_TYPE = {
  'Case Study':         0.70,
  'Geo Service Page':   0.72,
  'Service Page':       0.73,
  'Method Guide':       0.78,
  'Diagnostic Guide':   0.78,
  'Educational Guide':  0.78,
  'Hub Page':           0.75,
  'Buyer Guide':        0.75
};

// Article types designed for broad, even topic coverage rather than
// concentration around one query cluster. Forcing a fix-loop rewrite
// toward a single phrase actively damages these types — confirmed by
// a real regression on a Buyer Guide (Minton tile restoration cost)
// where the fix pass buried the price and produced tag-dump sentences.
// Scoring still runs for information; the fix-prompt is blocked.
var SEMANTIC_FIX_BLOCKED_TYPES = ['Buyer Guide', 'Method Guide', 'Educational Guide', 'Hub Page'];

function _isFixBlockedForType(articleType) {
  return SEMANTIC_FIX_BLOCKED_TYPES.indexOf(String(articleType || '').trim()) > -1;
}

function _getSemanticThreshold(articleType) {
  var key = String(articleType || '').trim();
  return SEMANTIC_THRESHOLDS_BY_TYPE.hasOwnProperty(key)
    ? SEMANTIC_THRESHOLDS_BY_TYPE[key]
    : SEMANTIC_LOW_THRESHOLD;
}

/* ============================================================
   HELPER — Parse W1.5C Enriched Plan into {sectionNumber: {entities, contentBrief, tsmRequirement}}
   Every section has an "Entities to include:" line plus a Content
   Brief and TSM Requirement — this is the primary scoring target
   per section, since it reflects what that section is actually
   meant to cover (more accurate than forcing every section toward
   one Primary Query Cluster phrase).
============================================================ */
function _parseEnrichedPlanEntities(enrichedPlanText) {
  var result = {};
  if (!enrichedPlanText) return result;

  var sectionBlocks = String(enrichedPlanText).split(/SECTION (\d+):/i);
  for (var i = 1; i < sectionBlocks.length; i += 2) {
    var sectionNum = sectionBlocks[i].trim();
    var block = sectionBlocks[i + 1] || '';

    var entry = { entities: '', contentBrief: '', tsmRequirement: '' };

    var entitiesMatch = block.match(/Entities to include:\s*(.+)/i);
    if (entitiesMatch) {
      var entities = entitiesMatch[1].split('\n')[0].trim();
      if (entities && entities.toLowerCase() !== 'none') {
        entry.entities = entities;
      }
    }

    var briefMatch = block.match(/Content Brief:\s*([\s\S]*?)(?=\n(?:Visual Pattern|List eligible|Images|Videos|Entities to include|Internal link):|$)/i);
    if (briefMatch) {
      entry.contentBrief = briefMatch[1].replace(/\s+/g, ' ').trim();
    }

    var tsmMatch = block.match(/TSM Requirement:\s*(.+)/i);
    if (tsmMatch) {
      entry.tsmRequirement = tsmMatch[1].split('\n')[0].trim();
    }

    var wordBudgetMatch = block.match(/Word Budget:\s*(\d+)[\s\S]{0,10}?(\d+)/i);
    if (wordBudgetMatch) {
      entry.wordBudgetMin = wordBudgetMatch[1];
      entry.wordBudgetMax = wordBudgetMatch[2];
    }

    if (entry.entities || entry.contentBrief) {
      result[sectionNum] = entry;
    }
  }
  return result;
}

function runStage15FSemanticAlignment() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    if (!sheet) throw new Error("'posts' sheet not found.");

    var row = sheet.getActiveRange().getRow();
    if (row < 2) throw new Error("Select a data row in the posts sheet first.");

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Read target query cluster (col 16 — Primary Query Cluster Owned)
    var clusterCol = _findColIndex(headers, 'Primary Query Cluster Owned');
    if (clusterCol === -1) throw new Error("Column 'Primary Query Cluster Owned' not found.");
    var queryCluster = String(sheet.getRange(row, clusterCol + 1).getValue() || '').trim();
    if (!queryCluster) throw new Error("Primary Query Cluster Owned is empty for this row — cannot score without a target.");

    // Strip trailing † governance-derived marker if present
    queryCluster = queryCluster.replace(/\s*†\s*$/, '').trim();

    var htmlCol = _findColIndex(headers, 'New HTML');
    if (htmlCol === -1) throw new Error("Column 'New HTML' not found.");
    // Always New HTML — no fallback to FT. FT has its own dedicated
    // scoring function (runStage15FAlignmentOnFT) so the two can never
    // silently score the same content under different labels.
    var html = String(sheet.getRange(row, htmlCol + 1).getValue() || '').trim();

    // Read Article Type for per-type threshold
    var articleTypeCol = _findColIndex(headers, 'Article Type');
    var articleType = articleTypeCol > -1 ? String(sheet.getRange(row, articleTypeCol + 1).getValue() || '').trim() : '';
    var threshold = _getSemanticThreshold(articleType);

    // Parse scoring units — H3-level where present, H2-level otherwise
    var sections = _extractScoringUnits(html);
    if (sections.length === 0) throw new Error("No H2 sections found in New HTML.");

    // Build embedding inputs: [queryCluster, unit1Text, unit2Text, ...]
    // Every section scores against Primary Query Cluster Owned — this is
    // the page-level ranking-relevance signal needed for recovery work,
    // not section-execution scoring.
    var inputs = [queryCluster];
    sections.forEach(function(s) { inputs.push(s.heading + '. ' + _stripTags(s.rawHtml)); });

    var vectors = _fetchOpenAIEmbeddings(inputs);
    var clusterVector = vectors[0];

    var results = [];
    var totalScore = 0;
    for (var i = 0; i < sections.length; i++) {
      var sectionVector = vectors[i + 1];
      var score = _cosineSimilarity(clusterVector, sectionVector);
      var rounded = Math.round(score * 100) / 100;
      totalScore += rounded;
      results.push({
        label: sections[i].label,
        heading: sections[i].heading,
        score: rounded,
        low: rounded < threshold
      });
    }

    var overallScore = Math.round((totalScore / sections.length) * 100) / 100;

    var breakdownParts = results.map(function(r) {
      return 'H2 ' + r.label + ': ' + r.score.toFixed(2) + (r.low ? ' (LOW)' : '');
    });
    var breakdownText = breakdownParts.join(' | ');

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

    // Write to columns 173, 174, 175
    sheet.getRange(row, 173).setValue(overallScore);
    sheet.getRange(row, 174).setValue(breakdownText);
    sheet.getRange(row, 175).setValue(today);

    var lowCount = results.filter(function(r) { return r.low; }).length;

    return {
      success: true,
      overallScore: overallScore,
      breakdown: breakdownText,
      lowCount: lowCount,
      totalSections: sections.length,
      threshold: threshold,
      articleType: articleType,
      message: 'Overall: ' + overallScore.toFixed(2) +
        ' — ' + sections.length + ' section(s) scored (threshold ' + threshold.toFixed(2) +
        (articleType ? ' for ' + articleType : '') + ')' +
        (lowCount > 0 ? ', ' + lowCount + ' flagged LOW' : ', none flagged')
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}


/**
 * Scores FT (New HTML With Semantic Adjustment) specifically, using the
 * exact same scoring logic as runStage15FSemanticAlignment — but this
 * function ALWAYS reads FT only (never falls back to New HTML) and NEVER
 * writes to columns 173/174/175 or anywhere else. Read-only, display-only,
 * so it can never disrupt the stored New HTML alignment score.
 */
function runStage15FAlignmentOnFT() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    if (!sheet) throw new Error("'posts' sheet not found.");

    var row = sheet.getActiveRange().getRow();
    if (row < 2) throw new Error("Select a data row in the posts sheet first.");

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var clusterCol = _findColIndex(headers, 'Primary Query Cluster Owned');
    if (clusterCol === -1) throw new Error("Column 'Primary Query Cluster Owned' not found.");
    var queryCluster = String(sheet.getRange(row, clusterCol + 1).getValue() || '').trim();
    if (!queryCluster) throw new Error("Primary Query Cluster Owned is empty for this row — cannot score without a target.");
    queryCluster = queryCluster.replace(/\s*†\s*$/, '').trim();

    // Always FT — no fallback to New HTML. If FT is empty, there is
    // nothing to score and we say so explicitly rather than silently
    // scoring New HTML under the FT label.
    var html = String(sheet.getRange(row, 176).getValue() || '').trim();
    if (!html) throw new Error("Column FT (New HTML With Semantic Adjustment) is empty — run Apply Fix first.");

    var articleTypeCol = _findColIndex(headers, 'Article Type');
    var articleType = articleTypeCol > -1 ? String(sheet.getRange(row, articleTypeCol + 1).getValue() || '').trim() : '';
    var threshold = _getSemanticThreshold(articleType);

    var sections = _extractScoringUnits(html);
    if (sections.length === 0) throw new Error("No H2 sections found in FT content.");

    var inputs = [queryCluster];
    sections.forEach(function(s) { inputs.push(s.heading + '. ' + _stripTags(s.rawHtml)); });

    var vectors = _fetchOpenAIEmbeddings(inputs);
    var clusterVector = vectors[0];

    var results = [];
    var totalScore = 0;
    for (var i = 0; i < sections.length; i++) {
      var sectionVector = vectors[i + 1];
      var score = _cosineSimilarity(clusterVector, sectionVector);
      var rounded = Math.round(score * 100) / 100;
      totalScore += rounded;
      results.push({
        label: sections[i].label,
        heading: sections[i].heading,
        score: rounded,
        low: rounded < threshold
      });
    }

    var overallScore = Math.round((totalScore / sections.length) * 100) / 100;
    var breakdownParts = results.map(function(r) {
      return 'H2 ' + r.label + ': ' + r.score.toFixed(2) + (r.low ? ' (LOW)' : '');
    });
    var breakdownText = breakdownParts.join(' | ');
    var lowCount = results.filter(function(r) { return r.low; }).length;

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

    // Dedicated FT columns — 178 (FV) and 179 (FW). Completely separate from
    // 173/174/175 (New HTML's score), so the two can never overwrite or
    // desync from each other.
    sheet.getRange(row, 178).setValue(overallScore);
    sheet.getRange(row, 179).setValue(breakdownText + ' — scored ' + today);

    return {
      success: true,
      overallScore: overallScore,
      breakdown: breakdownText,
      lowCount: lowCount,
      totalSections: sections.length,
      threshold: threshold,
      articleType: articleType,
      message: 'FT Overall: ' + overallScore.toFixed(2) +
        ' — ' + sections.length + ' section(s) scored (threshold ' + threshold.toFixed(2) +
        (articleType ? ' for ' + articleType : '') + ')' +
        (lowCount > 0 ? ', ' + lowCount + ' flagged LOW' : ', none flagged') +
        '. Saved to FT Semantic Alignment Score/Breakdown (col FV/FW) — New HTML\'s score (173/174) was not touched.'
    };

  } catch (e) {
    return { success: false, message: e.message };
  }
}
/* ============================================================
   HELPER — Extract H2 sections and their following text from HTML
   Returns [{ h2: "text", text: "h2 text + following paragraph text" }]
============================================================ */
function _extractH2Sections(html) {
  var sections = [];
  var h2Regex = /<h2[^>]*>([\s\S]*?)<\/h2>/gi;
  var matches = [];
  var m;
  while ((m = h2Regex.exec(html)) !== null) {
    matches.push({ index: m.index, endIndex: h2Regex.lastIndex, h2: _stripTags(m[1]) });
  }

  for (var i = 0; i < matches.length; i++) {
    var start = matches[i].endIndex;
    var searchArea = html.substring(start);

    var nextH2Idx = (i + 1 < matches.length) ? (matches[i + 1].index - start) : -1;
    var sectionCloseMatch = /<\/section>/i.exec(searchArea);
    var sectionCloseIdx = sectionCloseMatch ? sectionCloseMatch.index : -1;

    var end;
    if (nextH2Idx === -1 && sectionCloseIdx === -1) {
      end = html.length;
    } else if (nextH2Idx === -1) {
      end = start + sectionCloseIdx;
    } else if (sectionCloseIdx === -1) {
      end = start + nextH2Idx;
    } else {
      end = start + Math.min(nextH2Idx, sectionCloseIdx);
    }

    var sectionHtml = html.substring(start, end);
    var sectionText = _stripTags(sectionHtml);
    sections.push({
      h2: matches[i].h2,
      text: matches[i].h2 + '. ' + sectionText
    });
  }

  return sections;
}

function _stripTags(html) {
  return String(html)
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

/* ============================================================
   HELPER — Fetch embeddings from OpenAI in a single batched call
   Returns array of vectors in the same order as inputs.
============================================================ */
function _fetchOpenAIEmbeddings(inputs) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error("OPENAI_API_KEY not found in Script Properties.");

  var payload = {
    model: 'text-embedding-3-small',
    input: inputs
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch('https://api.openai.com/v1/embeddings', options);
  var code = response.getResponseCode();
  var body = response.getContentText();

  if (code !== 200) {
    throw new Error('OpenAI Embeddings API error (' + code + '): ' + body);
  }

  var data = JSON.parse(body);
  // data.data is an array of { embedding: [...], index: N } — sort by index to be safe
  var sorted = data.data.slice().sort(function(a, b) { return a.index - b.index; });
  return sorted.map(function(d) { return d.embedding; });
}

/* ============================================================
   HELPER — Cosine similarity between two equal-length vectors
============================================================ */
function _cosineSimilarity(a, b) {
  var dot = 0, normA = 0, normB = 0;
  for (var i = 0; i < a.length; i++) {
    dot += a[i] * b[i];
    normA += a[i] * a[i];
    normB += b[i] * b[i];
  }
  if (normA === 0 || normB === 0) return 0;
  return dot / (Math.sqrt(normA) * Math.sqrt(normB));
}

/* ============================================================
   HELPER — Find column index (0-based) by header name
============================================================ */
function _findColIndex(headers, name) {
  var wanted = String(name).trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim().toLowerCase() === wanted) return i;
  }
  return -1; 
}

/**
 * Build a fix prompt for any H2 sections scoring below threshold.
 * Call after runStage15FSemanticAlignment() has been run for the row.
 */
function buildSemanticAlignmentFixPrompt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var postIdColLow = _findColIndex(headers, 'Post ID');
  var postIdLow = postIdColLow > -1 ? String(sheet.getRange(row, postIdColLow + 1).getValue() || '').trim().replace(/\.0$/, '') : '';
  var originalHtml = '';
  if (postIdLow) {
    var exportSheetLow = ss.getSheetByName('site-export');
    if (exportSheetLow) {
      var exportDataLow = exportSheetLow.getDataRange().getValues();
      var exportHeadersLow = exportDataLow[0].map(function(h) { return String(h).trim(); });
      var exportIdColLow = exportHeadersLow.indexOf('ID');
      var exportHtmlColLow = exportHeadersLow.indexOf('Full Post HTML');
      if (exportIdColLow > -1 && exportHtmlColLow > -1) {
        for (var eiLow = 1; eiLow < exportDataLow.length; eiLow++) {
          if (String(exportDataLow[eiLow][exportIdColLow]).trim().replace(/\.0$/, '') === postIdLow) {
            originalHtml = String(exportDataLow[eiLow][exportHtmlColLow] || '').trim();
            break;
          }
        }
      }
    }
  }
  var articleTypeCol = _findColIndex(headers, 'Article Type');
  var articleType = articleTypeCol > -1 ? String(sheet.getRange(row, articleTypeCol + 1).getValue() || '').trim() : '';
  // Blanket block removed — SECTION JOB context (Content Brief/TSM
  // Requirement per section) now guards against the flattening
  // regression that previously required blocking these article types.

  var clusterCol = _findColIndex(headers, 'Primary Query Cluster Owned');
  var queryCluster = String(sheet.getRange(row, clusterCol + 1).getValue() || '').trim().replace(/\s*†\s*$/, '').trim();

  var htmlCol = _findColIndex(headers, 'New HTML');
  var html = String(sheet.getRange(row, htmlCol + 1).getValue() || '').trim();

  var articleShrinkShortfall = 0;
  var breakdown = String(sheet.getRange(row, 174).getValue() || '');
  var lowMatches = breakdown.match(/H2 ([\d]+[a-z]?): ([\d.]+) \(LOW\)/g) || [];

  if (lowMatches.length === 0) {
    return { success: false, message: 'No LOW-flagged sections found — run the score first, or all sections already pass.' };
  }

  var lowLabels = lowMatches.map(function(m) {
    return m.match(/H2 ([\d]+[a-z]?):/)[1];
  });

  var allUnits = _extractScoringUnits(html);
  var fixTargets = lowLabels.map(function(label) {
    for (var i = 0; i < allUnits.length; i++) {
      if (allUnits[i].label === label) return allUnits[i];
    }
    return null;
  }).filter(Boolean);

  // Pull each section's own Content Brief / TSM Requirement from
  // W1.5C Enriched Plan, so the fix keeps each section doing its
  // actual job (e.g. explaining grout cost) while nudging language
  // toward the target query — rather than flattening every section
  // toward one phrase, which caused a real regression on a Buyer Guide.
  var enrichedPlanCol = _findColIndex(headers, 'W1.5C Enriched Plan');
  var enrichedPlanText = enrichedPlanCol > -1 ? String(sheet.getRange(row, enrichedPlanCol + 1).getValue() || '') : '';
  var sectionBriefs = _parseEnrichedPlanEntities(enrichedPlanText);

  // Total-article ceiling check — the per-section budgets above cap each
  // flagged section individually, but if multiple sections are flagged at
  // once, independent per-section growth can still push the whole article
  // past its overall ceiling. Compute current total + remaining headroom
  // using the same Tier Matrix source of truth as W2B generation.
  var currentTotalWords = _stripTags(html).split(/\s+/).filter(function(w) { return w.length > 0; }).length;
  var articleCeiling = getTierWordCountCeiling(articleType);
  var totalCeilingLine = '';
  if (articleCeiling) {
    var headroom = articleCeiling.maxTarget - currentTotalWords;
    totalCeilingLine =
      'TOTAL ARTICLE LENGTH (Hard Lock — applies across ALL sections being revised together):\n' +
      'This article\'s current total length is approximately ' + currentTotalWords + ' words.\n' +
      'The governed ceiling for this article type is ' + articleCeiling.maxTarget + ' words.\n' +
      (headroom > 0
        ? 'That leaves approximately ' + headroom + ' words of headroom across ALL sections being rewritten in this pass combined — not per section.\n'
        : 'The article is already at or over its ceiling — this rewrite must not add any net length; where one section needs more words to close the alignment gap, cut an equivalent amount elsewhere in the same rewritten section.\n') +
      'If you are revising more than one section in this pass, divide the available headroom between them — do not let each section independently use its own full per-section allowance if that would push the combined total over the ceiling above.\n';
  }

  // Count how many flagged sub-units share the same parent H2 — when an H2's
  // word budget is split across multiple H3 sub-sections, each flagged H3
  // must only claim its fair share, not the full H2 budget repeated on each.
  var parentCounts = {};
  fixTargets.forEach(function(t) {
    parentCounts[t.parentH2Index] = (parentCounts[t.parentH2Index] || 0) + 1;
  });

  fixTargets.forEach(function(t) {
    var sectionNum = String(t.parentH2Index + 1);
    var plan = sectionBriefs[sectionNum];
    t.sectionBrief = plan ? plan.contentBrief : '';
    t.sectionTsm = plan ? plan.tsmRequirement : '';

    var rawMin = plan ? parseInt(plan.wordBudgetMin, 10) : NaN;
    var rawMax = plan ? parseInt(plan.wordBudgetMax, 10) : NaN;
    var siblingCount = parentCounts[t.parentH2Index] || 1;

    if (!isNaN(rawMin) && !isNaN(rawMax) && siblingCount > 1 && t.subIndex !== -1) {
      // This section's total budget is shared across multiple flagged H3
      // sub-units under the same H2 — divide it so the combined total for
      // all flagged siblings does not exceed the section's actual budget.
      t.wordBudgetMin = Math.round(rawMin / siblingCount);
      t.wordBudgetMax = Math.round(rawMax / siblingCount);
      t.wordBudgetShared = true;
      t.wordBudgetSiblingCount = siblingCount;
    } else {
      t.wordBudgetMin = plan ? plan.wordBudgetMin : '';
      t.wordBudgetMax = plan ? plan.wordBudgetMax : '';
      t.wordBudgetShared = false;
    }

    // Current word count of this specific flagged unit — used below to
    // distribute any required article-wide cut proportionally, so larger
    // sections absorb a bigger share of the reduction than small ones.
    t.currentWords = _stripTags(t.rawHtml).split(/\s+/).filter(function(w) { return w.length > 0; }).length;
  });

  // Distribute the article-wide overage (or headroom) across the flagged
  // units proportionally to their current length, so each section gets an
  // explicit numeric target rather than an implicit shared instruction.
  // The resulting target word count is then CLAMPED to that section's own
  // governed word budget range — a proportional share must never override
  // the section's own min/max, since that produces a contradictory prompt.
  var articleCeilingForShrink = getTierWordCountCeiling(articleType);
  if (articleCeilingForShrink) {
    var combinedFlaggedWords = fixTargets.reduce(function(sum, t) { return sum + t.currentWords; }, 0);
    var articleOverage = currentTotalWords - articleCeilingForShrink.maxTarget; // positive = over ceiling
    if (combinedFlaggedWords > 0) {
      var totalClampedReduction = 0;
      var totalRequestedReduction = Math.max(0, articleOverage);

      fixTargets.forEach(function(t) {
        var share = t.currentWords / combinedFlaggedWords;
        var rawDelta = Math.round(articleOverage * share); // positive = cut this many, negative = may add this many
        var rawTarget = t.currentWords - rawDelta;

        var budgetMin = parseInt(t.wordBudgetMin, 10);
        var budgetMax = parseInt(t.wordBudgetMax, 10);
        var hasBudget = !isNaN(budgetMin) && !isNaN(budgetMax);

        var clampedTarget = rawTarget;
        if (hasBudget) {
          if (clampedTarget > budgetMax) clampedTarget = budgetMax;
          if (clampedTarget < budgetMin) clampedTarget = budgetMin;
        }

        var clampedDelta = t.currentWords - clampedTarget;
        t.shrinkTarget = clampedDelta;
        t.shrinkTargetWasClamped = hasBudget && (clampedTarget !== rawTarget);

        if (clampedDelta > 0) totalClampedReduction += clampedDelta;
      });

      // If clamping every section to its own budget still can't reach the
      // full required reduction, flag the shortfall honestly rather than
      // silently under-delivering — this becomes visible in the prompt so
      // David knows a further pass may be needed.
      articleShrinkShortfall = totalRequestedReduction - totalClampedReduction;
    }
  }

  if (articleShrinkShortfall > 0) {
    totalCeilingLine += 'NOTE: Each flagged section below has an individual SPECIFIC CUT TARGET, but those targets are capped to stay within each section\'s own governed word budget — cutting them all fully still leaves this article approximately ' + articleShrinkShortfall + ' words over its ceiling. This is expected: the remaining reduction must come from sections NOT flagged in this pass, or from a further editing pass once this one is complete. Do not attempt to cut beyond a section\'s own budget floor to close this remaining gap.\n';
  }
  totalCeilingLine += '\n';

  var prompt =
    totalCeilingLine +
    'The following section(s) scored below the semantic alignment threshold against the target query cluster: "' + queryCluster + '"\n\n' +
    'Rewrite ONLY the prose paragraph text in these sections so the language more directly addresses the target query cluster — use closely related terms, phrasing, and reader framing a homeowner searching that exact phrase would expect to find, without keyword stuffing or repeating the phrase unnaturally.\n\n' +
    'Open each section with one plain, natural sentence that mirrors how a homeowner would actually search or ask about this, before moving into more technical or specialist phrasing.\n\n' +
    'PRIMARY CONCLUSION FIRST — HARD LOCK (overrides the instruction above where it applies):\n' +
    '- If this section\'s SECTION JOB is to deliver a clear safety verdict, prohibition, or a direct yes/no answer (e.g. "should you use X", "is Y safe", "can this be done"), that verdict MUST be the first sentence — stated plainly and without hedging.\n' +
    '- Do NOT open a safety-verdict section with a question, a conditional ("it depends on..."), or inspection/assessment guidance before the reader has been told the answer. A homeowner asking "can I steam clean travertine?" needs the answer first, not a checklist to work through before finding out.\n' +
    '- Only after the direct verdict is stated may the section explain the reasoning, the conditions that increase risk, or how to inspect for warning signs.\n' +
    '- This rule does NOT apply to purely descriptive or diagnostic sections that have no single clear verdict to give (e.g. "what does a clean floor look like") — for those, the homeowner-question opener above still applies as normal.\n\n' +
    'SECTION JOB PRESERVATION — CRITICAL:\n' +
    '- Each section below has its own specific job, shown as "SECTION JOB" before its HTML. You MUST keep the section doing that job — do not abandon its subject to chase the target query phrase.\n' +
    '- Nudge language toward the target query only where it fits naturally within the section\'s existing job. Do NOT flatten the section into a restatement of the target query.\n' +
    '- If a section\'s job is a specific cost factor, risk category, or technical explanation, that explanation must remain the substance of the section — the target query alignment is a language adjustment, not a subject change.\n' +
    '- Do NOT bury or remove key information (e.g. pricing, specific figures, named steps) to make room for target-query phrasing.\n\n' +
    'DO NOT PLAY IT SAFE — this section has already been rewritten once with only minor wording changes and the alignment score barely moved. A light paraphrase is not good enough. You must substantially rewrite the sentence structure and phrasing throughout, not just swap a few words.\n\n' +
    'STRUCTURAL PRESERVATION (applies ONLY to these specific elements):\n' +
    '- You MUST preserve every <figure>, <img>, <figcaption>, <a href="...">, and <blockquote> element EXACTLY as given — same tags, same attributes, same URLs, same alt text, same position within the section.\n' +
    '- Do NOT remove or relocate any image, link, or blockquote element.\n' +
    '- Subheadings (<h3>) may be reworded if it improves alignment, but must stay wrapped in the same <h3> tag.\n' +
    '- EVERY paragraph of body text MUST remain wrapped in its own <p>...</p> tags. Do NOT output plain unwrapped text — every block of prose needs its own opening <p> and closing </p>.\n' +
    '- If the source content uses a numbered list (<ol><li>...</li></ol>), preserve that exact list structure — every <li> must open and close correctly, and the list must remain nested exactly as given if it appears inside another list.\n\n' +
    'WORDING FREEDOM (applies to everything else):\n' +
    '- Outside of the elements listed above, you have full freedom to rewrite sentence structure, paragraph order, and phrasing.\n' +
    '- Technical or governed terms (e.g. specific product names, material names, or named techniques) may appear anywhere in the rewritten text — earlier, later, restructured around, or paired with plain-language explanation — as long as the term itself is not altered or removed entirely.\n' +
    '- Do not treat any sentence as fixed just because it contains a technical term. Rewrite the sentence around it.\n\n' +
    'NO NEW FACTS OR CLAIMS — HARD LOCK (this rule overrides "wording freedom" wherever the two conflict):\n' +
    '- "Freedom" above means freedom to reword and restructure EXISTING information only. It does NOT mean freedom to invent new information.\n' +
    '- Do NOT introduce any specific fact, cause, symptom, quantity, comparison, homeowner intention, homeowner statement, or attributed quote that is not already present in the original section text below or in the ORIGINAL SOURCE ARTICLE provided later in this prompt.\n' +
    '- This applies even if the invented detail sounds plausible, on-brand, or improves the narrative — plausible is not the same as true. If it is not in the original, it must not appear in the rewrite.\n' +
    '- Never attribute a direct statement or opinion to the homeowner (e.g. "the homeowner said...") unless that exact statement already appears in the original source article.\n' +
    '- You may rephrase an existing fact using different words, or reorder where it appears in the section — but you may not add a new fact alongside it, however small or incidental it may seem.\n' +
    '- Use the ORIGINAL SOURCE ARTICLE to check any claim you are unsure about before including it.\n' +
    '- If you find the target query alignment genuinely requires new supporting detail that is not in the original, do not invent it — instead, express the query-relevant framing using only the facts already present.\n\n' +
    'FACTUAL AUTHORITY ORDER — ABSOLUTE HARD LOCK:\n' +
    '- Use this authority order whenever instructions in this prompt appear to conflict: (1) ORIGINAL SOURCE ARTICLE, (2) original section HTML, (3) NO NEW FACTS OR CLAIMS rule, (4) structural and output rules, (5) SECTION JOB, (6) target-query and semantic-alignment instructions, (7) examples and suggested wording.\n' +
    '- Lower-ranked instructions must never add, strengthen, or imply a fact that is not supported by a higher-ranked source.\n' +
    '- The SECTION JOB describes the editorial purpose of the section. It is not factual evidence. The target query describes the language and reader intent to address. It is not factual evidence.\n' +
    '- If the SECTION JOB, target query, heading, or word budget appears to require a fact or mechanism more specific than the ORIGINAL SOURCE ARTICLE supports, do not try to satisfy that part of the instruction — use the closest factually supported wording instead. Do not mention this conflict in the output.\n' +
    '- CLAIM-SPECIFICITY TEST before including any factual clause: (a) Is this exact fact explicitly stated or directly paraphrasable from the ORIGINAL SOURCE ARTICLE? (b) Have I made the outcome more specific than the source? (c) Have I added a cause, mechanism, measurement, comparison, or absolute result the source does not state? (d) Have I converted an editorial instruction (like SECTION JOB wording) into a project fact? If (a) is no, or any of (b)/(c)/(d) is yes, remove or generalise the clause.\n' +
    '  Example — SUPPORTED: "The floor became far easier to maintain." NOT SUPPORTED: "Dirty mop water no longer collected in the repaired pits" (this invents a mechanism the source never states).\n' +
    '- This claim test applies to ALL editable text, not just paragraph prose — headings, list items, figure captions, link text, summaries, and conclusions are equally subject to it. An unsupported claim is still unsupported if it migrates into a heading or caption instead of a paragraph.\n' +
    '- No wording elsewhere in this prompt becomes factual evidence merely because it appears in an instruction, example, or the target query — examples illustrate STYLE only, never permission to state the facts they contain.\n' +
    '- SPECIFICITY RULE: a rewritten statement must remain at the same or a LOWER level of specificity than the source. A broad supported outcome (e.g. "easier to maintain") must never be sharpened into a specific physical mechanism (e.g. "water no longer pooled") unless the source states that specific mechanism.\n' +
    '- CONFLICT RESOLUTION: if the SECTION JOB requires a fact that is absent from the ORIGINAL SOURCE ARTICLE, silently discard that unsupported part of the SECTION JOB and use the nearest supported outcome instead. Do not attempt to imply the unsupported detail indirectly, and do not mention the conflict in the output.\n' +
    '- PRESERVATION VS FACTUAL CHECKING: this claim test applies to editable text only. Text inside any element marked for exact preservation (figure, img, figcaption, a href, blockquote) must NOT be changed to satisfy this rule, even if it contains wording that would otherwise fail the claim test — do not flag it, do not alter it.\n' +
    '- NEGATIVES AND ABSOLUTES ARE FACTUAL CLAIMS: words such as "prevented," "stopped," "eliminated," "no longer," "without," "always," "only," and "consistently" require the same explicit source support as any other factual claim — they are not safe just because they sound like a harmless process description.\n' +
    '- CROSS-SECTION RELEVANCE: facts may be drawn from anywhere in the ORIGINAL SOURCE ARTICLE, but must remain relevant to this section\'s SECTION JOB and must not be relocated in a way that distorts the article\'s chronological sequence.\n\n' +
    'HEADING VARIETY RULE — CRITICAL:\n' +
    '- If this section contains more than one heading (e.g. multiple <h3> subheadings), each heading MUST use a genuinely different sentence structure and opening pattern from the others.\n' +
    '- Do NOT start multiple headings with the same template, such as repeating "How the..." or "Why the..." across several headings in the same section — this reads as formulaic and repetitive to a human reader, even if each individual heading is well-optimised.\n' +
    '- Vary sentence construction naturally the way a skilled human editor would, so the section reads fluidly as a whole, not like a list of near-identical templates.\n\n' +
    'PARAGRAPH OPENER VARIETY — CRITICAL (applies across ALL sub-sections in this pass, not just within one):\n' +
    '- If you are revising multiple sub-sections that sit under the same parent H2, do NOT open every one of them with a homeowner-style question. Opening every single sub-section with "Which...", "What...", "How..." back-to-back reads as mechanical and repetitive, even though each individual question is well-formed.\n' +
    '- Vary the opening approach across these sub-sections the way a human editor would: one might open with a direct statement, another with a scenario ("If your..."), another with a question, another with a fact. Treat the sequence of openers across the whole set of sub-sections as one continuous piece of writing, not independent fragments each optimised in isolation.\n' +
    '- As a rough guide, no more than roughly one in three consecutive sub-sections should open with an explicit question — the rest should use varied natural openings instead.\n\n' +
    (originalHtml ? ('--- ORIGINAL SOURCE ARTICLE (ground truth — check any claim against this before using it) ---\n' + originalHtml + '\n--- END ORIGINAL SOURCE ARTICLE ---\n\n') : '') +
    'BANNED VOCABULARY — HARD LOCK:\n' +
    '- Use concrete, active verbs describing the actual task (e.g. inspect, clean, remove, extract, fill, seal, hone, protect) rather than abstract corporate language.\n' +
    '- NEVER use these banned verbs or their variants: boost, elevate, evolving, dominate, transform, leverage, utilize, maximize, optimize, enhance, revolutionize.\n' +
    '- WEAK example: "We optimize your floor\'s condition." STRONG example: "We clean the grout, fill the voids, and seal the surface."\n\n' +
    'HARD OUTPUT RULES:\n' +
    '- Return ONLY a JSON array, no markdown fences, no preamble, in this exact format:\n' +
    '[{"index": 0, "new": "the complete rewritten section HTML, including its own heading tag"}]\n' +
    '- Return exactly ' + fixTargets.length + ' array entr' + (fixTargets.length === 1 ? 'y' : 'ies') + ', one per section below, using the same index numbers shown for each section (starting at 0).\n' +
    '- Each "new" value must be the COMPLETE rewritten section as raw HTML — same heading tag level as the original, followed by the full rewritten body.\n' +
    '- Every double quote character inside the "new" string value MUST be escaped as \\" so the result is valid JSON. Use straight quotes only.\n' +
    '- Do NOT include any commentary, preamble, or explanation outside the JSON array.\n\n' +
    'FINAL OUTPUT VERIFICATION — MANDATORY (check the actual generated string, not your intended formatting):\n' +
    '- The response begins with [ and ends with ]. There is exactly one JSON object per section.\n' +
    '- Each "new" value begins literally with the opening heading tag, e.g. <h2> or <h3>, and that heading tag has a matching closing tag.\n' +
    '- Every prose block begins with a literal <p> and ends with a literal </p>. A newline character may appear only between complete HTML elements — it must never stand in place of an opening or closing tag.\n' +
    '  VALID: <h2>Heading</h2>\\n<p>First paragraph.</p>\\n<p>Second paragraph.</p>\n' +
    '  INVALID: Heading\\nFirst paragraph.\\nSecond paragraph.\n' +
    '- No prose appears outside an allowed HTML element. Every preserved figure, image, and figcaption is copied character-for-character.\n' +
    '- Every double quote inside the HTML string is escaped as \\". Every factual claim is directly supported by the ORIGINAL SOURCE ARTICLE, and no phrase from the SECTION JOB has been treated as evidence.\n' +
    '- If any check above fails, correct the output before returning it.\n\n' +
    'This FINAL DECISION ORDER overrides any inconsistent wording elsewhere in this prompt; the detailed rules above define how each step must be applied.\n\n' +
    'FINAL DECISION ORDER — apply in this sequence whenever any rules above appear to conflict:\n' +
    '1. Preserve locked HTML exactly (figure, img, figcaption, a href, blockquote).\n' +
    '2. Include no claim unsupported by the ORIGINAL SOURCE ARTICLE.\n' +
    '3. Keep every editable claim at the same or lower specificity than its source.\n' +
    '4. Silently discard unsupported requirements from the SECTION JOB, target query, or examples.\n' +
    '5. Preserve the section\'s editorial job using only the nearest supported facts.\n' +
    '6. Meet the word budget only where this can be done without repetition or new claims — accuracy overrides the minimum.\n' +
    '7. Validate literal HTML and JSON structure before submission.\n\n' +
    'SECTIONS TO REVISE (raw HTML — preserve all non-paragraph elements exactly):\n\n' +
    fixTargets.map(function(s, idx) {
      var indexLine = 'SECTION INDEX: ' + idx + '\n';
      var jobLine = s.sectionBrief
        ? 'SECTION JOB: ' + s.sectionBrief + (s.sectionTsm ? ' (' + s.sectionTsm + ')' : '') + '\n'
        : '';
      var budgetLine = s.wordBudgetMax
        ? 'WORD BUDGET (Hard Lock): This sub-section\'s share of the governed range is ' + s.wordBudgetMin + '\u2013' + s.wordBudgetMax + ' words' +
          (s.wordBudgetShared ? (' (the full section budget is split evenly across ' + s.wordBudgetSiblingCount + ' flagged sub-sections being revised together in this pass — this figure is already your fair share, not the whole section\'s allowance)') : '') +
          '. Your rewrite must stay within this range — substantial rewriting means changing sentence structure and phrasing, not adding length. If the current sub-section is already near ' + s.wordBudgetMax + ' words, your rewrite must not exceed it. WORD BUDGET VS FACTS (Hard Lock): never add factual detail merely to reach the minimum word count — expand only by explaining, reorganising, or restating already-supported information. If the supported material cannot naturally reach the minimum, factual accuracy overrides the minimum word count; a shorter, fully supported section is correct and preferred over a longer one that pads with unsupported detail.\n'
        : '';

      var shrinkLine = '';
      if (typeof s.shrinkTarget === 'number' && s.shrinkTarget !== 0) {
        var clampNote = s.shrinkTargetWasClamped
          ? ' (this target has been capped to stay within this sub-section\'s own word budget range stated above — do not exceed that range even though the proportional article-wide share would suggest a different number)'
          : '';
        if (s.shrinkTarget > 0) {
          shrinkLine = 'SPECIFIC CUT TARGET (Hard Lock): This sub-section is currently approximately ' + s.currentWords + ' words. As part of bringing the whole article back within its ceiling, this sub-section must come out at approximately ' + Math.max(1, s.currentWords - s.shrinkTarget) + ' words — a reduction of roughly ' + s.shrinkTarget + ' words from its current length' + clampNote + '. This is your specific share of the article-wide cut, calculated from this section\'s proportion of the total length being revised in this pass.\n';
        } else {
          shrinkLine = 'AVAILABLE HEADROOM (informational): This sub-section is currently approximately ' + s.currentWords + ' words. Based on its share of the article\'s remaining headroom, it could grow by up to approximately ' + Math.abs(s.shrinkTarget) + ' words if needed to close the alignment gap — but only use this if genuinely necessary, and never at the expense of the word budget range stated above.\n';
        }
      }
      if (s.subIndex === -1) {
        return indexLine + jobLine + budgetLine + shrinkLine + '<h2>' + s.heading + '</h2>\n' + s.rawHtml + '\n';
      }
      return indexLine + jobLine + budgetLine + shrinkLine + s.rawHtml + '\n'; // H3 sub-units already include their <h3> tag
    }).join('\n') +
    '\nReturn the revised HTML now, following the hard preservation and output rules above exactly. Do NOT include the "SECTION JOB" or "SECTION INDEX" lines in your output — they are context only.';
  return { success: true, prompt: prompt, lowCount: fixTargets.length, labels: lowLabels };
}

/**
 * FT-SPECIFIC TWIN of buildSemanticAlignmentFixPrompt(). Always reads HTML
 * from column 176 (FT) with NO fallback to New HTML, and always reads its
 * LOW-list from column 179 (FT Semantic Alignment Breakdown) — never 174.
 * This guarantees the two fix-prompt buttons can never cross-read each
 * other's data: New HTML's button only ever touches 173/174/98, FT's
 * button only ever touches 178/179/176.
 */
function buildSemanticAlignmentFixPromptForFT() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Always FT — no fallback to New HTML.
    var html = String(sheet.getRange(row, 176).getValue() || '').trim();
    if (!html) return { success: false, message: 'Column FT (New HTML With Semantic Adjustment) is empty — run Apply Fix first.' };

    // Always the FT breakdown column — never 174.
    var breakdownCol179 = String(sheet.getRange(row, 179).getValue() || '');
    var lowMatches = breakdownCol179.match(/H2 ([\d]+[a-z]?): [\d.]+ \(LOW\)/g) || [];
    if (lowMatches.length === 0) {
      return { success: false, message: 'No LOW-flagged sections found in the FT breakdown (col FW) — run "Generate FT Alignment Score" first.' };
    }
    var lowLabels = lowMatches.map(function(m) { return m.match(/H2 ([\d]+[a-z]?):/)[1]; });

    var postIdCol = _findColIndex(headers, 'Post ID');
    var postId = postIdCol > -1 ? String(sheet.getRange(row, postIdCol + 1).getValue() || '').trim().replace(/\.0$/, '') : '';
    var originalHtml = '';
    if (postId) {
      var exportSheet = ss.getSheetByName('site-export');
      if (exportSheet) {
        var exportData = exportSheet.getDataRange().getValues();
        var exportHeaders = exportData[0].map(function(h) { return String(h).trim(); });
        var exportIdCol = exportHeaders.indexOf('ID');
        var exportHtmlCol = exportHeaders.indexOf('Full Post HTML');
        if (exportIdCol > -1 && exportHtmlCol > -1) {
          for (var ei = 1; ei < exportData.length; ei++) {
            if (String(exportData[ei][exportIdCol]).trim().replace(/\.0$/, '') === postId) {
              originalHtml = String(exportData[ei][exportHtmlCol] || '').trim();
              break;
            }
          }
        }
      }
    }
    var clusterCol = _findColIndex(headers, 'Primary Query Cluster Owned');
    var queryCluster = clusterCol > -1
      ? String(sheet.getRange(row, clusterCol + 1).getValue() || '').trim().replace(/\s*†\s*$/, '').trim()
      : '';

    var articleTypeCol = _findColIndex(headers, 'Article Type');
    var articleType = articleTypeCol > -1 ? String(sheet.getRange(row, articleTypeCol + 1).getValue() || '').trim() : '';

    var enrichedPlanCol = _findColIndex(headers, 'W1.5C Enriched Plan');
    var enrichedPlanText = enrichedPlanCol > -1 ? String(sheet.getRange(row, enrichedPlanCol + 1).getValue() || '') : '';
    var sectionBriefs = _parseEnrichedPlanEntities(enrichedPlanText);

    var currentTotalWords = _stripTags(html).split(/\s+/).filter(function(w) { return w.length > 0; }).length;
    var articleCeiling = getTierWordCountCeiling(articleType);
    var totalCeilingLine = '';
    var headroom = 0;
    if (articleCeiling) {
      headroom = articleCeiling.maxTarget - currentTotalWords;
      totalCeilingLine =
        'TOTAL ARTICLE LENGTH (Hard Lock — applies across ALL sections being revised together):\n' +
        'This article\'s current total length (FT draft) is approximately ' + currentTotalWords + ' words.\n' +
        'The governed ceiling for this article type is ' + articleCeiling.maxTarget + ' words.\n' +
        (headroom > 0
          ? 'That leaves approximately ' + headroom + ' words of headroom across ALL sections being rewritten in this pass combined — not per section.\n'
          : 'The article is already at or over its ceiling — this rewrite must not add any net length; where one section needs more words to close the alignment gap, cut an equivalent amount elsewhere in the same rewritten section.\n') +
        'If you are revising more than one section in this pass, divide the available headroom between them — do not let each section independently use its own full per-section allowance if that would push the combined total over the ceiling above.\n';
    }

    var fixTargets = [];
    var allUnits = _extractScoringUnits(html);
    lowLabels.forEach(function(label) {
      for (var j = 0; j < allUnits.length; j++) {
        if (allUnits[j].label === label) { fixTargets.push(allUnits[j]); break; }
      }
    });

    var parentCounts = {};
    fixTargets.forEach(function(t) {
      parentCounts[t.parentH2Index] = (parentCounts[t.parentH2Index] || 0) + 1;
    });

    fixTargets.forEach(function(t) {
      var sectionNum = String(t.parentH2Index + 1);
      var plan = sectionBriefs[sectionNum];
      t.sectionBrief = plan ? plan.contentBrief : '';
      t.sectionTsm = plan ? plan.tsmRequirement : '';

      var rawMin = plan ? parseInt(plan.wordBudgetMin, 10) : NaN;
      var rawMax = plan ? parseInt(plan.wordBudgetMax, 10) : NaN;
      var siblingCount = parentCounts[t.parentH2Index] || 1;

      if (!isNaN(rawMin) && !isNaN(rawMax) && siblingCount > 1 && t.subIndex !== -1) {
        t.wordBudgetMin = Math.round(rawMin / siblingCount);
        t.wordBudgetMax = Math.round(rawMax / siblingCount);
        t.wordBudgetShared = true;
        t.wordBudgetSiblingCount = siblingCount;
      } else {
        t.wordBudgetMin = plan ? plan.wordBudgetMin : '';
        t.wordBudgetMax = plan ? plan.wordBudgetMax : '';
        t.wordBudgetShared = false;
      }

      t.currentWords = _stripTags(t.rawHtml).split(/\s+/).filter(function(w) { return w.length > 0; }).length;
    });

    var articleShrinkShortfall = 0;
    if (articleCeiling) {
      var combinedFlaggedWords = fixTargets.reduce(function(sum, t) { return sum + t.currentWords; }, 0);
      var articleOverage = currentTotalWords - articleCeiling.maxTarget;
      if (combinedFlaggedWords > 0) {
        var totalClampedReduction = 0;
        var totalRequestedReduction = Math.max(0, articleOverage);

        fixTargets.forEach(function(t) {
          var share = t.currentWords / combinedFlaggedWords;
          var rawDelta = Math.round(articleOverage * share);
          var rawTarget = t.currentWords - rawDelta;

          var budgetMin = parseInt(t.wordBudgetMin, 10);
          var budgetMax = parseInt(t.wordBudgetMax, 10);
          var hasBudget = !isNaN(budgetMin) && !isNaN(budgetMax);

          var clampedTarget = rawTarget;
          if (hasBudget) {
            if (clampedTarget > budgetMax) clampedTarget = budgetMax;
            if (clampedTarget < budgetMin) clampedTarget = budgetMin;
          }

          var clampedDelta = t.currentWords - clampedTarget;
          t.shrinkTarget = clampedDelta;
          t.shrinkTargetWasClamped = hasBudget && (clampedTarget !== rawTarget);

          if (clampedDelta > 0) totalClampedReduction += clampedDelta;
        });

        articleShrinkShortfall = totalRequestedReduction - totalClampedReduction;
      }
    }

    if (articleShrinkShortfall > 0) {
      totalCeilingLine += 'NOTE: Each flagged section below has an individual SPECIFIC CUT TARGET, but those targets are capped to stay within each section\'s own governed word budget — cutting them all fully still leaves this article approximately ' + articleShrinkShortfall + ' words over its ceiling. This is expected: the remaining reduction must come from sections NOT flagged in this pass, or from a further editing pass once this one is complete. Do not attempt to cut beyond a section\'s own budget floor to close this remaining gap.\n';
    }
    totalCeilingLine += '\n';

    var prompt =
      totalCeilingLine +
      'The following section(s) scored below the semantic alignment threshold against the target query cluster: "' + queryCluster + '"\n\n' +
      'ROLE: SEO content editor performing a targeted, surgical rewrite.\n' +
      'TASK: Rewrite ONLY the section(s) below to more directly and naturally address the target query cluster, while preserving each section\'s informational job and all governed facts.\n\n' +
      'PRIMARY CONCLUSION FIRST — HARD LOCK:\n' +
      '- If this section\'s SECTION JOB is to deliver a clear safety verdict, prohibition, or a direct yes/no answer (e.g. "should you use X", "is Y safe", "can this be done"), that verdict MUST be the first sentence — stated plainly and without hedging.\n' +
      '- Do NOT open a safety-verdict section with a question, a conditional ("it depends on..."), or inspection/assessment guidance before the reader has been told the answer. A homeowner asking "can I steam clean travertine?" needs the answer first, not a checklist to work through before finding out.\n' +
      '- Only after the direct verdict is stated may the section explain the reasoning, the conditions that increase risk, or how to inspect for warning signs.\n' +
      '- This rule does NOT apply to purely descriptive or diagnostic sections that have no single clear verdict to give (e.g. "what does a clean floor look like") — for those, open naturally as the section\'s job requires.\n\n' +
      'PARAGRAPH OPENER VARIETY — CRITICAL (applies across ALL sub-sections in this pass, not just within one):\n' +
      '- If you are revising multiple sub-sections that sit under the same parent H2, do NOT open every one of them with a homeowner-style question. Opening every single sub-section with "Which...", "What...", "How..." back-to-back reads as mechanical and repetitive, even though each individual question is well-formed.\n' +
      '- Vary the opening approach across these sub-sections the way a human editor would: one might open with a direct statement, another with a scenario ("If your..."), another with a question, another with a fact. Treat the sequence of openers across the whole set of sub-sections as one continuous piece of writing, not independent fragments each optimised in isolation.\n' +
      '- As a rough guide, no more than roughly one in three consecutive sub-sections should open with an explicit question — the rest should use varied natural openings instead.\n\n' +
      'STRUCTURAL PRESERVATION (applies ONLY to these specific elements):\n' +
      '- You MUST preserve every <figure>, <img>, <figcaption>, <a href="...">, and <blockquote> element EXACTLY as given — same tags, same attributes, same URLs, same alt text, same position within the section.\n' +
      '- Do NOT remove or relocate any image, link, or blockquote element.\n' +
      '- Subheadings (<h3>) may be reworded if it improves alignment, but must stay wrapped in the same <h3> tag.\n' +
      '- EVERY paragraph of body text MUST remain wrapped in its own <p>...</p> tags. Do NOT output plain unwrapped text — every block of prose needs its own opening <p> and closing </p>.\n' +
      '- If the source content uses a numbered list (<ol><li>...</li></ol>), preserve that exact list structure — every <li> must open and close correctly, and the list must remain nested exactly as given if it appears inside another list.\n\n' +
      'Do NOT play it safe — substantially rewrite sentence structure, opening lines, and phrasing rather than making minor word swaps.\n\n' +
      'NO NEW FACTS OR CLAIMS — HARD LOCK:\n' +
      '- You have freedom to reword and restructure EXISTING information only. You do NOT have freedom to invent new information.\n' +
      '- Do NOT introduce any specific fact, cause, symptom, quantity, comparison, homeowner intention, homeowner statement, or attributed quote that is not already present in the CURRENT SECTION TEXT below or in the ORIGINAL SOURCE ARTICLE provided at the end of this prompt.\n' +
      '- This applies even if the invented detail sounds plausible or improves the narrative — plausible is not the same as true.\n' +
      '- Never attribute a direct statement or opinion to the homeowner (e.g. "the homeowner said...") unless that exact statement already appears in the original source article.\n' +
      '- Use the ORIGINAL SOURCE ARTICLE at the end of this prompt to check any claim you are unsure about before including it.\n\n' +
      'FACTUAL AUTHORITY ORDER — ABSOLUTE HARD LOCK:\n' +
      '- Use this authority order whenever instructions in this prompt appear to conflict: (1) ORIGINAL SOURCE ARTICLE, (2) original section HTML, (3) NO NEW FACTS OR CLAIMS rule, (4) structural and output rules, (5) SECTION JOB, (6) target-query and semantic-alignment instructions, (7) examples and suggested wording.\n' +
      '- Lower-ranked instructions must never add, strengthen, or imply a fact that is not supported by a higher-ranked source. The SECTION JOB and target query describe editorial purpose and intent — neither is factual evidence.\n' +
      '- If the SECTION JOB, target query, heading, or word budget appears to require a fact or mechanism more specific than the ORIGINAL SOURCE ARTICLE supports, do not try to satisfy that part of the instruction — use the closest factually supported wording instead. Do not mention this conflict in the output.\n' +
      '- CLAIM-SPECIFICITY TEST before including any factual clause: (a) Is this exact fact explicitly stated or directly paraphrasable from the ORIGINAL SOURCE ARTICLE? (b) Have I made the outcome more specific than the source? (c) Have I added a cause, mechanism, measurement, comparison, or absolute result the source does not state? (d) Have I converted an editorial instruction into a project fact? If (a) is no, or any of (b)/(c)/(d) is yes, remove or generalise the clause.\n' +
      '  Example — SUPPORTED: "The floor became far easier to maintain." NOT SUPPORTED: "Dirty mop water no longer collected in the repaired pits" (this invents a mechanism the source never states).\n' +
      '- This claim test applies to ALL editable text, not just paragraph prose — headings, list items, figure captions, link text, summaries, and conclusions are equally subject to it. An unsupported claim is still unsupported if it migrates into a heading or caption instead of a paragraph.\n' +
      '- No wording elsewhere in this prompt becomes factual evidence merely because it appears in an instruction, example, or the target query — examples illustrate STYLE only, never permission to state the facts they contain.\n' +
      '- SPECIFICITY RULE: a rewritten statement must remain at the same or a LOWER level of specificity than the source. A broad supported outcome (e.g. "easier to maintain") must never be sharpened into a specific physical mechanism (e.g. "water no longer pooled") unless the source states that specific mechanism.\n' +
      '- CONFLICT RESOLUTION: if the SECTION JOB requires a fact that is absent from the ORIGINAL SOURCE ARTICLE, silently discard that unsupported part of the SECTION JOB and use the nearest supported outcome instead. Do not attempt to imply the unsupported detail indirectly, and do not mention the conflict in the output.\n' +
      '- PRESERVATION VS FACTUAL CHECKING: this claim test applies to editable text only. Text inside any element marked for exact preservation (figure, img, figcaption, a href, blockquote) must NOT be changed to satisfy this rule, even if it contains wording that would otherwise fail the claim test — do not flag it, do not alter it.\n' +
      '- NEGATIVES AND ABSOLUTES ARE FACTUAL CLAIMS: words such as "prevented," "stopped," "eliminated," "no longer," "without," "always," "only," and "consistently" require the same explicit source support as any other factual claim — they are not safe just because they sound like a harmless process description.\n' +
      '- CROSS-SECTION RELEVANCE: facts may be drawn from anywhere in the ORIGINAL SOURCE ARTICLE, but must remain relevant to this section\'s SECTION JOB and must not be relocated in a way that distorts the article\'s chronological sequence.\n\n' +
      (originalHtml ? ('--- ORIGINAL SOURCE ARTICLE (ground truth — check any claim against this before using it) ---\n' + originalHtml + '\n--- END ORIGINAL SOURCE ARTICLE ---\n\n') : '') +
      'BANNED VOCABULARY — HARD LOCK:\n' +
      '- Use concrete, active verbs describing the actual task (e.g. inspect, clean, remove, extract, fill, seal, hone, protect) rather than abstract corporate language.\n' +
      '- NEVER use these banned verbs or their variants: boost, elevate, evolving, dominate, transform, leverage, utilize, maximize, optimize, enhance, revolutionize.\n' +
      '- WEAK example: "We optimize your floor\'s condition." STRONG example: "We clean the grout, fill the voids, and seal the surface."\n\n' +
      'HARD OUTPUT RULES:\n' +
      '- Return ONLY a JSON array, no markdown fences, no preamble, in this exact format:\n' +
      '[{"index": 0, "new": "the complete rewritten section HTML, including its own heading tag"}]\n' +
      '- Return exactly ' + fixTargets.length + ' array entr' + (fixTargets.length === 1 ? 'y' : 'ies') + ', one per section below, using the same index numbers shown for each section (starting at 0).\n' +
      '- Each "new" value must be the COMPLETE rewritten section as raw HTML — same heading tag level as the original, followed by the full rewritten body.\n' +
      '- Every double quote character inside the "new" string value MUST be escaped as \\" so the result is valid JSON. Use straight quotes only.\n' +
      '- Do NOT include any commentary, preamble, or explanation outside the JSON array.\n\n' +
      'FINAL OUTPUT VERIFICATION — MANDATORY (check the actual generated string, not your intended formatting):\n' +
      '- The response begins with [ and ends with ]. There is exactly one JSON object per section.\n' +
      '- Each "new" value begins literally with the opening heading tag, e.g. <h2> or <h3>, and that heading tag has a matching closing tag.\n' +
      '- Every prose block begins with a literal <p> and ends with a literal </p>. A newline character may appear only between complete HTML elements — it must never stand in place of an opening or closing tag.\n' +
      '  VALID: <h2>Heading</h2>\\n<p>First paragraph.</p>\\n<p>Second paragraph.</p>\n' +
      '  INVALID: Heading\\nFirst paragraph.\\nSecond paragraph.\n' +
      '- No prose appears outside an allowed HTML element. Every preserved figure, image, and figcaption is copied character-for-character.\n' +
      '- Every double quote inside the HTML string is escaped as \\". Every factual claim is directly supported by the ORIGINAL SOURCE ARTICLE, and no phrase from the SECTION JOB has been treated as evidence.\n' +
      '- If any check above fails, correct the output before returning it.\n\n' +
      'This FINAL DECISION ORDER overrides any inconsistent wording elsewhere in this prompt; the detailed rules above define how each step must be applied.\n\n' +
      'FINAL DECISION ORDER — apply in this sequence whenever any rules above appear to conflict:\n' +
      '1. Preserve locked HTML exactly (figure, img, figcaption, a href, blockquote).\n' +
      '2. Include no claim unsupported by the ORIGINAL SOURCE ARTICLE.\n' +
      '3. Keep every editable claim at the same or lower specificity than its source.\n' +
      '4. Silently discard unsupported requirements from the SECTION JOB, target query, or examples.\n' +
      '5. Preserve the section\'s editorial job using only the nearest supported facts.\n' +
      '6. Meet the word budget only where this can be done without repetition or new claims — accuracy overrides the minimum.\n' +
      '7. Validate literal HTML and JSON structure before submission.\n\n' +
      'SECTIONS TO REVISE:\n\n' +
      fixTargets.map(function(s, idx) {
        var jobLine = s.sectionBrief
          ? 'SECTION JOB: ' + s.sectionBrief + (s.sectionTsm ? ' (' + s.sectionTsm + ')' : '') + '\n'
          : '';
        var budgetLine = s.wordBudgetMax
          ? 'WORD BUDGET (Hard Lock): This sub-section\'s share of the governed range is ' + s.wordBudgetMin + '\u2013' + s.wordBudgetMax + ' words' +
            (s.wordBudgetShared ? (' (the full section budget is split evenly across ' + s.wordBudgetSiblingCount + ' flagged sub-sections being revised together in this pass — this figure is already your fair share, not the whole section\'s allowance)') : '') +
            '. Your rewrite must stay within this range — substantial rewriting means changing sentence structure and phrasing, not adding length. If the current sub-section is already near ' + s.wordBudgetMax + ' words, your rewrite must not exceed it. WORD BUDGET VS FACTS (Hard Lock): never add factual detail merely to reach the minimum word count — expand only by explaining, reorganising, or restating already-supported information. If the supported material cannot naturally reach the minimum, factual accuracy overrides the minimum word count; a shorter, fully supported section is correct and preferred over a longer one that pads with unsupported detail.\n'
          : '';
        var shrinkLine = '';
        if (typeof s.shrinkTarget === 'number' && s.shrinkTarget !== 0) {
          var clampNote = s.shrinkTargetWasClamped
            ? ' (this target has been capped to stay within this sub-section\'s own word budget range stated above — do not exceed that range even though the proportional article-wide share would suggest a different number)'
            : '';
          if (s.shrinkTarget > 0) {
            shrinkLine = 'SPECIFIC CUT TARGET (Hard Lock): This sub-section is currently approximately ' + s.currentWords + ' words. As part of bringing the whole article back within its ceiling, this sub-section must come out at approximately ' + Math.max(1, s.currentWords - s.shrinkTarget) + ' words — a reduction of roughly ' + s.shrinkTarget + ' words from its current length' + clampNote + '. This is your specific share of the article-wide cut, calculated from this section\'s proportion of the total length being revised in this pass.\n';
          } else {
            shrinkLine = 'AVAILABLE HEADROOM (informational): This sub-section is currently approximately ' + s.currentWords + ' words. Based on its share of the article\'s remaining headroom, it could grow by up to approximately ' + Math.abs(s.shrinkTarget) + ' words if needed to close the alignment gap — but only use this if genuinely necessary, and never at the expense of the word budget range stated above.\n';
          }
        }
        var indexLine = 'SECTION INDEX: ' + idx + '\n';
        if (s.subIndex === -1) {
          return indexLine + jobLine + budgetLine + shrinkLine + '<h2>' + s.heading + '</h2>\n' + s.rawHtml + '\n';
        }
        return indexLine + jobLine + budgetLine + shrinkLine + s.rawHtml + '\n';
      }).join('\n') +
      '\n\nOUTPUT: Return your answer as the JSON array described above. Do NOT include the "SECTION JOB" or "SECTION INDEX" lines in your output — they are context only.';
    return { success: true, prompt: prompt, sectionCount: fixTargets.length };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/* ============================================================
   HELPER — Extract scoring units. If an H2 contains H3 subsections,
   each H3 becomes its own scoring unit (labelled 2a, 2b, 2c) to
   avoid diluting the embedding across multiple sub-topics. H2s
   with no H3s are scored as one unit, as before.
   Returns [{ label: "2" or "2a", heading: "text", rawHtml: "...",
              parentH2Index: 0-based }]
============================================================ */
function _extractScoringUnits(html) {
  var h2Sections = _extractH2SectionsRaw(html);
  var units = [];

  for (var i = 0; i < h2Sections.length; i++) {
    var sectionHtml = h2Sections[i].rawHtml;
    var h3Regex = /<h3[^>]*>([\s\S]*?)<\/h3>/gi;
    var h3Matches = [];
    var m;
    while ((m = h3Regex.exec(sectionHtml)) !== null) {
      h3Matches.push({ index: m.index, endIndex: h3Regex.lastIndex, h3: _stripTags(m[1]) });
    }

    if (h3Matches.length === 0) {
      // No H3s — score the whole H2 as one unit, as before
      units.push({
        label: String(i + 1),
        heading: h2Sections[i].h2,
        rawHtml: sectionHtml,
        fullOriginal: h2Sections[i].fullRawWithHeading,
        parentH2Index: i,
        subIndex: -1
      });
    } else {
      // Intro text before first H3 is NOT included in any sub-unit's
      // rawHtml — it lives before the <h3> tag, but the splice logic can
      // only match and replace starting FROM a heading tag. Including it
      // in sub-unit "a" caused ChatGPT to rewrite it, and the splice would
      // then insert that rewritten copy after the <h3> while leaving the
      // ORIGINAL intro paragraph untouched in place — producing duplicate
      // intro content. Excluding it here keeps the original intro intact,
      // unduplicated, and simply outside the scope of this fix pass.
      var letters = 'abcdefghijklmnopqrstuvwxyz';

      for (var j = 0; j < h3Matches.length; j++) {
        var start = h3Matches[j].endIndex;
        var end = (j + 1 < h3Matches.length) ? h3Matches[j + 1].index : sectionHtml.length;
        var subHtml = sectionHtml.substring(start, end).trim();
        var fullOriginalSub = sectionHtml.substring(h3Matches[j].index, end).trim();
        units.push({
          label: (i + 1) + letters.charAt(j),
          heading: h2Sections[i].h2 + ' — ' + h3Matches[j].h3,
          rawHtml: '<h3>' + h3Matches[j].h3 + '</h3>\n' + subHtml,
          fullOriginal: fullOriginalSub,
          parentH2Index: i,
          subIndex: j
        });
      }
    }
  }

  return units;
}

/* ============================================================
   HELPER — Extract H2 sections keeping RAW HTML body (not stripped)
   Used when sending content to the LLM so images/links/subheadings
   survive the round-trip.
============================================================ */
function _extractH2SectionsRaw(html) {
  var sections = [];
  // Capture the opening tag's attributes (group 1) separately from the
  // heading text (group 2), so commercial table headings — which always
  // carry inline styling, unlike every governed article H2 — can be
  // identified and excluded from section-boundary detection.
  var h2Regex = /<h2([^>]*)>([\s\S]*?)<\/h2>/gi;
  var allMatches = [];
  var m;
  while ((m = h2Regex.exec(html)) !== null) {
    var isCommercial = /style\s*=/i.test(m[1]);
    allMatches.push({
      index: m.index,
      endIndex: h2Regex.lastIndex,
      h2: _stripTags(m[2]),
      isCommercial: isCommercial
    });
  }

  // Only real governed H2s count as section boundaries — commercial table
  // headings must never shift the section index used for job-brief lookup
  // or semantic scoring.
  var matches = allMatches.filter(function(mm) { return !mm.isCommercial; });

  for (var i = 0; i < matches.length; i++) {
    var start = matches[i].endIndex;
    var searchArea = html.substring(start);

    var nextH2Idx = (i + 1 < matches.length) ? (matches[i + 1].index - start) : -1;
    var sectionCloseMatch = /<\/section>/i.exec(searchArea);
    var sectionCloseIdx = sectionCloseMatch ? sectionCloseMatch.index : -1;

    var end;
    if (nextH2Idx === -1 && sectionCloseIdx === -1) {
      end = html.length;
    } else if (nextH2Idx === -1) {
      end = start + sectionCloseIdx;
    } else if (sectionCloseIdx === -1) {
      end = start + nextH2Idx;
    } else {
      end = start + Math.min(nextH2Idx, sectionCloseIdx);
    }

    var rawSectionHtml = html.substring(start, end).trim();
    var fullRawWithHeading = html.substring(matches[i].index, end).trim();
    // Strip any embedded commercial product table (heading + table) from
    // the scored/returned section text. These are preserved verbatim per
    // a separate rule and must never pollute semantic scoring.
    rawSectionHtml = rawSectionHtml.replace(
      /<h2[^>]*style\s*=[^>]*>[\s\S]*?<\/h2>\s*<table[\s\S]*?<\/table>/gi,
      ''
    ).trim();

    sections.push({
      h2: matches[i].h2,
      rawHtml: rawSectionHtml,
      fullRawWithHeading: fullRawWithHeading
    });
  }

  return sections;
}

/**
 * SHARED SPLICE LOGIC — used by both applySemanticAlignmentFix (writes to
 * FT for the New HTML flow) and applySemanticAlignmentFixForFT (writes to
 * FT for the FT flow). Keeping this in one place guarantees both flows
 * splice identically. Pure function: takes a base HTML string and a
 * ChatGPT response, returns a new HTML string. Never touches the sheet.
 */
function _spliceSemanticFixIntoHtml(baseHtml, responseText, lowLabels) {
  var html = baseHtml;

  var blockRegex = /<h([23])[^>]*>([\s\S]*?)<\/h\1>([\s\S]*?)(?=<h[23][^>]*>|$)/gi;
  var revisedBlocks = [];
  var m;
  while ((m = blockRegex.exec(responseText)) !== null) {
    revisedBlocks.push({
      level: parseInt(m[1], 10),
      heading: _stripTags(m[2]),
      body: m[3].trim()
    });
  }

  if (revisedBlocks.length === 0) {
    return { success: false, message: "Could not find any <h2> or <h3> sections in the pasted response." };
  }
  if (revisedBlocks.length !== lowLabels.length) {
    return {
      success: false,
      message: "Expected " + lowLabels.length + " revised section(s) but found " +
        revisedBlocks.length + " in the pasted response. Paste only the revised sections, in the same order requested."
    };
  }

  var originalUnits = _extractScoringUnits(html);
  var appliedCount = 0;

  Logger.log('SPLICE DEBUG — total lowLabels=' + lowLabels.length + ', total revisedBlocks=' + revisedBlocks.length + ', total originalUnits=' + originalUnits.length);
  Logger.log('SPLICE DEBUG — originalUnits labels: ' + originalUnits.map(function(u) { return u.label; }).join(', '));

  for (var i = 0; i < lowLabels.length; i++) {
    var label = lowLabels[i];
    var original = null;
    for (var j = 0; j < originalUnits.length; j++) {
      if (originalUnits[j].label === label) { original = originalUnits[j]; break; }
    }
    if (!original) {
      Logger.log('SPLICE DEBUG — i=' + i + ', label=' + label + ' — NO ORIGINAL UNIT FOUND, skipping');
      continue;
    }

    var revised = revisedBlocks[i];
    var tag = 'h' + revised.level;

    var originalHeadingText = original.subIndex === -1
      ? original.heading
      : original.heading.split(' — ').pop();

    var stopPattern = (original.subIndex !== -1)
      ? '<h[23][^>]*>|<\\/li>|<\\/ol>|<\\/section>|$'
      : '<h[23][^>]*>|<\\/section>|$';

    var headingRegex = new RegExp(
      '(<' + tag + '[^>]*>\\s*' + _escapeRegex(originalHeadingText) + '\\s*<\\/' + tag + '>)([\\s\\S]*?)(?=' + stopPattern + ')',
      'i'
    );

    var replacement = '<' + tag + '>' + revised.heading + '</' + tag + '>\n' + revised.body + '\n';

    var matchResult = headingRegex.exec(html);
    Logger.log('SPLICE DEBUG — i=' + i + ', label=' + label + ', tag=' + tag + ', subIndex=' + original.subIndex +
      ', originalHeadingText=' + JSON.stringify(originalHeadingText) +
      ', revisedHeading=' + JSON.stringify(revised.heading) +
      ', matched=' + (matchResult !== null) +
      (matchResult ? (', matchStart=' + matchResult.index + ', matchLength=' + matchResult[0].length + ', matchedTextPreview=' + JSON.stringify(matchResult[0].substring(0, 80)) + '...' + JSON.stringify(matchResult[0].substring(matchResult[0].length - 80))) : ''));

    if (matchResult) {
      html = html.replace(headingRegex, replacement);
      appliedCount++;
      Logger.log('SPLICE DEBUG — i=' + i + ', label=' + label + ' — REPLACED. html length now=' + html.length);
    }
  }

  Logger.log('SPLICE DEBUG — FINAL appliedCount=' + appliedCount + ' of expected=' + lowLabels.length);

  return { success: true, html: html, appliedCount: appliedCount, expectedCount: lowLabels.length };
}

/**
 * Parse ChatGPT's revised sections response and replace the corresponding
 * H2 sections in the New HTML column. Expects sections labelled by H2
 * matching the format used in buildSemanticAlignmentFixPrompt().
 */
function applySemanticAlignmentFix(responseText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var htmlCol = _findColIndex(headers, 'New HTML');
    if (htmlCol === -1) throw new Error("Column 'New HTML' not found.");
    var html = String(sheet.getRange(row, htmlCol + 1).getValue() || '');
    var breakdownCol174 = String(sheet.getRange(row, 174).getValue() || '');
    var lowMatches = breakdownCol174.match(/H2 ([\d]+[a-z]?): [\d.]+ \(LOW\)/g) || [];
    if (lowMatches.length === 0) {
      throw new Error("No LOW-flagged sections found in the stored breakdown — run the score again first.");
    }
    var lowLabels = lowMatches.map(function(m) {
      return m.match(/H2 ([\d]+[a-z]?):/)[1];
    });
    var allUnits = _extractScoringUnits(html);
    var fixTargets = lowLabels.map(function(label) {
      for (var i = 0; i < allUnits.length; i++) {
        if (allUnits[i].label === label) return allUnits[i];
      }
      return null;
    }).filter(Boolean);
    if (fixTargets.length === 0) {
      throw new Error("Could not re-locate any of the LOW-flagged sections in the current HTML — the article may have changed since the score was run. Re-run Generate Alignment Score first.");
    }
    if (!responseText || !responseText.trim()) {
      throw new Error("No response provided.");
    }
    var cleaned = responseText.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
    var fixes;
    try {
      fixes = JSON.parse(cleaned);
    } catch (e) {
      try {
        var repaired = cleaned.replace(/"(new)":\s*"([\s\S]*?)"(?=\s*[,}])/g, function(match, key, value) {
          var fixedValue = value.replace(/(?<!\\)"/g, '\\"');
          return '"' + key + '": "' + fixedValue + '"';
        });
        fixes = JSON.parse(repaired);
      } catch (e2) {
        throw new Error("Could not parse response as JSON: " + e.message + " (repair attempt also failed: " + e2.message + ")");
      }
    }
    if (!Array.isArray(fixes) || fixes.length === 0) {
      throw new Error("No fixes found in parsed response.");
    }
    var brokenLinkPattern = /<a\s+href="\[https?:\/\/[^\]]+\]\(/i;
    var applied = [];
    var failed = [];
    fixes.forEach(function(fix) {
      var idx = fix.index;
      var label = 'Index ' + idx;
      if (typeof idx !== 'number' || idx < 0 || idx >= fixTargets.length) {
        failed.push(label + ' — invalid or out-of-range index');
        return;
      }
      var target = fixTargets[idx];
      var oldText = target.fullOriginal;
      var newText = fix['new'];
      if (!newText) { failed.push(label + ' — no "new" text provided'); return; }
      var structureError = _validateSemanticFixSectionHtml(newText);
      if (structureError) {
        failed.push(label + ' (section ' + target.label + ') — REJECTED (structure): ' + structureError);
        return;
      }
      if (brokenLinkPattern.test(newText)) {
        failed.push(label + ' — REJECTED: malformed markdown-in-HTML link syntax in "new" text');
        return;
      }
      var occurrences = html.split(oldText).length - 1;
      if (occurrences === 0) {
        failed.push(label + ' (section ' + target.label + ') — original section text not found in current HTML, skipped for safety');
        return;
      }
      if (occurrences > 1) {
        failed.push(label + ' (section ' + target.label + ') — original section text matches ' + occurrences + ' places, skipped for safety');
        return;
      }
      html = html.replace(oldText, newText);
      applied.push(label + ' (section ' + target.label + ')');
    });
    sheet.getRange(row, 176).setValue(html);
    SpreadsheetApp.flush();
    var verifyRead = String(sheet.getRange(row, 176).getValue() || '');
    var incomplete = applied.length < fixTargets.length;
    var message = (incomplete ? '⚠ PARTIAL APPLY — ' : '✔ ') +
      applied.length + ' of ' + fixTargets.length + ' section(s) updated — saved to New HTML With Semantic Adjustment (col FT). Read-back length: ' + verifyRead.length + '.';
    if (applied.length > 0) message += '\nApplied: ' + applied.join(', ');
    if (failed.length > 0) message += '\nSkipped: ' + failed.join(' | ');
    message += incomplete ? '' : ' Review, then click "Copy to New HTML" when ready.';
    return {
      success: true,
      incomplete: incomplete,
      message: message,
      appliedCount: applied.length
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * FT-SPECIFIC TWIN of applySemanticAlignmentFix(). Reads its LOW-list from
 * column 179 (FT Semantic Alignment Breakdown) instead of 174, and always
 * splices against FT's own current content (176) as the base — never New
 * HTML. Still writes the result back to FT (176), same as before.
 */
function applySemanticAlignmentFixForFT(responseText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    var html = String(sheet.getRange(row, 176).getValue() || '').trim();
    if (!html) throw new Error("Column FT (New HTML With Semantic Adjustment) is empty — nothing to apply the fix to.");
    var breakdownCol179 = String(sheet.getRange(row, 179).getValue() || '');
    var lowMatches = breakdownCol179.match(/H2 ([\d]+[a-z]?): [\d.]+ \(LOW\)/g) || [];
    if (lowMatches.length === 0) {
      throw new Error("No LOW-flagged sections found in the FT breakdown (col FW) — run 'Generate FT Alignment Score' again first.");
    }
    var lowLabels = lowMatches.map(function(m) {
      return m.match(/H2 ([\d]+[a-z]?):/)[1];
    });
    var allUnits = _extractScoringUnits(html);
    var fixTargets = lowLabels.map(function(label) {
      for (var i = 0; i < allUnits.length; i++) {
        if (allUnits[i].label === label) return allUnits[i];
      }
      return null;
    }).filter(Boolean);
    if (fixTargets.length === 0) {
      throw new Error("Could not re-locate any of the LOW-flagged sections in the current FT HTML — it may have changed since the score was run. Re-run Generate FT Alignment Score first.");
    }
    if (!responseText || !responseText.trim()) {
      throw new Error("No response provided.");
    }
    var cleaned = responseText.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
    var fixes;
    try {
      fixes = JSON.parse(cleaned);
    } catch (e) {
      try {
        var repaired = cleaned.replace(/"(new)":\s*"([\s\S]*?)"(?=\s*[,}])/g, function(match, key, value) {
          var fixedValue = value.replace(/(?<!\\)"/g, '\\"');
          return '"' + key + '": "' + fixedValue + '"';
        });
        fixes = JSON.parse(repaired);
      } catch (e2) {
        throw new Error("Could not parse response as JSON: " + e.message + " (repair attempt also failed: " + e2.message + ")");
      }
    }
    if (!Array.isArray(fixes) || fixes.length === 0) {
      throw new Error("No fixes found in parsed response.");
    }
    var brokenLinkPattern = /<a\s+href="\[https?:\/\/[^\]]+\]\(/i;
    var applied = [];
    var failed = [];
    fixes.forEach(function(fix) {
      var idx = fix.index;
      var label = 'Index ' + idx;
      if (typeof idx !== 'number' || idx < 0 || idx >= fixTargets.length) {
        failed.push(label + ' — invalid or out-of-range index');
        return;
      }
      var target = fixTargets[idx];
      var oldText = target.fullOriginal;
      var newText = fix['new'];
      if (!newText) { failed.push(label + ' — no "new" text provided'); return; }
      var structureError = _validateSemanticFixSectionHtml(newText);
      if (structureError) {
        failed.push(label + ' (section ' + target.label + ') — REJECTED (structure): ' + structureError);
        return;
      }
      if (brokenLinkPattern.test(newText)) {
        failed.push(label + ' — REJECTED: malformed markdown-in-HTML link syntax in "new" text');
        return;
      }
      var occurrences = html.split(oldText).length - 1;
      if (occurrences === 0) {
        failed.push(label + ' (section ' + target.label + ') — original section text not found in current FT HTML, skipped for safety');
        return;
      }
      if (occurrences > 1) {
        failed.push(label + ' (section ' + target.label + ') — original section text matches ' + occurrences + ' places, skipped for safety');
        return;
      }
      html = html.replace(oldText, newText);
      applied.push(label + ' (section ' + target.label + ')');
    });
    sheet.getRange(row, 176).setValue(html);
    SpreadsheetApp.flush();
    var incomplete = applied.length < fixTargets.length;
    var message = (incomplete ? '⚠ PARTIAL APPLY — ' : '✔ ') +
      applied.length + ' of ' + fixTargets.length + ' section(s) updated — saved back to FT.';
    if (applied.length > 0) message += '\nApplied: ' + applied.join(', ');
    if (failed.length > 0) message += '\nSkipped: ' + failed.join(' | ');
    message += incomplete ? '' : ' Re-run "Generate FT Alignment Score" to confirm improvement.';
    return {
      success: true,
      incomplete: incomplete,
      message: message
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Copy the reviewed content from New HTML With Semantic Adjustment (col 176)
 * into New HTML (the live column), once the user has confirmed the changes.
 */
function copySemanticAdjustmentToNewHtml() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    if (row < 2) throw new Error("Select a data row first — row 1 is the header row.");
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var htmlCol = _findColIndex(headers, 'New HTML');
    if (htmlCol === -1) throw new Error("Column 'New HTML' not found.");

    var adjusted = String(sheet.getRange(row, 176).getValue() || '').trim();
    if (!adjusted) throw new Error("New HTML With Semantic Adjustment (col FT) is empty — run a fix first.");

    sheet.getRange(row, htmlCol + 1).setValue(adjusted);
    sheet.getRange(row, 176).setValue('');
    SpreadsheetApp.flush();

    return { success: true, message: 'Copied to New HTML and cleared the adjustment column — ready for a fresh pass.' };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function _extractH2SectionRawHtml(html, index) {
  var h2Regex = /<h2[^>]*>([\s\S]*?)<\/h2>/gi;
  var matches = [];
  var m;
  while ((m = h2Regex.exec(html)) !== null) {
    matches.push({ index: m.index, endIndex: h2Regex.lastIndex });
  }
  if (index >= matches.length) return '';
  var start = matches[index].endIndex;
  var end = (index + 1 < matches.length) ? matches[index + 1].index : html.length;
  return html.substring(start, end).trim();
}

function _escapeRegex(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Validates that a rewritten section's "new" HTML uses literal heading and
 * paragraph tags rather than plain-text/newline separators. Returns null
 * if valid, or a pastable instruction string describing the exact defect
 * if invalid — so the failure message can be copied straight back to the
 * LLM as a correction request rather than a vague "redo it."
 */
function _validateSemanticFixSectionHtml(html) {
  if (typeof html !== 'string' || !html.trim()) {
    return 'Your last response is missing the "new" HTML string entirely. Return the corrected JSON with the complete rewritten section as a literal HTML string.';
  }
  var trimmed = html.trim();
  if (!/^<h[23][^>]*>/i.test(trimmed)) {
    return 'Your last response failed this check: the section does not begin with a literal <h2> or <h3> tag — it appears to use plain text instead. Return the corrected JSON now, with the "new" value starting exactly with the opening heading tag (e.g. <h2>Heading text</h2>), and every paragraph wrapped in its own <p>...</p> tags. Do not use \\n line breaks in place of HTML tags.';
  }
  if (!/<\/h[23]>/i.test(trimmed)) {
    return 'Your last response failed this check: the heading tag is missing its closing tag (</h2> or </h3>). Return the corrected JSON now with a literal closing heading tag.';
  }
  var bodyAfterHeading = trimmed.replace(/^<h[23][^>]*>[\s\S]*?<\/h[23]>/i, '');
  var hasParaTags = /<p[^>]*>[\s\S]*?<\/p>/i.test(bodyAfterHeading);
  if (!hasParaTags) {
    return 'Your last response failed this check: no literal <p>...</p> paragraph tags were found in the section body — it appears to use plain text with \\n line breaks instead. Return the corrected JSON now with every paragraph of prose wrapped in its own opening <p> and closing </p> tag.';
  }
  return null;
}

/**
 * Build a comparison prompt for Gemini to judge Current New HTML vs the
 * FT (New HTML With Semantic Adjustment) draft — both for search/alignment
 * quality and for structural HTML integrity (broken tags, missing <p>
 * wrapping, collapsed lists, dropped images).
 * Call this before deciding whether to copy FT into New HTML.
 */
function buildSemanticComparisonPrompt() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var htmlCol = _findColIndex(headers, 'New HTML');
    if (htmlCol === -1) return { success: false, message: "Column 'New HTML' not found." };

    var currentHtml = String(sheet.getRange(row, htmlCol + 1).getValue() || '').trim();
    if (!currentHtml) return { success: false, message: 'New HTML is empty — nothing to compare against.' };

    // Revision B is built the SAME WAY Apply Fix would build it: FT if it
    // already holds a full spliced draft (from a previous Apply Fix run),
    // otherwise a fresh splice of the pasted ChatGPT response directly
    // into live New HTML — in memory only, never written to the sheet.
    // This guarantees Gemini always sees the true fully-spliced document
    // (header/footer/sections/bio box all intact), never a bare fragment.
    var ftExisting = String(sheet.getRange(row, 176).getValue() || '').trim();
    var ftHtml;
    if (ftExisting) {
      ftHtml = ftExisting;
    } else {
      var responseBox = String(sheet.getRange(row, 176).getValue() || ''); // placeholder, overwritten below if unused
      return {
        success: false,
        message: 'Column FT is empty. Run Apply Fix first so a spliced draft exists in FT, then compare.'
      };
    }
    var postIdColCmp = _findColIndex(headers, 'Post ID');
    var postIdCmp = postIdColCmp > -1 ? String(sheet.getRange(row, postIdColCmp + 1).getValue() || '').trim().replace(/\.0$/, '') : '';
    var originalHtml = '';
    if (postIdCmp) {
      var exportSheetCmp = ss.getSheetByName('site-export');
      if (exportSheetCmp) {
        var exportDataCmp = exportSheetCmp.getDataRange().getValues();
        var exportHeadersCmp = exportDataCmp[0].map(function(h) { return String(h).trim(); });
        var exportIdColCmp = exportHeadersCmp.indexOf('ID');
        var exportHtmlColCmp = exportHeadersCmp.indexOf('Full Post HTML');
        if (exportIdColCmp > -1 && exportHtmlColCmp > -1) {
          for (var eic = 1; eic < exportDataCmp.length; eic++) {
            if (String(exportDataCmp[eic][exportIdColCmp]).trim().replace(/\.0$/, '') === postIdCmp) {
              originalHtml = String(exportDataCmp[eic][exportHtmlColCmp] || '').trim();
              break;
            }
          }
        }
      }
    }

    var clusterCol = _findColIndex(headers, 'Primary Query Cluster Owned');
    var queryCluster = clusterCol > -1
      ? String(sheet.getRange(row, clusterCol + 1).getValue() || '').trim().replace(/\s*†\s*$/, '').trim()
      : '';

    var articleTypeCol = _findColIndex(headers, 'Article Type');
    var articleType = articleTypeCol > -1 ? String(sheet.getRange(row, articleTypeCol + 1).getValue() || '').trim() : '';

    var materialCol = _findColIndex(headers, 'Stone Type');
    var material = materialCol > -1 ? String(sheet.getRange(row, materialCol + 1).getValue() || '').trim() : '';

    function wordCount(html) {
      var text = String(html).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
      return text.length > 0 ? text.split(' ').length : 0;
    }

    var currentWords = wordCount(currentHtml);
    var ftWords = wordCount(ftHtml);

    var prompt =
      'ROLE: Senior UK SEO Editor and HTML Quality Reviewer\n\n' +
      'TASK: Compare REVISION A (Current New HTML) against REVISION B (Semantic Fix Draft) for the same article, and give a clear recommendation.\n\n' +
      '--- PAGE CONTEXT ---\n' +
      'Material: ' + (material || 'Unknown') + '\n' +
      'Article Type: ' + (articleType || 'Unknown') + '\n' +
      'Target Query Cluster: ' + (queryCluster || 'Not specified') + '\n' +
      'Revision A word count (approx): ' + currentWords + '\n' +
      'Revision B word count (approx): ' + ftWords + '\n\n' +
      '--- WHAT REVISION B WAS ATTEMPTING ---\n' +
      'Revision B is a targeted rewrite of specific sections that scored below a semantic alignment threshold against the Target Query Cluster above. The rewrite aimed to more directly address that query while preserving each section\'s original informational job, and to bring the total article length back within its governed word budget.\n\n' +
      '--- WHAT TO EVALUATE ---\n' +
      '1. SEARCH RELEVANCE: Which revision more clearly and naturally addresses the Target Query Cluster? Look for genuine improvement in how directly the content answers what a homeowner searching that phrase would want — not just superficial keyword presence.\n' +
      '2. INFORMATION INTEGRITY: Check BOTH directions. (a) LOSS — did Revision B preserve the substance of each section (specific facts, named steps, technical detail, safety warnings) or did it thin out/generalise content while chasing the target phrase? (b) ADDITIONS vs Revision A — list every specific factual claim, named detail, direct quote, or attributed statement that appears in Revision B but is NOT present in Revision A, regardless of how plausible or well-written it seems. (c) TRUTH CHECK vs ORIGINAL SOURCE ARTICLE (provided below, if available) — for EVERY specific factual claim, quote, or attributed statement in BOTH Revision A and Revision B (not just new additions), check whether it is actually supported by the original source article. Flag anything in EITHER revision that is not supported by the original source, even if it has been present since Revision A and was never previously flagged — do not assume Revision A is automatically true just because it is the current live version.\n' +
      '3. READABILITY AND TONE: Is Revision B\'s prose natural and varied, or does it read as mechanically adjusted (repetitive openers, awkward phrasing, forced keyword insertion)?\n' +
      '4. STRUCTURAL HTML INTEGRITY (CRITICAL — check carefully): Scan Revision B for HTML defects that would break the page if published, specifically:\n' +
      '   - Every <li> has a matching closing </li>\n' +
      '   - Every paragraph of body text is wrapped in <p>...</p> tags — not left as bare unwrapped text\n' +
      '   - Nested <ol>/<li> structures are not collapsed or malformed\n' +
      '   - No <figure>, <img>, <figcaption>, <a href>, or <blockquote> elements have been dropped, altered, or duplicated compared to Revision A\n' +
      '   - No heading level has been changed (an <h3> must not have become an <h2> or vice versa)\n' +
      '   Flag ANY structural defect found, quote the exact broken fragment, and state which revision it appears in.\n' +
      '5. LENGTH: Is Revision B\'s length appropriate, or has it overshot/undershot to a degree that harms the page?\n\n' +
      '--- OUTPUT FORMAT (REQUIRED) ---\n' +
      'STRUCTURAL DEFECTS FOUND: [list any HTML integrity problems found in either revision, quoting the broken fragment — or write "None found" if the HTML is clean]\n' +
      'SEARCH RELEVANCE: [Revision A / Revision B / Roughly equal] — [one sentence why]\n' +
      'INFORMATION INTEGRITY: [Revision A / Revision B / Roughly equal] — [one sentence why]\n' +
      'UNVERIFIED ADDITIONS IN REVISION B: [list every specific claim, quote, or detail present in Revision B but absent from Revision A, quoting the exact text — or write "None found" if there are none]\n' +
      'READABILITY: [Revision A / Revision B / Roughly equal] — [one sentence why]\n' +
      'OVERALL RECOMMENDATION: [Use Revision A / Use Revision B / Neither — needs further work] — [2-3 sentence justification]\n' +
      'IF STRUCTURAL DEFECTS WERE FOUND: state explicitly whether they must be fixed before Revision B could be safely published, regardless of the SEO verdict above.\n\n' +
      (originalHtml ? ('--- ORIGINAL SOURCE ARTICLE (ground truth — use this to verify factual claims in BOTH revisions) ---\n' + originalHtml + '\n\n') : '') +
      '--- REVISION A (Current New HTML) ---\n' +
      currentHtml + '\n\n' +
      '--- REVISION B (Semantic Fix Draft) ---\n' +
      ftHtml + '\n\n' +
      '--- GENERATE NOW ---';

    return { success: true, prompt: prompt, currentWords: currentWords, ftWords: ftWords };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Saves the pasted Gemini comparison response verbatim to column 180 (FX —
 * Gemini Comparison Result). Simple archival save — no parsing, no
 * validation — the sidebar's displayComparisonResult() already handles
 * extracting/formatting the fields for on-screen display.
 */
function saveGeminiComparisonResult(responseText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    var text = String(responseText || '').trim();
    if (!text) return { success: false, message: 'No response text to save.' };

    sheet.getRange(row, 180).setValue(text);
    SpreadsheetApp.flush();

    return { success: true, message: 'Saved to Gemini Comparison Result (col FX).' };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

function buildFTUnverifiedAdditionsFixPrompt(findingsText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('posts');
  var row = sheet.getActiveRange().getRow();
  var html = String(sheet.getRange(row, 176).getValue() || '').trim();
  if (!html) {
    return { success: false, message: "ERROR: Column FT is empty for this row." };
  }
  if (!findingsText || !findingsText.trim()) {
    return { success: false, message: "ERROR: No unverified additions findings provided." };
  }
  var prompt =
    'You are removing specific unverified additions from an article, identified by comparing it against the original source project.\n\n' +
    'UNVERIFIED ADDITIONS TO ADDRESS (only fix items listed below — do not touch anything else):\n' +
    findingsText + '\n\n' +
    'RULES:\n' +
    '1. Only rewrite the specific fragments containing the flagged additions above. Leave everything else in the article completely unchanged.\n' +
    '2. Remove the unsupported claim, detail, or attributed statement so the sentence accurately reflects only what the original source project supports. Do NOT replace one unverified claim with a different invented detail — if removing it leaves less to say, that is the correct and safe outcome.\n' +
    '3. Keep the surrounding sentence and paragraph natural and grammatically correct after the removal. Do not leave dangling connectives ("However," "In addition,") that no longer make sense once the claim is removed.\n' +
    '4. Do NOT rewrite, wrap, or alter any <a href="..."> link, <figure>, <img>, or <blockquote> element unless the finding explicitly flags that exact element\'s text. Never convert HTML links into markdown-style syntax.\n' +
    '5. Return ONLY a JSON array, no markdown fences, no preamble, in this exact format:\n' +
    '[{"fixLabel": "short label for this fix", "old": "exact original fragment from the article below, including the unverified addition", "new": "the same fragment with the unverified addition removed or corrected"}]\n' +
    '6. Every double quote character inside the old/new string values MUST be escaped as \\" so the result is valid JSON. Use straight quotes only.\n' +
    '7. The "old" fragment must be copied verbatim from the article HTML below — exact characters — so it can be located by an automated find-and-replace. Include enough surrounding text to make each "old" fragment unique within the article.\n\n' +
    'ARTICLE HTML (column FT):\n' +
    html;
  return { success: true, prompt: prompt };
}

function applyFTUnverifiedAdditionsFix(rawJson) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    if (!rawJson || !rawJson.trim()) {
      return { success: false, message: "No response provided." };
    }
    var cleaned = rawJson.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
    var fixes;
    try {
      fixes = JSON.parse(cleaned);
    } catch (e) {
      try {
        var repaired = cleaned.replace(/"(old|new|fixLabel)":\s*"([\s\S]*?)"(?=\s*[,}])/g, function(match, key, value) {
          var fixedValue = value.replace(/(?<!\\)"/g, '\\"');
          return '"' + key + '": "' + fixedValue + '"';
        });
        fixes = JSON.parse(repaired);
      } catch (e2) {
        return { success: false, message: "Could not parse response as JSON: " + e.message + " (repair attempt also failed: " + e2.message + ")" };
      }
    }
    if (!Array.isArray(fixes) || fixes.length === 0) {
      return { success: false, message: "No fixes found in parsed response." };
    }
    var html = String(sheet.getRange(row, 176).getValue() || '');
    if (!html) {
      return { success: false, message: "Column FT is empty for this row." };
    }
    var applied = [];
    var failed = [];
    var brokenLinkPattern = /<a\s+href="\[https?:\/\/[^\]]+\]\(/i;
    fixes.forEach(function(fix) {
      var label = fix.fixLabel || '(unlabeled fix)';
      var oldText = fix.old;
      var newText = fix['new'];
      if (!oldText) { failed.push(label + ' — no OLD text provided'); return; }
      if (brokenLinkPattern.test(oldText) || brokenLinkPattern.test(newText || '')) {
        failed.push(label + ' — REJECTED: contains malformed markdown-in-HTML link syntax, not applied');
        return;
      }
      var occurrences = html.split(oldText).length - 1;
      if (occurrences === 0) {
        failed.push(label + ' — OLD fragment not found in article');
        return;
      }
      if (occurrences > 1) {
        failed.push(label + ' — OLD fragment matches ' + occurrences + ' places, skipped for safety');
        return;
      }
      html = html.replace(oldText, newText);
      applied.push(label);
    });
    sheet.getRange(row, 176).setValue(html);
    SpreadsheetApp.flush();
    var message = applied.length + '/' + fixes.length + ' fix(es) applied to FT.';
    if (applied.length > 0) message += '\nApplied: ' + applied.join(', ');
    if (failed.length > 0) message += '\nSkipped: ' + failed.join(' | ');
    return { success: applied.length > 0, message: message };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveAppliedRecommendationsToFT(html) {
  try {
    if (!html || !html.trim()) {
      return { success: false, message: "No HTML provided." };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };
    var cell = sheet.getRange(row, 176);
    cell.setNumberFormat('@');
    cell.setValue(html.trim());
    return { success: true, message: "Saved to column FT (176)." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}