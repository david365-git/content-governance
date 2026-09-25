/**
 * ================================================================================
 * ce_W8A_EditBrief.gs - W8A — EDIT BRIEF GENERATOR
 * ================================================================================
 *
 * Reads the W8_Analysis stored by W7 from posts_progress, applies TSM
 * article-type recall rules, and generates a targeted edit brief identifying
 * sections that need correction and why.
 *
 * Stores the edit brief to W8_Brief column in posts_progress.
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * ================================================================================
 */

function getW8EditBrief() {
  try {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet    = ss.getSheetByName('posts');
    var progressSheet = ss.getSheetByName('posts_progress');

    if (!postsSheet)    return { success: false, message: 'posts sheet not found.' };
    if (!progressSheet) return { success: false, message: 'posts_progress sheet not found.' };

    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2)  return { success: false, message: 'Select a data row first.' };

    // Get Post ID
    var postHeaders = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0];
    var postIdIdx   = postHeaders.map(function(h){ return String(h).trim(); }).indexOf('Post ID');
    var postId      = postIdIdx > -1 ? String(postsSheet.getRange(activeRow, postIdIdx + 1).getValue() || '').trim() : '';
    if (!postId) return { success: false, message: 'No Post ID found on active row.' };

    // Get Article Type
    var articleTypeIdx = postHeaders.map(function(h){ return String(h).trim(); }).indexOf('Article Type');
    var articleType    = articleTypeIdx > -1 ? String(postsSheet.getRange(activeRow, articleTypeIdx + 1).getValue() || '').trim() : '';
    if (!articleType) return { success: false, message: 'No Article Type found on active row.' };

    // Get current HTML from posts col 98 (New HTML)
    var html = String(postsSheet.getRange(activeRow, 98).getValue() || '').trim();
    if (!html) return { success: false, message: 'No HTML found in New HTML column (col 98). Run pipeline first.' };

    // Read W8_Analysis from posts sheet col 163 (FG)
    var analysisText = String(postsSheet.getRange(activeRow, 163).getValue() || '').trim();
    if (!analysisText) {
      return { success: false, message: 'No W8_Analysis found for this row. Run W7 first.' };
    }

    // Parse recall score from analysis text
    var recallScore    = 0;
    var recallMatch    = analysisText.match(/Recall Score[:\s]+([0-9.]+)/i);
    if (recallMatch) recallScore = parseFloat(recallMatch[1]) || 0;

    // Parse recommendation
    var recommendation = '';
    var recMatch       = analysisText.match(/RECOMMENDATION:\s*([A-Z ]+)/i);
    if (recMatch) recommendation = recMatch[1].trim();

    // Gate — refuse to run if score is already within acceptable range for article type
    var gateMessage = '';
    if (articleType === 'Hub Page' && recallScore <= 0.8) {
      gateMessage = 'W8 not needed. Hub Page recall score (' + recallScore + ') is already within the acceptable range (≤ 0.8). Proceed to publish.';
    } else if (articleType === 'Educational Guide' && recallScore >= 0.8) {
      gateMessage = 'W8 not needed. Educational Guide recall score (' + recallScore + ') is already at or above threshold (0.8). Proceed to publish.';
    } else if (articleType === 'Diagnostic Guide' && recallScore >= 0.8) {
      gateMessage = 'W8 not needed. Diagnostic Guide recall score (' + recallScore + ') is already at or above threshold (0.8). Proceed to publish.';
    } else if (articleType === 'Method Guide' && recallScore >= 0.8) {
      gateMessage = 'W8 not needed. Method Guide recall score (' + recallScore + ') is already at or above threshold (0.8). Proceed to publish.';
    } else if (articleType === 'Buyer Guide' && recallScore >= 0.7) {
      gateMessage = 'W8 not needed. Buyer Guide recall score (' + recallScore + ') is already at or above threshold (0.7). Proceed to publish.';
    } else if (articleType === 'Case Study') {
      gateMessage = 'W8 not needed. Case Study articles are not scored by recall. Proceed to publish.';
    } else if (recommendation === 'PUBLISH') {
      gateMessage = 'W8 not needed. W7 recommendation is PUBLISH. Proceed to publish.';
    }

    if (gateMessage) {
      return { success: false, message: gateMessage };
    }

    // Parse missing PQC terms
    var missingTerms = [];
    var missingMatch = analysisText.match(/Missing PQC Terms[:\s]+([^\n]+)/i);
    if (missingMatch && missingMatch[1].trim().toLowerCase() !== 'none') {
      missingTerms = missingMatch[1].split(',').map(function(t){ return t.trim(); }).filter(function(t){ return t.length > 0; });
    }

    // Parse word count line
    var wordCountLine = '';
    var wcMatch = analysisText.match(/Word Count[:\s]+([^\n]+)/i);
    if (wcMatch) wordCountLine = wcMatch[1].trim();

    // Apply TSM article-type recall rules to determine brief type
    var brief = buildBriefForArticleType(articleType, recallScore, missingTerms, wordCountLine, html);

    // Store W8_Brief to posts sheet col 164 (FH)
    var briefCell = postsSheet.getRange(activeRow, 164);
    briefCell.setNumberFormat('@');
    briefCell.setValue(brief);

    return {
      success:        true,
      brief:          brief,
      articleType:    articleType,
      recallScore:    recallScore,
      recommendation: recommendation,
      message:        'Edit brief generated and stored to W8_Brief.'
    };

  } catch(e) {
    return { success: false, message: 'getW8EditBrief error: ' + e.message };
  }
}

/**
 * W8 MEDIA GAP TOOL
 * Builds a Gemini prompt that scans the article's current HTML (same
 * source as W8A — col 98, New HTML) for stretches of continuous prose
 * exceeding 300 words with no visual break, and recommends one of five
 * non-video media types per flagged stretch. Never recommends
 * photorealistic AI imagery — only illustration/diagram/icon/table types,
 * or real photography for genuine before/after portfolio claims.
 */
function buildW8MediaGapPrompt(overrideHtml) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    var html = String(overrideHtml || '').trim();
    if (!html) {
      html = String(sheet.getRange(row, 98).getValue() || '').trim();
    }
    if (!html) return { success: false, message: 'No HTML found — New HTML (col 98) is empty and no override was provided.' };

    var prompt =
      'ROLE: Content Structure Editor reviewing a floor-care article for visual pacing and media selection.\n\n' +
      'CONTEXT: This is a professional UK stone/tile floor restoration website. Readers are homeowners in "solve-a-problem" mode, often on mobile. Long unbroken stretches of prose hurt readability, scroll depth and engagement.\n\n' +
      'TASK: Scan the HTML below and identify every stretch of continuous body text that exceeds 300 words with no visual break. A "visual break" is any existing <figure>, <img>, <blockquote>, <h2>, <h3>, <table>, or list (<ol>/<ul>) — these reset the running word count.\n\n' +
      'For each flagged stretch, recommend ONE media type from this list, chosen strictly by what that specific paragraph content is doing:\n\n' +
      '1. BEFORE/AFTER COMPARISON — only when the text describes a transformation or dramatic outcome contrast. MUST use real photography, never AI-generated — do not suggest this type unless real project photos would plausibly exist for this content.\n' +
      '2. CALLOUT / WARNING BOX — when the text states an explicit prohibition, safety rule, or "never do X" instruction.\n' +
      '3. DIAGRAM / CROSS-SECTION — when explaining a hidden mechanism, process, or how something behaves beneath the visible surface (e.g. how heat affects sealer, how moisture moves through stone).\n' +
      '4. COMPARISON TABLE — when comparing two or more named options, products, or approaches side by side.\n' +
      '5. ICON-BASED STEP SUMMARY — when the text describes a sequential multi-step process that could be condensed into a visual checklist.\n\n' +
      'HARD RULE: Do NOT recommend photorealistic AI-generated imagery under any circumstance. AI-generated visuals must always be clearly illustrated/diagrammatic/iconographic — never an attempt to look like a real photograph. Only genuine before/after portfolio content should use real photography, and only if real project photos plausibly exist for that specific claim.\n\n' +
      'Do NOT suggest video. Do NOT flag headings, short transitional sentences, or existing lists/tables as needing a break — only continuous prose paragraphs exceeding 300 words.\n\n' +
      'For each flagged stretch, provide:\n' +
      '1. WORD COUNT: approximate word count of the unbroken stretch\n' +
      '2. LOCATION: quote the first ~15 words and the last ~15 words of the flagged stretch\n' +
      '3. RECOMMENDED MEDIA TYPE: one of the five categories above\n' +
      '4. RATIONALE: one sentence explaining why this category fits this specific content\n' +
      '5. BRIEF: for illustration/diagram/icon/table types, a ready-to-use AI generation prompt (landscape, clearly non-photorealistic, specific to the actual subject matter — not generic). For before/after, instead note what real project photo would be needed (do not generate a prompt).\n\n' +
      'OUTPUT FORMAT (repeat for each flagged stretch):\n' +
      '---\n' +
      'FLAGGED STRETCH [N]\n' +
      'WORD COUNT: [number]\n' +
      'LOCATION: "[first ~15 words]..." ... "...[last ~15 words]"\n' +
      'RECOMMENDED MEDIA TYPE: [category name]\n' +
      'RATIONALE: [one sentence]\n' +
      'BRIEF: [prompt, or note on real photo needed]\n' +
      '---\n\n' +
      'If no stretch exceeds 300 words, state "No wall-of-text sections found — visual pacing is adequate" and stop.\n\n' +
      'ARTICLE HTML:\n' +
      html;

    return { success: true, prompt: prompt };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Saves the finished HTML (after manually inserting suggested media and
 * generating WordPress image URLs) to col 181 (FY — W8 Media Gap HTML).
 * Purely archival/staging — you will manually copy this into New HTML
 * yourself once satisfied, matching the W8B/W8C review-before-push pattern.
 */
function saveW8MediaGapHtml(html) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    var text = String(html || '').trim();
    if (!text) return { success: false, message: 'No HTML to save.' };

    sheet.getRange(row, 181).setValue(text);
    SpreadsheetApp.flush();

    return { success: true, message: 'Saved to W8 Media Gap HTML (col FY).' };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/* ============================================================
   BRIEF BUILDER — applies TSM recall rules per article type
============================================================ */
function buildBriefForArticleType(articleType, recallScore, missingTerms, wordCountLine, html) {

  var lines = [];
  lines.push('W8 EDIT BRIEF');
  lines.push('Article Type: ' + articleType);
  lines.push('Recall Score: ' + recallScore);
  lines.push('');

  // Filter out stemming artefacts from missing terms
  var STEMMING_ARTEFACTS = ['surfac', 'choos', 'turn', 'tri', 'use', 'caus', 'provid', 'includ', 'requir'];
  var genuineGaps = missingTerms.filter(function(t) {
    return STEMMING_ARTEFACTS.indexOf(t.toLowerCase()) === -1;
  });
  var stemmingArtefacts = missingTerms.filter(function(t) {
    return STEMMING_ARTEFACTS.indexOf(t.toLowerCase()) > -1;
  });

  // ── HUB PAGE ──
  if (articleType === 'Hub Page') {
    lines.push('RECALL RULE: Hub pages route at summary level. Deep entity vocabulary belongs in spoke pages.');
    lines.push('Score above 0.8 indicates the hub is doing spoke page jobs.');
    lines.push('');
    lines.push('BRIEF TYPE: HUB BOUNDARY CORRECTION');
    lines.push('INSTRUCTION: Do not add content. Identify and elevate or remove sections containing deep entity vocabulary, treatment descriptions, product categories, procedural language, or sealer type comparisons. Each flagged section must be reduced to summary-level framing with a link to the appropriate spoke page.');
    lines.push('');

    // Scan HTML sections for spoke-level content signals
    var flaggedSections = identifyHubSpokeDrift(html);
    if (flaggedSections.length > 0) {
      lines.push('FLAGGED SECTIONS:');
      flaggedSections.forEach(function(s) {
        lines.push('');
        lines.push('Section: ' + s.id);
        lines.push('H2: ' + s.h2);
        lines.push('Issue: ' + s.issue);
        lines.push('Instruction: ' + s.instruction);
      });
    } else {
      lines.push('No sections automatically flagged. Review manually for treatment-level vocabulary.');
    }
    lines.push('');
    lines.push('WORD COUNT NOTE: Do not pad or cut to match competitor benchmark. Word count reduction is an expected outcome of hub boundary correction, not a target.');
  }

  // ── EDUCATIONAL GUIDE ──
  else if (articleType === 'Educational Guide') {
    lines.push('RECALL RULE: Deep entity vocabulary expected. Score above 0.8 is good coverage.');
    lines.push('');
    if (recallScore >= 0.8) {
      lines.push('BRIEF TYPE: COVERAGE CONFIRMED — minor gap fill only');
      lines.push('INSTRUCTION: Coverage is strong. Address genuine missing terms only. Do not restructure.');
    } else if (recallScore >= 0.6) {
      lines.push('BRIEF TYPE: BORDERLINE — targeted gap fill');
      lines.push('INSTRUCTION: Coverage is borderline. Add entity vocabulary for the following genuine gaps.');
    } else {
      lines.push('BRIEF TYPE: COVERAGE GAPS — substantive gap fill needed');
      lines.push('INSTRUCTION: Coverage is below target. Expand entity vocabulary across flagged sections.');
    }
    if (genuineGaps.length > 0) {
      lines.push('');
      lines.push('GENUINE MISSING TERMS: ' + genuineGaps.join(', '));
      lines.push('Add natural coverage for each term in the most relevant existing section.');
    }
  }

  // ── DIAGNOSTIC GUIDE ──
  else if (articleType === 'Diagnostic Guide') {
    lines.push('RECALL RULE: Defect vocabulary depth expected.');
    lines.push('');
    if (recallScore >= 0.8) {
      lines.push('BRIEF TYPE: COVERAGE CONFIRMED — minor gap fill only');
      lines.push('INSTRUCTION: Defect vocabulary is strong. Address genuine missing terms only.');
    } else {
      lines.push('BRIEF TYPE: DEFECT VOCABULARY GAPS');
      lines.push('INSTRUCTION: Expand defect and condition vocabulary. Each named defect must include: what it is, what the homeowner sees, what is done about it.');
    }
    if (genuineGaps.length > 0) {
      lines.push('');
      lines.push('GENUINE MISSING TERMS: ' + genuineGaps.join(', '));
    }
  }

  // ── METHOD GUIDE ──
  else if (articleType === 'Method Guide') {
    lines.push('RECALL RULE: Process vocabulary depth expected.');
    lines.push('');
    if (recallScore >= 0.8) {
      lines.push('BRIEF TYPE: COVERAGE CONFIRMED — minor gap fill only');
      lines.push('INSTRUCTION: Process vocabulary is strong. Address genuine missing terms only.');
    } else {
      lines.push('BRIEF TYPE: PROCESS VOCABULARY GAPS');
      lines.push('INSTRUCTION: Expand process and step vocabulary. Each stage must be named and described.');
    }
    if (genuineGaps.length > 0) {
      lines.push('');
      lines.push('GENUINE MISSING TERMS: ' + genuineGaps.join(', '));
    }
  }

  // ── BUYER GUIDE ──
  else if (articleType === 'Buyer Guide') {
    lines.push('RECALL RULE: Decision-support vocabulary expected.');
    lines.push('');
    if (recallScore >= 0.7) {
      lines.push('BRIEF TYPE: COVERAGE CONFIRMED — minor gap fill only');
      lines.push('INSTRUCTION: Decision-support vocabulary is strong. Address genuine missing terms only.');
    } else {
      lines.push('BRIEF TYPE: DECISION VOCABULARY GAPS');
      lines.push('INSTRUCTION: Expand evaluation criteria, cost and contractor vocabulary.');
    }
    if (genuineGaps.length > 0) {
      lines.push('');
      lines.push('GENUINE MISSING TERMS: ' + genuineGaps.join(', '));
    }
  }

  // ── ALL OTHER TYPES ──
  else {
    lines.push('RECALL RULE: Standard entity vocabulary coverage.');
    lines.push('');
    if (genuineGaps.length > 0) {
      lines.push('GENUINE MISSING TERMS: ' + genuineGaps.join(', '));
      lines.push('INSTRUCTION: Add natural coverage for each term in the most relevant existing section.');
    } else {
      lines.push('INSTRUCTION: No genuine coverage gaps identified. Review manually before rewriting.');
    }
  }

// ── ENTITY DUMP DETECTION (mechanical pre-scan — informational only) ──
  // The mechanical regex over-flags normal prose using verbs outside its whitelist.
  // Genuine entity-dump judgment now happens in the Stage 3B deferred LLM audit (CHECK 8).
  // This section only surfaces candidates for manual review — it is not a hard instruction.
  var entityDumps = detectEntityDumps(html);
  if (entityDumps.length > 0) {
    lines.push('');
    lines.push('POSSIBLE ENTITY DUMP PARAGRAPHS (mechanical pre-scan — verify against Stage 3B CHECK 8 before rewriting):');
    lines.push('The following paragraphs were mechanically flagged as possible comma-separated lists.');
    lines.push('This scan produces false positives on normal prose. Only rewrite if Stage 3B CHECK 8 confirms a genuine issue.');
    lines.push('');
    entityDumps.forEach(function(dump, i) {
      lines.push('CANDIDATE ' + (i + 1) + ':');
      lines.push('Section: ' + dump.sectionId);
      lines.push('Paragraph (first 120 chars): ' + dump.snippet);
      lines.push('');
    });
  }

  // Stemming artefacts note
  if (stemmingArtefacts.length > 0) {
    lines.push('');
    lines.push('STEMMING ARTEFACTS (ignore — not genuine gaps): ' + stemmingArtefacts.join(', '));
  }

  // Word count note
  if (wordCountLine) {
    lines.push('');
    lines.push('WORD COUNT: ' + wordCountLine);
    lines.push('Competitor benchmark is a ceiling guide only. Do not pad or cut to match it.');
  }

  return lines.join('\n');
}

/* ============================================================
   HUB / SPOKE DRIFT DETECTOR
   Scans HTML sections for treatment-level vocabulary that
   belongs on spoke pages rather than the hub.
============================================================ */
function identifyHubSpokeDrift(html) {
  var flagged = [];

  // Signals that indicate spoke-level content in a hub section
  var SPOKE_SIGNALS = [
    { pattern: /impregnating sealer|topical sealer|colour.enhanc|acrylic (coating|sealer)|wax system|urethane (sealer|coating)/i, issue: 'Sealer type descriptions — product category detail belongs on sealing spoke page', instruction: 'Remove sealer type names and finish descriptions. Keep only the diagnostic observation that old coatings can hide the true surface.' },
    { pattern: /pH.neutral|mild surfactant|wring(er)? bucket|mop (routine|method|water)|residue.free daily clean/i, issue: 'Mopping method detail — procedural cleaning instruction belongs on cleaning spoke page', instruction: 'Remove mopping method and cleaning product references. Keep only the framing that care routine needs to match floor age, finish and treatment history.' },
    { pattern: /sealer selection|which sealer|choose (a |the )?sealer|sealer (decision|choice|matching)/i, issue: 'Sealer selection language — decision belongs on sealing spoke page', instruction: 'Remove sealer selection guidance. Keep only the observation that texture changes how the floor responds to protection.' },
    { pattern: /solvent action|acrylic removal|wax dissolution|chemical penetration|coating stripping|sealer removal/i, issue: 'Build-up treatment method — correction procedure belongs on treatment spoke page', instruction: 'Remove treatment procedure language. Keep only the diagnostic signal that old coating history can hide the true slate surface.' },
    { pattern: /slurry extraction|wet vacuum|redeposition prevention|residue removal/i, issue: 'Cleaning method detail — extraction procedure belongs on cleaning spoke page', instruction: 'Remove extraction method language. Keep only the framing that links to the cleaning spoke page.' }
  ];

  // Extract sections from HTML
  var sectionRe = /<section[^>]+id="([^"]+)"[^>]*>([\s\S]*?)<\/section>/gi;
  var sm;
  while ((sm = sectionRe.exec(html)) !== null) {
    var sectionId      = sm[1];
    var sectionContent = sm[2];

    // Skip hub-intro navigation section
    if (sectionId === 'hub-intro') continue;

    // Extract H2
    var h2 = '';
    var h2m = sectionContent.match(/<h2[^>]*>([\s\S]*?)<\/h2>/i);
    if (h2m) h2 = h2m[1].replace(/<[^>]+>/g, '').trim();

    // Strip HTML tags for text analysis
    var plainText = sectionContent.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

    // Check each spoke signal
    SPOKE_SIGNALS.forEach(function(signal) {
      if (signal.pattern.test(plainText)) {
        // Avoid duplicate section entries
        var alreadyFlagged = flagged.some(function(f) { return f.id === sectionId && f.issue === signal.issue; });
        if (!alreadyFlagged) {
          flagged.push({
            id:          sectionId,
            h2:          h2,
            issue:       signal.issue,
            instruction: signal.instruction
          });
        }
      }
    });
  }

  return flagged;
}

/* ============================================================
   LOAD W8 BRIEF FOR SIDEBAR DISPLAY
============================================================ */
function loadW8BriefFromSheet() {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet = ss.getSheetByName('posts');

    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return { success: false, message: 'Select a data row first.' };

    var brief = String(postsSheet.getRange(activeRow, 164).getValue() || '').trim();
    if (!brief) return { success: false, message: 'No W8_Brief found for this row. Run W8A first.' };

    return { success: true, brief: brief };

  } catch(e) {
    return { success: false, message: 'loadW8BriefFromSheet error: ' + e.message };
  }
}

