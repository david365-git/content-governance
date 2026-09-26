/**
 * ================================================================================
 * ce_AtomicAuditFixer.gs — ATOMIC AUDIT FIX ENGINE
 * ================================================================================
 */

function runAtomicAuditWithContext(html) {
  try {
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var sheet       = ss.getSheetByName('posts');
    var row         = sheet.getActiveRange().getRow();
    if (row < 2) return { error: 'Select a data row first.' };

    var articleType = String(sheet.getRange(row, 8).getValue() || 'General').trim();
    var hubUrl      = String(sheet.getRange(row, 15).getValue() || '').trim();

    var report = runMechanicalAuditOnHTML(html, '', articleType, hubUrl);

    var failLines = report.split('\n').filter(function(l) { return l.indexOf('FAIL') > -1; });
    var allPass   = failLines.length === 0;

    if (allPass) {
      return { success: true, allPass: true, report: report, failures: [] };
    }

    var failures = [];
    failLines.forEach(function(line) {
      var failure = buildAtomicFailure(line, html, hubUrl, articleType);
      if (!failure) return;
      if (Array.isArray(failure)) {
        failure.forEach(function(f) { failures.push(f); });
      } else {
        failures.push(failure);
      }
    });

    return { success: true, allPass: false, report: report, failures: failures };

  } catch(e) {
    return { error: 'ATOMIC AUDIT ERROR: ' + e.toString() };
  }
}

function atomicHasRealLateralLink(html, hubUrl) {

  function normaliseUrl(url) {
    return String(url || "")
      .trim()
      .replace(/^https?:\/\/(?:www\.)?/i, "")
      .replace(/[?#].*$/, "")
      .replace(/\/+$/, "")
      .toLowerCase();
  }

  var hubNorm = normaliseUrl(hubUrl);

  var links = [];
  var re =
    /href=["'](https?:\/\/(?:www\.)?abbeyfloorcare\.co\.uk[^"']*)["']/gi;

  var match;

  while ((match = re.exec(String(html || ""))) !== null) {
    links.push(match[1]);
  }

  return links.some(function(link) {
    return normaliseUrl(link) !== hubNorm;
  });
}

function buildAtomicFailure(failLine, html, hubUrl, articleType) {
  var failure = {
    checkId:     '',
    description: failLine.trim(),
    fragment:    '',
    fixPrompt:   '',
    canAutoFix:  false,
    autoFix:     null
  };

  // ── CHECK 2E — H1 in body ──
  if (failLine.indexOf('CHECK 2E') > -1) {
    failure.checkId    = '2E';
    failure.canAutoFix = true;
    failure.autoFix    = 'removeH1FromHtml';
    failure.description = 'H1 tag found in body — will be removed automatically.';
    return failure;
  }

  // ── CHECK 3G — Diagnostic captions ──
  // NOTE: returns an ARRAY of failures (one per non-diagnostic caption found),
  // not a single failure — the audit line can report multiple failing captions
  // at once, and all of them need to be included in the same batch.
  if (failLine.indexOf('CHECK 3G') > -1) {
    var figRe = /<figure[^>]*class="[^"]*wp-caption[^"]*"[^>]*>[\s\S]*?<\/figure>/gi;
    var figMatch;
    var failingFigures = [];
    while ((figMatch = figRe.exec(html)) !== null) {
      var figBlock = figMatch[0];
      var capMatch = figBlock.match(/<figcaption[^>]*>([\s\S]*?)<\/figcaption>/i);
      if (!capMatch) continue;
      var capText = capMatch[1].replace(/<[^>]+>/g, '').trim().toLowerCase();
      var DIAGNOSTIC_SIGNALS = [
        'your ', 'you ', 'this is ', 'like this', 'indicat', 'at this stage',
        'need', 'have ', 'suggests', 'means', 'shows ', 'if your', 'floors '
      ];
      var isDiagnostic = DIAGNOSTIC_SIGNALS.some(function(signal) {
        return capText.indexOf(signal) > -1;
      });
      if (!isDiagnostic) {
        failingFigures.push(figBlock);
      }
    }
    if (failingFigures.length === 0) return null;

    var multiFailures = failingFigures.map(function(fig) {
      return {
        checkId: '3G',
        description: 'CHECK 3G — Diagnostic caption',
        fragment: fig,
        fixPrompt: buildStrictRepairPrompt(
          'CHECK 3G — Diagnostic caption',
          'The figcaption is purely descriptive. Rewrite it to connect the image to the reader\'s problem using one of these patterns:\n' +
          '- "If your floor looks like this..." \n' +
          '- "This is [condition] — [what it means for the reader]"\n' +
          '- "Floors at this stage need..."\n' +
          '- "Dark patches like these indicate..."\n' +
          'Caption must be under 20 words. Change ONLY the figcaption text.',
          fig
        ),
        canAutoFix: false,
        autoFix: null
      };
    });

    return multiFailures;
  }

  // ── CHECK 3 — Missing figcaption ──
  if (failLine.indexOf('CHECK 3') > -1 && failLine.indexOf('3a:') > -1) {
    failure.checkId = '3a';
    var headingScanRe = /<h([1-6])[^>]*>[\s\S]*?<\/h\1>/gi;
    var headingMatches3a = [];
    var headingScanMatch;
    while ((headingScanMatch = headingScanRe.exec(html)) !== null) {
      headingMatches3a.push({ level: parseInt(headingScanMatch[1], 10), tag: headingScanMatch[0], index: headingScanMatch.index });
    }
    var skipIndex = -1;
    for (var hi = 1; hi < headingMatches3a.length; hi++) {
      if (headingMatches3a[hi].level - headingMatches3a[hi - 1].level > 1) {
        skipIndex = hi;
        break;
      }
    }
    if (skipIndex === -1) return null;
    var skippingHeading = headingMatches3a[skipIndex].tag;
    var correctLevel = headingMatches3a[skipIndex - 1].level + 1;
    failure.fragment = skippingHeading;
    failure.fixPrompt = buildStrictRepairPrompt(
      'CHECK 3 — Heading hierarchy skip',
      'This heading skips a level in the document\'s heading hierarchy. Change ONLY the heading tag level ' +
      'to h' + correctLevel + ' (keep the exact same text content) so it follows the preceding heading correctly. ' +
      'Do not change the heading text itself.',
      skippingHeading
    );
    return failure;
  }
  if (failLine.indexOf('CHECK 3') > -1 && failLine.indexOf('3c:') > -1) {
    failure.checkId = '3';
    var figRe2 = /<figure[^>]*>[\s\S]*?<\/figure>/gi;
    var figMatch2;
    var failingFig = null;
    while ((figMatch2 = figRe2.exec(html)) !== null) {
      var fb = figMatch2[0];
      if (!/David_Allen\.jpg/i.test(fb) && !/<figcaption/i.test(fb) && !/video-embed/i.test(fb)) {
        failingFig = fb;
        break;
      }
    }
    if (!failingFig) return null;
    failure.fragment = failingFig;
    failure.fixPrompt = buildStrictRepairPrompt(
      'CHECK 3 — Missing figcaption',
      'Add a figcaption immediately after the img tag and before the closing figure tag.\n' +
      'Caption must connect the image to the reader\'s problem diagnostically and be under 20 words.\n' +
      'Change ONLY by adding the figcaption.',
      failingFig
    );
    return failure;
  }

  // ── CHECK 2B — Bare URL anchor text ──
  if (failLine.indexOf('CHECK 2B') > -1) {
    failure.checkId = '2B';
    var urlMatch = failLine.match(/https?:\/\/[^\s|]+/);
    if (!urlMatch) return null;
    var bareUrl = urlMatch[0].trim();
    var anchorRe = new RegExp('<a[^>]+href="' + bareUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*>' + bareUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '<\/a>', 'i');
    var anchorMatch = html.match(anchorRe);
    if (!anchorMatch) {
      var genericAnchor = html.match(/<a[^>]+href="[^"]*"[^>]*>https?:\/\/[^<]+<\/a>/i);
      if (!genericAnchor) return null;
      failure.fragment = genericAnchor[0];
    } else {
      failure.fragment = anchorMatch[0];
    }
    failure.fixPrompt = buildStrictRepairPrompt(
      'CHECK 2B — Bare URL anchor text',
      'Replace the anchor text with descriptive text (3-8 words) describing the linked page.\n' +
      'Change ONLY the anchor text. Do not change href or any other attribute.',
      failure.fragment
    );
    return failure;
  }

  // ── CHECK 9 — Missing CTA ──
  if (failLine.indexOf('CHECK 9') > -1) {
    failure.checkId = '9';
    var footerMatch = html.match(/<footer[\s\S]*?<\/footer>/i);
    if (!footerMatch) return null;
    failure.fragment = footerMatch[0];
    failure.fixPrompt = buildStrictRepairPrompt(
      'CHECK 9 — Missing CTA link',
      'Add a link to /contact as the last sentence in the footer paragraph.\n' +
      'Anchor text: "Contact us to arrange a no-obligation floor assessment."\n' +
      'Change ONLY the footer content by adding the contact link.',
      footerMatch[0]
    );
    return failure;
  }

  // ── CHECK 2 — Lateral internal link ──
    if (failLine.indexOf('CHECK 2') > -1 && failLine.indexOf('No lateral internal link found') > -1) {

    if (atomicHasRealLateralLink(html, hubUrl)) {
      return null;
    }

    failure.checkId = '2-lateral';
    var targetSection = null;
    var maxSectionNum = -1;
    var sectionScanRe = /<section[^>]*id="section-(\d+)"[^>]*>[\s\S]*?<\/section>/gi;
    var sectionScanMatch;
    while ((sectionScanMatch = sectionScanRe.exec(html)) !== null) {
      var thisSectionNum = parseInt(sectionScanMatch[1], 10);
      if (thisSectionNum > maxSectionNum) {
        maxSectionNum = thisSectionNum;
        targetSection = sectionScanMatch[0];
      }
    }
    var lastParaMatch2 = targetSection ? targetSection.match(/<p[^>]*>[\s\S]*?<\/p>/gi) : null;
    if (lastParaMatch2) {
      lastParaMatch2 = lastParaMatch2.filter(function(p) { return !/<iframe/i.test(p) && !/<center>/i.test(p); });
    }
    if (!lastParaMatch2 || lastParaMatch2.length === 0) return null;
    var lastPara2 = lastParaMatch2[lastParaMatch2.length - 1];
    failure.fragment = lastPara2;
    failure.fixPrompt = buildStrictRepairPrompt(
      'CHECK 2 — Lateral internal link missing',
      'Add one contextual internal link to a related page within the same silo — NOT the hub page.\n' +
      'Use descriptive anchor text (3-8 words). Change ONLY by adding the link.',
      lastPara2
    );
    return failure;
  }

  // ── CHECK 2 — Hub link ──
  if (failLine.indexOf('CHECK 2') > -1 && failLine.indexOf('Hub page link missing') > -1) {
    failure.checkId = '2-hub';
    var sec1Match = html.match(/<section[^>]*id="section-1"[^>]*>[\s\S]*?<\/section>/i);
    if (!sec1Match) return null;
    var lastParaMatch = sec1Match[0].match(/<p[^>]*>[\s\S]*?<\/p>/gi);
    if (!lastParaMatch || lastParaMatch.length === 0) return null;
    var lastPara = lastParaMatch[lastParaMatch.length - 1];
    failure.fragment = lastPara;
    failure.fixPrompt = buildStrictRepairPrompt(
      'CHECK 2 — Hub link missing',
      'Add a contextual internal link to the hub page: ' + hubUrl + '\n' +
      'Use descriptive anchor text (3-8 words). Change ONLY by adding the link.',
      lastPara
    );
    return failure;
  }

  // ── CHECK 2 — Excess internal links ──
  if (failLine.indexOf('CHECK 2') > -1 && failLine.indexOf('exceeds the maximum') > -1) {
    var maxLinksMatch = failLine.match(/maximum of (\d+)/i);
    var maxLinksAllowed = maxLinksMatch ? parseInt(maxLinksMatch[1], 10) : null;
    if (maxLinksAllowed === null) return null;

    var hubNorm = hubUrl ? hubUrl.replace(/#.*$/, '').replace(/[/]+$/, '').toLowerCase() : '';
    var linkRe = /<a\s+href="(https?:\/\/(?:www\.)?abbeyfloorcare\.co\.uk[^"]*)"[^>]*>[\s\S]*?<\/a>/gi;
    var seenHrefs = {};
    var candidates = [];
    var linkMatch;
    while ((linkMatch = linkRe.exec(html)) !== null) {
      var href = linkMatch[1];
      var hrefNorm = href.replace(/#.*$/, '').replace(/[/]+$/, '').toLowerCase();
      var isHub = hubNorm && (hrefNorm === hubNorm || hrefNorm.indexOf(hubNorm) > -1 || hubNorm.indexOf(hrefNorm) > -1);
      if (isHub) continue; // hub link is mandatory, never a removal candidate
      if (seenHrefs[hrefNorm]) continue;
      seenHrefs[hrefNorm] = true;

      var linkIndex = linkMatch.index;
      var paraStart = html.lastIndexOf('<p', linkIndex);
      var paraEnd = html.indexOf('</p>', linkIndex);
      if (paraStart === -1 || paraEnd === -1) continue;
      var paragraph = html.substring(paraStart, paraEnd + 4);
      candidates.push({ href: href, paragraph: paragraph });
    }

    if (candidates.length === 0) return null;

    var totalCounted = candidates.length + (hubNorm ? 1 : 0);
    var excessCount = totalCounted - maxLinksAllowed;
    if (excessCount <= 0) return null;

    var candidateListText = candidates.map(function(c, i) {
      return 'CANDIDATE ' + i + ' (href: ' + c.href + '):\n' + c.paragraph;
    }).join('\n\n');

    return candidates.map(function(c, i) {
      var f = {
        checkId:     '2-excess-' + i,
        description: 'CHECK 2 — Excess internal links (candidate ' + i + ')',
        fragment:    c.paragraph,
        fixPrompt:   '',
        canAutoFix:  false,
        autoFix:     null
      };
      f.fixPrompt = buildStrictRepairPrompt(
        'CHECK 2 — Excess internal links',
        'This article has ' + totalCounted + ' total internal links but the maximum permitted is ' + maxLinksAllowed +
        ' (hub link excluded from this count). Exactly ' + excessCount + ' of the ' + candidates.length +
        ' lateral link candidate(s) below must be removed — converted from a link to plain text, keeping the anchor text and surrounding sentence otherwise unchanged. ' +
        'Keep the candidate(s) that support a specific claim or causal step in the project narrative; remove the candidate(s) that merely re-explain a mechanism already stated inline in the same paragraph. ' +
        'If THIS specific candidate should be REMOVED, return the paragraph with the <a>...</a> tag replaced by its plain anchor text only — no other change. ' +
        'If THIS specific candidate should be KEPT, return the paragraph completely unchanged.\n\n' +
        'ALL CANDIDATES IN THIS DECISION (for context — judge consistently across all of them):\n' + candidateListText,
        c.paragraph
      );
      return f;
    });
  }

  return null;
}

function buildStrictRepairPrompt(checkName, rule, fragment) {
  return 'ROLE: HTML micro-patch engine.\n\n' +
    'TASK: Fix only the single failing HTML element provided below.\n\n' +
    'NON-NEGOTIABLE RULES:\n' +
    '1. Return ONLY the corrected replacement element — nothing else.\n' +
    '2. Zero-Modification Policy: Do not alter any character, whitespace, attribute, class, id, src, href, or media element that is not directly required by the fix.\n' +
    '3. No prose, no explanation, no markdown fences. Raw HTML only.\n' +
    '4. Preserve all existing tag names, attributes, URLs, image src values, classes and ids exactly.\n' +
    '5. If the fix cannot be made within this element, return exactly: ESCALATION REQUIRED\n\n' +
    'FAILURE: ' + checkName + '\n\n' +
    'RULE TO FIX:\n' + rule + '\n\n' +
    'TARGET ELEMENT:\n' + fragment + '\n\n' +
    'REQUIRED OUTPUT:\n' +
    'Corrected replacement element only. No preamble. No explanation. Raw HTML starting with the opening tag.';
}

function removeH1FromHtml(html) {
  return html.replace(/<h1[^>]*>[\s\S]*?<\/h1>/gi, '');
}

function applyAtomicFix(fullHtml, oldFragment, newFragment) {
  try {
    if (fullHtml.indexOf(oldFragment) === -1) {
      return { success: false, message: 'Original fragment not found in HTML — it may have already been changed. Re-run the audit.' };
    }
    if (!newFragment || newFragment.trim() === '' || newFragment.trim() === 'ESCALATION REQUIRED') {
      return {
        success: false,
        message: newFragment.trim() === 'ESCALATION REQUIRED'
          ? 'ChatGPT flagged this as requiring escalation — manual review needed.'
          : 'Empty fragment returned — no fix applied.'
      };
    }
    var patchedHtml = fullHtml.replace(oldFragment, newFragment);
    var beforeWithoutFragment = fullHtml.replace(oldFragment, '[[PATCHED]]');
    var afterWithoutFragment  = patchedHtml.replace(newFragment, '[[PATCHED]]');
    if (beforeWithoutFragment !== afterWithoutFragment) {
      return { success: false, message: 'Diff check failed — changes detected outside the target fragment. Fix rejected.' };
    }
    return { success: true, patchedHtml: patchedHtml, message: 'Fix applied successfully — diff check passed.' };
  } catch(e) {
    return { success: false, message: 'APPLY FIX ERROR: ' + e.toString() };
  }
}

function applyAutoFix(fullHtml, checkId) {
  try {
    if (checkId === '2E') {
      var fixed = removeH1FromHtml(fullHtml);
      return { success: true, patchedHtml: fixed, message: 'H1 tag removed automatically.' };
    }
    return { success: false, message: 'No auto-fix available for check ' + checkId };
  } catch(e) {
    return { success: false, message: 'AUTO-FIX ERROR: ' + e.toString() };
  }
}

/* ============================================================
   BATCH ATOMIC FIX — combines all non-auto-fixable failures
   into a single JSON-based prompt, so all fixes come back in
   one round-trip instead of one element at a time.
============================================================ */

function buildBatchAtomicFixPrompt(failures) {
  // failures: array of {checkId, description, fragment, fixPrompt} —
  // only include ones that have a fragment and fixPrompt (skip auto-fixable).
  var items = failures.filter(function(f) { return f && f.fragment && f.fixPrompt && !f.canAutoFix; });

  if (items.length === 0) {
    return { success: false, message: "No fixable failures to batch." };
  }

  var itemBlocks = items.map(function(f, idx) {
    // Extract the rule text between "RULE TO FIX:" and "TARGET ELEMENT:" from the existing fixPrompt
    var ruleMatch = f.fixPrompt.match(/RULE TO FIX:\n([\s\S]*?)\n\nTARGET ELEMENT:/);
    var rule = ruleMatch ? ruleMatch[1].trim() : '(rule text unavailable)';
    return 'ITEM ' + idx + ':\n' +
      'CHECK: ' + f.checkId + ' — ' + f.description + '\n' +
      'RULE TO FIX:\n' + rule + '\n\n' +
      'TARGET ELEMENT:\n' + f.fragment;
  }).join('\n\n' + '='.repeat(50) + '\n\n');

  var prompt =
`ROLE: HTML micro-patch engine.

TASK: Fix EACH of the failing HTML elements listed below independently. Each item is a separate, unrelated fix — do not let one item's fix affect another.

NON-NEGOTIABLE RULES (apply to every item):
1. Zero-Modification Policy: for each item, do not alter any character, whitespace, attribute, class, id, src, href, or media element that is not directly required by that item's specific fix.
2. Preserve all existing tag names, attributes, URLs, image src values, classes and ids exactly, except where the rule explicitly requires a change.
3. If a fix cannot be made within an item's element, set that item's "new" value to the exact string "ESCALATION REQUIRED" instead of guessing.
4. Return ONLY a JSON array, no markdown fences, no preamble, in this exact format:
[
  {"index": 0, "checkId": "exact check id from the item", "old": "exact original TARGET ELEMENT text, copied verbatim", "new": "the corrected replacement element, or ESCALATION REQUIRED"}
]
5. Every double quote character inside the old/new string values MUST be escaped as \\" so the result is valid JSON. Use straight quotes only.
6. The "old" value must be copied verbatim from the TARGET ELEMENT of that item — exact characters — so it can be located by an automated find-and-replace.
7. Include one array entry for every item below — do not skip any item.

ITEMS TO FIX:

${itemBlocks}

Return the complete JSON array covering all ${items.length} item(s) now.`;

  return { success: true, prompt: prompt, count: items.length };
}

function applyBatchAtomicFix(fullHtml, rawJson) {
  try {
    if (!rawJson || !rawJson.trim()) {
      return { success: false, message: "No response provided." };
    }

    var fixes;
    var cleaned = rawJson.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
    try {
      fixes = JSON.parse(cleaned);
    } catch (e) {
      try {
        var repaired = cleaned.replace(/"(old|new|checkId)":\s*"([\s\S]*?)"(?=\s*[,}])/g, function(match, key, value) {
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

    var html = String(fullHtml);
    var applied = [];
    var failed = [];
    var escalated = [];
    var brokenLinkPattern = /<a\s+href="\[https?:\/\/[^\]]+\]\(/i;

    fixes.forEach(function(fix) {
      var label = 'Item ' + (fix.index !== undefined ? fix.index : '?') + ' (' + (fix.checkId || 'unknown check') + ')';
      var oldText = fix.old;
      var newText = fix.new;

      if (!oldText) { failed.push(label + ' — no OLD text provided'); return; }

      if (String(newText).trim() === 'ESCALATION REQUIRED') {
        escalated.push(label);
        return;
      }

      if (brokenLinkPattern.test(oldText) || brokenLinkPattern.test(newText || '')) {
        failed.push(label + ' — REJECTED: malformed markdown-in-HTML link syntax');
        return;
      }

      var occurrences = html.split(oldText).length - 1;
      if (occurrences === 0) {
        failed.push(label + ' — OLD fragment not found in HTML');
        return;
      }
      if (occurrences > 1) {
        failed.push(label + ' — OLD fragment matches ' + occurrences + ' places, skipped for safety');
        return;
      }

      var patchedHtml = html.replace(oldText, newText);
      var beforeCheck = html.replace(oldText, '[[PATCHED]]');
      var afterCheck = patchedHtml.replace(newText, '[[PATCHED]]');
      if (beforeCheck !== afterCheck) {
        failed.push(label + ' — diff check failed, changes detected outside target fragment');
        return;
      }

      html = patchedHtml;
      applied.push(label);
    });

    var message = applied.length + '/' + fixes.length + ' fix(es) applied.';
    if (applied.length > 0) message += '\nApplied: ' + applied.join(', ');
    if (escalated.length > 0) message += '\nEscalated (needs manual review): ' + escalated.join(', ');
    if (failed.length > 0) message += '\nSkipped: ' + failed.join(' | ');
    if (applied.length > 0) {
      var ssBatch  = SpreadsheetApp.getActiveSpreadsheet();
      var sheetBatch = ssBatch.getSheetByName('posts');
      var rowBatch = sheetBatch.getActiveRange().getRow();
      if (rowBatch >= 2) {
        var saveCell = sheetBatch.getRange(rowBatch, 104);
        saveCell.setNumberFormat('@');
        saveCell.setValue(html);
        message += '\nSaved to Column CZ.';
      }
    }
    return { success: applied.length > 0, patchedHtml: html, message: message, escalatedCount: escalated.length, failedCount: failed.length };

  } catch(e) {
    return { success: false, message: 'BATCH APPLY ERROR: ' + e.toString() };
  }
}