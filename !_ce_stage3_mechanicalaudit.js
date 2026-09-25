/**
 * ================================================================================
 * ce_Stage3_MechanicalAudit.gs - STAGE3 MECHANICALAUDIT
 * ================================================================================
 *
 * Deterministic HTML and schema validation
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * Updated: March 2026
 *
 * CHANGES FROM PREVIOUS VERSION:
 *   - CHECK 15 auditSchemaURLs: EXEMPT array extended with two additional prefixes:
 *     1. https://schema.org/ — schema.org enumeration values (e.g. InStock)
 *        are valid non-domain URLs and must not be flagged
 *     2. https://www.google.com/maps/place/ — canonical sameAs target for
 *        Google Business Profile Place ID URLs — valid external authority reference
 * ================================================================================
 */

function runMechanicalAuditWithContext(html) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName("posts");
    const row     = sheet.getActiveRange().getRow();
    if (row < 2) return "ERROR: Select a data row first.";
    const articleType = String(sheet.getRange(row, 8).getValue() || "General").trim();
    const hubUrl      = String(sheet.getRange(row, 15).getValue() || "").trim();
    return runMechanicalAuditOnHTML(html.toString(), "", articleType, hubUrl);
  } catch(e) {
    return "AUDIT ERROR: " + e.toString();
  }
}

function runMechanicalAudit() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName("posts");
    const row     = sheet.getActiveRange().getRow();
    if (row < 2) return "ERROR: Select a data row first.";

    const html = sheet.getRange(row, 98).getValue();
    if (!html || html.toString().trim() === "") {
      return "ERROR: No HTML found in col 98 (New HTML) for this row. Push re-hydrated HTML first.";
    }

    let schema = String(sheet.getRange(row, 92).getValue() || "");
    if (schema) {
      schema = String(schema)
        .replace(/%22/g, '"').replace(/%3A/g, ':').replace(/%2F/g, '/')
        .replace(/%7B/g, '{').replace(/%7D/g, '}').replace(/%26/g, '&')
        .replace(/%5B/g, '[').replace(/%5D/g, ']')
        .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
        .replace(/\[("?https?:\/\/[^\]"]*"?)\]/g, '$1')
        .replace(/\[(https?:\/\/)/g, '$1');
      const sm = schema.match(/<script[^>]*>([\s\S]*?)<\/scr\x69pt>/i);
      if (sm) schema = sm[1].trim();
    }

    const articleType = String(sheet.getRange(row, 8).getValue() || "General").trim();
    const hubUrl      = String(sheet.getRange(row, 15).getValue() || "").trim();

    return runMechanicalAuditOnHTML(html.toString(), schema ? schema.toString() : "", articleType, hubUrl);
  } catch(e) {
    return "AUDIT ERROR: " + e.toString();
  }
}

function runMechanicalAuditOnHTML(html, schema, articleType, hubUrl) {
  articleType = articleType || "General";
  hubUrl = hubUrl || "";
  const results = [];
  let passCount = 0;
  let totalChecks = 5;

  /* ── CHECK 2 ── */
  const hubTypes2 = ["Hub Page"];
  const check2 = hubTypes2.indexOf(articleType) > -1
    ? auditQuickLinks(html)
    : auditInternalLinks(html, hubUrl, articleType);
  results.push("CHECK 2 [Internal links]: " + check2.label + " — " + check2.detail);
  if (check2.pass) passCount++;

  /* ── CHECK 3 ── */
  const check3 = auditHtmlStructure(html, articleType);
  results.push("CHECK 3 [Content integrity]: " + check3.label + " — " + check3.detail);
  if (check3.pass) passCount++;

  /* ── CHECK 3F ── */
  (function() {
    try {
      var ss          = SpreadsheetApp.getActiveSpreadsheet();
      var postsSheet  = ss.getSheetByName('posts');
      var activeRow   = postsSheet.getActiveRange().getRow();
      if (activeRow < 2) {
        results.push('CHECK 3F [Figure inventory]: NOT APPLICABLE — no active row');
        return;
      }
      var postHeaders = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0]
                                  .map(function(h) { return String(h).trim(); });
      var postIdIdx   = postHeaders.indexOf('Post ID');
      var postId      = postIdIdx > -1
        ? String(postsSheet.getRange(activeRow, postIdIdx + 1).getValue() || '').trim().replace(/\.0$/, '')
        : '';
      if (!postId) {
        results.push('CHECK 3F [Figure inventory]: NOT APPLICABLE — no Post ID in active row');
        return;
      }
      // Read approved figure inventory from column ER (148) — set by W2B/W4B pipeline
      var figInvJson = String(postsSheet.getRange(activeRow, 148).getValue() || '').trim();
      var originalSrcs = [];
      if (figInvJson) {
        try {
          var figInv = JSON.parse(figInvJson);
          if (Array.isArray(figInv)) {
            figInv.forEach(function(f) {
              if (f.src && originalSrcs.indexOf(f.src) === -1) originalSrcs.push(f.src);
            });
          }
        } catch(e) {
          results.push('CHECK 3F [Figure inventory]: NOT APPLICABLE — could not parse column ER JSON: ' + e.toString());
          return;
        }
      }

      // Also read original videos from site-export (videos are not tracked in ER)
      var originalVideos = [];
      var exportSheet = ss.getSheetByName('site-export');
      if (exportSheet) {
        var exportData = exportSheet.getDataRange().getValues();
        var expHeaders = exportData[0].map(function(h) { return String(h).trim(); });
        var expIdIdx   = expHeaders.indexOf('ID');
        var expHtmlIdx = expHeaders.indexOf('Full Post HTML');
        var originalHtml = '';
        for (var ei = 1; ei < exportData.length; ei++) {
          if (String(exportData[ei][expIdIdx]).trim().replace(/\.0$/, '') === postId) {
            originalHtml = String(exportData[ei][expHtmlIdx] || '');
            break;
          }
        }
        if (originalHtml) {
          var iframeRe = /src=["']([^"']*(?:youtube|vimeo|wistia|loom|embed)[^"']*)["']/gi;
          var sm;
          while ((sm = iframeRe.exec(originalHtml)) !== null) {
            if (originalVideos.indexOf(sm[1]) === -1) originalVideos.push(sm[1]);
          }
        }
      }

      if (originalSrcs.length === 0 && originalVideos.length === 0) {
        results.push('CHECK 3F [Figure inventory]: NOT APPLICABLE — no images in column ER and no videos in site-export');
        return;
      }
      totalChecks++;
      var IMG_DOMAIN = 'abbeyfloorcare.co.uk';
      var SKIP  = /logo|icon|avatar|sprite|pixel|tracking|David_Allen\.jpg/i;
      var THUMB = /-\d+x\d+(@\d+x)?\.(jpg|jpeg|png|webp|gif|avif)$/i;
      var missingImgs  = originalSrcs.filter(function(src)  { return html.indexOf(src) === -1; });
      var missingVids  = originalVideos.filter(function(src) { return html.indexOf(src) === -1; });
      var newSrcs = [];
      var newSrcRe = /src=["']([^"']+)["']/gi;
      var nsm;
      while ((nsm = newSrcRe.exec(html)) !== null) {
        var ns = nsm[1];
        if (ns.indexOf(IMG_DOMAIN) > -1 && !SKIP.test(ns) && !THUMB.test(ns)) {
          if (newSrcs.indexOf(ns) === -1) newSrcs.push(ns);
        }
      }
      // Exclude W2C-injected images from hallucination check
      // These are identified by data-w4b-skip="true" on their img tag
      var w2cSrcs = [];
      var w2cImgRe = /<img[^>]*data-w4b-skip\s*=\s*["']true["'][^>]*>/gi;
      var w2cMatch;
      while ((w2cMatch = w2cImgRe.exec(html)) !== null) {
        var w2cSrc = w2cMatch[0].match(/src=["']([^"']+)["']/i);
        if (w2cSrc) w2cSrcs.push(w2cSrc[1]);
      }
      var hallucinatedImgs = newSrcs.filter(function(src) {
        return originalSrcs.indexOf(src) === -1 && w2cSrcs.indexOf(src) === -1;
      });
      var newVideos = [];
      var newIframeRe = /src=["']([^"']*(?:youtube|vimeo|wistia|loom|embed)[^"']*)["']/gi;
      while ((nsm = newIframeRe.exec(html)) !== null) {
        if (newVideos.indexOf(nsm[1]) === -1) newVideos.push(nsm[1]);
      }
      var hallucinatedVids = newVideos.filter(function(src) {
        if (originalVideos.indexOf(src) > -1) return false;
        // Exempt intentional W9 video embeds — figure.video-embed is never hallucinated
        var embedPattern = new RegExp('figure[^>]*class="video-embed"[\\s\\S]*?' + src.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
        if (embedPattern.test(html)) return false;
        return true;
      });
      var failures3f = [];
      if (missingImgs.length > 0) {
        failures3f.push('MISSING: ' + missingImgs.length + ' original image(s) deleted: ' +
          missingImgs.map(function(s) { return s.split('/').pop(); }).join(', '));
      }
      if (missingVids.length > 0) {
        failures3f.push('MISSING: ' + missingVids.length + ' original video(s) deleted');
      }
      if (hallucinatedImgs.length > 0) {
        failures3f.push('HALLUCINATED: ' + hallucinatedImgs.length + ' new image(s) added: ' +
          hallucinatedImgs.map(function(s) { return s.split('/').pop(); }).join(', '));
      }
      if (hallucinatedVids.length > 0) {
        failures3f.push('HALLUCINATED: ' + hallucinatedVids.length + ' new video(s) added');
      }
      if (failures3f.length > 0) {
        results.push('CHECK 3F [Figure inventory]: FAIL — ' + failures3f.join(' | '));
      } else {
        var total3f = originalSrcs.length + originalVideos.length;
        results.push('CHECK 3F [Figure inventory]: PASS — all ' + total3f + ' original asset(s) present' +
          (originalVideos.length > 0 ? ' (including ' + originalVideos.length + ' video)' : ''));
        passCount++;
      }
    } catch(e3f) {
      results.push('CHECK 3F [Figure inventory]: ERROR — ' + e3f.toString());
    }
  })();
/* ── CHECK 3G — DIAGNOSTIC CAPTIONS ── */
  (function() {
    try {
      totalChecks++;
      var figRe = /<figure[^>]*class="[^"]*wp-caption[^"]*"[^>]*>[\s\S]*?<\/figure>/gi;
      var figMatch;
      var nonDiagnostic = [];
      var DIAGNOSTIC_SIGNALS = [
        'your ', 'you ', 'this is ', 'like this', 'indicat', 'at this stage',
        'need', 'have ', 'suggests', 'means', 'shows ', 'if your', 'floors ',
        'reveals', 'confirms', 'beneath', 'remains', 'stage need', 'when ',
        'once ', 'after ', 'before ', 'why ', 'what ', 'how '
      ];
      while ((figMatch = figRe.exec(html)) !== null) {
        var figBlock = figMatch[0];
        var capMatch = figBlock.match(/<figcaption[^>]*>([\s\S]*?)<\/figcaption>/i);
        if (!capMatch) continue;
        var capText = capMatch[1].replace(/<[^>]+>/g, '').trim().toLowerCase();
        if (!capText) continue;
        var isDiagnostic = DIAGNOSTIC_SIGNALS.some(function(signal) {
          return capText.indexOf(signal) > -1;
        });
        if (!isDiagnostic) {
          nonDiagnostic.push('"' + capText.substring(0, 60) + '"');
        }
      }
      if (nonDiagnostic.length === 0) {
        results.push('CHECK 3G [Diagnostic captions]: PASS — all wp-caption figcaptions contain diagnostic language');
        passCount++;
      } else {
        results.push('CHECK 3G [Diagnostic captions]: FAIL — ' + nonDiagnostic.length + ' caption(s) appear purely descriptive: ' + nonDiagnostic.join(' | ') + '. Fix: rewrite to connect image to reader\'s problem.');
      }
    } catch(e3g) {
      results.push('CHECK 3G [Diagnostic captions]: ERROR — ' + e3g.toString());
    }
  })();
  /* ── CHECK 2B ── */
  (function() {
    try {
      totalChecks++;
      var bareUrlRe = /<a[^>]+href="([^"]+)"[^>]*>(https?:\/\/[^<]+)<\/a>/gi;
      var bareUrls  = [];
      var bm;
      while ((bm = bareUrlRe.exec(html)) !== null) {
        var anchorText = bm[2].trim();
        if (/^https?:\/\//i.test(anchorText)) {
          bareUrls.push(anchorText.substring(0, 80));
        }
      }
      if (bareUrls.length === 0) {
        results.push('CHECK 2B [Anchor text]: PASS — no bare URL anchor text found');
        passCount++;
      } else {
        results.push('CHECK 2B [Anchor text]: FAIL — ' + bareUrls.length + ' link(s) use bare URL as anchor text: ' +
          bareUrls.slice(0, 3).join(' | ') +
          '. Fix: replace with descriptive anchor text (3–8 words describing the target page).');
      }
    } catch(e2b) {
      results.push('CHECK 2B [Anchor text]: ERROR — ' + e2b.toString());
    }
  })();

  /* ── CHECK 2E ── */
  (function() {
    try {
      totalChecks++;
      if (/<h1[\s>]/i.test(html)) {
        results.push('CHECK 2E [H1 in body]: FAIL — <h1> tag found in body HTML. H1 is set in WordPress — remove it from the article body.');
      } else {
        results.push('CHECK 2E [H1 in body]: PASS — no <h1> tag in body HTML');
        passCount++;
      }
    } catch(e2e) {
      results.push('CHECK 2E [H1 in body]: ERROR — ' + e2e.toString());
    }
  })();

  /* ── CHECK 6 ── */
  const check6 = auditHubIntro(html);
  results.push("CHECK 6 [Hub-intro length]: " + check6.label + " — " + check6.detail);
  if (check6.pass) passCount++;

  /* ── CHECK 8 ── */
  const depthTypes8 = ["Hub Page", "Educational Guide", "Method Guide", "Diagnostic Guide", "Buyer Guide"];
  const check8 = depthTypes8.indexOf(articleType) > -1
    ? auditProcessSections(html)
    : { pass: true, label: "NOT APPLICABLE", detail: "Process section depth not required for " + articleType };
  results.push("CHECK 8 [Process section depth]: " + check8.label + " — " + check8.detail);
  if (check8.pass) passCount++;

  /* ── CHECK 9 ── */
  const check9 = auditFooterCTA(html);
  results.push("CHECK 9 [CTA link]: " + check9.label + " — " + check9.detail);
  if (check9.pass) passCount++;

  /* ── CHECK 3E MOVED — now CHECK 8 in the Stage 3B deferred LLM audit (entity dump / comma-list judgment requires linguistic reasoning, not regex) ── */

  /* ── SCHEMA CHECKS ── */
  if (schema && schema.trim() !== "") {
    results.push("");
    results.push("── SCHEMA AUDIT ─────────────────────────────────");
    totalChecks += 12;

    var schemaJSON = schema.trim();
    var scriptMatch = schemaJSON.match(/<script[^>]*>([\s\S]*?)<\/script>/i);
    if (scriptMatch) schemaJSON = scriptMatch[1].trim();

    const check10 = auditSchemaValidJSON(schemaJSON);
    results.push("CHECK 10 [Schema — valid JSON]: " + check10.label + " — " + check10.detail);
    if (check10.pass) passCount++;

    if (check10.pass) {
      var parsed = JSON.parse(schemaJSON);

      const check11 = auditSchemaContext(parsed);
      results.push("CHECK 11 [Schema — @context]: " + check11.label + " — " + check11.detail);
      if (check11.pass) passCount++;

      const check12 = auditSchemaType(parsed);
      results.push("CHECK 12 [Schema — @type present]: " + check12.label + " — " + check12.detail);
      if (check12.pass) passCount++;

      const check13 = auditSchemaRequiredFields(parsed);
      results.push("CHECK 13 [Schema — required fields]: " + check13.label + " — " + check13.detail);
      if (check13.pass) passCount++;

      const check14 = auditSchemaEncodedChars(schemaJSON);
      results.push("CHECK 14 [Schema — HTML encoding]: " + check14.label + " — " + check14.detail);
      if (check14.pass) passCount++;

      const check15 = auditSchemaURLs(parsed);
      results.push("CHECK 15 [Schema — URL format]: " + check15.label + " — " + check15.detail);
      if (check15.pass) passCount++;

      const check16 = auditSchemaBracketURLs(schemaJSON);
      results.push("CHECK 16 [Schema — bracket-wrapped URLs]: " + check16.label + " — " + check16.detail);
      if (check16.pass) passCount++;

      const check20 = auditSchemaPublisherAnchor(parsed);
      results.push("CHECK 20 [Schema — publisher @id anchor]: " + check20.label + " — " + check20.detail);
      if (check20.pass) passCount++;

      const check21 = auditSchemaAuthorPerson(parsed);
      results.push("CHECK 21 [Schema — author @type Person]: " + check21.label + " — " + check21.detail);
      if (check21.pass) passCount++;

      const check22 = auditSchemaNoUrlOnArticle(parsed);
      results.push("CHECK 22 [Schema — no url field on Article]: " + check22.label + " — " + check22.detail);
      if (check22.pass) passCount++;

      const check23 = auditSchemaHasPartFragmentTypes(parsed);
      results.push("CHECK 23 [Schema — hasPart fragment/@type]: " + check23.label + " — " + check23.detail);
      if (check23.pass) passCount++;

      const check24 = auditSchemaRootFields(parsed);
      results.push("CHECK 24 [Schema — illegal root fields]: " + check24.label + " — " + check24.detail);
      if (check24.pass) passCount++;

    } else {
      results.push("CHECK 11–16, 20–24 [Schema]: BLOCKED — fix invalid JSON (Check 10) first.");
    }
  }

  /* ── Summary ── */
  const allPass = (passCount === totalChecks);
  results.push("");
  results.push("MECHANICAL AUDIT: " + passCount + "/" + totalChecks + " PASS");
  if (allPass) {
    results.push("STATUS: CLEAR — proceed to Stage 3B LLM judgment audit.");
  } else {
    results.push("STATUS: FIX REQUIRED — correct failures before Stage 3B.");
  }

  results.push("");
  results.push("DEFERRED TO STAGE 3B (LLM):");
  results.push("  Check 1  — Cluster classification and opening tone");
  results.push("  Check 4  — Section opening tone (symptom first)");
  results.push("  Check 5  — Cluster tone across full header");
  results.push("  Check 7  — Named defect completion");
  if (schema && schema.trim() !== "") {
    results.push("  Check 17 — Schema type vs article type (Schema Alignment Validation Sheet)");
    results.push("  Check 18 — Prohibited schema types absent");
    results.push("  Check 19 — Entity role fields match governed entity");
  }

  return results.join("\n");
}

/* ============================================================
   CHECK 2 (SPOKE) — INTERNAL LINKS
============================================================ */
function auditInternalLinks(html, hubUrl, articleType) {
  const governedMaxLinks = getMaxLateralLinks(articleType);
  const MIN_LINKS = (governedMaxLinks !== null) ? governedMaxLinks : 5;
  const MAX_LINKS = (governedMaxLinks !== null) ? governedMaxLinks : 8;
  const internalRe = /href="(https?:\/\/(?:www\.)?abbeyfloorcare\.co\.uk[^"]*)"/gi;
  const internalLinks = [];
  let m;
  while ((m = internalRe.exec(html)) !== null) {
    internalLinks.push(m[1]);
  }
  const uniqueLinks = internalLinks.filter(function(l, i, a) { return a.indexOf(l) === i; });

  if (internalLinks.length === 0) {
    return { pass: false, label: "FAIL", detail: "No internal abbeyfloorcare.co.uk links found. This article requires " + MIN_LINKS + "-" + MAX_LINKS + " contextual internal links from the silo map, plus the hub link." };
  }
  let hubFound = false;
  if (hubUrl) {
    const hubNorm = hubUrl.replace(/#.*$/, "").replace(/[/]+$/, "").toLowerCase();
    hubFound = internalLinks.some(function(l) {
      const lNorm = l.replace(/#.*$/, "").replace(/[/]+$/, "").toLowerCase();
      return lNorm === hubNorm || lNorm.indexOf(hubNorm) > -1 || hubNorm.indexOf(lNorm) > -1;
    });
  }
  const lateralLinks = internalLinks.filter(function(l) {
    if (!hubUrl) return true;
    const hubNorm = hubUrl.replace(/#.*$/, "").replace(/[/]+$/, "").toLowerCase();
    const lNorm = l.replace(/#.*$/, "").replace(/[/]+$/, "").toLowerCase();
    return lNorm.indexOf(hubNorm) === -1 && hubNorm.indexOf(lNorm) === -1;
  });
  const uniqueLateralLinks = lateralLinks.filter(function(l, i, a) { return a.indexOf(l) === i; });

  const failures = [];
  if (hubUrl && !hubFound) failures.push("Hub page link missing — expected link to: " + hubUrl);
  if (uniqueLateralLinks.length === 0) failures.push("No lateral internal link found — add contextual links to related pages in the silo");
  if (uniqueLinks.length < MIN_LINKS) failures.push("Only " + uniqueLinks.length + " internal link(s) found — Stage 2A requires " + MIN_LINKS + "-" + MAX_LINKS + " contextual internal links from the silo map (hub link + lateral links combined)");
  if (uniqueLinks.length > MAX_LINKS) failures.push(uniqueLinks.length + " internal links found — exceeds the maximum of " + MAX_LINKS);

  const selfLinkRe = /href="([^"]+)"/gi;
  const selfLinks = [];
  let slm;
  while ((slm = selfLinkRe.exec(html)) !== null) {
    const href = slm[1].replace(/#.*$/, "").replace(/[/]+$/, "").toLowerCase();
    const ss2  = SpreadsheetApp.getActiveSpreadsheet();
    const sheet2 = ss2.getSheetByName("posts");
    const activeRow2 = sheet2.getActiveRange().getRow();
    const ownUrl2 = activeRow2 >= 2
      ? String(sheet2.getRange(activeRow2, 3).getValue() || "").trim().replace(/\/$/, "").toLowerCase()
      : "";
    if (ownUrl2 && href === ownUrl2) selfLinks.push(slm[1]);
  }
  if (selfLinks.length > 0) {
    failures.push("Self-referencing link(s) detected: " + selfLinks.slice(0, 3).join(", "));
  }
  if (failures.length > 0) return { pass: false, label: "FAIL", detail: failures.join(" | ") };
  const hubLabel = hubFound ? "Hub link ✓" : "Hub URL not set — skipped";
  return { pass: true, label: "PASS", detail: hubLabel + " | " + uniqueLinks.length + " total internal link(s), " + uniqueLateralLinks.length + " lateral" };
}

/* ============================================================
   CHECK 2 (HUB) — QUICK-LINK ANCHORS
============================================================ */
function auditQuickLinks(html) {
  // Check 1 — nav anchor links in hub-intro
  const sectionIds = [];
  const sectionIdRe = /<section[^>]+id="([^"]+)"/gi;
  let m;
  while ((m = sectionIdRe.exec(html)) !== null) sectionIds.push(m[1].toLowerCase());
  const quickLinkRe = /href="#([^"]+)"/gi;
  const linkedAnchors = [];
  while ((m = quickLinkRe.exec(html)) !== null) linkedAnchors.push(m[1].toLowerCase());
  if (linkedAnchors.length === 0) return { pass: false, label: "FAIL", detail: "No quick-link anchors found." };
  if (sectionIds.length === 0) return { pass: false, label: "FAIL", detail: "No section id attributes found." };
  const broken = linkedAnchors.filter(function(a) { return sectionIds.indexOf(a) === -1; });
  if (broken.length > 0) return { pass: false, label: "FAIL", detail: "Broken anchor(s): #" + broken.join(", #") };

  // Check 2 — contextual silo map links in section body content (excludes hub-intro nav)
  // Strip hub-intro section before counting contextual links
  var bodyHtml = html.replace(/<section[^>]*id="hub-intro"[^>]*>[\s\S]*?<\/section>/i, "");
  const internalRe = /href="(https?:\/\/(?:www\.)?abbeyfloorcare\.co\.uk[^"]*)"/gi;
  const internalLinks = [];
  while ((m = internalRe.exec(bodyHtml)) !== null) {
    internalLinks.push(m[1]);
  }
  const uniqueLinks = internalLinks.filter(function(l, i, a) { return a.indexOf(l) === i; });

  // Dynamic link ceiling — count sections in approved plan
  var sectionCount = (html.match(/<section[^>]+id="section-\d+"/gi) || []).length;
  var minLinks = 5;
  var maxLinks = Math.max(8, sectionCount);

  const failures = [];
  if (broken.length > 0) failures.push("Broken nav anchor(s): #" + broken.join(", #"));
  if (uniqueLinks.length < minLinks) failures.push("Only " + uniqueLinks.length + " contextual silo link(s) found in body — minimum " + minLinks + " required");
  if (uniqueLinks.length > maxLinks) failures.push(uniqueLinks.length + " contextual silo links found — maximum " + maxLinks + " permitted (based on " + sectionCount + " sections)");

  if (failures.length > 0) return { pass: false, label: "FAIL", detail: failures.join(" | ") };
  return {
    pass: true,
    label: "PASS",
    detail: linkedAnchors.length + " nav anchor(s) valid: #" + linkedAnchors.join(", #") +
            " | " + uniqueLinks.length + " contextual silo link(s) in body content"
  };
}

/* ============================================================
   CHECK 3 — CONTENT INTEGRITY
============================================================ */
function auditHtmlStructure(html, articleType) {
  articleType = articleType || "General";
  const failures = [];
  const headingRe = /<h([2-4])[^>]*>/gi;
  let hm;
  let lastLevel = 1;
  let hierarchyErrors = [];
  while ((hm = headingRe.exec(html)) !== null) {
    const level = parseInt(hm[1], 10);
    if (level > lastLevel + 1) hierarchyErrors.push("H" + level + " follows H" + lastLevel);
    lastLevel = level;
  }
  if (hierarchyErrors.length > 0) failures.push("3a: Heading hierarchy skip — " + hierarchyErrors.join(", "));
  const h2Re = /<h2[^>]*>([\s\S]*?)<\/h2>/gi;
  const h2Texts = [];
  let h2m;
  while ((h2m = h2Re.exec(html)) !== null) {
    const text = h2m[1].replace(/<[^>]+>/g, "").trim().toLowerCase();
    if (h2Texts.indexOf(text) > -1) failures.push("3b: Duplicate H2 — [" + text + "]");
    else h2Texts.push(text);
  }
  const figureRe = /<figure[\s\S]*?<\/figure>/gi;
  let fm;
  let figCount = 0;
  let missingCaption = 0;
  while ((fm = figureRe.exec(html)) !== null) {
    const figBlock = fm[0];
    const altMatch = figBlock.match(/alt="([^"]*)"/i);
    const isDecorative = altMatch && altMatch[1].trim() === "";
    const isDavidAllen = /David_Allen\.jpg/i.test(figBlock);
    if (!isDecorative && !isDavidAllen) {
      figCount++;
      if (!/<figcaption/i.test(figBlock)) missingCaption++;
    }
  }
  if (missingCaption > 0) failures.push("3c: " + missingCaption + " figure(s) missing <figcaption>");
  if (failures.length === 0) {
    return { pass: true, label: "PASS", detail: "Heading hierarchy ✓ | No duplicate H2s ✓ | " + figCount + " figure(s) captioned ✓" };
  }
  return { pass: false, label: "FAIL", detail: failures.join(" | ") };
}

/* ============================================================
   CHECK 6 — HUB-INTRO ≤ 15 WORDS
============================================================ */
function auditHubIntro(html) {
  const hubIntroMatch = html.match(/<section[^>]*(?:hub-intro|hub-navigation|How to use)[^>]*>([\s\S]*?)<\/section>/i);
  if (!hubIntroMatch) return { pass: true, label: "NOT APPLICABLE", detail: "No hub-intro section found" };
  const sectionContent = hubIntroMatch[1];
  const firstPMatch = sectionContent.match(/<p[^>]*>([\s\S]*?)<\/p>/i);
  if (!firstPMatch) return { pass: false, label: "FAIL", detail: "Hub-intro section contains no <p> element" };
  const rawText = firstPMatch[1].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim();
  const words   = rawText.split(" ").filter(function(w) { return w.length > 0; });
  const count   = words.length;
  if (count <= 15) return { pass: true, label: "PASS", detail: "Orienting sentence is " + count + " words: \"" + rawText + "\"" };
  return { pass: false, label: "FAIL", detail: "Orienting sentence is " + count + " words (limit 15): \"" + rawText + "\"" };
}

/* ============================================================
   CHECK 8 — PROCESS SECTIONS ≥ 3 PARAGRAPHS BEFORE FIGURE
============================================================ */
function auditProcessSections(html) {
  const PROCESS_KEYWORDS = ["cleaning","honing","polishing","sealing","restoration","burnishing","stripping","re-grouting","recolouring","regrout","recolour","repair","treatment"];
  const sectionRe = /<section[^>]+id="([^"]+)"[^>]*>([\s\S]*?)<\/section>/gi;
  const failures  = [];
  const checked   = [];
  let m;
  while ((m = sectionRe.exec(html)) !== null) {
    const sectionId      = m[1].toLowerCase();
    const sectionContent = m[2];
    const isProcess = PROCESS_KEYWORDS.some(function(kw) { return sectionId.indexOf(kw) !== -1; });
    if (!isProcess) continue;
    checked.push(sectionId);
    const figureIndex = sectionContent.search(/<figure/i);
    if (figureIndex === -1) {
      const pCount = (sectionContent.match(/<p[\s>]/gi) || []).length;
      if (pCount < 3) failures.push("#" + sectionId + " has " + pCount + "/3 required paragraphs (no figure present)");
    } else {
      const contentBeforeFigure = sectionContent.substring(0, figureIndex);
      const pCount = (contentBeforeFigure.match(/<p[\s>]/gi) || []).length;
      if (pCount < 3) {
        const totalP = (sectionContent.match(/<p[\s>]/gi) || []).length;
        const contentAfterFigure = sectionContent.substring(figureIndex);
        const isAtEnd = contentAfterFigure.length < (sectionContent.length * 0.25);
        if (!(isAtEnd && totalP >= 3)) failures.push("#" + sectionId + " has " + pCount + "/3 required paragraphs before <figure>");
      }
    }
  }
  if (checked.length === 0) return { pass: true, label: "NOT APPLICABLE", detail: "No process sections found by id keyword matching" };
  if (failures.length === 0) return { pass: true, label: "PASS", detail: "All " + checked.length + " process section(s) have ≥ 3 paragraphs: " + checked.map(function(id){ return "#" + id; }).join(", ") };
  return { pass: false, label: "FAIL", detail: failures.join(" | ") };
}

/* ============================================================
   CHECK 9 — CTA LINK
============================================================ */
function auditFooterCTA(html) {
  const ctaRe = /<a[^>]+href="([^"]*contact[^"]*)"/i;
  const ctaMatch = html.match(ctaRe);
  if (ctaMatch) return { pass: true, label: "PASS", detail: "Contact link found: " + ctaMatch[1] };
  const anyLink = html.match(/<a[^>]+href="(https?:\/\/(?:www\.)?abbeyfloorcare\.co\.uk[^"]*)"/i);
  if (anyLink) return { pass: true, label: "PASS", detail: "Internal link present — no /contact link found, consider adding CTA" };
  return { pass: false, label: "FAIL", detail: "No links found in HTML — add a CTA or internal link" };
}

/* ============================================================
   CHECK 10 — VALID JSON
============================================================ */
function auditSchemaValidJSON(schemaJSON) {
  try {
    JSON.parse(schemaJSON);
    return { pass: true, label: "PASS", detail: "Schema parses as valid JSON" };
  } catch(e) {
    return { pass: false, label: "FAIL", detail: "Invalid JSON — " + e.message };
  }
}

/* ============================================================
   CHECK 11 — @context
============================================================ */
function auditSchemaContext(parsed) {
  var ctx = parsed["@context"] || "";
  if (ctx === "https://schema.org") return { pass: true, label: "PASS", detail: "@context is correct" };
  return { pass: false, label: "FAIL", detail: "@context is \"" + ctx + "\" — must be exactly \"https://schema.org\"" };
}

/* ============================================================
   CHECK 12 — @type PRESENT
============================================================ */
function auditSchemaType(parsed) {
  var graph = parsed["@graph"];
  if (!Array.isArray(graph) || graph.length === 0) return { pass: false, label: "FAIL", detail: "@graph is missing or empty." };
  var missing = [];
  graph.forEach(function(entity, i) {
    var t = entity["@type"] || "";
    if (!t || t.toString().trim() === "") missing.push("entity[" + i + "]");
  });
  if (missing.length === 0) return { pass: true, label: "PASS", detail: "@type present on all " + graph.length + " graph entities" };
  return { pass: false, label: "FAIL", detail: "@type missing on graph " + missing.join(", ") };
}

/* ============================================================
   CHECK 13 — REQUIRED FIELDS
============================================================ */
function auditSchemaRequiredFields(parsed) {
  var graph = parsed["@graph"];
  if (!Array.isArray(graph) || graph.length === 0) return { pass: false, label: "FAIL", detail: "@graph missing." };
  var primary = null;
  for (var i = 0; i < graph.length; i++) {
    var t = graph[i]["@type"] || "";
    if (t === "Article" || t === "Service") { primary = graph[i]; break; }
  }
  if (!primary) return { pass: false, label: "FAIL", detail: "No Article or Service entity found in @graph." };
  var isService = primary["@type"] === "Service";
  var missing = [];
  if (isService) { if (!primary["name"] || primary["name"].toString().trim() === "") missing.push("name"); }
  else { if (!primary["headline"] || primary["headline"].toString().trim() === "") missing.push("headline"); }
  if (!primary["description"] || primary["description"].toString().trim() === "") missing.push("description");
  if (!primary["mainEntityOfPage"]) missing.push("mainEntityOfPage");
  if (missing.length === 0) return { pass: true, label: "PASS", detail: "Required fields present on primary entity" };
  return { pass: false, label: "FAIL", detail: "Missing required fields: " + missing.join(", ") };
}

/* ============================================================
   CHECK 14 — HTML ENCODED CHARS
============================================================ */
function auditSchemaEncodedChars(schemaJSON) {
  var encoded = schemaJSON.match(/(%22|%3A|%2F|%7B|%7D|&amp;|&quot;)/g);
  if (!encoded) return { pass: true, label: "PASS", detail: "No HTML-encoded characters found" };
  var unique = encoded.filter(function(v, i, a) { return a.indexOf(v) === i; });
  return { pass: false, label: "FAIL", detail: "HTML-encoded characters found: " + unique.join(", ") };
}

/* ============================================================
   CHECK 15 — URL DOMAIN
   EXEMPT prefixes (must not be flagged as non-domain URLs):
     - https://schema.org and http://schema.org (exact match)
     - https://schema.org/ (enumeration values e.g. InStock)
     - https://www.google.com/maps/place/ (sameAs GBP Place ID URL)
============================================================ */
function auditSchemaURLs(parsed) {
  var badURLs = [];

  function isExempt(url) {
    // Exact schema.org root
    if (url === "https://schema.org" || url === "http://schema.org") return true;
    // schema.org enumeration values (e.g. https://schema.org/InStock)
    if (url.indexOf("https://schema.org/") === 0) return true;
    if (url.indexOf("http://schema.org/") === 0) return true;
    // Google Business Profile sameAs URL
    if (url.indexOf("https://www.google.com/maps/place/") === 0) return true;
    // YouTube URLs — valid in VideoObject schema
    if (url.indexOf("https://www.youtube.com/") === 0) return true;
    if (url.indexOf("https://img.youtube.com/") === 0) return true;
    return false;
  }

  function walk(obj) {
    if (typeof obj === "string") {
      if (obj.indexOf("http") === 0 &&
          obj.indexOf("https://www.abbeyfloorcare.co.uk") !== 0 &&
          !isExempt(obj)) {
        badURLs.push(obj.substring(0, 80));
      }
    } else if (typeof obj === "object" && obj !== null) {
      Object.keys(obj).forEach(function(k) { walk(obj[k]); });
    }
  }
  walk(parsed);

  if (badURLs.length === 0) return { pass: true, label: "PASS", detail: "All URLs match abbeyfloorcare.co.uk domain or are exempt schema/maps references" };
  return { pass: false, label: "FAIL", detail: "Non-domain URLs found: " + badURLs.join(" | ") + ". Fix: replace with correct abbeyfloorcare.co.uk URLs." };
}

/* ============================================================
   CHECK 16 — BRACKET-WRAPPED URLS
============================================================ */
function auditSchemaBracketURLs(schemaJSON) {
  var brackets = schemaJSON.match(/"\[https?:\/\//g);
  if (!brackets) return { pass: true, label: "PASS", detail: "No bracket-wrapped URLs found" };
  return { pass: false, label: "FAIL", detail: brackets.length + " bracket-wrapped URL(s) found. Fix: remove leading [ bracket." };
}

/* ============================================================
   CHECK 20 — PUBLISHER @ID ANCHOR
============================================================ */
function auditSchemaPublisherAnchor(parsed) {
  const EXPECTED = "https://www.abbeyfloorcare.co.uk/#localbusiness";
  const graph = parsed["@graph"] || (parsed["@type"] ? [parsed] : []);
  const problems = [];
  graph.forEach(function(entity) {
    if (!entity.publisher) return;
    const pub   = entity.publisher;
    const pubId = (typeof pub === "object") ? (pub["@id"] || "") : "";
    if (pubId && pubId !== EXPECTED) problems.push("publisher @id is \"" + pubId + "\" — expected \"" + EXPECTED + "\"");
  });
  if (problems.length === 0) return { pass: true, label: "PASS", detail: "publisher @id matches #localbusiness anchor" };
  return { pass: false, label: "FAIL", detail: problems.join("; ") };
}

/* ============================================================
   CHECK 21 — AUTHOR @TYPE PERSON
============================================================ */
function auditSchemaAuthorPerson(parsed) {
  const graph = parsed["@graph"] || (parsed["@type"] ? [parsed] : []);
  const problems = [];
  graph.forEach(function(entity) {
    if (!entity.author) return;
    const author     = entity.author;
    const authorType = (typeof author === "object") ? (author["@type"] || "") : "";
    if (authorType && authorType !== "Person") problems.push("author @type is \"" + authorType + "\" — expected \"Person\"");
    const authorName = (typeof author === "object") ? (author["name"] || "") : "";
    if (authorType === "Person" && authorName !== "David Allen") problems.push("author name is \"" + authorName + "\" — expected \"David Allen\"");
  });
  if (problems.length === 0) return { pass: true, label: "PASS", detail: "author @type is Person with correct name" };
  return { pass: false, label: "FAIL", detail: problems.join("; ") };
}

/* ============================================================
   CHECK 22 — NO URL FIELD ON ARTICLE
============================================================ */
function auditSchemaNoUrlOnArticle(parsed) {
  const graph = parsed["@graph"] || (parsed["@type"] ? [parsed] : []);
  const problems = [];
  graph.forEach(function(entity) {
    const types = Array.isArray(entity["@type"]) ? entity["@type"] : [entity["@type"]];
    if (types.includes("Article") && entity.hasOwnProperty("url")) {
      problems.push("Article entity has a direct 'url' field — remove it");
    }
  });
  if (problems.length === 0) return { pass: true, label: "PASS", detail: "No redundant url field on Article entity" };
  return { pass: false, label: "FAIL", detail: problems.join("; ") };
}

/* ============================================================
   CHECK 23 — HASPART FRAGMENT TYPES
============================================================ */
function auditSchemaHasPartFragmentTypes(parsed) {
  const graph = parsed["@graph"] || [];
  const problems = [];
  graph.forEach(function(entity) {
    if (!entity.hasPart) return;
    const parts = Array.isArray(entity.hasPart) ? entity.hasPart : [entity.hasPart];
    parts.forEach(function(part) {
      const id   = part["@id"] || "";
      const type = part["@type"] || "";
      if (id.indexOf("#article") > -1 && type !== "Article") problems.push("hasPart @id ends #article but @type is \"" + type + "\" — expected Article");
      if (id.indexOf("#service") > -1 && type !== "WebPage")  problems.push("hasPart @id ends #service but @type is \"" + type + "\" — expected WebPage");
    });
  });
  if (problems.length === 0) return { pass: true, label: "PASS", detail: "hasPart fragment/@type combinations are correct" };
  return { pass: false, label: "FAIL", detail: problems.join(" | ") };
}

/* ============================================================
   CHECK 24 — ILLEGAL ROOT FIELDS
============================================================ */
function auditSchemaRootFields(parsed) {
  const ALLOWED = ['@context', '@graph'];
  const found   = Object.keys(parsed);
  const illegal = found.filter(function(key) { return ALLOWED.indexOf(key) === -1; });
  if (illegal.length === 0) return { pass: true, label: 'PASS', detail: 'Root level contains only @context and @graph' };
  return { pass: false, label: 'FAIL', detail: 'Illegal root-level fields: [' + illegal.join(', ') + ']. Only "@context" and "@graph" permitted at root.' };
}

/* ============================================================
   DEFERRED CHECK PROMPT BUILDER (Checks 1, 4, 5, 7)
   Called after mechanical audit passes.
   Returns a prompt string for ChatGPT to evaluate editorial
   checks that cannot be done mechanically.
============================================================ */
function buildDeferredCheckPrompt(html) {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheet       = ss.getSheetByName('posts');
  var row         = sheet.getActiveRange().getRow();
  var articleType = String(sheet.getRange(row, 8).getValue() || 'General').trim();
  var d           = getActiveRowDataMap();
  var cluster     = String(d['Cluster'] || '').trim();
  var material    = String(d['Stone Type'] || '').trim();

  // Extract header block
  var headerMatch = html.match(/<header[^>]*>([\s\S]*?)<\/header>/i);
  var headerHtml  = headerMatch ? headerMatch[0] : '';

  // Extract section 1 opening paragraph
  var sec1Match   = html.match(/<section[^>]*id="section-1"[^>]*>([\s\S]*?)<\/section>/i);
  var sec1Opening = '';
  if (sec1Match) {
    var firstP = sec1Match[1].match(/<p[^>]*>([\s\S]*?)<\/p>/i);
    if (firstP) sec1Opening = firstP[0];
  }

  // Extract all H2s and their first paragraphs
  var sectionOpeners = [];
  var secRe = /<section[^>]*id="([^"]+)"[^>]*>([\s\S]*?)<\/section>/gi;
  var sm;
  while ((sm = secRe.exec(html)) !== null) {
    var h2m = sm[2].match(/<h2[^>]*>([\s\S]*?)<\/h2>/i);
    var pm  = sm[2].match(/<p[^>]*>([\s\S]*?)<\/p>/i);
    if (h2m && pm) {
      sectionOpeners.push({
        id:    sm[1],
        h2:    h2m[1].replace(/<[^>]+>/g, '').trim(),
        opens: pm[1].replace(/<[^>]+>/g, '').trim().substring(0, 200)
      });
    }
  }

  var sectionOpenerBlock = sectionOpeners.map(function(s) {
    return 'Section "' + s.id + '" — H2: "' + s.h2 + '"\nOpening: "' + s.opens + '"';
  }).join('\n\n');

  return 'DEFERRED EDITORIAL AUDIT — CHECKS 1, 4, 5, 7\n' +
    'ROLE: Senior UK SEO & Editorial Quality Auditor\n\n' +
    'ARTICLE TYPE: ' + articleType + '\n' +
    'MATERIAL: ' + material + '\n' +
    'CLUSTER: ' + (cluster || 'not set — derive from header tone') + '\n\n' +
    'Return ONLY the four check results in the exact format shown below.\n' +
    'No preamble. No explanation. No markdown.\n\n' +
    '------------------------------------------------------------\n' +
    'CHECK 1 — CLUSTER CLASSIFICATION AND OPENING TONE\n' +
    '------------------------------------------------------------\n' +
    'Read the header paragraph below. Determine which cluster the opening tone matches:\n' +
    'CLUSTER A: Opens with reassurance — floor can be saved, problem is not permanent\n' +
    'CLUSTER B: Opens with confidence — problem is correctable, pigment/colour is recoverable\n' +
    'CLUSTER C: Opens with authority — specialist knowledge is required, wrong approach causes damage\n\n' +
    'HEADER:\n' + headerHtml + '\n\n' +
    'Does the opening tone match the stated cluster? If no cluster is stated, name the cluster you detect.\n\n' +
    'OUTPUT FORMAT:\n' +
    'CHECK 1: PASS|FAIL|REVIEW\n' +
    'CLUSTER DETECTED: [A/B/C/D]\n' +
    'OBSERVATION: [one sentence]\n\n' +
    '------------------------------------------------------------\n' +
    'CHECK 4 — SECTION OPENING TONE (SYMPTOM FIRST)\n' +
    '------------------------------------------------------------\n' +
    'Each section opening sentence must orient the reader to their problem before introducing mechanism or technical terms.\n' +
    'Review the section openers below. Flag any section whose first sentence opens with mechanism, entity name, or process rather than homeowner symptom.\n\n' +
    'SECTION OPENERS:\n' + sectionOpenerBlock + '\n\n' +
    'OUTPUT FORMAT:\n' +
    'CHECK 4: PASS|FAIL|REVIEW\n' +
    'OBSERVATION: [one sentence — name any failing sections by id]\n\n' +
    '------------------------------------------------------------\n' +
    'CHECK 5 — CLUSTER TONE ACROSS FULL HEADER\n' +
    '------------------------------------------------------------\n' +
    'Read the full header block. Check that the cluster tone is maintained across ALL sentences — not just the first.\n' +
    'CLUSTER B headers must not contain comfort language anywhere: "reassuring", "the good news is", "fortunately", "rest assured".\n' +
    'CLUSTER A headers must not open with damage or mechanism.\n' +
    'CLUSTER C headers must not qualify expertise or lead with homeowner uncertainty.\n\n' +
    'HEADER:\n' + headerHtml + '\n\n' +
    'OUTPUT FORMAT:\n' +
    'CHECK 5: PASS|FAIL|REVIEW\n' +
    'OBSERVATION: [one sentence]\n\n' +
    '------------------------------------------------------------\n' +
    'CHECK 7 — NAMED DEFECT COMPLETION\n' +
    '------------------------------------------------------------\n' +
    'Any named technical defect must be explained to completion before the next paragraph:\n' +
    '  ELEMENT 1: What it is (definition)\n' +
    '  ELEMENT 2: What the homeowner sees (symptom)\n' +
    '  ELEMENT 3: What is done about it (correction)\n\n' +
    'Scan the full article HTML below for named defects.\n' +
    'Named defects include: delamination, sealer failure, efflorescence, filler collapse, colour loss, grout haze, lippage, spalling, micro-scratching, residue lock-in.\n\n' +
    'ARTICLE HTML:\n' + html + '\n\n' +
    'OUTPUT FORMAT:\n' +
    'CHECK 7: PASS|FAIL|REVIEW\n' +
    'OBSERVATION: [one sentence — name any incomplete defects]\n\n' +
    '------------------------------------------------------------\n' +
    'CHECK 8 — ENTITY DUMP / COMMA-LIST PROSE CHECK\n' +
    '------------------------------------------------------------\n' +
    'Some paragraphs list several technical terms separated by commas. This is only a problem if the\n' +
    'paragraph reads as a bare list of nouns with no verb anywhere in the sentence — e.g.\n' +
    '"Cracks, chips, worn lanes, failed coatings, dark grout" with no verb attached.\n' +
    'It is NOT a problem if the terms sit inside a normal sentence with a verb anywhere in it, even if\n' +
    'the verb comes after the list — e.g. "Cracks, chips, worn lanes, failed coatings, and dark grout\n' +
    'all influence each other" is normal prose, not a dump.\n\n' +
    'Scan the full article HTML above for any paragraph containing 5 or more comma-separated technical\n' +
    'or descriptive terms with NO verb anywhere in that sentence.\n\n' +
    'OUTPUT FORMAT:\n' +
    'CHECK 8: PASS|FAIL|REVIEW\n' +
    'OBSERVATION: [one sentence — quote the first few words of any genuine dump found, or state none found]\n\n' +
    '------------------------------------------------------------\n' +
    'REQUIRED OUTPUT — return exactly these lines, nothing else:\n' +
    'CHECK 1: [PASS|FAIL|REVIEW]\n' +
    'CLUSTER DETECTED: [A/B/C/D]\n' +
    'OBSERVATION 1: [one sentence]\n' +
    'CHECK 4: [PASS|FAIL|REVIEW]\n' +
    'OBSERVATION 4: [one sentence]\n' +
    'CHECK 5: [PASS|FAIL|REVIEW]\n' +
    'OBSERVATION 5: [one sentence]\n' +
    'CHECK 7: [PASS|FAIL|REVIEW]\n' +
    'OBSERVATION 7: [one sentence]\n' +
    'CHECK 8: [PASS|FAIL|REVIEW]\n' +
    'OBSERVATION 8: [one sentence]';
}

/* ============================================================
   PARSE DEFERRED CHECK RESPONSE
   Parses ChatGPT response from deferred checks.
   Returns structured result object.
============================================================ */
function parseDeferredCheckResponse(response) {
  var result = {
    check1: '', cluster: '', obs1: '',
    check4: '', obs4: '',
    check5: '', obs5: '',
    check7: '', obs7: '',
    check8: '', obs8: '',
    allPass: false,
    hasFailures: false
  };

  var normalised = response
    .replace(/\r?\n/g, ' ')
    .replace(/(CHECK \d+:|CLUSTER DETECTED:|OBSERVATION \d+:)/g, '\n$1')
    .trim();
  var lines = normalised.split('\n');
  lines.forEach(function(line) {
    var t = line.trim();
    if (/^CHECK 1:/i.test(t))         result.check1  = t.replace(/^CHECK 1:\s*/i, '').trim().toUpperCase();
    if (/^CLUSTER DETECTED:/i.test(t)) result.cluster = t.replace(/^CLUSTER DETECTED:\s*/i, '').trim();
    if (/^OBSERVATION 1:/i.test(t))    result.obs1    = t.replace(/^OBSERVATION 1:\s*/i, '').trim();
    if (/^CHECK 4:/i.test(t))          result.check4  = t.replace(/^CHECK 4:\s*/i, '').trim().toUpperCase();
    if (/^OBSERVATION 4:/i.test(t))    result.obs4    = t.replace(/^OBSERVATION 4:\s*/i, '').trim();
    if (/^CHECK 5:/i.test(t))          result.check5  = t.replace(/^CHECK 5:\s*/i, '').trim().toUpperCase();
    if (/^OBSERVATION 5:/i.test(t))    result.obs5    = t.replace(/^OBSERVATION 5:\s*/i, '').trim();
    if (/^CHECK 7:/i.test(t))          result.check7  = t.replace(/^CHECK 7:\s*/i, '').trim().toUpperCase();
    if (/^OBSERVATION 7:/i.test(t))    result.obs7    = t.replace(/^OBSERVATION 7:\s*/i, '').trim();
    if (/^CHECK 8:/i.test(t))          result.check8  = t.replace(/^CHECK 8:\s*/i, '').trim().toUpperCase();
    if (/^OBSERVATION 8:/i.test(t))    result.obs8    = t.replace(/^OBSERVATION 8:\s*/i, '').trim();
  });

  result.hasFailures = (result.check1 === 'FAIL' || result.check4 === 'FAIL' ||
                        result.check5 === 'FAIL' || result.check7 === 'FAIL' ||
                        result.check8 === 'FAIL');
  result.allPass = !result.hasFailures &&
                   result.check1 !== '' && result.check4 !== '' &&
                   result.check5 !== '' && result.check7 !== '' &&
                   result.check8 !== '';

  return result;
}