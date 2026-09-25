/**
 * ================================================================================
 * ce_W5C_SchemaBuilder.gs - W5C SCHEMABUILDER
 * ================================================================================
 *
 * Deterministic schema generation and location context
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * Updated: March 2026
 *
 * CHANGES FROM PREVIOUS VERSION:
 *   - `mentions` added to all Article-based primary entities
 *     Populated from Col 6 Supporting Entities Core (comma-separated)
 *     Each entity mapped to {"@type":"Thing","name":"..."}
 *   - `about` made precise — uses Primary Entity (Col 2) not generic stone type
 *     Fires for isPartOf pages (Educational Guide, Method Guide etc.)
 *   - `sameAs` added to LocalBusiness node
 *     Points to canonical Google Business Profile Place ID URL
 *   - `hasOfferCatalog` added to LocalBusiness node
 *     Lists 10 core services matching GBP service categories
 *   - Fragment IDs now read from SAV sheet Hub Spoke Fragment ID column
 *     Falls back to hardcoded #article / #service if column not present
 * ================================================================================
 */

/* ============================================================
   W5C PRE-FLIGHT
============================================================ */
function getW5CPreflightData() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { needsParentArea: false };

    var articleType = String(sheet.getRange(row, 8).getValue() || "").trim();
    if (articleType !== "Geo Service Page") return { needsParentArea: false };

    var parentArea = String(sheet.getRange(row, 116).getValue() || "").trim();
    if (parentArea) return { needsParentArea: false };

    var pageUrl = String(sheet.getRange(row, 3).getValue() || "").trim();
    return { needsParentArea: true, pageUrl: pageUrl };
  } catch(e) {
    return { needsParentArea: false };
  }
}

function pushParentAreaToSheet(locality, parentArea) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("posts");
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };
    var locCell = sheet.getRange(row, 115);
    locCell.setNumberFormat("@");
    locCell.setValue((locality || "").trim());
    var parCell = sheet.getRange(row, 116);
    parCell.setNumberFormat("@");
    parCell.setValue((parentArea || "").trim());
    return { success: true, message: "Saved — Locality: " + locality.trim() + " | Parent Area: " + parentArea.trim() };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/* ============================================================
   W0B — LOCATION CONTEXT GENERATOR
============================================================ */
function buildLocationContextPrompt() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName("posts");
    const row     = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                         .map(function(h) { return String(h).trim(); });
    function getCol(name) {
      var idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
    }

    const articleType = getCol("Article Type");
    if (articleType !== "Geo Service Page" && articleType !== "Case Study") {
      return { success: false, message: "W0B only available for Geo Service Page and Case Study article types." };
    }

    const locality   = getCol("Locality");
    const parentArea = getCol("Parent Area");
    const stoneType  = getCol("Stone Type") || "stone";

    if (!locality)   return { success: false, message: "Enter Locality first and save to sheet." };
    if (!parentArea) return { success: false, message: "Enter Parent Area first and save to sheet." };

    var p = "";
    p += "ROLE: Expert stone floor restoration surveyor with deep knowledge of UK architectural layouts.\n\n";
    p += "TASK: Generate a practical property and structural context block for an SEO case study about " + stoneType + " floor cleaning and restoration in " + locality + ", " + parentArea + ".\n\n";
    p += "RESEARCH AND ANSWER the following:\n";
    p += "1. What are the predominant property types in " + locality + ", " + parentArea + "? (e.g. Victorian terraces, Edwardian semis, period conversions, modern rear extensions)\n";
    p += "2. What are the typical layouts and room uses where " + stoneType + " floors are found in these specific properties?\n";
    p += "3. What practical, real-world challenges do these specific property styles present for a " + stoneType + " floor? (e.g. heavy foot traffic from garden transitions, subfloor moisture in older builds, multi-use family spaces, high-traffic hallway and kitchen routes)\n";
    p += "4. What is the broader geographic area or postcode district that anchors " + locality + " as a location? (one phrase only — no administrative history)\n\n";
    p += "OUTPUT FORMAT:\n";
    p += "Write exactly two blocks of plain prose. No bullet points. No headings. No markdown.\n\n";
    p += "BLOCK 1 — PROPERTY CONTEXT (3-4 sentences):\n";
    p += "This block must read naturally when inserted into a professional floor restoration article, focusing on the architecture and floor usage.\n";
    p += "Do not mention Abbey Floor Care or any company name.\n";
    p += "Do not include pricing or service offers.\n";
    p += "Do not include historical facts about railway openings, bridge construction dates, administrative boundaries, or village origins.\n";
    p += "The paragraph must naturally blend the locality name, the housing stock (age and type), and where " + stoneType + " floors are found in these homes.\n";
    p += "The first word of Block 1 must be the locality name: " + locality + ".\n\n";
    p += "BLOCK 2 — MATERIAL AND PRACTICAL CHALLENGE NOTE (1-2 sentences):\n";
    p += "Write one or two sentences explaining how the specific lifestyle or structural nature of these " + locality + " homes affects the wear and condition of " + stoneType + " floors.\n";
    p += "Focus on the house and the stone — not local history. Examples: how extensions create high-traffic grit zones, how older subfloors handle moisture, how family use patterns affect surface wear.\n";
    p += "Do not mention Abbey Floor Care or any company name.\n";
    p += "Do not include pricing or service offers.\n\n";
    p += "OUTPUT CONSTRAINT:\n";
    p += "Output Block 1 followed by a single blank line followed by Block 2.\n";
    p += "No labels, no headings, no markdown. Plain prose only.\n";
    p += "The first word of your entire output must be the locality name: " + locality + ".";;

    return { success: true, prompt: p, locality: locality, parentArea: parentArea, stoneType: stoneType };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}


/* ============================================================
   MAIN SCHEMA GENERATOR
============================================================ */
function generateSchemaForActiveRow() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("posts");
    const row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: "ERROR: Select a data row first." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                         .getValues()[0]
                         .map(function(h) { return String(h).trim(); });

    // ── Core fields by column index ──
    const pageUrl     = String(sheet.getRange(row, 3).getValue()   || "").trim();
    const stoneType   = String(sheet.getRange(row, 7).getValue()   || "").trim();
    const articleType = String(sheet.getRange(row, 8).getValue()   || "").trim();
    const feedsHub    = String(sheet.getRange(row, 15).getValue()  || "").trim();
    const locality    = String(sheet.getRange(row, 115).getValue() || "").trim();
    const parentArea  = String(sheet.getRange(row, 116).getValue() || "").trim();
    const logoUrl     = String(sheet.getRange(row, 86).getValue()  || "").trim()
                        || "https://www.abbeyfloorcare.co.uk/wp-content/uploads/abbey-logo-1.png";

    // ── Fields by header name ──
    function colByHeader(name) {
      const idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
    }
    const headline          = colByHeader("New H1");
    const descriptionRaw    = colByHeader("New Meta Description");
    const description       = descriptionRaw.indexOf('Yoast Keyphrase:') > -1
      ? descriptionRaw.split('Yoast Keyphrase:')[0].trim()
      : descriptionRaw;
    const primaryEntity_    = colByHeader("Primary Entity");
    const supportingEntities = colByHeader("Supporting Entities Core");

    // ── Featured image ──
    let featuredImg = String(sheet.getRange(row, 85).getValue() || "").trim();
    if (!featuredImg) {
      const pageHtml  = String(sheet.getRange(row, 98).getValue() || "");
      const imgMatches = pageHtml.match(/<img[^>]+src=["']([^"']+)["'][^>]*>/gi) || [];
      const SKIP = /logo|icon|avatar|sprite|pixel|tracking/i;
      for (let m = 0; m < imgMatches.length; m++) {
        const srcMatch = imgMatches[m].match(/src=["']([^"']+)["']/i);
        if (srcMatch && srcMatch[1] && !SKIP.test(srcMatch[1])) {
          featuredImg = srcMatch[1];
          break;
        }
      }
    }

    // ── Validation ──
    if (!pageUrl)     return { success: false, message: "ERROR: URL column (col 3) empty for this row." };
    if (!articleType) return { success: false, message: "ERROR: Article Type (col 8) empty for this row." };
    if (!headline)    return { success: false, message: "ERROR: New H1 column empty — run W5 first." };
    if (!description) return { success: false, message: "ERROR: New Meta Description empty — run W5 first." };

    // ── Normalise URL ──
    const cleanUrl     = pageUrl.replace(/\/?$/, "/");
    const SERVICE_TYPES = ["Service Page", "Geo Service Page"];
    const isService    = SERVICE_TYPES.indexOf(articleType) > -1;

    // ── Read SAV sheet ──
    const savSheet = ss.getSheetByName("Schema Alignment Validation Sheet");
    let primaryType       = "Article";
    let primaryRel        = "";
    let secondaryRel      = "";
    let includeAreaServed = false;
    let fragmentId        = isService ? "#service" : "#article";

    if (savSheet) {
      const savData    = savSheet.getDataRange().getValues();
      const savHeaders = savData[0].map(function(h) { return String(h).trim(); });

      function savGet(rowArr, colName) {
        const idx = savHeaders.indexOf(colName);
        return idx > -1 ? String(rowArr[idx] || "").trim() : "";
      }

      for (let i = 1; i < savData.length; i++) {
        if (savGet(savData[i], "Article Type").toLowerCase() === articleType.toLowerCase()) {
          primaryType       = savGet(savData[i], "Allowed Primary Schema Type") || "Article";
          primaryRel        = savGet(savData[i], "Primary Relationship Property");
          secondaryRel      = savGet(savData[i], "Secondary Relationship Property");
          includeAreaServed = savGet(savData[i], "Must Include Area Served") === "Yes";
          // Read fragment ID from SAV sheet if present
          var savFragment   = savGet(savData[i], "Hub Spoke Fragment ID");
          if (savFragment)  fragmentId = savFragment.charAt(0) === "#" ? savFragment : "#" + savFragment;
          break;
        }
      }
    }

    const pageId = cleanUrl + fragmentId;
    const bcId   = cleanUrl + "#breadcrumb";

    // ── Constants ──
    const AUTHOR    = { "@type": "Person", "name": "David Allen" };
    const PUBLISHER = { "@id": "https://www.abbeyfloorcare.co.uk/#localbusiness" };
    const LB_ID     = "https://www.abbeyfloorcare.co.uk/#localbusiness";
    const today     = new Date().toISOString().split("T")[0];

    // ── Parse mentions from Supporting Entities Core ──
    // Returns array of {"@type":"Thing","name":"..."} objects
    function buildMentions(raw) {
      if (!raw) return [];
      return raw.split(",")
        .map(function(e) { return e.trim(); })
        .filter(function(e) { return e.length > 0; })
        .map(function(e) { return { "@type": "Thing", "name": e }; });
    }
    const mentionsArray = buildMentions(supportingEntities);

    // ── Build precise `about` from Primary Entity ──
    // Primary Entity format is "[Material] [Intent]" e.g. "Marble Stain Removal"
    // Use as-is for the about name — more precise than generic stone type
    function buildAbout(primaryEntityVal, stoneTypeVal) {
      var name = primaryEntityVal || (stoneTypeVal + " Floor Restoration");
      return { "@type": "Thing", "name": name };
    }

    // ── Dates ──
    function parseSheetDate(raw) {
      if (!raw || raw === "") return null;
      if (raw instanceof Date) return raw;
      var s = String(raw).trim();
      var dmyMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
      if (dmyMatch) {
        var day = parseInt(dmyMatch[1], 10);
        var mon = parseInt(dmyMatch[2], 10) - 1;
        var yr  = parseInt(dmyMatch[3], 10);
        if (yr < 100) yr += 2000;
        return new Date(yr, mon, day);
      }
      var d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    }
    let datePublished = today;
    let dateModified  = today;
    try {
      const postIdColIdx = headers.indexOf("Post ID");
      const postIdRaw    = postIdColIdx > -1 ? sheet.getRange(row, postIdColIdx + 1).getValue() : "";
      const postId       = String(postIdRaw || "").trim().replace(/\.0$/, "");
      let exportSheet    = ss.getSheetByName("site-export");
      if (!exportSheet) {
        const allSheets = ss.getSheets();
        for (let s = 0; s < allSheets.length; s++) {
          if (allSheets[s].getSheetId() === 228922537) { exportSheet = allSheets[s]; break; }
        }
      }
      if (exportSheet && postId) {
        const exportData = exportSheet.getDataRange().getValues();
        const expHeaders = exportData[0].map(function(h) { return String(h).trim(); });
        const expIdIdx   = expHeaders.indexOf("ID");
        const pubColIdx  = expHeaders.indexOf("Published Date");
        const modColIdx  = expHeaders.indexOf("Last Updated Date");
        if (expIdIdx > -1) {
          for (let e = 1; e < exportData.length; e++) {
            var rowId = String(exportData[e][expIdIdx]).trim().replace(/\.0$/, "");
            if (rowId === postId) {
              var pubDate = parseSheetDate(exportData[e][pubColIdx]);
              if (pubDate) datePublished = pubDate.toISOString().split("T")[0];
              // dateModified intentionally left as today's date (see const today above) —
              // W5C is only ever run after a genuine content change, so today is correct.
              break;
            }
          }
        }
      }
    } catch(dateErr) {
      console.error("W5C date lookup failed: " + dateErr.stack);
    }

    // ── BreadcrumbList ──
    const breadcrumb = w5cBuildBreadcrumbList(cleanUrl, bcId);

    // ── LocalBusiness entity ──
    // Now includes sameAs (canonical GBP Place ID URL) and hasOfferCatalog
    const localBusiness = {
      "@type":  "LocalBusiness",
      "@id":    LB_ID,
      "name":   "Abbey Floor Care",
      "url":    "https://www.abbeyfloorcare.co.uk",
      "logo":   logoUrl,
      "sameAs": "https://www.google.com/maps/place/?q=place_id:ChIJPQ6EKRMCekgRkXMuCPQH8U0",
      "hasOfferCatalog": {
        "@type": "OfferCatalog",
        "name":  "Stone Floor Restoration Services",
        "itemListElement": [
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Marble Floor Cleaning" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Marble Floor Polishing" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Marble Restoration" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Natural Stone Floor Restoration" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Limestone Floor Cleaning" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Travertine Cleaning" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Slate Floor Cleaning" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Stain Removal" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Grout Sealing" } },
          { "@type": "Offer", "itemOffered": { "@type": "Service", "name": "Tile And Grout Cleaning" } }
        ]
      }
    };

    // ── Primary entity ──
    let primaryEntityObj = {};

    if (primaryType === "Service") {
      // Build service name from URL slug
      const slugParts  = cleanUrl.replace(/\/$/, "").split("/").filter(Boolean);
      const lastSlug   = slugParts[slugParts.length - 1] || "";
      const serviceName = lastSlug
        .replace(/-(?:expert|solutions?|services?|near|me|specialist|company|professional|local|near-me|experts?|providers?|trusted|certified|recommended)(?=-|$)/gi, "")
        .replace(/-+$/, "")
        .replace(/-/g, " ")
        .replace(/\b\w/g, function(c) { return c.toUpperCase(); })
        .replace(/\bAnd\b/g, "and")
        .replace(/\bIn\b/g, "in")
        .replace(/\bOf\b/g, "of")
        .replace(/\bFor\b/g, "for")
        .trim()
        || headline;

      primaryEntityObj = {
        "@type":            "Service",
        "@id":              pageId,
        "name":             serviceName,
        "description":      description,
        "provider":         PUBLISHER,
        "offers": {
          "@type":       "Offer",
          "availability": "https://schema.org/InStock",
          "areaServed":  "United Kingdom",
          "seller":      { "@id": LB_ID }
        },
        "mainEntityOfPage": { "@type": "WebPage", "@id": cleanUrl }
      };

      if (featuredImg)      primaryEntityObj["image"]       = featuredImg;
      if (includeAreaServed) {
        var areaName = (locality && parentArea) ? locality + ", " + parentArea
                     : parentArea || locality || "United Kingdom";
        primaryEntityObj["areaServed"] = areaName;
      }

      // about — precise primary entity name
      primaryEntityObj["about"] = buildAbout(primaryEntity_, stoneType);

      // mentions — from Supporting Entities Core
      if (mentionsArray.length > 0) {
        primaryEntityObj["mentions"] = mentionsArray;
      }

      // isRelatedTo — related hub content
      const related = w5cGetRelatedPages(stoneType, row, 5);
      if (related.length > 0) {
        primaryEntityObj["isRelatedTo"] = related.map(function(p) {
          return { "@id": p.url.replace(/\/?$/, "/") + p.fragment };
        });
      }

    } else {
      // Article-based pages
      primaryEntityObj = {
        "@type":            "Article",
        "@id":              pageId,
        "headline":         headline,
        "description":      description,
        "image":            featuredImg,
        "author":           AUTHOR,
        "publisher":        PUBLISHER,
        "mainEntityOfPage": { "@type": "WebPage", "@id": cleanUrl },
        "datePublished":    datePublished,
        "dateModified":     dateModified
      };

      // mentions — from Supporting Entities Core (all Article-based pages)
      if (mentionsArray.length > 0) {
        primaryEntityObj["mentions"] = mentionsArray;
      }

      if (primaryRel === "hasPart") {
        // Hub Page — editorial spokes only
        const HUB_EXCLUDE = ["Service Page", "Geo Service Page"];
        const spokes = w5cGetRelatedPages(stoneType, row, 5, HUB_EXCLUDE);
        if (spokes.length > 0) {
          primaryEntityObj["hasPart"] = spokes.map(function(p) {
            const spokeUrl  = p.url.replace(/\/?$/, "/");
            const spokeId   = spokeUrl + p.fragment;
            const spokeType = p.fragment === "#service" ? "WebPage" : "Article";
            return {
              "@type": spokeType,
              "@id":   spokeId,
              "name":  w5cCleanTitle(p.title)
            };
          });
        }
        // about on hub — uses primary entity
        primaryEntityObj["about"] = buildAbout(primaryEntity_, stoneType);

      } else if (primaryRel === "isPartOf") {
        // Spoke Page — declare parent hub
        if (feedsHub) {
          const hubId = feedsHub.replace(/\/?$/, "/") + "#article";
          primaryEntityObj["isPartOf"] = { "@type": "Article", "@id": hubId };
        }
        // about — precise primary entity name (replaces generic stone type)
        primaryEntityObj["about"] = buildAbout(primaryEntity_, stoneType);

      } else if (primaryRel === "isRelatedTo") {
        // Service-type article fallback
        const related = w5cGetRelatedPages(stoneType, row, 5);
        if (related.length > 0) {
          primaryEntityObj["isRelatedTo"] = related.map(function(p) {
            return { "@id": p.url.replace(/\/?$/, "/") + p.fragment };
          });
        }
        primaryEntityObj["about"] = buildAbout(primaryEntity_, stoneType);
      }
    }

    // ── VideoObject nodes — parse from New HTML ──
    const videoNodes = [];
    try {
      const pageHtml = String(sheet.getRange(row, 98).getValue() || "");
      const iframeRe = /<iframe[^>]+src=["']([^"']*(?:youtube\.com\/embed|youtu\.be)\/([a-zA-Z0-9_-]{11})[^"']*)["'][^>]*(?:title=["']([^"']*)["'])?[^>]*>/gi;
      let iframeMatch;
      while ((iframeMatch = iframeRe.exec(pageHtml)) !== null) {
        const videoId    = iframeMatch[2];
        const videoTitle = iframeMatch[3] || headline;
        if (!videoId) continue;

        // Get upload date from YouTube Upload Date column, fall back to dateModified
        const ytDateRaw  = colByHeader("YouTube Upload Date");
        const parsedYtDate = ytDateRaw ? parseSheetDate(ytDateRaw) : null;
        const uploadDate = parsedYtDate
          ? parsedYtDate.toISOString().split("T")[0]
          : dateModified;

        videoNodes.push({
          "@type":        "VideoObject",
          "@id":          cleanUrl + "#video-" + videoId,
          "name":         videoTitle,
          "description":  description,
          "thumbnailUrl": "https://img.youtube.com/vi/" + videoId + "/maxresdefault.jpg",
          "uploadDate":   uploadDate,
          "contentUrl":   "https://www.youtube.com/watch?v=" + videoId,
          "embedUrl":     "https://www.youtube.com/embed/" + videoId,
          "publisher":    PUBLISHER
        });
      }
    } catch(vidErr) {
      console.error("W5C VideoObject parse failed: " + vidErr.toString());
    }

    // ── Assemble graph ──
    const graphNodes = [ primaryEntityObj, breadcrumb, localBusiness ].concat(videoNodes);
    const graph = {
      "@context": "https://schema.org",
      "@graph":   graphNodes
    };

    // ── Minify and wrap ──
    const minified = JSON.stringify(graph);
    const wrapped  = '<script type="application/ld+json">' + minified + "<\/script>";

    // ── Push to col 92 ──
    const cell = sheet.getRange(row, 92);
    try {
      cell.setPlainTextValue(wrapped);
    } catch(e) {
      cell.setNumberFormat("@");
      cell.setValue(wrapped);
    }

    logPipelineResume("W5C — Schema Built", "");
    return {
      success: true,
      message: "Schema built and pushed to row " + row + ".\n" +
               "Article Type: " + articleType + "\n" +
               "Primary @type: " + primaryType + "\n" +
               "Relationship: " + (primaryRel || "none") + "\n" +
               "mentions: " + mentionsArray.length + " entities\n" +
               "Run W5B audit to verify.",
      preview: wrapped
    };

  } catch(e) {
    return { success: false, message: "BUILD ERROR: " + e.toString() };
  }
}

/* ============================================================
   BREADCRUMBLIST BUILDER
============================================================ */
function w5cBuildBreadcrumbList(cleanUrl, bcId) {
  const SITE = "https://www.abbeyfloorcare.co.uk";
  const path = cleanUrl.replace(SITE, "").replace(/^\/|\/$/g, "");
  const segments = path ? path.split("/").filter(function(s) { return s.length > 0; }) : [];

  const items = [ { "@type": "ListItem", "position": 1, "name": "Home", "item": SITE + "/" } ];
  let accumulated = SITE;

  const SEGMENT_LABELS = {
    "home-garden":   "Home & Garden",
    "tile-cleaning": "Tile Cleaning",
    "tile-care":     "Tile Care",
    "marble-care":   "Marble Care",
    "stone-care":    "Stone Care"
  };

  segments.forEach(function(seg, i) {
    accumulated += "/" + seg;
    const label = SEGMENT_LABELS[seg] ||
                  seg.replace(/-/g, " ")
                     .replace(/\b\w/g, function(c) { return c.toUpperCase(); });
    items.push({
      "@type":    "ListItem",
      "position": i + 2,
      "name":     label,
      "item":     accumulated + "/"
    });
  });

  return {
    "@type":           "BreadcrumbList",
    "@id":             bcId,
    "itemListElement": items
  };
}

/* ============================================================
   TITLE CLEANER — used for hasPart spoke names
============================================================ */
function w5cCleanTitle(t) {
  return t
    .replace(/[\u{1F000}-\u{1FFFF}]/gu, "")
    .replace(/[\u{2600}-\u{27BF}]/gu, "")
    .replace(/[\u{1F300}-\u{1F9FF}]/gu, "")
    .replace(/\u2728|\u2605|\u2764/gu, "")
    .replace(/\s*[\(\[]\d{4}[^\)\]]*[\)\]]/g, "")
    .replace(/\s*\(\d{4}\)\s*/g, " ")
    .replace(/\s*[-–|]\s*(Updated|New|Latest|Current|\d{4}).*$/i, "")
    .replace(/\s+\d{4}\s*$/, "")
    .replace(/\s+0[\d\s]{4,15}$/, "")
    .replace(/\s+[Bb]y\s+Abbey\s+Floor\s+Care\s*$/i, "")
    .replace(/\s+/g, " ")
    .trim();
}

/* ============================================================
   RELATED PAGES FETCHER
============================================================ */
function w5cGetRelatedPages(stoneType, currentRow, limit, excludeTypes) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("posts");
  const data  = sheet.getDataRange().getValues();
  const SERVICE_TYPES = ["Service Page", "Geo Service Page"];
  const excluded = Array.isArray(excludeTypes) ? excludeTypes : [];
  const results  = [];

  for (let i = 1; i < data.length; i++) {
    if (i + 1 === currentRow) continue;
    if (String(data[i][6] || "").trim() !== stoneType) continue;
    const url         = String(data[i][2]  || "").trim();
    const title       = String(data[i][79] || "").trim();
    const articleType = String(data[i][7]  || "").trim();
    if (!url || !title) continue;
    if (excluded.indexOf(articleType) > -1) continue;
    const fragment = SERVICE_TYPES.indexOf(articleType) > -1 ? "#service" : "#article";
    results.push({ url: url, title: title, articleType: articleType, fragment: fragment });
    if (results.length >= limit) break;
  }
  return results;
}

/* ============================================================
   W4B METADATA FETCHER
============================================================ */
function getW4BMetadata() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName("posts");
    const row     = sheet.getActiveRange().getRow();
    if (row < 2) return { stoneType: "", articleType: "", primaryTerm: "" };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                         .map(function(h) { return String(h).trim(); });
    function colVal(name) {
      var idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || "").trim() : "";
    }
    return {
      stoneType:   colVal("Stone Type"),
      articleType: colVal("Article Type"),
      primaryTerm: colVal("Primary Search Term") || colVal("Primary Entity") || ""
    };
  } catch(e) {
    return { stoneType: "", articleType: "", primaryTerm: "" };
  }
}