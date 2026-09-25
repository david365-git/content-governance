/* ============================================================
   ce_Stage2C_ImageSuggestion.gs
   STAGE 2C — IMAGE SUGGESTION
   Reads Article Type and the W1.5C Enriched Plan (column CW / 101).
   Parses sections, identifies candidates needing a visual,
   and builds the LLM prompt requesting JSON suggestions.
   Runs after W2B, before W3.
============================================================ */
/* ============================================================
   STAGE 2C — HTML SOURCE RESOLVER
   Maps a friendly source key to its posts sheet column number.
   Used by W2C so the user can choose which HTML version to
   draw image suggestions from.
============================================================ */
function getW2CSourceColumn(sourceKey) {
  var SOURCES = {
    'new_html':        { col: 98,  label: 'New HTML (col CT)' },
    'humanised_html':  { col: 104, label: 'Humanised HTML (col CZ)' },
    'w2b_raw_html':    { col: 151, label: 'W2B Raw HTML (col EU)' },
    'w8b_html':        { col: 165, label: 'W8B_HTML (col FI)' },
    'w8c_html':        { col: 166, label: 'W8C_HTML (col FJ)' },
    'html_semtc_adj':  { col: 176, label: 'New HTML With Semantic Adjustment (col FT)' }
  }
  var key = String(sourceKey || 'w2b_raw_html').trim();
  return SOURCES[key] || SOURCES['w2b_raw_html'];
}


function buildImageSuggestionPrompt(sourceKey) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const row   = sheet.getActiveRange().getRow();
  const d     = getActiveRowDataMap();

  const articleType = String(d["Article Type"] || "General").trim();

  const source = getW2CSourceColumn(sourceKey);
  const html = String(sheet.getRange(row, source.col).getValue() || "").trim();
  if (!html) {
    return { success: false, message: "Column " + source.label + " is empty for this row — save that stage's output first." };
  }

  // Extract H2 and H3 headings with their immediately following prose
  var h2Blocks = [];
  var currentH2 = null;
  var h2Count = 0;
  var h3CountInSection = 0;

  // Split HTML into heading blocks
  var headingRegex = /<(h2|h3)([^>]*)>([\s\S]*?)<\/(h2|h3)>([\s\S]*?)(?=<h2|<h3|<\/section|<footer|$)/gi;
  var match;

  while ((match = headingRegex.exec(html)) !== null) {
    var level        = match[1].toLowerCase();
    var headingRaw   = match[3];
    var followingRaw = match[5];

    var headingText = headingRaw.replace(/<[^>]+>/g, '').trim();
    if (!headingText) continue;

    // Extract up to 2 <p> tag contents following the heading
    var pMatches = [];
    var pRegex   = /<p[^>]*>([\s\S]*?)<\/p>/gi;
    var pMatch;
    var pCount   = 0;
    while ((pMatch = pRegex.exec(followingRaw)) !== null && pCount < 2) {
      var pText = pMatch[1].replace(/<[^>]+>/g, '').trim();
      if (pText) {
        pMatches.push(pText);
        pCount++;
      }
    }

    if (level === 'h2') {
      currentH2 = headingText;
      h2Count++;
      h3CountInSection = 0;
      h2Blocks.push({
        level:   'h2',
        heading: headingText,
        prose:   pMatches.join(' '),
        h2Index: h2Count
      });
    } else if (level === 'h3' && currentH2) {
      h3CountInSection++;
      h2Blocks.push({
        level:    'h3',
        heading:  headingText,
        prose:    pMatches.join(' '),
        parentH2: currentH2,
        h2Index:  h2Count
      });
    }
  }

  if (h2Blocks.length === 0) {
    return { success: false, message: "No H2 or H3 headings found in " + source.label + " content." };
  }

  // Build candidate block for the prompt
  var candidateBlock = "";
  var h2Counter  = 0;
  var h3Counters = {};

  h2Blocks.forEach(function(b) {
    if (b.level === 'h2') {
      h2Counter++;
      h3Counters[h2Counter] = 0;
      candidateBlock += "H2 " + h2Counter + ": " + b.heading + "\n";
      if (b.prose) {
        candidateBlock += "Opening prose: " + b.prose.substring(0, 300) + (b.prose.length > 300 ? "..." : "") + "\n";
      }
      candidateBlock += "\n";
    } else if (b.level === 'h3') {
      h3Counters[b.h2Index] = (h3Counters[b.h2Index] || 0) + 1;
      candidateBlock += "H3 " + b.h2Index + "." + h3Counters[b.h2Index] + ": " + b.heading + "\n";
      if (b.prose) {
        candidateBlock += "Opening prose: " + b.prose.substring(0, 300) + (b.prose.length > 300 ? "..." : "") + "\n";
      }
      candidateBlock += "\n";
    }
  });

var typeInstruction = "ARTICLE TYPE: " + articleType + "\n" +
    "For every suggestion, set \"type\": \"ai-photo\" and provide an \"ai_prompt\" field — a detailed photorealistic image description that serves two purposes simultaneously: (1) as a prompt for AI image generation if needed, and (2) as a visual search brief to find a matching real photograph from an existing library of thousands of job site images. All images must be landscape orientation, 16:9 aspect ratio. Describe the scene as if briefing a photographer: the specific floor condition visible, the viewing angle and composition, the lighting quality, the dominant colours and tones of the stone and grout, what is happening to the surface at this exact stage, and any specific visual details that make this image recognisable and distinct from a generic floor photo. Write as a single detailed paragraph. Do NOT describe abstract diagrams, charts, or illustrations.\n" +
    "\nALSO provide an ALTERNATIVE non-photographic media option for every suggestion, so the photorealistic option above and this alternative can be compared before choosing. Set \"alt_media_type\" to ONE of the following five categories, chosen strictly by what this specific heading's content is doing:\n" +
    "1. BEFORE/AFTER COMPARISON — only when the content describes a transformation or dramatic outcome contrast. MUST use real photography, never AI-generated — do not suggest this type unless real project photos would plausibly exist for this content.\n" +
    "2. CALLOUT / WARNING BOX — when the content states an explicit prohibition, safety rule, or \"never do X\" instruction.\n" +
    "3. DIAGRAM / CROSS-SECTION — when explaining a hidden mechanism, process, or how something behaves beneath the visible surface (e.g. how heat affects sealer, how moisture moves through stone).\n" +
    "4. COMPARISON TABLE — when comparing two or more named options, products, or approaches side by side.\n" +
    "5. ICON-BASED STEP SUMMARY — when the content describes a sequential multi-step process that could be condensed into a visual checklist.\n" +
    "HARD RULE: \"alt_media_brief\" must NEVER describe a photorealistic image — it must always be clearly illustrated/diagrammatic/iconographic, EXCEPT for category 1 (Before/After), where it must instead be a one-sentence note describing what real project photo would be needed (do not generate an AI prompt for that case). For categories 2-5, \"alt_media_brief\" should be a ready-to-use AI generation prompt for that non-photorealistic visual, landscape orientation, specific to this heading's actual content — not generic.\n";

  var prompt = "STAGE 2C — IMAGE SUGGESTION\n" +
    "ROLE: Senior UK SEO & Stone Restoration Content Strategist with visual content expertise.\n" +
    "TASK: Review the H2 and H3 headings and opening prose below from the published article HTML. For each heading that would genuinely benefit from a supporting image placed immediately after it, output a suggestion. Not every heading requires an image — only suggest where a visual adds real diagnostic, illustrative, or engagement value. For H3 headings, suggest images wherever genuinely useful — the suggestions will be reviewed and only the best selected.\n\n" +
    typeInstruction + "\n" +
    "------------------------------------------------------------\n" +
    "H2 AND H3 HEADINGS AND OPENING PROSE FROM ARTICLE:\n" +
    "------------------------------------------------------------\n" +
    candidateBlock +
    "------------------------------------------------------------\n" +
    "OUTPUT RULES:\n" +
    "- Return ONLY a JSON array. No preamble, no markdown fences, no explanation.\n" +
    "- One object per suggested image.\n" +
    "- Maximum " + h2Blocks.length + " suggestions total.\n" +
    "- Suggest images for all H3 headings where genuinely useful — the suggestions will be reviewed and only the best selected.\n" +
    "- \"h2_text\" must be the exact H2 or H3 heading text from the list above — copied character for character.\n" +
    "- \"heading_level\" must be either \"h2\" or \"h3\".\n" +
    "- \"position\" must be \"after_h2\" for H2 images and \"after_h3\" for H3 images.\n" +
    "- \"alt\" must be 5-15 words, descriptive, no keyword stuffing.\n" +
    "- \"caption\" must follow the diagnostic caption pattern — connect the image to the reader's situation, not just describe it (e.g. \"If your floor looks like this...\").\n" +
    "- \"slug\" must be lowercase, hyphenated, no file extension.\n" +
    "- \"ai_prompt\" must be comprehensive enough to serve two purposes simultaneously:\n" +
    "  PURPOSE 1 — AI IMAGE GENERATION: detailed enough that an AI image generator can produce a useful photorealistic image from it alone.\n" +
    "  PURPOSE 2 — LIBRARY SEARCH AID: descriptive enough that a person scanning thousands of existing photos can use it as a mental image to recognise a match. Include: the specific floor condition or stage being shown, the viewing angle and composition (wide shot, close-up, top-down), the dominant colours and tones, what distinguishes this image from a generic floor photo, and any specific visual details that make this image recognisable. Write as a single detailed paragraph.\n\n" +
    "REQUIRED JSON SHAPE (one example object):\n" +
    "[\n" +
    "  {\n" +
    "    \"h2_text\": \"exact H2 or H3 heading text here\",\n" +
    "    \"heading_level\": \"h2\" or \"h3\",\n" +
    "    \"slug\": \"example-descriptive-slug\",\n" +
    "    \"alt\": \"descriptive alt text here\",\n" +
    "    \"caption\": \"diagnostic caption here\",\n" +
    "    \"type\": \"ai-photo\",\n" +
    "    \"ai_prompt\": \"detailed photorealistic image description here\",\n" +
    "    \"alt_media_type\": \"one of: Before/After Comparison | Callout/Warning Box | Diagram/Cross-Section | Comparison Table | Icon-Based Step Summary\",\n" +
    "    \"alt_media_brief\": \"AI prompt for the non-photorealistic alternative, or a note on the real photo needed if alt_media_type is Before/After\",\n" +
    "    \"position\": \"after_h2\" or \"after_h3\"\n" +
    "  }\n" +
    "]\n\n" +
    "--- GENERATE NOW ---";

  return {
    success: true,
    prompt:        prompt,
    candidateCount: h2Blocks.length,
    articleType:   articleType
  };
}

/* ============================================================
   STAGE 2C — SAVE SUGGESTIONS
   Validates pasted JSON and writes it to posts column FK (167).
============================================================ */
function saveW2CSuggestions(raw) {
  if (!raw || !raw.trim()) {
    return { success: false, message: "No JSON provided." };
  }

  // Strip markdown code fences if present
  var cleaned = raw.trim()
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/, '')
    .trim();

  // Validate JSON
  try {
    JSON.parse(cleaned);
  } catch (e) {
    return { success: false, message: "Invalid JSON — " + e.toString() };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();
  posts.getRange(row, 167).setValue(cleaned); // FK — Image Suggestions (JSON)

  return { success: true, message: "Saved to column FK (167)." };
}

/* ============================================================
   STAGE 2C — SAVE SUGGESTIONS TO IMAGE LIBRARY SHEET
   Reads JSON from column FK (167) for the active row.
   Writes one row per suggestion to the Image Library sheet.
   Replaces any existing rows for this Post ID on re-run.
============================================================ */
function saveW2CSuggestionsToImageLibrary() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();

  // Get Post ID
  var headers = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0]
                     .map(function(h) { return String(h).trim(); });
  var postIdIdx = headers.indexOf('Post ID');
  if (postIdIdx === -1) return { success: false, message: "'Post ID' column not found in posts sheet." };
  var postId = String(posts.getRange(row, postIdIdx + 1).getValue() || "").trim();
  if (!postId) return { success: false, message: "No Post ID found for this row." };

  // Get JSON from column FK (167)
  var json = String(posts.getRange(row, 167).getValue() || "").trim();
  if (!json) return { success: false, message: "Column FK (Image Suggestions JSON) is empty — run W2C Generate and Save first." };

  var suggestions;
  try {
    suggestions = JSON.parse(json);
  } catch (e) {
    return { success: false, message: "Invalid JSON in column FK — " + e.toString() };
  }

  if (!Array.isArray(suggestions) || suggestions.length === 0) {
    return { success: false, message: "No suggestions found in column FK JSON." };
  }

  // Get Image Library sheet
  var libSheet = ss.getSheetByName('Image Library');
  if (!libSheet) return { success: false, message: "'Image Library' sheet not found — create it first." };

  // Remove existing rows for this Post ID (now 8 columns)
  var lastRow = libSheet.getLastRow();
  if (lastRow > 1) {
    var existing = libSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var r = existing.length - 1; r >= 0; r--) {
      if (String(existing[r][0]).trim() === postId) {
        libSheet.deleteRow(r + 2);
      }
    }
  }

  // Write new rows — prepend Post ID to slug, include H2 Heading
  var newRows = suggestions.map(function(s) {
    var slug = s.slug ? postId + '-' + s.slug : postId;
    var h2   = s.h2_text || '';
    return [
      postId,
      (s.heading_level === 'h3' ? 'H3- ' : 'H2- ') + h2,  // Heading with level prefix
      slug,         // Slug with Post ID prefix
      s.alt        || '',
      s.caption    || '',
      s.ai_prompt  || '',
      '',           // Image URL — to be filled manually
      'Pending'     // Status
    ];
  });

  if (newRows.length > 0) {
    libSheet.getRange(libSheet.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
  }

  return { success: true, message: newRows.length + " suggestion(s) saved to Image Library for Post ID " + postId + "." };
}

/* ============================================================
   STAGE 2C — INJECT REAL IMAGES FROM IMAGE LIBRARY INTO FL HTML
   Reads Image Library rows where Status = Ready for this Post ID.
   Replaces each matching IMAGE SUGGESTION comment block in
   column FL (168) with a real <figure> tag.
   Updates Status to Injected in the Image Library sheet.
============================================================ */
function injectImagesFromLibraryToFL() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();

  // Get Post ID
  var headers = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0]
                     .map(function(h) { return String(h).trim(); });
  var postIdIdx = headers.indexOf('Post ID');
  if (postIdIdx === -1) return { success: false, message: "'Post ID' column not found in posts sheet." };
  var postId = String(posts.getRange(row, postIdIdx + 1).getValue() || "").trim();
  if (!postId) return { success: false, message: "No Post ID found for this row." };

  // Get FL HTML from column 168
  var html = String(posts.getRange(row, 168).getValue() || "").trim();
  if (!html) return { success: false, message: "Column FL is empty — run Reconcile first." };

  // Get Image Library sheet
  var libSheet = ss.getSheetByName('Image Library');
  if (!libSheet) return { success: false, message: "'Image Library' sheet not found." };

  var lastRow = libSheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "Image Library sheet is empty." };

  var libData = libSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // Find Ready rows for this Post ID
  var readyRows = [];
  libData.forEach(function(r, i) {
    if (String(r[0]).trim() === postId && String(r[9]).trim().toLowerCase() === 'ready') {
      readyRows.push({
        sheetRow:   i + 2,
        slug:       String(r[2]).trim(),   // col C
        alt:        String(r[3]).trim(),   // col D
        caption:    String(r[4]).trim(),   // col E
        imageUrl:   String(r[7]).trim(),   // col H
        figureHtml: String(r[11]).trim()   // col L — Figure HTML from ChatGPT
      });
    }
  });

  if (readyRows.length === 0) {
    return { success: false, message: "No rows with Status 'Ready' found in Image Library for Post ID " + postId + ". Add image URLs and set Status to Ready first." };
  }

  var injected = [];
  var notFound = [];

  readyRows.forEach(function(r) {
    if (!r.imageUrl) {
      notFound.push(r.slug + " (no URL)");
      return;
    }

    // Use Figure HTML from column L if available, otherwise build basic figure tag
    var figureTag;
    if (r.figureHtml) {
      figureTag = '\n' + r.figureHtml + '\n';
    } else {
      figureTag =
        '\n<figure class="wp-caption aligncenter" style="width: 700px">' +
        '<img src="' + r.imageUrl + '" alt="' + r.alt + '" data-w4b-skip="true" />' +
        '<figcaption class="wp-caption-text">' + r.caption + '</figcaption>' +
        '</figure>\n';
    }

    // Find and replace the matching comment block by slug
    // Strip Post ID prefix from slug for matching against comment blocks
    var matchSlug = r.slug.replace(/^\d+-/, '');
    var escapedSlug = matchSlug.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var commentRegex = new RegExp(
      '<!--\\s*IMAGE SUGGESTION[\\s\\S]*?slug:\\s*[^\\n]*' + escapedSlug + '[\\s\\S]*?-->',
      'i'
    );

    if (!commentRegex.test(html)) {
      notFound.push(r.slug + " (comment block not found in FL HTML)");
      return;
    }

    html = html.replace(commentRegex, figureTag);
    injected.push(r.slug);

    // Update Status to Injected in Image Library
    libSheet.getRange(r.sheetRow, 10).setValue('Injected');
  });

  // Save updated HTML back to column FL (168)
  if (injected.length > 0) {
    posts.getRange(row, 168).setValue(html);
  }

  var message = injected.length + " image(s) injected into FL HTML.";
  if (notFound.length > 0) {
    message += " Not found: " + notFound.join(", ") + ".";
  }

  return { success: true, message: message, injected: injected, notFound: notFound };
}

/* ============================================================
   STAGE 2C — BUILD FIGURE HTML PROMPT
   Reads Image Library rows for the active Post ID where
   Image URL (col H) is populated. Builds a ChatGPT prompt
   asking for complete figure HTML with natural varied captions.
   Rows with Caption Link URL (col K) get linked captions.
   Rows without get plain captions.
============================================================ */
function buildFigureHtmlPrompt() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();

  // Get Post ID
  var headers   = posts.getRange(1, 1, 1, posts.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });
  var postIdIdx = headers.indexOf('Post ID');
  if (postIdIdx === -1) return { success: false, message: "'Post ID' column not found." };
  var postId = String(posts.getRange(row, postIdIdx + 1).getValue() || "").trim();
  if (!postId) return { success: false, message: "No Post ID found for this row." };

  // Get Image Library sheet
  var libSheet = ss.getSheetByName('Image Library');
  if (!libSheet) return { success: false, message: "'Image Library' sheet not found." };

  var lastRow = libSheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "Image Library sheet is empty." };

  // Read all 12 columns (A-L)
  var libData = libSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // Filter rows for this Post ID with an Image URL
  var imageRows = libData.filter(function(r) {
    return String(r[0]).trim() === postId &&
           String(r[7]).trim() !== ''; // col H — Image URL
  });

  if (imageRows.length === 0) {
    // Diagnostic — show what we actually found
    var debugInfo = "Post ID searched: '" + postId + "'. ";
    debugInfo += "Total rows in library: " + libData.length + ". ";
    if (libData.length > 0) {
      debugInfo += "First row Post ID: '" + String(libData[0][0]).trim() + "'. ";
      debugInfo += "First row Image URL: '" + String(libData[0][7]).trim() + "'.";
    }
    return { success: false, message: "No rows with Image URL found. " + debugInfo };
  }

  // Build image context block for the prompt
  var imageBlock = "";
  var linkUrls   = [];

  imageRows.forEach(function(r, i) {
    var slug       = String(r[2]).trim();  // col C
    var h2         = String(r[1]).trim();  // col B
    var alt        = String(r[3]).trim();  // col D
    var caption    = String(r[4]).trim();  // col E
    var aiPrompt   = String(r[5]).trim();  // col F
    var imageUrl   = String(r[7]).trim();  // col H
    var captionLink = String(r[10]).trim(); // col K — Caption Link URL

    if (captionLink && linkUrls.indexOf(captionLink) === -1) {
      linkUrls.push(captionLink);
    }

    imageBlock += "IMAGE " + (i + 1) + ":\n";
    imageBlock += "  slug: " + slug + "\n";
    imageBlock += "  h2_heading: " + h2 + "\n";
    imageBlock += "  image_url: " + imageUrl + "\n";
    imageBlock += "  alt: " + alt + "\n";
    imageBlock += "  suggested_caption: " + caption + "\n";
    imageBlock += "  ai_prompt_description: " + aiPrompt.substring(0, 200) + "...\n";
    imageBlock += "  caption_link_url: " + (captionLink || "NONE") + "\n";
    imageBlock += "\n";
  });

  // Build link variation instruction if multiple images share a URL
  var linkVariationBlock = "";
  if (linkUrls.length > 0) {
    linkVariationBlock =
      "------------------------------------------------------------\n" +
      "ANCHOR TEXT VARIATION RULES (CRITICAL):\n" +
      "------------------------------------------------------------\n" +
      "Some images link to the same URL. Follow these rules strictly:\n\n" +
      "1. FIRST LINK to any URL must use the most keyword-rich, descriptive anchor text — " +
      "include the material type, the intervention and the location if available from the slug or H2 heading.\n\n" +
      "2. SUBSEQUENT LINKS to the same URL must vary the anchor text naturally — " +
      "rephrase using different words that describe the same destination page. " +
      "Draw from the H2 heading, the slug, and the image description to find natural variation.\n\n" +
      "3. NEVER repeat the exact same anchor text twice for the same URL.\n\n" +
      "4. Anchor text must read naturally within the caption sentence — " +
      "it must not look like a keyword string or an SEO tag.\n\n" +
      "5. The caption sentence must make sense without the link — " +
      "the link wraps a natural phrase, it does not replace the caption.\n\n" +
      "URLS THAT APPEAR MORE THAN ONCE IN THIS SET:\n" +
      linkUrls.join("\n") + "\n\n";
  }

  var prompt =
    "FIGURE HTML GENERATION\n" +
    "ROLE: Senior UK SEO Content Specialist for a natural stone floor restoration website.\n" +
    "TASK: For each image below, write a complete HTML figure block.\n\n" +
    "SITE: Abbey Floor Care (abbeyfloorcare.co.uk) — UK natural stone and tile floor restoration.\n\n" +
    "------------------------------------------------------------\n" +
    "FIGURE HTML FORMAT (use exactly this structure for every image):\n" +
    "------------------------------------------------------------\n" +
    "<figure class='wp-caption aligncenter' style='width: 700px'>\n" +
    "<img src='[image_url]' alt='[alt text]' data-w4b-skip='true' />\n" +
    "<figcaption class='wp-caption-text'>[caption text]</figcaption>\n" +
    "</figure>\n\n" +
    "RULES FOR EVERY FIGURE:\n" +
    "- The data-w4b-skip=\"true\" attribute MUST be present on every img tag — do not omit it.\n" +
    "- Alt text must be descriptive, specific, 5-15 words, no keyword stuffing.\n" +
    "- You may improve the suggested alt text if it better describes the image based on the ai_prompt_description.\n" +
    "- Caption must be plain English, diagnostic in tone — connect the image to the reader's situation.\n" +
    "- Maximum 20 words for caption text excluding any link anchor text.\n" +
    "- Where caption_link_url is provided, wrap a natural phrase within the caption in an anchor tag:\n" +
    "  <a href=\"[caption_link_url]\">natural anchor text phrase</a>\n" +
    "- Where caption_link_url is NONE, write a plain caption with no link.\n\n" +
    linkVariationBlock +
    "------------------------------------------------------------\n" +
    "OUTPUT FORMAT:\n" +
    "------------------------------------------------------------\n" +
    "Return ONLY a JSON array. No preamble, no markdown fences, no explanation.\n" +
    "One object per image, in the same order as the images below.\n\n" +
    "[\n" +
    "  {\n" +
    "    \"slug\": \"exact slug from image list below\",\n" +
    "    \"alt\": \"final alt text used in the img tag\",\n" +
    "    \"caption_html\": \"final caption HTML including any anchor tag\",\n" +
    "    \"figure_html\": \"complete figure HTML block as one line\"\n" +
    "  }\n" +
    "]\n\n" +
    "------------------------------------------------------------\n" +
    "IMAGES:\n" +
    "------------------------------------------------------------\n" +
    imageBlock +
    "--- GENERATE NOW ---";

  return {
    success:    true,
    prompt:     prompt,
    imageCount: imageRows.length
  };
}



/* ============================================================
   STAGE 2C — SAVE FIGURE HTML FROM CHATGPT TO IMAGE LIBRARY
   Parses JSON response, matches on slug, writes figure_html
   to column L (12) in the Image Library sheet.
============================================================ */
function saveFigureHtmlToImageLibrary(raw) {
  if (!raw || !raw.trim()) {
    return { success: false, message: "No output provided." };
  }

  var cleaned = raw.trim()
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/, '')
    .trim();

  // Attempt to fix unescaped quotes inside JSON string values
  // by replacing smart quotes and normalising line breaks
  cleaned = cleaned
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/\r\n|\r/g, '\n');

  var parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    parsed = [];
    var blocks = cleaned.split(/"slug"\s*:/);
    blocks.shift();
    blocks.forEach(function(block) {
      var slugMatch   = block.match(/^\s*"([^"]+)"/);
      var altMatch    = block.match(/"alt"\s*:\s*"([^"]+)"/);
      var figureMatch = block.match(/"figure_html"\s*:\s*"([\s\S]*?)"\s*}/);
      if (slugMatch) {
        parsed.push({
          slug:        slugMatch[1] || '',
          alt:         altMatch ? altMatch[1] : '',
          figure_html: figureMatch ? figureMatch[1].replace(/\\"/g, '"') : ''
        });
      }
    });
    if (parsed.length === 0) {
      return { success: false, message: "Could not parse response — regenerate the prompt and try again." };
    }
  }

  if (!Array.isArray(parsed) || parsed.length === 0) {
    return { success: false, message: "No items found in JSON response." };
  }

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var libSheet = ss.getSheetByName('Image Library');
  if (!libSheet) return { success: false, message: "'Image Library' sheet not found." };

  var lastRow = libSheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "Image Library sheet is empty." };

  var libData = libSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  var saved    = [];
  var notFound = [];

  parsed.forEach(function(item) {
    var slug = String(item.slug || "").trim();
    if (!slug) return;

    var matched = false;
    libData.forEach(function(r, i) {
      if (String(r[2]).trim() === slug) { // col C — Slug
        libSheet.getRange(i + 2, 12).setValue(item.figure_html || ''); // col L
        matched = true;
        saved.push(slug);
      }
    });

    if (!matched) notFound.push(slug);
  });

  var message = saved.length + " figure HTML block(s) saved to Image Library.";
  if (notFound.length > 0) {
    message += " Not matched: " + notFound.join(", ") + ".";
  }

  return { success: true, message: message };
}

/* ============================================================
   STAGE 2C — RECONCILE
   Reads W2B Raw HTML (column 151) and Image Suggestions JSON
   (column 167), injects an HTML comment block at the specified
   position for each suggestion, and saves the merged result to
   column 168 (FL — HTML with Image Suggestions).
============================================================ */
function reconcileImageSuggestions(sourceKey) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();

  var source = getW2CSourceColumn(sourceKey);
  var html = String(posts.getRange(row, source.col).getValue() || "").trim();
  var json = String(posts.getRange(row, 167).getValue() || "").trim(); // FK — Image Suggestions (JSON)

  if (!html) {
    return { success: false, message: "Column " + source.label + " is empty for this row." };
  }
  if (!json) {
    return { success: false, message: "Column FK (Image Suggestions JSON) is empty for this row." };
  }

  var suggestions;
  try {
    suggestions = JSON.parse(json);
  } catch (e) {
    return { success: false, message: "Invalid JSON in column FK — " + e.toString() };
  }

  if (!Array.isArray(suggestions) || suggestions.length === 0) {
    return { success: false, message: "No suggestions found in column FK JSON." };
  }

  var notFound = [];

  suggestions.forEach(function(s) {
    var h2Text = s.h2_text;
    if (!h2Text) return;

    // Build comment block
    var commentBlock = '\n<!-- IMAGE SUGGESTION\n' +
      'slug: '            + (s.slug           || '') + '\n' +
      'alt: '             + (s.alt            || '') + '\n' +
      'caption: '         + (s.caption        || '') + '\n' +
      'type: '            + (s.type           || '') + '\n' +
      'ai_prompt (photorealistic): ' + (s.ai_prompt || '') + '\n' +
      'ALT MEDIA TYPE: '  + (s.alt_media_type  || '') + '\n' +
      'ALT MEDIA BRIEF: ' + (s.alt_media_brief || '') + '\n' +
      '-->\n';

    // Determine heading level from suggestion — default to h2
    var headingLevel = String(s.heading_level || 'h2').toLowerCase();
    var tagName      = (headingLevel === 'h3') ? 'h3' : 'h2';

    var escaped      = h2Text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var headingRegex = new RegExp(
      '(<' + tagName + '[^>]*>[^<]*' + escaped + '[^<]*<\/' + tagName + '>)',
      'i'
    );

    if (!headingRegex.test(html)) {
      notFound.push(h2Text);
      return;
    }

    html = html.replace(headingRegex, '$1' + commentBlock);
  });

  posts.getRange(row, 168).setValue(html); // FL — HTML with Image Suggestions

  var message = "Reconciled and saved to column FL (168).";
  if (notFound.length > 0) {
    message += " H2s not found in HTML: " + notFound.join(" | ") + ".";
  }

  return { success: true, message: message, notFound: notFound };
}

/* ============================================================
   STAGE 2C — SAVE FINAL HTML
   Saves the pasted HTML (with images added in place of the
   suggestion comments) to column FL (168), overwriting the
   placeholder version.
============================================================ */
function saveW2CFinalHtmlToSheet(html) {
  if (!html || !html.trim()) {
    return { success: false, message: "No HTML provided." };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();
  posts.getRange(row, 168).setValue(html.trim()); // FL — HTML with Image Suggestions

  return { success: true, message: "Final HTML saved to column FL (168)." };
}

/* ============================================================
   STAGE 2C — LOAD FINAL HTML
   Reads column FL (168) — HTML with Image Suggestions — for
   loading into the W3 humanisation input.
============================================================ */
function getW2CFinalHtmlFromSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();
  var html  = String(posts.getRange(row, 168).getValue() || "").trim(); // FL

  if (!html) {
    return { success: false, message: "Column FL (HTML with Image Suggestions) is empty for this row." };
  }

  return { success: true, html: html };
}

/* ============================================================
   STAGE 2C — CHECK FL HAS CONTENT
   Called before opening the FL viewer modal.
============================================================ */
function checkW2CFLHasContent() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();
  var html  = String(posts.getRange(row, 168).getValue() || "").trim();

  if (!html) {
    return { success: false, message: "Column FL is empty for this row — run Reconcile first." };
  }
  return { success: true };
}

/* ============================================================
   STAGE 2C — OPEN FL VIEWER MODAL
============================================================ */
function openW2CFLViewerModal() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  var row   = posts.getActiveRange().getRow();
  var html  = String(posts.getRange(row, 168).getValue() || "").trim();

  var template = HtmlService.createTemplateFromFile('ce_W2C_FL_Viewer');
  template.htmlContent = html;
  var output = template.evaluate()
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(output, 'Column FL — HTML with Image Suggestions');
}