/**
 * ================================================================================
 * ce_VideoManager.gs — VIDEO INVENTORY & DECISION MANAGER
 * ================================================================================
 *
 * Scans existing article HTML for video iframes, writes detected videos to
 * EC (135) Video Inventory, and reads/writes placement decisions to ED (136)
 * Video Decisions.
 *
 * Column map:
 *   Col 135 (EC) — Video Inventory  (JSON written by scanVideosForActiveRow)
 *   Col 136 (ED) — Video Decisions  (JSON written by saveVideoDecisions)
 *
 * Called from sidebar UI: ce_Sidebar_VideoManager.html
 *
 * Part of Abbey Floor Care Content Pipeline
 * ================================================================================
 */

var VIDEO_INVENTORY_COL  = 133; // EC
var VIDEO_DECISIONS_COL  = 134; // ED

// ---------------------------------------------------------------------------
// SCAN — called when user opens the Video Manager panel
// Reads existing HTML from site-export or current HTML column,
// detects all iframes, extracts section list, writes Video Inventory to EC.
// Returns data object for sidebar rendering.
// ---------------------------------------------------------------------------
function scanVideosForActiveRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row   = sheet.getActiveRange().getRow();

  if (row < 2) {
    return { error: 'Please select a data row (not the header row).' };
  }

  // --- Read existing HTML ---
  // Read HTML — prefer col 104 (Humanised), fall back to col 98 (New HTML), then col 102 (W2B output)
  var humanisedHtml  = sheet.getRange(row, 104).getValue().toString().trim();
  var newHtml        = sheet.getRange(row, 98).getValue().toString().trim();
  var w2bHtml        = sheet.getRange(row, 102).getValue().toString().trim();
  var sourceHtml     = humanisedHtml || newHtml || w2bHtml;

  if (!sourceHtml) {
    return { error: 'No HTML found in col 104 (CZ), col 98 (CT), or col 102 (CX) for this row.' };
  }

  // --- Extract video iframes ---
  var videos = extractVideosFromHtml(sourceHtml);

  // --- Extract section list from HTML ---
  var sections = extractSectionsFromHtml(sourceHtml);

  // --- Detect header video ---
  var headerVideoUrl = detectHeaderVideo(sourceHtml);

  // --- Read existing decisions if any ---
  var existingDecisionsRaw = sheet.getRange(row, VIDEO_DECISIONS_COL).getValue().toString().trim();
  var existingDecisions = {};
  if (existingDecisionsRaw) {
    try {
      existingDecisions = JSON.parse(existingDecisionsRaw);
    } catch(e) {
      existingDecisions = {};
    }
  }

  // --- Write Video Inventory to EC ---
  var inventory = {
    scanned_at: new Date().toISOString(),
    source: humanisedHtml ? 'humanised' : newHtml ? 'new-html' : 'w2b-output',
    videos: videos,
    sections: sections,
    header_video_url: headerVideoUrl
  };

  sheet.getRange(row, VIDEO_INVENTORY_COL).setValue(JSON.stringify(inventory));

  return {
    success: true,
    row: row,
    videos: videos,
    sections: sections,
    headerVideoUrl: headerVideoUrl,
    existingDecisions: existingDecisions,
    highlightedHtml: buildHighlightedHtml(sourceHtml, videos)
  };
}

// ---------------------------------------------------------------------------
// SAVE DECISIONS — called when user confirms their video decisions
// Writes JSON to ED (136)
// ---------------------------------------------------------------------------
function saveVideoDecisions(row, decisions) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!row || row < 2) {
    return { error: 'Invalid row.' };
  }

  // decisions is an array of objects:
  // { url, action, section_number, section_title, youtube_id }
  // action: 'header' | 'keep' | 'move' | 'remove'

  var decisionsJson = JSON.stringify({
    saved_at: new Date().toISOString(),
    decisions: decisions
  });

  sheet.getRange(row, VIDEO_DECISIONS_COL).setValue(decisionsJson);

  return { success: true, saved: decisions.length };
}

// ---------------------------------------------------------------------------
// READ DECISIONS — called by pipeline stages to get placement instructions
// ---------------------------------------------------------------------------
function getVideoDecisionsForRow(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var raw   = sheet.getRange(row, VIDEO_DECISIONS_COL).getValue().toString().trim();

  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch(e) {
    return null;
  }
}

// ---------------------------------------------------------------------------
// EXTRACT VIDEOS — parses HTML and returns array of video objects
// ---------------------------------------------------------------------------
function extractVideosFromHtml(html) {
  var videos  = [];
  var pattern = /<iframe[^>]+src=["']([^"']*(?:youtube|youtu\.be|vimeo)[^"']*)["'][^>]*>/gi;
  var match;
  var index   = 0;

  while ((match = pattern.exec(html)) !== null) {
    var iframeSrc  = match[1];
    var fullIframe = match[0];
    var youtubeId  = extractYoutubeId(iframeSrc);
    var position   = detectVideoPosition(html, match.index);

    videos.push({
      index:        index,
      url:          iframeSrc,
      youtube_id:   youtubeId,
      iframe_html:  fullIframe,
      position:     position,  // 'header' | 'section-N' | 'unknown'
      char_offset:  match.index
    });

    index++;
  }

  return videos;
}

// ---------------------------------------------------------------------------
// EXTRACT SECTIONS — reads H2 headings from section blocks
// ---------------------------------------------------------------------------
function extractSectionsFromHtml(html) {
  var sections = [];
  var pattern  = /<section[^>]*id=["']section-(\d+)["'][^>]*>[\s\S]*?<h2[^>]*>([\s\S]*?)<\/h2>/gi;
  var match;

  while ((match = pattern.exec(html)) !== null) {
    var sectionNum = parseInt(match[1], 10);
    var h2Text     = match[2].replace(/<[^>]+>/g, '').trim();

    sections.push({
      number: sectionNum,
      title:  h2Text
    });
  }

  return sections;
}

// ---------------------------------------------------------------------------
// DETECT HEADER VIDEO — checks if a video exists inside <header> block
// ---------------------------------------------------------------------------
function detectHeaderVideo(html) {
  var headerMatch = html.match(/<header[\s\S]*?<\/header>/i);
  if (!headerMatch) return null;

  var headerHtml  = headerMatch[0];
  var iframeMatch = headerHtml.match(/src=["']([^"']*(?:youtube|youtu\.be|vimeo)[^"']*)["']/i);
  return iframeMatch ? iframeMatch[1] : null;
}

// ---------------------------------------------------------------------------
// DETECT VIDEO POSITION — determines whether a video is in header or section
// ---------------------------------------------------------------------------
function detectVideoPosition(html, charOffset) {
  // Find the nearest opening tag context before the iframe
  var preceding = html.substring(0, charOffset);

  // Check if inside <header>
  var lastHeader  = preceding.lastIndexOf('<header');
  var lastHClose  = preceding.lastIndexOf('</header>');
  if (lastHeader > lastHClose) return 'header';

  // Check which section it's in
  var sectionPattern = /<section[^>]*id=["']section-(\d+)["']/gi;
  var match;
  var lastSection = null;
  while ((match = sectionPattern.exec(preceding)) !== null) {
    lastSection = match[1];
  }

  return lastSection ? 'section-' + lastSection : 'unknown';
}

// ---------------------------------------------------------------------------
// EXTRACT YOUTUBE ID from URL
// ---------------------------------------------------------------------------
function extractYoutubeId(url) {
  var match = url.match(/(?:embed\/|v=|youtu\.be\/)([a-zA-Z0-9_-]{11})/);
  return match ? match[1] : null;
}

// ---------------------------------------------------------------------------
// BUILD HIGHLIGHTED HTML — wraps each iframe block in a highlight span
// Returns HTML string with video blocks marked for display in sidebar
// ---------------------------------------------------------------------------
function buildHighlightedHtml(html, videos) {
  if (!videos || videos.length === 0) return html;

  var colours = ['#FFD700', '#90EE90', '#87CEEB', '#FFB6C1', '#DDA0DD'];
  var result  = html;
  var offset  = 0;

  // Sort by char_offset ascending so replacements don't shift positions
  var sorted = videos.slice().sort(function(a, b) { return a.char_offset - b.char_offset; });

  sorted.forEach(function(video, i) {
    var colour    = colours[i % colours.length];
    var label     = 'VIDEO ' + (video.index + 1) + ' [' + video.position + ']';
    var highlight =
      '<span style="background:' + colour + ';padding:2px 4px;font-weight:bold;font-size:11px;">' +
      label + '</span>' +
      video.iframe_html;

    var pos    = result.indexOf(video.iframe_html);
    if (pos !== -1) {
      result = result.substring(0, pos) + highlight + result.substring(pos + video.iframe_html.length);
    }
  });

  return result;
}