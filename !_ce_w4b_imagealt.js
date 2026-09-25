/**
 * ================================================================================
 * ce_W4B_ImageAlt.gs - W4B IMAGEALT
 * ================================================================================
 * 
 * Image alt text and caption generation
 * 
 * Part of Abbey Floor Care Content Pipeline v77+
 * Reorganised: July 2026
 * ================================================================================
 */



/* ── W4B: Load existing HTML from sheet for alt/caption processing ── */
function getHtmlForW4B() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { html: null };

    var html = String(sheet.getRange(row, 98).getValue() || '').trim();
    return { html: html || null };
  } catch(e) {
    return { html: null };
  }
}

/* ── Insert LLM-returned alt/caption JSON into HTML and push to sheet ── */
function pushAltCaptionUpdates(rawJson) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'ERROR: Select a data row first.' };

    var clean = rawJson.replace(/```json|```/gi, '').trim();
    var updates;
    try {
      updates = JSON.parse(clean);
    } catch(e) {
      return { success: false, message: 'ERROR: Could not parse JSON — ' + e.message };
    }

    if (!Array.isArray(updates) || updates.length === 0) {
      return { success: false, message: 'ERROR: Expected a JSON array of image updates.' };
    }

    var html = String(sheet.getRange(row, 104).getValue() || '');
    if (!html) html = String(sheet.getRange(row, 98).getValue() || '');
    if (!html) return { success: false, message: 'ERROR: No HTML found in col 104 (CZ) or col 98 (CT).' };

    var updated  = 0;
    var captioned = 0;

    updates.forEach(function(item) {
      var filename = item.filename || (item.src ? item.src.split('/').pop() : null);
      if (!filename) return;

      // Escape filename for regex
      var esc = filename.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');

      // Match the full <img> tag containing this filename in src attribute
      var tagPattern = '<img((?:[^>]*?))\\s+src=(?:"[^"]*' + esc + '[^"]*"|\'[^\']*' + esc + '[^\']*\')([^>]*?)\\/?>'; 
      var imgRegex   = new RegExp(tagPattern, 'i');
      var match      = html.match(imgRegex);
      if (!match) return;

      // Skip images marked with data-w4b-skip="true" — these were injected by W2C
      if (/data-w4b-skip\s*=\s*["']true["']/i.test(match[0])) return;

      var fullTag = match[0];
      var before  = match[1]; // attributes before src
      var after   = match[2]; // attributes after src

      // Get the actual src URL from the matched tag
      var srcMatch = fullTag.match(/src=(?:"([^"]*)"|'([^']*)')/i);
      var actualSrc = srcMatch ? (srcMatch[1] || srcMatch[2]) : item.src;

      // Strip any existing alt from before/after
      var cleanBefore = before.replace(/\s*alt=(?:"[^"]*"|'[^']*')/gi, '');
      var cleanAfter  = after.replace(/\s*alt=(?:"[^"]*"|'[^']*')/gi, '');

      // Build alt attribute
      var altVal  = (item.alt !== null && item.alt !== undefined) ? String(item.alt) : '';
      var altAttr = ' alt="' + altVal.replace(/"/g, '&quot;') + '"';

      var newImg = '<img' + cleanBefore + ' src="' + actualSrc + '"' + altAttr + cleanAfter + '>';

      // Never wrap an image in <figure> if it sits inside a <td> — this is a
      // commercial product table cell, and wrapping would break the preserved
      // table markup required by Rule 10, regardless of what caption the LLM sent.
      var imgPosInHtml = html.indexOf(fullTag);
      var precedingHtml = imgPosInHtml > -1 ? html.substring(0, imgPosInHtml) : '';
      var lastTdOpen  = precedingHtml.lastIndexOf('<td');
      var lastTdClose = precedingHtml.lastIndexOf('</td>');
      var isInsideTableCell = lastTdOpen > lastTdClose;

      // Wrap in figure/figcaption if caption provided and not already in a figure
      var replacement = newImg;
      if (!isInsideTableCell && item.caption && String(item.caption).trim()) {
        var captionText = String(item.caption).trim();
        // Check if already inside a figure — if so update the figcaption, otherwise wrap
        var figPattern = '(<figure[^>]*>[\\s\\S]*?)' + esc + '([\\s\\S]*?<\\/figure>)';
        var figCheck   = new RegExp(figPattern, 'i');
        if (figCheck.test(html)) {
          // Already in a figure — replace the existing figcaption text
          html = html.replace(figCheck, function(m) {
            return m.replace(/<figcaption[^>]*>[\s\S]*?<\/figcaption>/i,
              '<figcaption class="wp-caption-text">' + captionText + '</figcaption>');
          });
          captioned++;
        } else {
          replacement = '<figure class="wp-caption aligncenter" style="width: 700px">' + newImg + '<figcaption class="wp-caption-text">' + captionText + '</figcaption></figure>';
          captioned++;
        }
      }

      html = html.replace(fullTag, replacement);
      updated++;
    });

    if (updated === 0) {
      return { success: false, message: 'WARNING: No matching img tags found. Check the src filenames match the HTML.' };
    }

    var cell = sheet.getRange(row, 104);
    try {
      cell.setPlainTextValue(html);
    } catch(e) {
      cell.setNumberFormat('@');
      cell.setValue(html);
    }

    logPipelineResume("W4B — Alt Caption Update", "");
    return {
      success: true,
      message: updated + ' image(s) updated, ' + captioned + ' caption(s) added. HTML pushed to sheet.'
    };

  } catch(e) {
    return { success: false, message: 'PUSH ERROR: ' + e.toString() };
  }
}