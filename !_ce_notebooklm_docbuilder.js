/**
 * ================================================================================
 * ce_NotebookLM_DocBuilder.gs — NOTEBOOKLM SOURCE CONTENT BUILDER
 * ================================================================================
 *
 * Builds a structured "Project Dossier" text block for NotebookLM.
 * Returns the formatted text to the sidebar for copying directly into
 * NotebookLM as a "Copied Text" source.
 *
 * No Drive or Docs permissions required.
 *
 * Part of Abbey Floor Care Content Pipeline
 * ================================================================================
 */

function buildNotebookLMDoc() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(function(h) { return String(h).trim(); });

    function colVal(name) {
      var idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || '').trim() : '';
    }

    // ── Read governed fields ──
    var postId         = colVal('Post ID').replace(/\.0$/, '');
    var h1             = colVal('New H1');
    var metaDesc       = colVal('New Meta Description');
    var stoneType      = colVal('Stone Type');
    var articleType    = colVal('Article Type');
    var locality       = colVal('Locality');
    var primaryCluster = colVal('Primary Query Cluster Owned');
    var supportingEnts = colVal('Supporting Entities Core');

    // ── Read New HTML (col 98) ──
    var newHtml = String(sheet.getRange(row, 98).getValue() || '').trim();
    if (!newHtml) {
      return { success: false, message: 'No HTML found in New HTML column (col 98). Run W4 first.' };
    }

    // ── Extract alt text from images in New HTML ──
    var altTexts = [];
    var altRe = /<img[^>]+alt="([^"]+)"[^>]*>/gi;
    var altMatch;
    var SKIP_ALT = /david.allen|logo|icon|avatar/i;
    while ((altMatch = altRe.exec(newHtmlCleaned)) !== null) {
      var alt = altMatch[1].trim();
      if (alt && !SKIP_ALT.test(alt)) altTexts.push(alt);
    }

    // ── Strip HTML tags for plain text body ──
    var newHtmlCleaned = newHtml
      .replace(/David Allen/gi, 'Abbey Floor Care')
      .replace(/In my experience/gi, 'In Abbey Floor Care\'s experience')
      .replace(/my experience/gi, 'Abbey Floor Care\'s experience');
    var plainText = newHtmlCleaned
      .replace(/<header[\s\S]*?<\/header>/gi, '')
      .replace(/<footer[\s\S]*?<\/footer>/gi, '')
      .replace(/<div[^>]*abbey-bio-box[^>]*>[\s\S]*?<\/div>/gi, '')
      .replace(/<figure[^>]*class="video-embed"[^>]*>[\s\S]*?<\/figure>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&mdash;/g, '—')
      .replace(/&nbsp;/g, ' ')
      .replace(/\s{2,}/g, ' ')
      .trim();

    // ── Build structured text ──
    var today   = new Date();
    var dateStr = today.getFullYear() + '-' +
                  String(today.getMonth() + 1).padStart(2, '0') + '-' +
                  String(today.getDate()).padStart(2, '0');
    var docName = postId + ' — ' + dateStr;

    var output = '';
    output += 'PROJECT FACT SHEET\n';
    output += '==================\n\n';
    output += 'Official Project Title: ' + (h1 || 'Not set') + '\n';
    output += 'Company: Abbey Floor Care — specialist natural stone and tile restoration company, run by David Allen with over 30 years of hands-on experience.\n';
    output += 'Core Summary: ' + (metaDesc || 'Not set') + '\n';
    output += 'Location / Territory: ' + (locality || 'Not set') + '\n';
    output += 'Stone Classification: ' + (stoneType || 'Not set') + '\n';
    output += 'Article Type: ' + (articleType || 'Not set') + '\n';
    output += 'Project Goal: ' + (primaryCluster || 'Not set') + '\n\n';

    output += 'TECHNICAL SPECIFICATIONS\n';
    output += '========================\n\n';
    if (supportingEnts) {
      supportingEnts.split(',').forEach(function(e) {
        output += '• ' + e.trim() + '\n';
      });
    } else {
      output += 'No supporting entities recorded.\n';
    }
    output += '\n';

    output += 'VISUAL STORYBOARD\n';
    output += '=================\n\n';
    if (altTexts.length > 0) {
      altTexts.forEach(function(alt, i) {
        output += 'Scene ' + (i + 1) + ': ' + alt + '\n';
      });
    } else {
      output += 'No image alt text found.\n';
    }
    output += '\n';

    output += 'FULL CONTENT REFERENCE\n';
    output += '======================\n\n';
    output += plainText;

    return {
      success:  true,
      message:  'Content ready — copy and paste into NotebookLM as a Copied Text source.',
      docName:  docName,
      content:  output
    };

  } catch(e) {
    return { success: false, message: 'BUILD ERROR: ' + e.toString() };
  }
}
function generateNotebookLMDocName() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return 'Select a data row first.';

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                       .getValues()[0]
                       .map(function(h) { return String(h).trim(); });

    function colVal(name) {
      var idx = headers.indexOf(name);
      return idx > -1 ? String(sheet.getRange(row, idx + 1).getValue() || '').trim() : '';
    }

    var postId      = colVal('Post ID').replace(/\.0$/, '');
    var material    = colVal('Stone Type');
    var articleType = colVal('Article Type');
    var locality    = colVal('Locality');
    var primaryTerm = colVal('Primary Query Cluster Owned');

    var topic = locality || primaryTerm || 'General';

    var today   = new Date();
    var dateStr = today.getFullYear() + '-' +
                  String(today.getMonth() + 1).padStart(2, '0') + '-' +
                  String(today.getDate()).padStart(2, '0');

    return postId + ' — ' + material + ' — ' + articleType + ' — ' + topic + ' — ' + dateStr;

  } catch(e) {
    return 'Error generating name: ' + e.toString();
  }
}