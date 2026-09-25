/**
 * ================================================================================
 * ce_W8B_SectionRewrite.gs - W8B — SECTION REWRITE
 * ================================================================================
 *
 * Combines the W8_Brief with the current New HTML to build a governed
 * rewrite instruction for ChatGPT. The instruction targets only the flagged
 * sections identified in the brief — leaving all other sections untouched.
 *
 * Output is pasted back and saved to W8B_HTML column in posts_progress.
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * ================================================================================
 */

function buildW8BRewritePrompt() {
  try {
    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet    = ss.getSheetByName('posts');
    var progressSheet = ss.getSheetByName('posts_progress');

    if (!postsSheet)    return { success: false, message: 'posts sheet not found.' };
    if (!progressSheet) return { success: false, message: 'posts_progress sheet not found.' };

    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return { success: false, message: 'Select a data row first.' };

    // Get Post ID
    var postHeaders = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).trim(); });
    var postIdIdx = postHeaders.indexOf('Post ID');
    var postId    = postIdIdx > -1 ? String(postsSheet.getRange(activeRow, postIdIdx + 1).getValue() || '').trim() : '';
    if (!postId) return { success: false, message: 'No Post ID found on active row.' };

    // Get Article Type
    var articleTypeIdx = postHeaders.indexOf('Article Type');
    var articleType    = articleTypeIdx > -1 ? String(postsSheet.getRange(activeRow, articleTypeIdx + 1).getValue() || '').trim() : '';

    // Get current HTML from posts col 98
    var html = String(postsSheet.getRange(activeRow, 98).getValue() || '').trim();
    if (!html) return { success: false, message: 'No HTML found in New HTML column (col 98).' };

    // Read W8_Brief from posts_progress
    var progressData    = progressSheet.getDataRange().getValues();
    var progressHeaders = progressData[0].map(function(h){ return String(h).trim(); });
    var progressColIdx  = {};
    for (var i = 0; i < progressHeaders.length; i++) {
      progressColIdx[progressHeaders[i]] = i;
    }

    // Read W8_Brief from posts sheet col 164 (FH)
    var brief = String(postsSheet.getRange(activeRow, 164).getValue() || '').trim();
    if (!brief) return { success: false, message: 'No W8_Brief found for this row. Run W8A first.' };

    // Build governed rewrite prompt
    var prompt = buildW8BPrompt(brief, articleType, html);

    return {
      success:  true,
      prompt:   prompt,
      message:  'W8B rewrite prompt ready — open a NEW ChatGPT chat before pasting.'
    };

  } catch(e) {
    return { success: false, message: 'buildW8BRewritePrompt error: ' + e.message };
  }
}

/* ============================================================
   PROMPT BUILDER
============================================================ */
function buildW8BPrompt(brief, articleType, html) {
  return 'W8B — TARGETED SECTION REWRITE\n\n' +
    'Read the following and respond as instructed:\n' +
    'ROLE: You are an expert UK SEO content editor working on a natural stone floor care website.\n\n' +

    'TASK: Apply the edit brief below to the HTML article provided. ' +
    'Rewrite ONLY the sections identified in the brief. ' +
    'All other sections must be returned completely unchanged — do not reword, restructure or improve them. ' +
    'Return the complete article HTML with only the flagged sections corrected.\n\n' +

    'ARTICLE TYPE: ' + articleType + '\n\n' +

    'CRITICAL RULES:\n' +
    '1. Return complete HTML only — no markdown, no explanation, no preamble.\n' +
    '2. Do not add, remove or reorder sections.\n' +
    '3. Do not change any images, figure tags, captions, links or schema.\n' +
    '4. Do not change the header, footer, bio box or hub-intro navigation.\n' +
    '5. Do not change any section not named in the brief.\n' +
    '6. Apply each instruction in the brief exactly as written.\n' +
    '7. Preserve all existing internal links in flagged sections — update surrounding prose only.\n\n' +

    '── EDIT BRIEF ─────────────────────────────────\n\n' +
    brief + '\n\n' +

    '── ARTICLE HTML ────────────────────────────────\n\n' +
    html;
}

/* ============================================================
   SAVE W8B OUTPUT
============================================================ */
function saveW8BHtml(html) {
  try {
    if (!html || html.trim() === '') {
      return { success: false, message: 'No HTML to save.' };
    }

    var ss            = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet    = ss.getSheetByName('posts');
    var progressSheet = ss.getSheetByName('posts_progress');

    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return { success: false, message: 'Select a data row first.' };

    var postHeaders = postsSheet.getRange(1, 1, 1, postsSheet.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).trim(); });
    var postIdIdx = postHeaders.indexOf('Post ID');
    var postId    = postIdIdx > -1 ? String(postsSheet.getRange(activeRow, postIdIdx + 1).getValue() || '').trim() : '';
    if (!postId) return { success: false, message: 'No Post ID found on active row.' };

    var progressData    = progressSheet.getDataRange().getValues();
    var progressHeaders = progressData[0].map(function(h){ return String(h).trim(); });
    var progressColIdx  = {};
    for (var i = 0; i < progressHeaders.length; i++) {
      progressColIdx[progressHeaders[i]] = i;
    }

    // Save W8B_HTML to posts sheet col 165 (FI)
    var cell = postsSheet.getRange(activeRow, 165);
    cell.setNumberFormat('@');
    cell.setValue(html.trim());
    return { success: true, message: 'W8B HTML saved to posts sheet col 165 (FI).' };

  } catch(e) {
    return { success: false, message: 'saveW8BHtml error: ' + e.message };
  }
}

/* ============================================================
   LOAD W8B HTML FOR W8C
============================================================ */
function loadW8BHtmlForW8C() {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var postsSheet = ss.getSheetByName('posts');

    var activeRow = postsSheet.getActiveRange().getRow();
    if (activeRow < 2) return { success: false, message: 'Select a data row first.' };

    var html = String(postsSheet.getRange(activeRow, 165).getValue() || '').trim();
    if (!html) return { success: false, message: 'No W8B HTML found for this row. Run W8B first.' };

    return { success: true, html: html };

  } catch(e) {
    return { success: false, message: 'loadW8BHtmlForW8C error: ' + e.message };
  }
}