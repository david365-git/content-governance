/**
 * ================================================================================
 * ce_W8C_Humanisation.gs - W8C — HUMANISATION
 * ================================================================================
 *
 * Mirrors the W3 humanisation pattern exactly.
 * Reads humanisation rules from prompts sheet D42.
 * Loads W8B HTML from posts_progress W8B_HTML column.
 * Output saved to W8C_HTML column in posts_progress.
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * ================================================================================
 */

function getW8CHumanisationRules() {
  try {
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var promptSheet = ss.getSheetByName('prompts');
    if (!promptSheet) return { success: false, message: 'prompts sheet not found.' };

    var humanisePrompt = String(promptSheet.getRange('D42').getValue() || '').trim();
    if (!humanisePrompt) return { success: false, message: 'No humanisation prompt found in prompts sheet D42.' };

    return { success: true, prompt: humanisePrompt };

  } catch(e) {
    return { success: false, message: 'getW8CHumanisationRules error: ' + e.message };
  }
}

function saveW8CHtml(html) {
  try {
    if (!html || html.trim() === '') {
      return { success: false, message: 'No HTML to save.' };
    }

    // Safety net — remove any stray David_Allen.jpg figures
    // Bio box is injected at push time in ce_Utilities.gs
    if (/David_Allen\.jpg/i.test(html)) {
      html = html.replace(/<figure[^>]*>[\s\S]*?David_Allen\.jpg[\s\S]*?<\/figure>/gi, '');
      html = html.replace(/<img[^>]*David_Allen\.jpg[^>]*\/?>/gi, '');
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

    // Save W8C_HTML to posts sheet col 166 (FJ)
    var cell = postsSheet.getRange(activeRow, 166);
    cell.setNumberFormat('@');
    cell.setValue(html.trim());
    return { success: true, message: 'W8C HTML saved to posts sheet col 166 (FJ).' };

  } catch(e) {
    return { success: false, message: 'saveW8CHtml error: ' + e.message };
  }
}
