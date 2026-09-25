/**
 * ================================================================================
 * ce_stage3_humanisation.gs - STAGE 3 — HUMANISATION
 * ================================================================================
 *
 * Reads the humanisation prompt from prompts sheet D42 and combines it
 * with the pasted W2B HTML for copying into a fresh ChatGPT chat.
 * Saves the humanised output to Column CZ (104) for audit purposes.
 *
 * Bio box is NOT injected here — it is injected at push time in
 * pushHtmlToActiveRow() in ce_Utilities.gs, which is the final step.
 * The LLM outputs only a BIO_PARAGRAPH: marker — the script builds
 * the complete bio box HTML when pushing to col 98.
 *
 * Reads:  prompts sheet D42 — humanisation prompt
 *         HTML pasted in sidebar textarea
 * Saves:  Col 104 (CZ) — Humanised HTML (audit copy)
 *
 * Part of Abbey Floor Care Content Pipeline v77+
 * Updated: March 2026
 * ================================================================================
 */

function buildHumanisationPrompt(w2bHtml) {
  try {
    if (!w2bHtml || w2bHtml.trim() === '') {
      return { success: false, message: 'ERROR: No HTML provided. Paste W2B output first.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const promptSheet = ss.getSheetByName('prompts');
    if (!promptSheet) {
      return { success: false, message: 'ERROR: prompts sheet not found.' };
    }

    const humanisePrompt = String(promptSheet.getRange('D42').getValue() || '').trim();
    if (!humanisePrompt) {
      return { success: false, message: 'ERROR: No humanisation prompt found in prompts sheet D42.' };
    }

    const combined = humanisePrompt + '\n\n---\n\n' + w2bHtml.trim();

    return {
      success: true,
      prompt: combined,
      message: 'Humanisation prompt ready — open a NEW ChatGPT chat before pasting.'
    };

  } catch(e) {
    return { success: false, message: 'ERROR: ' + e.message };
  }
}

function getHumanisationRulesOnly() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const promptSheet = ss.getSheetByName('prompts');
    if (!promptSheet) {
      return { success: false, message: 'ERROR: prompts sheet not found.' };
    }
    const humanisePrompt = String(promptSheet.getRange('D42').getValue() || '').trim();
    if (!humanisePrompt) {
      return { success: false, message: 'ERROR: No humanisation prompt found in prompts sheet D42.' };
    }
    return { success: true, prompt: humanisePrompt };
  } catch(e) {
    return { success: false, message: 'ERROR: ' + e.message };
  }
}

function saveHumanisedHtml(html) {
  try {
    if (!html || html.trim() === '') {
      return { success: false, message: 'No HTML to save.' };
    }

    // Safety net — remove any David_Allen.jpg standalone figures
    // Bio box is injected at push time in ce_Utilities.gs — not here
    var htmlWithoutBioBox = html.replace(/<div[^>]*abbey-bio-box[^>]*>[\s\S]*?<\/div>/i, '');
    if (/David_Allen\.jpg/i.test(htmlWithoutBioBox)) {
      html = html.replace(/<figure[^>]*>[\s\S]*?David_Allen\.jpg[\s\S]*?<\/figure>/gi, '');
      html = html.replace(/<img[^>]*David_Allen\.jpg[^>]*\/?>/gi, '');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    // Save to Column CZ (104) — humanised HTML audit copy
    const cell = sheet.getRange(row, 104);
    cell.setNumberFormat('@');
    cell.setValue(html.trim());

    logPipelineResume("W3 — Humanised HTML", "");
    return {
      success: true,
      message: 'Humanised HTML saved to Column CZ — bio box will be injected at push.'
    };

  } catch(e) {
    return { success: false, message: 'ERROR: ' + e.message };
  }
}

function getHumanisedHtmlForW4B() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('posts');
    const row = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    const html = String(sheet.getRange(row, 104).getValue() || '').trim();
    if (!html) {
      return { success: false, message: 'No humanised HTML found in Column CZ. Run W3 first.' };
    }

    return { success: true, html: html };

  } catch(e) {
    return { success: false, message: 'ERROR: ' + e.message };
  }
}

// Legacy alias
function getHtmlForW4B() {
  return getHumanisedHtmlForW4B();
}