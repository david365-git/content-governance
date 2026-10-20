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

    const mediaRule = '\n\nHARD MEDIA EXCLUSION: Humanise article prose only. Do not rewrite or reorder <figure> blocks, <figcaption>, <img> alt attributes, <iframe> or video captions. Preserve complete media blocks verbatim and in their original positions. Do not add or remove images or videos. The saving step will restore all original media blocks and reject a mismatched inventory.\n';
    const combined = humanisePrompt + mediaRule + '\n\n---\n\n' + w2bHtml.trim();

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

/**
 * Restore media unchanged from GI after the language model has humanised prose.
 * Full figures include their captions, image alts and embedded video data.
 * Refuse incomplete/reordered inventories before writing CZ.
 */
function ce_restoreHumanisationMedia_(source, candidate) {
  var pattern = /<figure\b[^>]*>[\s\S]*?<\/figure>|<iframe\b[^>]*>[\s\S]*?<\/iframe>|<img\b[^>]*>/gi;
  function inventory(s) {
    var items=[], m;
    var re=new RegExp(pattern.source,'gi');
    while((m=re.exec(s))!==null) {
      var kind=/^<figure\b/i.test(m[0])?'figure':/^<iframe\b/i.test(m[0])?'iframe':'img';
      var asset=(m[0].match(/(?:src|data-src)\s*=\s*["']([^"']+)["']/i)||[])[1]||'';
      items.push({kind:kind,html:m[0],asset:asset});
    }
    return items;
  }
  var originals=inventory(source), revised=inventory(candidate);
  if(originals.length!==revised.length) {
    throw new Error('Humanisation changed media count ('+originals.length+' to '+revised.length+'). CZ was not saved.');
  }
  for(var i=0;i<originals.length;i++) {
    if(originals[i].kind!==revised[i].kind) {
      throw new Error('Humanisation changed media order at item '+(i+1)+'. CZ was not saved.');
    }
    // Check an available asset identity; otherwise preserve position/kind.
    if(originals[i].asset && revised[i].asset &&
       originals[i].asset!==revised[i].asset) {
      throw new Error('Humanisation changed media source at item '+(i+1)+'. CZ was not saved.');
    }
  }
  var index=0;
  var restored=candidate.replace(pattern,function(){return originals[index++].html;});
  return restored;
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

    // Media is outside humanisation scope: take its unchanged markup from GI.
    var originalGi=String(sheet.getRange(row,191).getValue()||'').trim();
    if(!originalGi) return {success:false,message:'GI source HTML missing; CZ was not saved.'};
    html=ce_restoreHumanisationMedia_(originalGi,html);

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