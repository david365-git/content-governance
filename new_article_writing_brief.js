/**
 * V2 ARTICLE WRITING BRIEF — SHADOW/READ-ONLY
 * Builds a content-quality-first brief for one posts row. Does not call AI
 * or edit any sheet, old HTML, source evidence, or pipeline state.
 */
function new_buildArticleWritingBrief(rowNumber) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var posts = ss.getSheetByName('posts');
    if (!posts) return {success:false,message:'posts sheet missing'};
    var row = Number(rowNumber);
    if (!Number.isInteger(row) || row < 2 || row > posts.getLastRow())
      return {success:false,message:'Valid posts row required'};
    var headers = posts.getRange(1,1,1,posts.getLastColumn()).getValues()[0].map(String);
    var values = posts.getRange(row,1,1,headers.length).getValues()[0];
    var d = {};
    headers.forEach(function(h,i){ d[String(h).trim()] = values[i]; });
    function val(k){return String(d[k] == null ? '' : d[k]).trim();}
    var articleType = val('Article Type'), material = val('Stone Type');
    var postId = val('Post ID');
    if (!postId && d['ID'] != null) postId = val('ID');
    if (!articleType || !material || !postId)
      return {success:false,message:'Post ID, Article Type and Stone Type must be present'};

    // Use authoritative original, NOT a later W2B/W3 output. Never fall back silently.
    var exportSheet = ss.getSheetByName('site-export');
    if (!exportSheet) return {success:false,message:'site-export sheet missing; original article not verified'};
    var exported = exportSheet.getDataRange().getValues();
    if (!exported.length) return {success:false,message:'site-export is empty'};
    var eh = exported[0].map(function(v){return String(v).trim();});
    var idCol = eh.indexOf('ID'), htmlCol = eh.indexOf('Full Post HTML');
    if (idCol < 0 || htmlCol < 0)
      return {success:false,message:'site-export ID/Full Post HTML columns missing'};
    var matches = exported.slice(1).filter(function(v){return String(v[idCol]).trim() === postId;});
    if (matches.length !== 1) return {success:false,message:'Expected one source article for ID '+postId+'; found '+matches.length};
    var originalHtml = String(matches[0][htmlCol] || '').trim();
    if (!originalHtml) return {success:false,message:'Original HTML is empty for ID '+postId};

    var typeRules = '';
    if (typeof getArticleTypeExpectations === 'function') {
      typeRules = String(getArticleTypeExpectations(articleType) || '').trim();
      if (!ss.getSheetByName('Article Type Control Sheet') || !typeRules || /error:|not found|no expectations defined/i.test(typeRules))
        return {success:false,message:'Article-type control missing for '+articleType+': '+typeRules};
    } else {
      return {success:false,message:'Article Type Control Sheet reader unavailable'};
    }
    var angle = val('Problem Angle'), cluster = val('Primary Query Cluster Owned');
    if (!cluster) return {success:false,message:'Primary Query Cluster Owned missing'};
    var axes = ['ac_entry_condition','ac_homeowner_perception','ac_material_behaviour',
      'ac_constraint','ac_process_emphasis','ac_result_type','ac_narrative_archetype'];
    var selection = axes.map(function(k){return k+': '+(val(k)||'[not classified]');}).join('\n');
    var acLibrary = getAcValuesForMaterial(material, articleType);
    if (!acLibrary || !acLibrary.values || !Object.keys(acLibrary.values).length)
      return {success:false,message:'AC library not available for this article type/material'};
    var libraryKeys = ['ac_what_they_found','ac_why_they_called',
      'ac_what_the_floor_was_doing','ac_what_forced_careful_decisions',
      'ac_what_governed_the_approach','ac_what_changed_for_the_homeowner',
      'ac_story_direction'];
    var acProblems = axes.filter(function(k,i){
      return val(k) && (acLibrary.values[libraryKeys[i]] || []).indexOf(val(k)) === -1;
    });
    if (acProblems.length)
      return {success:false,message:'Invalid AC choices: '+acProblems.join(', ')};
    selection = axes.map(function(k,i){
      return String((acLibrary.labels || {})[libraryKeys[i]] || libraryKeys[i])+': '+(val(k)||'[not classified]');
    }).join('\n');
    // Article-specific truth-source columns GL:HA, preserved with their live row-1 headings.
    // Read by position so duplicate/renamed header keys cannot silently suppress facts.
    // Column 194 = GL, column 209 = HA.
    var truthStart = 194, truthEnd = 209;
    if (headers.length < truthEnd)
      return {success:false,message:'Article truth-source range GL:HA is missing or incomplete'};
    var truthFields = [];
    for (var ti = truthStart - 1; ti < truthEnd; ti++) {
      var truthHeader = String(headers[ti] || '').trim();
      var truthValue = String(values[ti] == null ? '' : values[ti]).trim();
      if (!truthHeader)
        return {success:false,message:'Truth-source column '+(ti+1)+' (GL:HA) has no heading'};
      truthFields.push({heading:truthHeader,value:truthValue});
    }
    // Keep the ten organised fact categories (GL:GU) plus evidence
    // relationships, coverage, media, uncertainties and status (GW:HA).
    // GV holds the full v10 package and repeats the same fact records:
    // validate it via bc_caseEvidencePromptBlock_ below but do not send its
    // bulky JSON in the writing-brief prompt.
    if (truthFields[10].heading !== 'Case Study Evidence Package')
      return {success:false,message:'Expected Case Study Evidence Package in GV; evidence layout changed'};
    var promptTruthFields = truthFields.filter(function(x,i){return i !== 10;});
    var truthFilled = promptTruthFields.filter(function(x){return !!x.value;});
    var truthBlock = promptTruthFields.map(function(x){
      return x.heading+': '+(x.value || '[not documented]');
    }).join('\n\n');
    var evidence = '';
    if (articleType === 'Case Study') {
      if (typeof bc_caseEvidencePromptBlock_ !== 'function')
        return {success:false,message:'Case Study evidence reader not available'};
      evidence = bc_caseEvidencePromptBlock_(row);
      if (!evidence || !evidence.includes('AUTHORITATIVE CASE STUDY FACTS'))
        return {success:false,message:'Case Study evidence package unavailable'};
    }
    if (typeof getTechnicalMatrix !== 'function')
      return {success:false,message:'Technical Matrix reader unavailable'};
    var technical = getTechnicalMatrix(material);
    var media = {
      images:(originalHtml.match(/<img\b[^>]*>/gi)||[]).map(function(tag){
        return {src:(tag.match(/\bsrc\s*=\s*["']([^"']+)["']/i)||[])[1]||'',
          alt:(tag.match(/\balt\s*=\s*["']([^"']*)["']/i)||[])[1]||''};
      }).filter(function(x){return !!x.src;}),
      videos:(originalHtml.match(/<(?:iframe|video|source)\b[^>]*>/gi)||[]).map(function(tag){
        return (tag.match(/\bsrc\s*=\s*["']([^"']+)["']/i)||[])[1]||'';
      }).filter(Boolean)
    };
    var scopeKeys = ['Primary Entity','Supporting Entities Core','Page Rewrite Brief',
      'Rewrite Governance Summary','Cannibalisation Guardrail','Safe Handoff Pages',
      'Observed Query Cluster','Entity Role','Feeds Hub','Article Angle',
      'Rewrite Status','Page Governance Summary','Peripheral Entities Link Out'];
    var scope = scopeKeys.map(function(k){return k+': '+(val(k)||'[not supplied]');}).join('\n');
    var benchmarkQueries = typeof bc_getQueriesFromSheet === 'function'
      ? String(bc_getQueriesFromSheet(postId,'GSC Benchmark') || '').trim() : '';
    var monitorQueries = typeof bc_getQueriesFromSheet === 'function'
      ? String(bc_getQueriesFromSheet(postId,'GSC Monitor') || '').trim() : '';
    benchmarkQueries = benchmarkQueries || val('Benmrk Queries');
    monitorQueries = monitorQueries || val('Montr Queries');
    scope += '\nGSC pre-migration benchmark queries: '+(benchmarkQueries || '[unavailable]');
    scope += '\nGSC monitor queries: '+(monitorQueries || '[unavailable]');

    var locality = articleType === 'Geo Service Page' ?
      'Locality: '+val('Locality')+'; Parent Area: '+val('Parent Area') : '';
    // The legacy reader above validates the current v10 case-study package.
    // It reads the same posts evidence as GL:HA, so do not insert it twice.
    // "Case Study Evidence Package" sheet is an append-only BEFORE/AFTER
    // archive; current writing authority is the selected posts row.
    var sourceContext = articleType === 'Case Study' ?
      'ORIGINAL ARTICLE NARRATIVE CONTEXT (not authority for extra project facts):\n'+originalHtml+
      '\nCommercial confidentiality: generic products must remain generic; never infer product brand, supplier, specific chemistry or equipment.' :
      'ORIGINAL ARTICLE AS DISTINCTIVE SOURCE CONTEXT:\n'+originalHtml;
    var prompt = [
      'ROLE: Senior UK stone and tile restoration editor. Prepare a UNIQUE writing brief, not the finished article.',
      'GOVERNING PRINCIPLE: Stay true to the details and experience recorded in the original article, but never allow its existing structure, wording or limitations to restrict the quality of the new article.',
      'PRIMARY AIM: An outstanding, distinctively useful article that answers the intended reader better than a generic summary. Rankings are not guaranteed.',
      'ARTICLE: '+val('Title')+' | TYPE: '+articleType+' | MATERIAL: '+material,
      'OWNED SEARCH INTENT / CLUSTER: '+cluster,
      'APPROVED PROBLEM ANGLE: '+(angle||'[not yet set; flag for review]'),
      'PRIMARY SEARCH TERM: '+val('Primary Search Term'),
      'ARTICLE-TYPE RULES (authoritative):\n'+typeRules,
      'EXISTING CLASSIFICATION CONTEXT:\n'+selection,
      'ARTICLE-SPECIFIC TRUTH SOURCES — POSTS GL:GU AND GW:HA (GV package validated but omitted as duplicate):\n'+truthBlock,
      postId === '22021' && articleType === 'Case Study' ? 'OWNER-CONFIRMED CORRECTION (Penkhull only): The hallway floor had no damp-proof membrane. This is a direct owner confirmation, not an inference from the article. Preserve this in the brief, distinguishing it from source-quoted facts.' : '',
      'TRUTH RULE: Case Study Evidence Package sheet is a historical BEFORE/AFTER archive, not an additional independent truth source. Treat the selected posts row GL:HA as the current account of this article. Distinguish documented facts from blanks; never fill unknowns with assumptions. If original narrative or derived evidence conflicts with GL:HA, explicitly flag the conflict and follow GL:HA for the writing brief. Technical reference and SEO suggestions must not override these facts.',
      'SEARCH, GOVERNANCE AND SCOPE:\n'+scope,
      locality,
      'TECHNICAL MATRIX (general reference, never evidence of work performed):\n'+technical,
      'ORIGINAL MEDIA ASSET INVENTORY:\n'+JSON.stringify(media),
      sourceContext,
      'TASK: Analyse the original and produce a writing brief under the following headings:',
      '1. Distinctive value: what only THIS source article can credibly tell readers; quote short source anchors for traceability.',
      articleType === 'Case Study' ? '2. Documented project details: client objective, conditions, decisions, methods, results and limits; distinguish proven from unknown.' : '2. Article-specific source insights and reader concerns; do not invent a client or restoration project.',
      '3. Reader questions and intent: which concerns matter most and why; do not add unsupported search-volume or competitor claims.',
      '4. Editorial opportunity: stronger explanation, useful decisions, natural technical context and what to omit as generic filler.',
      angle ? '5. Suggested article structure: follow the approved Article Angle ('+angle+') as the structural direction, while selecting useful sections for THIS article. W1.5A–W1.5E approve the final structure; no mandatory section count.' : '5. Suggested article structure: no Article Angle is approved yet. Suggest a suitable article-type-specific direction and flag it for explicit approval; do not treat the suggestion as an approved angle. W1.5A–W1.5E approve the final structure.',
      '6. Evidence and claim boundaries: what must stay true; what must NOT be invented. Never infer tests, dates, work performed, customer reactions, results, products, brands, suppliers, chemistry or equipment.',
      '7. Brief-specific success test: what a reviewer should look for to distinguish an exceptional article from a generic AI article.',
      'Do not draft final HTML. Do not select internal links. Do not invent SERP findings, external sources or project facts.',
      'Write in practical homeowner-facing UK English. Technical Matrix and governed CX link plan are applied by their authoritative downstream stages.'
    ].join('\n\n');
    return {success:true,row:row,postId:postId,articleType:articleType,
      source:'site-export',truthSourceColumns:'GL:HA',truthFieldsPresent:truthFilled.length,originalCharacters:originalHtml.length,prompt:prompt,
      media:media,warnings:axes.filter(function(k){return !val(k);}).map(function(k){return 'Missing classification: '+k;})};
  } catch(e) {return {success:false,message:'V2 brief error: '+String(e)};}
}



/**
 * Generate the ACTUAL editorial brief. Called only after the user presses
 * the explicit Generate with AI button; does not write to Sheets.
 */
function new_generateArticleWritingBriefForRow(rowNumber) {
  var assembled = new_buildArticleWritingBrief(rowNumber);
  if (!assembled.success) return assembled;
  if (typeof bc_sendPromptViaOpenAI !== 'function')
    return {success:false,message:'Existing OpenAI request function is unavailable'};
  try {
    var chosenModel = typeof MODEL_CHEAP !== 'undefined' ? MODEL_CHEAP : undefined;
    if (!chosenModel) return {success:false,message:'MODEL_CHEAP is not configured'};
    // The preview preserves the full archive context. The API request only
    // needs key SEO queries and technical reference M: repeated N:S guidance
    // and hundreds of incidental GSC queries waste input context.
    var aiPrompt = assembled.prompt
      .replace(/GSC pre-migration benchmark queries:[\s\S]*?(?=\nGSC monitor queries:)/,
        'GSC pre-migration benchmark queries: Use the primary search term and owned intent stated above.\n')
      .replace(/GSC monitor queries:[\s\S]*?(?=\n\nTECHNICAL MATRIX)/,
        'GSC monitor queries: Omitted from this editorial brief; preserve search intent above.')
      .replace(/(\[M [^\n]*[\s\S]*?)(?=\n\[N [^\n]*)[\s\S]*?(?=\n\nORIGINAL MEDIA ASSET INVENTORY:)/,
        '$1');
    // Shape the output, not the project truth: retain full evidence in the
    // source input, but avoid repeating each fact in several brief sections.
    aiPrompt = [
      'OUTPUT CONTRACT — CONCISE EDITORIAL WRITING BRIEF ONLY:',
      'Target 900–1,300 words, maximum 1,500. No article prose.',
      'Use these headings exactly: (1) Distinctive project value,',
      '(2) Verified facts and owner confirmations, (3) Reader intent,',
      '(4) Editorial angle and structure, (5) Evidence-led media plan,',
      '(6) Claim boundaries and unknowns, (7) Final quality checklist.',
      'Write each project fact once in its most useful section; refer back',
      'rather than restating entire fact lists or the restoration sequence.',
      'Include the project-specific before, actions, results and limits;',
      'distinguish owner-confirmed facts from source-quoted facts.',
      'Give a practical 4–6-section H2 outline, each with one-sentence intent;',
      'make media placement concise, without reproducing every URL.',
      'State the essential prohibitions and important unknowns once, grouped.',
      'Do not omit a material constraint, source-supported result, or',
      'evidence-integrity safeguard merely to meet the word target.',
      'Project truth and technical restrictions always outrank brevity.',
      '',
      aiPrompt
    ].join('\\n');
    var api = bc_sendPromptViaOpenAI(aiPrompt, 4500, chosenModel);
    if (!api || !api.success || api.truncated || api.finishReason === 'length')
      return {success:false,message:'AI brief incomplete. Reason: '+(api&&api.message||api&&api.error||'not supplied')+
        '; finishReason='+String(api&&api.finishReason||'unknown')+
        '; truncated='+String(!!(api&&api.truncated))+
        '; receivedCharacters='+String(api&&api.text?String(api.text).length:0)+
        '. No data was changed.',cost:Number(api && api.cost || 0)};
    var brief = String(api.text || '').trim();
    if (!brief) return {success:false,message:'AI returned an empty brief. No data was changed.',cost:Number(api.cost || 0)};
    return {success:true,row:assembled.row,brief:brief,cost:Number(api.cost||0),
      message:'Generated with AI. Preview only; nothing saved to the spreadsheet.'};
  } catch(e) {
    return {success:false,message:'AI writing-brief generation failed: '+String(e)};
  }
}

/**
 * Manual entry point for the existing Posts menu.
 * Produces a read-only prompt preview; AI generation requires a separate click.
 */
function new_previewWritingBriefActiveRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var posts = ss.getSheetByName('posts');
  if (!posts) { SpreadsheetApp.getUi().alert('Posts sheet not found.'); return; }
  var row = posts.getActiveRange().getRow();
  var result = new_buildArticleWritingBrief(row);
  if (!result.success) {
    SpreadsheetApp.getUi().alert('New Writing Brief — row ' + row + '\n\n' + result.message);
    return;
  }
  var escaped = String(result.prompt || '').replace(/&/g, '&amp;')
    .replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  var html = '<!doctype html><html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;padding:12px}' +
    'textarea{width:100%;height:75vh;box-sizing:border-box;font:12px/1.45 Consolas,monospace}' +
    'button{margin:8px 0;padding:9px 16px}</style></head><body>' +
    '<strong>Row ' + result.row + ' — ' + String(result.articleType).replace(/</g, '&lt;') +
    '</strong><p>Read-only writing brief. Nothing has been changed.</p>' +
    '<button onclick="var e=document.getElementById(\'brief\');e.select();' +
    'document.execCommand(\'copy\');">Copy text</button>' +
    '<button id="generate" onclick="generateBrief()">Generate writing brief with AI (API cost)</button>' +
    '<span id="status" style="margin-left:10px"></span>' +
    '<textarea id="brief" readonly>' + escaped + '</textarea>' +
    '<script>function generateBrief(){' +
    'var b=document.getElementById("generate"), t=document.getElementById("status");' +
    'b.disabled=true;t.textContent="Generating AI writing brief...";' +
    'google.script.run.withSuccessHandler(function(r){' +
    'b.disabled=false;if(!r||!r.success){t.textContent=(r&&r.message)||"Generation failed";return;}' +
    'document.getElementById("brief").value=r.brief;' +
    't.textContent="Actual AI brief (not saved). API cost: $"+Number(r.cost||0).toFixed(4);' +
    '}).withFailureHandler(function(e){b.disabled=false;t.textContent=String(e);})' +
    '.new_generateArticleWritingBriefForRow(' + result.row + ');' +
    '}</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(920).setHeight(650),
    'NEW article writing brief — row ' + row
  );
}
