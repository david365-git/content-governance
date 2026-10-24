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
    var acProblems = axes.filter(function(k){
      return val(k) && (acLibrary.values[k] || []).indexOf(val(k)) === -1;
    });
    if (acProblems.length)
      return {success:false,message:'Invalid AC choices: '+acProblems.join(', ')};
    selection = axes.map(function(k){
      return String((acLibrary.labels || {})[k] || k)+': '+(val(k)||'[not classified]');
    }).join('\n');
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
      'Observed Query Cluster','Entity Role','Feeds Hub','Authority Brief','Article Angle',
      'GSC Benchmark Queries','GSC Monitor Queries','Montr Queries'];
    var scope = scopeKeys.map(function(k){return k+': '+(val(k)||'[not supplied]');}).join('\n');
    var locality = articleType === 'Geo Service Page' ?
      'Locality: '+val('Locality')+'; Parent Area: '+val('Parent Area') : '';
    var sourceContext = articleType === 'Case Study' ?
      'AUTHORITATIVE CASE STUDY EVIDENCE:\n'+evidence+
      '\nORIGINAL ARTICLE NARRATIVE CONTEXT (not authority for extra project facts):\n'+originalHtml+
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
      '5. Suggested article structure: each proposed section must have a clear job; no mandatory section count or template.',
      '6. Evidence and claim boundaries: what must stay true; what must NOT be invented. Never infer tests, dates, work performed, customer reactions, results, products, brands, suppliers, chemistry or equipment.',
      '7. Brief-specific success test: what a reviewer should look for to distinguish an exceptional article from a generic AI article.',
      'Do not draft final HTML. Do not select internal links. Do not invent SERP findings, external sources or project facts.',
      'Write in practical homeowner-facing UK English. Technical Matrix and governed CX link plan are applied by their authoritative downstream stages.'
    ].join('\n\n');
    return {success:true,row:row,postId:postId,articleType:articleType,
      source:'site-export',originalCharacters:originalHtml.length,prompt:prompt,
      media:media,warnings:axes.filter(function(k){return !val(k);}).map(function(k){return 'Missing classification: '+k;})};
  } catch(e) {return {success:false,message:'V2 brief error: '+String(e)};}
}
