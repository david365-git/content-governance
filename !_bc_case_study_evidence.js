
/**
 * Enrich Evidence Media & Links from exact site-export HTML (no AI/API).
 * Retains existing links; recovers missing image/video tags; records the
 * source section and surrounding original text as placement context, not
 * independently verified visual identification.
 */
function bc_caseEvidenceEnrichMedia_(row, source, headers) {
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  var col=headers.indexOf('Evidence Media & Links')+1;
  if(!col)throw new Error('Evidence Media & Links column missing');
  var cell=sh.getRange(row,col),previous=String(cell.getValue()||'');
  var record;
  try{record=previous?JSON.parse(previous):{};}catch(e){throw new Error('Existing media JSON invalid: '+e.message);}
  if(!record||Array.isArray(record)||typeof record!=='object')throw new Error('Invalid media inventory');
  if(!Array.isArray(record.imagesAndVideos))record.imagesAndVideos=[];
  if(!Array.isArray(record.links))record.links=[];
  function decode(v){return String(v||'').replace(/&amp;/gi,'&').replace(/&quot;/gi,'"').replace(/&#39;/gi,"'").trim();}
  function attr(tag,name){
    var m=new RegExp('(?:^|\\s)'+name+'\\s*=\\s*(["\\x27])([\\s\\S]*?)\\1','i').exec(tag);
    return m?decode(m[2]):'';
  }
  function plain(s){return String(s||'').replace(/<[^>]*>/g,' ').replace(/&nbsp;|&#160;/gi,' ').replace(/&amp;/gi,'&').replace(/\s+/g,' ').trim();}
  function sectionAt(index){
    var h=/<h[1-6]\b[^>]*>([\s\S]*?)<\/h[1-6]>/gi,m,last='';
    while((m=h.exec(source))&&m.index<index)last=plain(m[1]);
    return last.slice(0,220);
  }
  function nearby(index,tagLen){
    var before=source.slice(Math.max(0,index-450),index),after=source.slice(index+tagLen,index+tagLen+450);
    var fm=/<figcaption\b[^>]*>([\s\S]*?)<\/figcaption>/i.exec(after);
    if(fm)return {text:plain(fm[1]).slice(0,350),basis:'source figcaption'};
    var p=/<p\b[^>]*>([\s\S]*?)<\/p>/gi,m,last='';
    while((m=p.exec(before)))last=plain(m[1]);
    if(last)return {text:last.slice(-260),basis:'preceding paragraph'};
    var next=/<p\b[^>]*>([\s\S]*?)<\/p>/i.exec(after);
    return {text:next?plain(next[1]).slice(0,260):'',basis:next?'following paragraph':'no adjacent text'};
  }
  function stage(text){
    var t=String(text||'').toLowerCase();
    var phases=[['before',/\bbefore\b|\bpre[- ]restoration\b|\buntreated\b/],
      ['inspection',/\binspect(?:ion|ed)?\b|\btest patch\b|\bsurvey\b/],
      ['cleaning',/\bscrubb?(?:ing|ed)?\b|\bcleaning\b|\brins(?:e|ing)\b|\bextract(?:ion|ing)\b/],
      ['repair',/\brepair(?:ing|ed)?\b|\bfilling\b|\bfiller\b|\bregrout\b/],
      ['surface refinement',/\bgrind(?:ing)?\b|\bhon(?:e|ing)\b|\bpolish(?:ing|ed)?\b/],
      ['sealing',/\bseal(?:ing|er|ed)?\b|\bwax(?:ing|ed)?\b/],
      ['after',/\bafter\b|\bfinished\b|\bcompleted\b|\brestored floor\b/]];
    var found=phases.filter(function(x){return x[1].test(t);});
    return found.length===1?found[0][0]:'uncertain';
  }
  var tags=/<img\b[^>]*>|<iframe\b[^>]*>|<video\b[^>]*>|<source\b[^>]*>/gi,m,seen={},items=[];
  while((m=tags.exec(source))){
    var tag=m[0],kind=/^<img\b/i.test(tag)?'image':/^<iframe\b/i.test(tag)?'video':/^<video\b/i.test(tag)?'video':'video-source';
    var url=attr(tag,'src')||attr(tag,'data-src');
    if(!url||seen[kind+'|'+url])continue;
    seen[kind+'|'+url]=true;
    var original=record.imagesAndVideos.filter(function(x){return decode(x.url)===url;})[0]||{};
    var context=nearby(m.index,tag.length),section=sectionAt(m.index);
    var caption=String(original.originalCaption||'');
    var alt=attr(tag,'alt')||String(original.alt||'');
    var label=caption||alt;
    var phase=stage(label);
    // Source prose may suggest a placement section, but it never verifies pixels.
    var item={};
    Object.keys(original).forEach(function(k){item[k]=original[k];});
    item.order=items.length+1;
    item.type=kind==='video-source'?'video':kind;
    item.url=url;
    item.alt=alt;
    item.title=attr(tag,'title')||String(original.title||'');
    item.originalCaption=caption;
    item.provenance='Original HTML and metadata; image contents not independently verified';
    item.sourcePosition=m.index;
    item.sourceSection=section;
    item.nearbySourceText=context.text;
    item.nearbyTextBasis=context.basis;
    item.suggestedStage=phase;
    item.stageBasis=phase==='uncertain'?'No unambiguous stage in existing description':'Original caption or alt text only; not image-verified';
    item.placementConfidence=phase==='uncertain'?'unconfirmed':'source-description-only';
    items.push(item);
  }
  // Preserve any previously indexed media not found by the tag extractor,
  // such as poster images or embedded services not represented by src tags.
  record.imagesAndVideos.forEach(function(x){
    if(!x||!x.url)return;
    if(!items.some(function(y){return y.url===x.url&&y.type===x.type})){
      var retained={};Object.keys(x).forEach(function(k){retained[k]=x[k];});
      retained.placementConfidence='unconfirmed';
      retained.sourcePosition=null;
      retained.stageBasis='Previously indexed; not matched to source media tag';
      items.push(retained);
    }
  });
  record.imagesAndVideos=items;
  record.mediaAudit={
    sourceImageTags:(source.match(/<img\b/gi)||[]).length,
    capturedImageUrls:items.filter(function(x){return x.type==='image'&&x.sourcePosition!==null&&x.sourcePosition!==undefined;}).length,
    unconfirmedPlacement:items.filter(function(x){return x.placementConfidence==='unconfirmed';}).length,
    missingSourceImageUrls:0,
    descriptionBasis:'Source HTML captions/alt and surrounding context only; not actual image inspection'
  };
  var json=JSON.stringify(record);
  if(json.length>49000)throw new Error('Media inventory too long for Sheets cell; no media update made');
  if(json!==previous)cell.setValue(json);
  return record.mediaAudit;
}
function bc_caseEvidenceMediaActiveRow(){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts'),row=sh.getActiveRange().getRow();
  if(row<2)throw new Error('Select a case-study row');
  var heads=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  function c(n){var i=heads.indexOf(n);if(i<0)throw new Error('Missing column '+n);return i+1;}
  if(String(sh.getRange(row,c('Article Type')).getValue()).trim()!=='Case Study')throw new Error('Selected row is not a Case Study');
  var postId=String(sh.getRange(row,c('Post ID')).getDisplayValue()).trim();
  var ex=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('site-export');
  var ids=ex.getRange(2,1,ex.getLastRow()-1,1).getDisplayValues(),match=[];
  ids.forEach(function(v,i){if(String(v[0]).trim()===postId)match.push(i+2);});
  if(match.length!==1)throw new Error('Expected exactly one site-export source for '+postId);
  var src=String(ex.getRange(match[0],15).getValue()||'');
  if(!src)throw new Error('Original HTML missing');
  return bc_caseEvidenceEnrichMedia_(row,src,heads);
}

/**
 * Content Intelligence: source-quoted case-study evidence, GL:GV.
 * Evidence quotes must match original HTML in site-export Column O, by Post ID.
 */
function bc_caseEvidenceDigest_(source) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source, Utilities.Charset.UTF_8);
  return bytes.map(function(b){ return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function bc_caseEvidenceText_(html) {
  return String(html || '').replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ').replace(/&nbsp;|&#160;/gi, ' ')
    .replace(/&amp;/gi, '&').replace(/&quot;/gi, '"')
    .replace(/&apos;|&#39;/gi, "'").replace(/\s+/g, ' ').trim();
}
/**
 * Recover a quotation only when the source contains the same contiguous words.
 * Accept harmless case differences and an extra leading "the/a/an"; never
 * fabricate missing words or use a semantic/fuzzy match as source evidence.
 */
function bc_caseEvidenceRepairQuote_(quote, sourceText) {
  var clean=function(v){return String(v||'').replace(/[’‘]/g,"'").replace(/[“”]/g,'"').replace(/\s+/g,' ').trim();};
  var src=clean(sourceText), q=clean(quote);
  var candidates=[q];
  var withoutArticle=q.replace(/^(?:the|a|an)\s+/i,'');
  if(withoutArticle!==q && withoutArticle.split(/\s+/).length>=3)candidates.push(withoutArticle);
  for(var i=0;i<candidates.length;i++){
    var part=candidates[i];
    if(part.length<12)continue;
    var at=src.toLowerCase().indexOf(part.toLowerCase());
    if(at>=0)return src.slice(at,at+part.length);
  }
  return null;
}
function bc_caseEvidenceExtract_(row) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if (!sh) throw new Error('posts sheet missing');
  row = Number(row);
  if (!row || row < 2) throw new Error('Valid posts row required');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  function col(name) { var i = headers.indexOf(name); if (i < 0) throw new Error('Missing heading: ' + name); return i + 1; }
  var type = String(sh.getRange(row, col('Article Type')).getValue() || '').trim();
  if (type !== 'Case Study') return { success:true, skipped:true, message:'Not a Case Study', cost:0 };
  var postId = String(sh.getRange(row, col('Post ID')).getDisplayValue() || '').trim();
  if (!postId) throw new Error('Post ID missing in posts row ' + row);
  var exportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('site-export');
  if (!exportSheet) throw new Error('site-export sheet missing');
  var exportLastRow = exportSheet.getLastRow();
  var ids = exportLastRow > 1 ? exportSheet.getRange(2, 1, exportLastRow - 1, 1).getDisplayValues() : [];
  var matchedRow = -1;
  for (var k = 0; k < ids.length; k++) {
    if (String(ids[k][0] || '').trim() === postId) {
      if (matchedRow !== -1) throw new Error('Duplicate Post ID ' + postId + ' in site-export');
      matchedRow = k + 2;
    }
  }
  if (matchedRow === -1) throw new Error('Post ID ' + postId + ' not found in site-export');
  var source = String(exportSheet.getRange(matchedRow, 15).getValue() || ''); // site-export column O: Full Post HTML
  if (!source.trim()) {
    sh.getRange(row, col('Evidence Status')).setValue('SOURCE REQUIRED — site-export O is empty');
    return {success:false, cost:0, message:'SOURCE REQUIRED — site-export Full Post HTML missing'};
  }
  var digest = bc_caseEvidenceDigest_(source);
  var statusCell = sh.getRange(row, col('Evidence Status'));
  var oldStatus = String(statusCell.getValue() || '');
  if (oldStatus === 'SOURCE-QUOTED | sha256:' + digest) {
    bc_caseEvidenceEnrichMedia_(row,source,headers);
    return {success:true, skipped:true, cost:0, message:'Source-quoted facts unchanged; media inventory checked'};
  }
  var labels = ['Project Location & Property','Observed Conditions','Client Requirements',
    'Materials & Construction','Inspection & Tests','Restoration Methods',
    'Products & Equipment','Technical Findings','Project Results','Maintenance Handover'];
  var prompt = [
    'Extract ONLY documented facts from this original case-study article.',
    'Return ONLY a valid JSON object with exactly these keys: ' + JSON.stringify(labels) + '.',
    'Each key maps to an array of objects {"fact":"concise source-supported fact","quote":"exact contiguous quotation from the visible original article"}',
    'Every quote MUST be copied exactly from the source visible text. Never infer a performed test, cause, property detail, client conversation, or project outcome.',
    'If a category has no documented facts, return an empty array. Inspection & Tests must not infer test results from general professional practice.',
    'Distinguish original construction materials from repair products, and maintain documented chronological order for methods.',
    'Treat statements of general advice as advice, not work performed.',
    'ORIGINAL ARTICLE:\n' + bc_caseEvidenceText_(source)
  ].join('\n');
  var normalizedSource = bc_caseEvidenceText_(source).replace(/[’‘]/g,"'").replace(/[“”]/g,'"');
  var output = null;
  var totalCost = 0;
  var lastError = '';
  for (var attempt = 0; attempt < 2; attempt++) {
    var retryPrompt = attempt === 0 ? prompt : prompt + '\nVALIDATION CORRECTION: Your previous response contained this quotation which is not an exact substring of ORIGINAL ARTICLE: ' + lastError + '\nRegenerate the complete JSON. Copy quotations character-for-character from ORIGINAL ARTICLE. Do not paraphrase, omit words, or change punctuation inside quotes.';
    var result = bc_sendPromptViaOpenAI(retryPrompt, 5000, MODEL_CHEAP);
    if (!result || !result.success) throw new Error('Evidence extraction API failed: ' + (result && result.message || 'unknown error'));
    totalCost += Number(result.cost || 0);
    var raw = String(result.text || '').trim().replace(/^\x60\x60\x60(?:json)?\s*/i,'').replace(/\s*\x60\x60\x60$/,'');
    var parsed;
    try { parsed = JSON.parse(raw); } catch(e) {
      lastError = 'Invalid JSON: ' + e.message;
      continue;
    }
    try {
      var proposed = [];
      labels.forEach(function(label) {
        var items = parsed[label];
        if (!Array.isArray(items)) throw new Error('Missing evidence array: ' + label);
        var checked = [];
        items.forEach(function(item) {
          var fact = String(item && item.fact || '').trim();
          var quote = String(item && item.quote || '').replace(/\s+/g,' ').trim();
          var normalizedQuote = quote.replace(/[’‘]/g,"'").replace(/[“”]/g,'"');
          if (!fact || !quote) {
            throw new Error('Incomplete evidence in ' + label);
          }
          // Preserve exact matches; only repair case or an extra leading article
          // by copying the real contiguous phrase back from the source.
          if (normalizedSource.indexOf(normalizedQuote) < 0) {
            var repaired=bc_caseEvidenceRepairQuote_(quote,normalizedSource);
            if (!repaired)throw new Error('Unverifiable quote in ' + label + ': ' + quote.slice(0,120));
            quote=repaired;
          }
          checked.push({fact:fact,quote:quote});
        });
        proposed.push(JSON.stringify(checked));
      });
      output = proposed;
      break;
    } catch(e) { lastError = e.message; }
  }
  if (!output) throw new Error('Evidence extraction failed after 2 attempts: ' + lastError);
  // Fail closed: no partial evidence is written on model error or quote mismatch.
  for (var j=0;j<labels.length;j++) col(labels[j]);
  sh.getRange(row, col(labels[0]), 1, labels.length).setValues([output]);
  bc_caseEvidenceEnrichMedia_(row,source,headers);
  statusCell.setValue('SOURCE-QUOTED | sha256:' + digest);
  return {success:true, skipped:false, cost:totalCost, message:'Evidence quotations matched to source (not independently fact-checked)'};
}
function bc_caseEvidenceActiveRow() {
  return bc_caseEvidenceExtract_(SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getRow());
}

/**
 * Shared read-only evidence block for W1.5 and W2B.
 * Do not trust a prior extraction unless its original HTML fingerprint matches.
 */
function bc_caseEvidencePromptBlock_(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('posts');
  row = Number(row);
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  function c(name) { var idx=headers.indexOf(name); if(idx<0) throw new Error('Evidence column missing: '+name); return idx+1; }
  if (String(sh.getRange(row,c('Article Type')).getValue()).trim() !== 'Case Study') return '';
  var id=String(sh.getRange(row,c('Post ID')).getDisplayValue()).trim();
  var ex=ss.getSheetByName('site-export');
  var ids=ex.getRange(2,1,Math.max(ex.getLastRow()-1,1),1).getDisplayValues();
  var match=-1;
  ids.forEach(function(v,i){if(String(v[0]).trim()===id){if(match>=0)throw new Error('Duplicate site-export Post ID '+id);match=i+2;}});
  if(match<0)throw new Error('No site-export source for Post ID '+id);
  var original=String(ex.getRange(match,15).getValue()||'');
  if(!original.trim())throw new Error('Original case study HTML is empty in site-export O');
  var status=String(sh.getRange(row,c('Evidence Status')).getValue()||'');
  if(status!=='SOURCE-QUOTED | sha256:'+bc_caseEvidenceDigest_(original)){
    throw new Error('Case Study Evidence is missing or out of date. Run Extract Case Study Evidence in Content Intelligence first.');
  }
  var labels=['Project Location & Property','Observed Conditions','Client Requirements','Materials & Construction',
    'Inspection & Tests','Restoration Methods','Products & Equipment','Technical Findings','Project Results','Maintenance Handover'];
  var lines=[];
  labels.forEach(function(label){
    var raw=String(sh.getRange(row,c(label)).getValue()||'[]');
    var items=JSON.parse(raw);
    if(!Array.isArray(items))throw new Error('Invalid case-study evidence in '+label);
    var entries=items.map(function(x){
      if(!x||!x.fact||!x.quote)throw new Error('Incomplete case-study evidence in '+label);
      return '- '+String(x.fact)+' [SOURCE: '+String(x.quote)+']';
    });
    lines.push(label+':\n'+(entries.length?entries.join('\n'):'(Not documented)'));
  });
  return '\nCASE STUDY PROJECT EVIDENCE (source-quoted, not independently fact-checked):\n'
    +'All project-specific claims must stay within the evidence below and confirmed owner corrections. Do not turn general technical notes into observations or performed tests. An empty category does not prove an activity never happened. Governed headings and SEO cannot add project facts.\n'
    +lines.join('\n')+'\nEND CASE STUDY EVIDENCE\n';
}


/**
 * Extract evidence for visible Case Study rows in the currently filtered posts view.
 * Safe to rerun: bc_caseEvidenceExtract_ skips source-matched completed rows.
 * A time budget lets the operator rerun in batches without Apps Script timeouts.
 */
function bc_caseEvidenceFilteredBatch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('posts');
  if (!sh) throw new Error('posts sheet missing');
  var last = sh.getLastRow();
  if (last < 2) return SpreadsheetApp.getUi().alert('No posts to process.');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var typeCol = headers.indexOf('Article Type') + 1;
  var stoneCol = headers.indexOf('Stone Type') + 1;
  if (!typeCol || !stoneCol) throw new Error('Article Type or Stone Type heading missing');
  var vals = sh.getRange(2,1,last-1,Math.max(typeCol,stoneCol)).getDisplayValues();
  var visible=[], hidden=0, types={};
  for (var i=0;i<vals.length;i++) {
    var row=i+2;
    if (sh.isRowHiddenByUser(row) || sh.isRowHiddenByFilter(row)) { hidden++; continue; }
    var stone=String(vals[i][stoneCol-1]||'').trim();
    if (stone) types[stone]=true;
    if (String(vals[i][typeCol-1]||'').trim()==='Case Study') visible.push(row);
  }
  // The custom Stone Type sidebar hides unmatched rows rather than using a Sheets filter.
  // Refuse an unfiltered or mixed-stone selection to avoid accidentally processing all materials.
  if (!hidden || Object.keys(types).length !== 1) {
    SpreadsheetApp.getUi().alert('First filter the posts sheet to ONE Stone Type using the Stone Type filter. No rows were processed.');
    return;
  }
  var started=Date.now(), done=0, skipped=0, failed=[], pending=0, cost=0;
  for (var j=0;j<visible.length;j++) {
    if (Date.now()-started>240000) { pending=visible.length-j; break; }
    try {
      var result=bc_caseEvidenceExtract_(visible[j]);
      cost+=Number(result.cost||0);
      if (result.skipped) skipped++;
      else if (result.success) done++;
      else failed.push('Row '+visible[j]+': '+(result.message||'Unsuccessful'));
    } catch(err) { failed.push('Row '+visible[j]+': '+err.message); }
  }
  SpreadsheetApp.getUi().alert(
    'Case Study Evidence — '+Object.keys(types)[0]+'\n'
    +'Newly completed: '+done+'\nAlready completed: '+skipped
    +'\nFailures: '+failed.length+'\nRemaining for next run: '+pending
    +'\nAPI cost: $'+cost.toFixed(4)
    +(failed.length?'\n'+failed.slice(0,8).join('\n'):'')
    +'\nRun the same menu item again to continue; completed rows are skipped.'
  );
}


/** Refresh only the eleven Google Love test case-study media inventories.
 * No AI calls. Does not alter the first ten evidence fields or published posts.
 */
function bc_caseEvidenceMediaTest11(){
 var ids=['48109','7952','5774','21838','21746','8027','8336','12107','11452','22942','11945'];
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('posts'),ex=ss.getSheetByName('site-export');
 var heads=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
 function c(n){var i=heads.indexOf(n);if(i<0)throw Error('Missing header '+n);return i+1;}
 var rows=sh.getRange(2,c('Post ID'),sh.getLastRow()-1,1).getDisplayValues(),map={};
 rows.forEach(function(v,i){var id=String(v[0]).trim();if(ids.indexOf(id)>=0){if(map[id])throw Error('Duplicate posts ID '+id);map[id]=i+2;}});
 var erows=ex.getRange(2,1,ex.getLastRow()-1,1).getDisplayValues(),sources={};
 erows.forEach(function(v,i){var id=String(v[0]).trim();if(ids.indexOf(id)>=0){if(sources[id])throw Error('Duplicate export ID '+id);sources[id]=i+2;}});
 var results=[];
 ids.forEach(function(id){
  if(!map[id]||!sources[id])throw Error('Missing source or posts row for '+id);
  if(String(sh.getRange(map[id],c('Article Type')).getValue()).trim()!=='Case Study')throw Error('Wrong article type '+id);
  var html=String(ex.getRange(sources[id],15).getValue()||'');
  if(!html)throw Error('Missing original HTML '+id);
  var m=bc_caseEvidenceEnrichMedia_(map[id],html,heads);
  results.push({postId:id,sourceImages:m.sourceImageTags,recordedImages:m.capturedImageUrls,uncertainPlacement:m.unconfirmedPlacement});
 });
 console.log(JSON.stringify(results));
 return results;
}
