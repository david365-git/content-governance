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

/**
 * Additive evidence package: the ten existing columns keep {fact,quote}.
 * Project claims remain source-bound; this package is not proof beyond the article.
 */
function bc_caseEvidenceBuildPackage_(postId, source, digest, labels, output) {
  var plain=bc_caseEvidenceText_(source);
  var normalized=plain.replace(/[’‘]/g,"'").replace(/[“”]/g,'"').toLowerCase();
  var facts=[], order=0;
  labels.forEach(function(label, categoryIndex) {
    JSON.parse(output[categoryIndex]).forEach(function(item) {
      var quote=String(item.quote||'');
      var q=quote.replace(/[’‘]/g,"'").replace(/[“”]/g,'"').toLowerCase();
      var position=normalized.indexOf(q);
      if(position<0) throw new Error('Cannot locate evidence quote for package: '+label);
      var fact=String(item.fact||'');
      var status='documented_statement';
      if(/^CLIENT REQUEST:/i.test(fact)) status='client_requested';
      else if(/^RESTORATION OBJECTIVE:/i.test(fact)) status='restoration_objective';
      else if(label==='Maintenance Handover') status='recommended';
      else if(/\b(?:not found|no major|no evidence|absent|ruled out)\b/i.test(fact)) status='negative_finding';
      else if(/\b(?:considered|possible|might|could|option)\b/i.test(fact)) status='considered_or_possible';
      else if(/\b(?:where needed|if needed|conditional)\b/i.test(fact)) status='conditional';
      else if(label==='Restoration Methods'||label==='Products & Equipment') status='documented_action';
      facts.push({id:'F'+('000'+(++order)).slice(-4),category:label,fact:fact,quote:quote,
        sourceOrder:position,status:status,scope:'As stated in quotation',
        certainty:'Documented in source; not independently verified',relatedFactIds:[]});
    });
  });
  facts.sort(function(a,b){return a.sourceOrder-b.sourceOrder;});
  // IDs are assigned in source order so references remain understandable.
  facts.forEach(function(f,i){f.id='F'+('000'+(i+1)).slice(-4);});
  var media=[], links=[];
  var figureRe=/<figure\b[^>]*>[\s\S]*?<\/figure>/gi, match, index=0;
  while((match=figureRe.exec(source))!==null){
    var block=match[0],img=block.match(/<img\b[^>]*>/i),iframe=block.match(/<iframe\b[^>]*>/i);
    var tag=img?img[0]:(iframe?iframe[0]:'');
    var src=tag.match(/\bsrc\s*=\s*["']([^"']+)["']/i);
    var alt=tag.match(/\balt\s*=\s*["']([^"']*)["']/i);
    var title=tag.match(/\btitle\s*=\s*["']([^"']*)["']/i);
    var caption=block.match(/<figcaption\b[^>]*>([\s\S]*?)<\/figcaption>/i);
    media.push({order:++index,type:iframe?'video':(img?'image':'figure'),url:src?src[1]:'',
      alt:alt?alt[1]:'',title:title?title[1]:'',
      originalCaption:caption?bc_caseEvidenceText_(caption[1]):'',
      provenance:'Original article caption/alt claim; image not independently verified'});
  }
  var linkRe=/<a\b([^>]*)>([\s\S]*?)<\/a>/gi, lm;
  while((lm=linkRe.exec(source))!==null){
    var href=lm[1].match(/\bhref\s*=\s*["']([^"']+)["']/i);
    if(href)links.push({order:links.length+1,url:href[1],anchor:bc_caseEvidenceText_(lm[2]),
      sourcePosition:lm.index,role:'Original link, not automatically approved for reuse'});
  }
  var notStated=['Area dimensions','Job duration','Number of visits','Exact product brands',
    'Measured test results','Direct client quotation','Price'];
  var limits=[];
  notStated.forEach(function(label) {
    // Explicitly mark these as unverified, not necessarily absent.
    limits.push({item:label,status:'Not established by this package; do not invent'});
  });
  return {version:'v9',postId:String(postId),sourceDigest:digest,
    provenance:'Source-quoted project account; not independently verified',
    facts:facts,media:media,links:links,
    chronology:facts.filter(function(f){return f.category==='Restoration Methods';})
      .map(function(f){return {factId:f.id,sourceOrder:f.sourceOrder,status:f.status};}),
    relationships:[], // Never infer problem/action/outcome causal links from proximity.
    unknowns:limits,
    coverage:{status:'Review required: source sentences are not yet exhaustively reconciled',
      extractedFactCount:facts.length,sourceQuotePositionsLocated:true}};
}
function bc_caseEvidenceExtraColumns_(packageData, source, labels, prior) {
  var paragraphs=String(source||'').replace(/<\/(?:p|h[1-6]|li|blockquote)>/gi,'\n')
    .replace(/<[^>]+>/g,' ').split(/\n+/).map(function(p){return bc_caseEvidenceText_(p);})
    .filter(function(p){return p.length>=50;});
  var checked=paragraphs.map(function(p,i){
    var pnorm=p.toLowerCase().replace(/\s+/g,' ');
    var matching=packageData.facts.filter(function(f){
      return pnorm.indexOf(String(f.quote||'').toLowerCase().replace(/\s+/g,' '))>=0;
    }).map(function(f){return f.id;});
    return {paragraph:i+1,preview:p.slice(0,180),factIds:matching,
      status:matching.length?'Contains extracted quotation':'Needs manual classification'};
  });
  var uncovered=checked.filter(function(p){return !p.factIds.length;});
  // Separate review of *unmatched* passages and source-supported relationships.
  // Missing or invalid review is recorded as unresolved, never silently passed.
  var relationships={status:'REVIEW REQUIRED',links:[]}, decisions={}, reviewCost=0;
  var factBrief=packageData.facts.map(function(f){
    return {id:f.id,category:f.category,fact:f.fact,quote:f.quote};
  });
  var auditPrompt=[
    'Audit a source-quoted floor-restoration case study evidence package.',
    'For EACH unmatched passage below classify as one of:',
    'PROJECT_FACT_MISSING, GENERAL_ADVICE, SEO_OR_HEADING, DUPLICATE, MEDIA_CAPTION, UNCERTAIN.',
    'Only GENERAL_ADVICE, SEO_OR_HEADING, DUPLICATE, MEDIA_CAPTION may be excluded.',
    'If the passage contains substantive job details absent from FACTS, choose PROJECT_FACT_MISSING.',
    'If uncertain choose UNCERTAIN. Do not invent source details.',
    'Identify only causal/decision relationships directly expressed in quoted original article wording.',
    'Relationship fromId and toId MUST be existing fact IDs; provide exact short contiguous sourceQuote which explicitly supports their connection.',
    'Do not connect facts merely because they are close or their order suggests cause.',
    'Return JSON only: {"classifications":[{"paragraph":number,"classification":"ENUM","reason":"brief"}],',
    '"relationships":[{"fromId":"F0001","toId":"F0002","type":"reason_for|led_to|limited|resulted_in","sourceQuote":"exact original quote"}]}.',
    'EXTRACTED FACTS:'+JSON.stringify(factBrief),
    'UNMATCHED PASSAGES:'+JSON.stringify(uncovered.map(function(p){return {paragraph:p.paragraph,text:paragraphs[p.paragraph-1]};})),
    'ORIGINAL ARTICLE TEXT:'+bc_caseEvidenceText_(source)
  ].join('\n');
  var audit=bc_sendPromptViaOpenAI(auditPrompt,5000,MODEL_CHEAP);
  if(!audit||!audit.success)throw new Error('Evidence coverage/relationship audit failed: '+(audit&&audit.message||'unknown'));
  reviewCost+=Number(audit.cost||0);
  var auditData;
  try {
    var raw=String(audit.text||'').trim().replace(/^\x60\x60\x60(?:json)?\s*/i,'').replace(/\s*\x60\x60\x60$/,'');
    auditData=JSON.parse(raw);
    if(!Array.isArray(auditData.classifications)||!Array.isArray(auditData.relationships))
      throw new Error('Missing classifications or relationships arrays');
  } catch(e){throw new Error('Invalid evidence coverage/relationship audit JSON: '+e.message);}
  var allowed=['PROJECT_FACT_MISSING','GENERAL_ADVICE','SEO_OR_HEADING','DUPLICATE','MEDIA_CAPTION','UNCERTAIN'];
  (auditData.classifications||[]).forEach(function(c){
    var n=Number(c.paragraph),classification=String(c.classification||'');
    if(!uncovered.some(function(p){return p.paragraph===n;})||allowed.indexOf(classification)<0)
      throw new Error('Invalid evidence paragraph classification '+n+' '+classification);
    if(decisions[n])throw new Error('Duplicate coverage classification '+n);
    decisions[n]={classification:classification,reason:String(c.reason||'').slice(0,300)};
  });

  // Repair missing project evidence, but never convert general explanation into
  // a performed project activity or site finding. Any uncertain material remains flagged.
  var missing=uncovered.filter(function(p){
    return decisions[p.paragraph]&&decisions[p.paragraph].classification==='PROJECT_FACT_MISSING';
  });
  var corrections=[], repairReview='No missing project facts proposed';
  if(missing.length){
    var repairPrompt=[
      'Check the following purported missing facts against the exact original article.',
      'Distinguish documented project observations/actions/results from generic restoration theory.',
      'Return JSON only: {"additions":[{"paragraph":number,"category":"exact category name","fact":"plain UK English factual statement","quote":"exact contiguous source quotation"}],"notProjectFacts":[{"paragraph":number,"reason":"why this is generic theory, duplicated or unsupported"}]}.',
      'Use only these ten categories: '+JSON.stringify(labels),
      'Only add DISTINCT project-specific evidence not already in EXISTING FACTS. Do not upgrade guesses, conditional work, client speech, or future options.',
      'For generic explanation, put paragraph in notProjectFacts. If unsure, leave the paragraph unresolved.',
      'CANDIDATES:'+JSON.stringify(missing.map(function(p){return {paragraph:p.paragraph,text:paragraphs[p.paragraph-1]};})),
      'EXISTING FACTS:'+JSON.stringify(packageData.facts.map(function(f){return {category:f.category,fact:f.fact,quote:f.quote};})),
      'ORIGINAL ARTICLE:'+bc_caseEvidenceText_(source)
    ].join('\n');
    var repairApi=bc_sendPromptViaOpenAI(repairPrompt,2500,MODEL_CHEAP);
    if(!repairApi||!repairApi.success)throw new Error('Evidence gap repair failed: '+(repairApi&&repairApi.message||'unknown'));
    reviewCost+=Number(repairApi.cost||0);
    var repaired;
    try {
      repaired=JSON.parse(String(repairApi.text||'').trim().replace(/^\x60\x60\x60(?:json)?\s*/i,'').replace(/\s*\x60\x60\x60$/,''));
      if(!Array.isArray(repaired.additions)||!Array.isArray(repaired.notProjectFacts))throw new Error('Missing repair arrays');
    }catch(e){throw new Error('Invalid missing-fact repair JSON: '+e.message);}
    var missingIds={};missing.forEach(function(p){missingIds[p.paragraph]=true;});
    repaired.notProjectFacts.forEach(function(item){
      if(!missingIds[Number(item.paragraph)])throw new Error('Invalid repair exclusion reference');
      decisions[Number(item.paragraph)]={classification:'GENERAL_ADVICE',reason:'Repair review: '+String(item.reason||'').slice(0,250)};
    });
    repaired.additions.forEach(function(item){
      var number=Number(item.paragraph), category=String(item.category||'');
      if(!missingIds[number]||labels.indexOf(category)<0)throw new Error('Invalid missing-fact addition category/paragraph');
      var fact=String(item.fact||'').trim();
      var quote=bc_caseEvidenceRepairQuote_(item.quote,bc_caseEvidenceText_(source));
      if(!fact||!quote)throw new Error('Unverifiable missing-fact addition');
      if(packageData.facts.some(function(f){return f.fact.toLowerCase()===fact.toLowerCase()||f.quote.toLowerCase()===quote.toLowerCase();}))
        return;
      corrections.push({category:category,fact:fact,quote:quote,paragraph:number});
    });
    repairReview='Proposed '+corrections.length+' additional evidence facts for independent validation';
  }
  // Independently verify every proposed addition and every proposed relationship
  // before accepting them into the evidence record.
  var proposals=auditData.relationships.map(function(link){return link;});
  var gatePrompt=[
    'Independently review source evidence. Return JSON only:',
    '{"acceptedAdditions":[number indexes],"acceptedRelationships":[number indexes]}.',
    'Accept an addition only if the quote supports ALL of the fact, including this job, location, status, scope and certainty; reject general theory and duplicated findings.',
    'Accept a relationship only where the quoted wording explicitly supports that exact connection between both referenced facts, not simply proximity.',
    'Be conservative. Return zero accepted indexes when unsupported.',
    'ADDitions (zero-based indexes):'+JSON.stringify(corrections),
    'RELATIONSHIPS (zero-based indexes):'+JSON.stringify(proposals),
    'EXISTING FACTS:'+JSON.stringify(packageData.facts.map(function(f){return {id:f.id,fact:f.fact,quote:f.quote};})),
    'ORIGINAL:'+bc_caseEvidenceText_(source)
  ].join('\n');
  var gateApi=bc_sendPromptViaOpenAI(gatePrompt,1800,MODEL_CHEAP);
  if(!gateApi||!gateApi.success)throw new Error('Independent repair and relationship verification failed');
  reviewCost+=Number(gateApi.cost||0);
  var gate;
  try {
    gate=JSON.parse(String(gateApi.text||'').trim().replace(/^\x60\x60\x60(?:json)?\s*/i,'').replace(/\s*\x60\x60\x60$/,''));
    if(!Array.isArray(gate.acceptedAdditions)||!Array.isArray(gate.acceptedRelationships))throw new Error('Missing accepted index arrays');
  }catch(e){throw new Error('Invalid independent evidence verification: '+e.message);}
  function acceptedIndexes(values,max) {
    var result={};
    values.forEach(function(n){if(!Number.isInteger(n)||n<0||n>=max)throw new Error('Invalid verification index '+n);result[n]=true;});
    return result;
  }
  var acceptedAdd=acceptedIndexes(gate.acceptedAdditions,corrections.length);
  var acceptedRel=acceptedIndexes(gate.acceptedRelationships,proposals.length);
  // Add accepted facts, rebuild source-ordered IDs and retain both old and new fact references.
  var acceptedFacts=corrections.filter(function(c,i){return acceptedAdd[i];});
  if(acceptedFacts.length){
    acceptedFacts.forEach(function(c){packageData.facts.push({category:c.category,fact:c.fact,quote:c.quote,
      sourceOrder:bc_caseEvidenceText_(source).toLowerCase().indexOf(c.quote.toLowerCase()),
      status:'source_documented',scope:'As stated in quotation',
      certainty:'Documented in source; not independently verified',relatedFactIds:[]});});
    packageData.facts.sort(function(a,b){return a.sourceOrder-b.sourceOrder;});
    var remap={};
    packageData.facts.forEach(function(f,i){if(f.id)remap[f.id]='F'+('000'+(i+1)).slice(-4);});
    packageData.facts.forEach(function(f,i){f.id='F'+('000'+(i+1)).slice(-4);});
    proposals.forEach(function(link){link.fromId=remap[link.fromId]||link.fromId;link.toId=remap[link.toId]||link.toId;});
    checked.forEach(function(p){p.factIds=p.factIds.map(function(id){return remap[id]||id;});});
    acceptedFacts.forEach(function(c){
      var p=checked.filter(function(x){return x.paragraph===c.paragraph;})[0];
      if(p){var f=packageData.facts.filter(function(x){return x.fact===c.fact&&x.quote===c.quote;})[0];if(f)p.factIds.push(f.id);}
    });
    packageData.chronology=packageData.facts.filter(function(f){return f.category==='Restoration Methods';})
      .map(function(f){return {factId:f.id,sourceOrder:f.sourceOrder,status:f.status};});
  }
  var ids={}; packageData.facts.forEach(function(f){ids[f.id]=true;});

  var normalizedSource=bc_caseEvidenceText_(source);
  relationships.rejected=[];
  proposals.forEach(function(link,index){
    if(!acceptedRel[index])return;
    var reason='';
    if(!ids[link.fromId]||!ids[link.toId]||link.fromId===link.toId)
      reason='Invalid or identical fact IDs';
    else if(['reason_for','led_to','limited','resulted_in'].indexOf(String(link.type))<0)
      reason='Invalid relationship type';
    else {
      var quote=bc_caseEvidenceRepairQuote_(link.sourceQuote,normalizedSource);
      if(!quote)reason='Relationship quotation not found verbatim in original source';
    }
    if(reason){
      relationships.rejected.push({proposalIndex:index,fromId:link.fromId,toId:link.toId,
        reason:reason,proposedQuote:String(link.sourceQuote||'').slice(0,180)});
      return; // Reject only this relationship, never accept unsupported evidence.
    }
    relationships.links.push({fromId:link.fromId,toId:link.toId,type:link.type,
      sourceQuote:quote,verification:'Source quotation found; semantic connection requires independent review'});
  });
  relationships.status=relationships.rejected.length?
    'REVIEW REQUIRED — '+relationships.rejected.length+' proposed relationship(s) rejected; accepted links are source-quoted':
    (relationships.links.length?'Proposed source-quoted relationships; review causal meaning':'No directly supported relationships identified');
  checked.forEach(function(p){
    if(p.factIds.length){p.status='EXTRACTED';return;}
    var d=decisions[p.paragraph];
    p.classification=d?d.classification:'UNCERTAIN';
    p.reason=d?d.reason:'Audit did not classify this passage';
    p.status=(d&&['GENERAL_ADVICE','SEO_OR_HEADING','DUPLICATE','MEDIA_CAPTION'].indexOf(d.classification)>=0)?
      'EXCLUDED WITH REASON':'REVIEW REQUIRED';
  });
  var unresolved=checked.filter(function(p){return p.status==='REVIEW REQUIRED';});
  var coverage={state:unresolved.length?'REVIEW REQUIRED':'CLASSIFIED — RELATIONSHIP AND FACT REVIEW STILL ADVISED',
    mappedParagraphs:checked.filter(function(p){return p.status==='EXTRACTED';}).length,
    excludedParagraphs:checked.filter(function(p){return p.status==='EXCLUDED WITH REASON';}).length,
    unresolvedParagraphs:unresolved.length,totalParagraphs:checked.length,
    notes:'Source paragraph classification is AI-reviewed, not independent factual verification. Missing facts require extraction correction.',
    ledger:checked};
  // Compare earlier extraction after verification, never silently discard its material facts.
  var dropped=[];
  (prior||[]).forEach(function(old){
    if(!packageData.facts.some(function(f){return f.quote.toLowerCase()===String(old.quote||'').toLowerCase()||f.fact.toLowerCase()===String(old.fact||'').toLowerCase();}))
      dropped.push({category:old.category,fact:old.fact,quote:old.quote,status:'Requires review — possibly omitted, paraphrased, merged or previously unsupported'});
  });
  if(dropped.length){coverage.state='REVIEW REQUIRED — PREVIOUS EVIDENCE CHANGED';coverage.previousEvidenceChanges=dropped;}

  var gap={unknowns:packageData.unknowns,
    repairReview:repairReview,acceptedMissingFacts:acceptedFacts.length,
    unresolvedSourcePassages:unresolved.map(function(p){return {paragraph:p.paragraph,preview:p.preview,classification:p.classification,reason:p.reason};}),
    contradictions:[],status:'No independent contradiction audit performed'};
  // Main package stores factual records; companion columns own the other datasets.
  var core={version:packageData.version,postId:packageData.postId,
    sourceDigest:packageData.sourceDigest,provenance:packageData.provenance,
    facts:packageData.facts,chronology:packageData.chronology};
  var media={imagesAndVideos:packageData.media,links:packageData.links,
    note:'Source captions and alt text are claims, not independent image verification'};
  gap.previousEvidenceChanges=dropped;
  var parts=[core,relationships,coverage,media,gap].map(function(o){return JSON.stringify(o);});
  parts.forEach(function(v,i){if(v.length>49000)throw new Error('Evidence column '+(i+1)+' exceeds safe 49000-character limit; no evidence written');});
  return {parts:parts,cost:reviewCost};
}

/**
 * Append-only run history in the existing "Case Study Evidence Package" tab.
 * A:P mirror posts' 16 evidence columns. Q:W identify the run and snapshot.
 * Never overwrite a prior history row; fail closed if the archive is unavailable.
 */
function bc_caseEvidenceArchive_(sh,row,postId,phase,runId) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var history=ss.getSheetByName('Case Study Evidence Package');
  if(!history)throw new Error('Case Study Evidence Package history sheet missing; source evidence not overwritten');
  var expected=['Project Location & Property','Observed Conditions','Client Requirements & Restoration Objectives',
    'Materials & Construction','Inspection & Tests','Restoration Methods','Products & Equipment',
    'Technical Findings','Project Results','Maintenance Handover','Case Study Evidence Package',
    'Evidence Relationships','Evidence Coverage','Evidence Media & Links','Evidence Gaps & Uncertainties',
    'Evidence Status','Post ID','Posts Row','Run ID','Snapshot Type','Snapshot Time','Evidence Version','Source Digest'];
  var existing=history.getRange(1,1,1,23).getDisplayValues()[0];
  expected.forEach(function(h,i){if(existing[i]!==h)throw new Error('Archive header mismatch at '+(i+1)+': '+h);});
  var heads=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  var first=heads.indexOf(expected[0])+1;
  if(first<1||expected.slice(0,16).some(function(h,i){return heads[first-1+i]!==h;}))
    throw new Error('Evidence columns not contiguous in posts; archive aborted');
  var cells=sh.getRange(row,first,1,16).getValues()[0].map(function(v){return String(v==null?'':v);});
  var status=cells[15], match=status.match(/SOURCE-QUOTED-(v\d+)\s*\|\s*sha256:([a-f0-9]+)/i);
  var version=match?match[1]:(status?'legacy/unknown':'none');
  var digest=match?match[2]:'';
  var appended=cells.concat([String(postId),row,String(runId),String(phase),new Date(),version,digest]);
  var lock=LockService.getDocumentLock()||LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error('Could not lock evidence history for append');
  try {
    var next=Math.max(2,history.getLastRow()+1);
    if(next>history.getMaxRows())history.insertRowsAfter(history.getMaxRows(),Math.max(100,next-history.getMaxRows()));
    history.getRange(next,1,1,23).setValues([appended]);
    SpreadsheetApp.flush();
  } finally {lock.releaseLock();}
  return next;
}
/** Manually preserve a current row (including v9) without rerunning the AI. */
function bc_caseEvidenceSnapshotActiveRow() {
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if(!sh)throw new Error('posts sheet missing');
  var row=SpreadsheetApp.getActiveRange().getRow();
  if(row<2)throw new Error('Select a posts data row');
  var h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  var idcol=h.indexOf('Post ID')+1;
  if(!idcol)throw new Error('Post ID column missing');
  var id=String(sh.getRange(row,idcol).getDisplayValue()).trim();
  if(!id)throw new Error('Post ID missing');
  var runId='MANUAL-'+Utilities.getUuid();
  return bc_caseEvidenceArchive_(sh,row,id,'MANUAL SNAPSHOT',runId);
}

var BC_CASE_EVIDENCE_VERSION_ = 'v9';
function bc_caseEvidenceStatus_(digest) { return 'SOURCE-QUOTED-' + BC_CASE_EVIDENCE_VERSION_ + ' | sha256:' + digest; }
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
  if (oldStatus === bc_caseEvidenceStatus_(digest)) {
    return {success:true, skipped:true, cost:0, message:'Source-quoted evidence unchanged; use bc_caseEvidenceSnapshotActiveRow to archive current results'};
  }
  var labels = ['Project Location & Property','Observed Conditions','Client Requirements & Restoration Objectives',
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
    'CLASSIFICATION AND CERTAINTY — MANDATORY:',
    'Project Location & Property: actual property facts only; exclude generic council, postcode and regional geography that does not describe this property.',
    'Observed Conditions: what was actually seen or recorded at this site; do not place treatment decisions here.',
    'Client Requirements & Restoration Objectives: include explicitly documented client requests labelled CLIENT REQUEST:, and documented professional restoration aims labelled RESTORATION OBJECTIVE:. Never present a restoration objective as something the homeowner asked for. If neither is documented, use [].',
    'RESTORATION OBJECTIVES MUST BE EXTRACTED: Inspect the original article for explicit project aims, planned outcomes and reasons the work was undertaken (for example stabilising loose tiles, removing old coatings, retaining the historic surface). Add each distinct documented objective to Client Requirements & Restoration Objectives as RESTORATION OBJECTIVE: even when no client request is recorded. Do not infer objectives solely from the completed treatment; require an explicit original-source quotation showing the aim or need.',
    'Materials & Construction: ONLY documented installed materials and construction; exclude hypothetical moisture paths, restoration systems, future sealers and generic technical principles.',
    'Inspection & Tests: separate what was inspected or CONSIDERED from findings actually confirmed; never imply a test occurred merely because it would be normal practice.',
    'Restoration Methods: actions actually performed, in documented order; conditional treatments must remain conditional, never imply they always occurred.',
    'Products & Equipment: products/tools documented as used, not items merely recommended or hypothetical.',
    'Technical Findings: actual on-site findings, meaningful negative findings and explicitly supported professional decisions; exclude unverified mechanisms and general explanations.',
    'Project Results: ONLY documented results and limitations; distinguish observed outcome from generic benefits.',
    'Maintenance Handover: include documented maintenance recommendations as RECOMMENDATIONS; only say a personal handover occurred when explicitly recorded.',
    'Preserve meaningful negative results (e.g. no major adhesive field), declined or unnecessary treatment, and differences between considered possibilities and confirmed findings.',
    'DISTINCT METHOD RETENTION: Extract each separately documented completed restoration step, including old sealer or coating stripping, even if a broader cleaning or sealing summary is already recorded. Do not collapse stripping, cleaning, rinsing, repairs and sealing into one fact if doing so loses a distinct actual operation.',
    'REPAIR OPTIONS CONSIDERED: Retain documented repair alternatives and assessment decisions, including replacement-tile matching considered but not performed. Label options as considered, not completed. Distinguish local resetting from replacement and preserve why a more extensive repair was unnecessary where documented.',
    'FACT–QUOTE ENTAILMENT: every fact must be directly supported by its own quote in meaning, certainty, attribution and action. An exact copied quote is NOT sufficient when the fact changes its meaning.',
    'Do not repeat generic advice or hypothetical examples as project findings. Do not invent a fact to fill any category. Use [] for categories with no actual applicable evidence.',
    'Each category must be accurate and concise; retain material distinct findings without keyword repetition.',
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
  // Second-pass semantic validation within extraction, not an extra pipeline stage.
  // Exact quote matching alone cannot establish correct category or certainty.
  var reviewPrompt = [
    'You are an independent case-study evidence verifier. Inspect the SOURCE and the EXTRACTED JSON.',
    'Return ONLY a JSON object with exactly these same ten category keys and arrays of {fact,quote}.',
    'For every item, verify: (1) its quote directly entails the fact, (2) its category is appropriate,',
    '(3) its certainty and attribution have not been strengthened, and (4) it concerns this job rather than general advice.',
    'Drop unsupported, generic geographical padding and misleading inferred client requirements.',
    'Keep negative findings and distinct material job details. Separate considered conditions from confirmed discoveries.',
    'COVERAGE REVIEW: Compare every material passage in the source against the candidate evidence. Restore any omitted distinct project observation, negative finding, decision, completed operation, limitation or result with an exact quote; remove only duplicate statements, generic advice or SEO filler. Do not invent connective facts.',
    'Do not remove unique documented completed steps while deduplicating broader method summaries: preserve old sealer stripping separately from cleaning and applying new sealer.',
    'Retain documented repair options considered (such as replacement-tile matching), clearly as considerations rather than completed repairs.',
    'In Client Requirements & Restoration Objectives, prefix each supported fact with CLIENT REQUEST: when explicitly attributed to the client, or RESTORATION OBJECTIVE: when a documented professional aim. Never infer client speech, wishes or an undocumented objective.',
    'Do not return an empty Client Requirements & Restoration Objectives array if the original explicitly states restoration aims or why the treatment was necessary. Verify that each such objective is retained with its exact quote. Keep client requests and professional objectives separately labelled.',
    'General maintenance RECOMMENDATIONS may appear in Maintenance Handover, but do not imply a personal handover.',
    'Materials & Construction must not include theoretical moisture paths or future protective treatments.',
    'Keep conditional methods conditional. Never infer tests, chemicals used, owner statements or outcomes.',
    'You may move an existing fact to a better category or reword it to match its original quote,',
    'but NEVER invent quotes or facts. Preserve exact contiguous source quotations.',
    'If a category has no supported items, use [].',
    'ORIGINAL ARTICLE:\n' + bc_caseEvidenceText_(source),
    'EXTRACTED JSON:\n' + JSON.stringify(JSON.parse('{' + labels.map(function(label,i){
      return JSON.stringify(label)+':'+output[i];
    }).join(',') + '}'))
  ].join('\n');
  var validatedOutput = null;
  var reviewError = '';
  for (var reviewAttempt = 0; reviewAttempt < 2; reviewAttempt++) {
    var reviewApi = bc_sendPromptViaOpenAI(reviewPrompt + (reviewAttempt ? '\nPRIOR OUTPUT INVALID: ' + reviewError + '. Return valid, source-supported JSON only.' : ''), 5000, MODEL_CHEAP);
    if (!reviewApi || !reviewApi.success) throw new Error('Evidence semantic validation API failed: ' + (reviewApi && reviewApi.message || 'unknown'));
    totalCost += Number(reviewApi.cost || 0);
    try {
      var reviewRaw=String(reviewApi.text||'').trim().replace(/^\x60\x60\x60(?:json)?\s*/i,'').replace(/\s*\x60\x60\x60$/,'');
      var reviewParsed=JSON.parse(reviewRaw);
      var checkedReview=[];
      labels.forEach(function(label) {
        if (!Array.isArray(reviewParsed[label])) throw new Error('Invalid reviewed category: '+label);
        var items=reviewParsed[label].map(function(item) {
          var fact=String(item && item.fact || '').trim();
          var quote=String(item && item.quote || '').replace(/\s+/g,' ').trim();
          if (!fact||!quote) throw new Error('Incomplete reviewed fact in '+label);
          var nq=quote.replace(/[’‘]/g,"'").replace(/[“”]/g,'"');
          if (normalizedSource.indexOf(nq)<0) {
            var repaired=bc_caseEvidenceRepairQuote_(quote,normalizedSource);
            if (!repaired) throw new Error('Unsupported reviewed quote in '+label+': '+quote.slice(0,90));
            quote=repaired;
          }
          return {fact:fact,quote:quote};
        });
        checkedReview.push(JSON.stringify(items));
      });
      validatedOutput=checkedReview;
      break;
    } catch(e) { reviewError=e.message; }
  }
  if (!validatedOutput) throw new Error('Evidence semantic validation failed: '+reviewError);
  output=validatedOutput;
  // Fail closed: no partial evidence is written on model error or quote mismatch.
  for (var j=0;j<labels.length;j++) col(labels[j]);
  var packageData=bc_caseEvidenceBuildPackage_(postId,source,digest,labels,output);
  var extraLabels=['Case Study Evidence Package','Evidence Relationships','Evidence Coverage',
    'Evidence Media & Links','Evidence Gaps & Uncertainties'];
  extraLabels.forEach(function(label,i){
    if(col(label)!==col(labels[labels.length-1])+i+1) throw new Error('Evidence supporting columns must immediately follow Maintenance Handover: '+label);
  });
  if(col('Evidence Status')!==col(extraLabels[extraLabels.length-1])+1)
    throw new Error('Evidence Status must immediately follow Evidence Gaps & Uncertainties');
  var priorFacts=[];
  try {
    var priorRaw=String(sh.getRange(row,col('Case Study Evidence Package')).getValue()||'');
    var priorObj=priorRaw?JSON.parse(priorRaw):null;
    if(priorObj&&String(priorObj.postId)===postId&&Array.isArray(priorObj.facts))priorFacts=priorObj.facts;
  }catch(ignorePrior) { throw new Error('Cannot read previous evidence package safely: '+ignorePrior.message); }
  var checkedExtras=bc_caseEvidenceExtraColumns_(packageData,source,labels,priorFacts);
  totalCost+=checkedExtras.cost;
  var extras=checkedExtras.parts;
  // Archive current evidence BEFORE replacing it, then preserve the finished run.
  // A failed pre-run archive aborts without overwriting the previous evidence.
  var runId=Utilities.getUuid();
  bc_caseEvidenceArchive_(sh,row,postId,'BEFORE',runId);
  sh.getRange(row,col(labels[0]),1,16).setValues([output.concat(extras,[bc_caseEvidenceStatus_(digest)])]);
  bc_caseEvidenceArchive_(sh,row,postId,'AFTER',runId);
  return {success:true, skipped:false, cost:totalCost, message:'Evidence checked; BEFORE and AFTER archived under run '+runId};
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
  if(status!==bc_caseEvidenceStatus_(bc_caseEvidenceDigest_(original))){
    throw new Error('Case Study Evidence is missing or out of date. Run Extract Case Study Evidence in Content Intelligence first.');
  }
  var labels=['Project Location & Property','Observed Conditions','Client Requirements & Restoration Objectives','Materials & Construction',
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
 * W0 filtered Stone Type batch. Only visible Case Studies for one stone type
 * are returned; completed evidence is not scheduled again.
 * The row extraction remains source-fingerprint aware for safety.
 */
function bc_caseEvidenceFilteredRows() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if (!sh) return {success:false,message:'posts sheet missing'};
  // Supports both standard Sheets filters and the custom Stone Type filter that hides rows.
  var last = sh.getLastRow();
  if (last < 2) return {success:true,stoneType:'Filtered selection',rows:[],alreadyCompleted:0};
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  function col(name) { var idx=headers.indexOf(name); if(idx<0) throw new Error('Missing heading: '+name); return idx+1; }
  var typeCol=col('Article Type'), stoneCol=col('Stone Type'), statusCol=col('Evidence Status');
  var types=sh.getRange(2,typeCol,last-1,1).getDisplayValues();
  var stones=sh.getRange(2,stoneCol,last-1,1).getDisplayValues();
  var statuses=sh.getRange(2,statusCol,last-1,1).getDisplayValues();
  var rows=[], completed=0, stoneTypes={}, hidden=0;
  for (var i=0;i<last-1;i++) {
    var row=i+2;
    if (sh.isRowHiddenByFilter(row) || sh.isRowHiddenByUser(row)) { hidden++; continue; }
    var stone=String(stones[i][0]||'').trim();
    if (stone) stoneTypes[stone]=true;
    if (String(types[i][0]||'').trim()!=='Case Study') continue;
    // Include completed rows: extraction checks the source hash and safely skips unchanged evidence.
    rows.push(row);
  }
  var names=Object.keys(stoneTypes);
  if (!hidden || names.length!==1) return {success:false,message:'Filter the posts sheet to exactly one Stone Type first. Visible types: '+(names.join(', ')||'none')};
  return {success:true,stoneType:names[0],rows:rows,alreadyCompleted:completed};
}
/** Public endpoint for client-side sequential batching. */
function bc_caseEvidenceRunRowBatch(row) {
  return bc_caseEvidenceExtract_(row);
}

/** Direct Row Analysis menu command: processes visible Case Studies, skipping unchanged evidence. */
function bc_caseEvidenceFilteredBatch() {
  var info = bc_caseEvidenceFilteredRows();
  var ui = SpreadsheetApp.getUi();
  if (!info.success) { ui.alert(info.message); return; }
  var started = Date.now(), done = 0, skipped = 0, failed = [], cost = 0, remaining = 0;
  for (var i=0; i<info.rows.length; i++) {
    if (Date.now()-started > 240000) { remaining = info.rows.length-i; break; }
    var row = info.rows[i];
    try {
      var result = bc_caseEvidenceExtract_(row);
      cost += Number(result.cost || 0);
      if (result.skipped) skipped++;
      else if (result.success) done++;
      else failed.push('Row '+row+': '+(result.message||'Unsuccessful'));
    } catch (e) { failed.push('Row '+row+': '+(e.message||e)); }
  }
  ui.alert('Case Study Evidence — '+info.stoneType+'\n'
    +'Extracted: '+done+'\nAlready complete: '+skipped
    +'\nFailed: '+failed.length+'\nRemaining: '+remaining
    +'\nAPI cost: $'+cost.toFixed(4)
    +(failed.length?'\n'+failed.slice(0,5).join('\n'):'')
    +'\nRerun to continue; unchanged evidence is skipped.');
}

/** Open a live, modeless batch-progress window for the current Stone Type. */
function bc_caseEvidenceOpenBatchProgress() {
  var html = HtmlService.createHtmlOutputFromFile('bc_CaseEvidenceBatchProgress')
    .setWidth(450).setHeight(340);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Case Study Evidence — Batch');
}
