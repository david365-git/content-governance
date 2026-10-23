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
 * A:G identify the run and snapshot. H:W mirror posts' 16 evidence columns.
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
  var historyHeadings=['Post ID','Posts Row','Run ID','Snapshot Type','Snapshot Time','Evidence Version','Source Digest'].concat(expected.slice(0,16));
  var existing=history.getRange(1,1,1,23).getDisplayValues()[0];
  historyHeadings.forEach(function(h,i){if(existing[i]!==h)throw new Error('Archive header mismatch at '+(i+1)+': '+h);});
  var heads=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  var first=heads.indexOf(expected[0])+1;
  if(first<1||expected.slice(0,16).some(function(h,i){return heads[first-1+i]!==h;}))
    throw new Error('Evidence columns not contiguous in posts; archive aborted');
  var cells=sh.getRange(row,first,1,16).getValues()[0].map(function(v){return String(v==null?'':v);});
  var status=cells[15], match=status.match(/SOURCE-QUOTED-(v\d+)\s*\|\s*sha256:([a-f0-9]+)/i);
  var version=match?match[1]:(status?'legacy/unknown':'none');
  var digest=match?match[2]:'';
  var appended=[String(postId),row,String(runId),String(phase),new Date(),version,digest].concat(cells);
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

var BC_CASE_EVIDENCE_VERSION_ = 'v10';
function bc_caseEvidenceStatus_(digest) { return 'SOURCE-QUOTED-' + BC_CASE_EVIDENCE_VERSION_ + ' | sha256:' + digest; }
function bc_caseEvidenceExtract_(row){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('posts');
 row=Number(row);if(row<2)throw Error('Invalid posts row');
 var h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
 function c(n){var i=h.indexOf(n);if(i<0)throw Error('Missing '+n);return i+1;}
 if(String(sh.getRange(row,c('Article Type')).getValue())!=='Case Study')return {success:true,skipped:true,cost:0};
 var id=String(sh.getRange(row,c('Post ID')).getDisplayValue()).trim();
 var ex=ss.getSheetByName('site-export'),matches=[];
 ex.getRange(2,1,Math.max(1,ex.getLastRow()-1),1).getDisplayValues().forEach(function(v,i){if(String(v[0]).trim()===id)matches.push(i+2);});
 if(matches.length!==1)throw Error('Source ID lookup failed: '+id);
 var html=String(ex.getRange(matches[0],15).getValue()||''),digest=bc_caseEvidenceDigest_(html);
 if(!html.trim())throw Error('Original article empty');
 var labels=['Project Location & Property','Observed Conditions','Client Requirements & Restoration Objectives','Materials & Construction','Inspection & Tests','Restoration Methods','Products & Equipment','Technical Findings','Project Results','Maintenance Handover'];
 var first=c(labels[0]),extras=['Case Study Evidence Package','Evidence Relationships','Evidence Coverage','Evidence Media & Links','Evidence Gaps & Uncertainties','Evidence Status'];
 labels.concat(extras).forEach(function(n,i){if(c(n)!==first+i)throw Error('Evidence headings out of order: '+n);});
 var run=Utilities.getUuid();bc_caseEvidenceArchive_(sh,row,id,'BEFORE',run);
 var paras=html.replace(/<\/(?:p|h[1-6]|li|blockquote)>/gi,'\n').replace(/<[^>]+>/g,' ').split(/\n+/).map(function(t){return bc_caseEvidenceText_(t);}).filter(function(t){return t.length>=25;});
 var prompt=['Extract every distinct documented fact about THIS floor restoration project, in one pass.','Read every numbered paragraph, including paragraphs with several facts.','Retain property, observed conditions, client requests (only when explicit), restoration objectives, materials, actual inspections, negative findings, distinct completed methods, equipment, results and maintenance.','Separate stripping, repairing, cleaning, rinsing, extraction and sealing where documented.','Never turn advice or possible treatment into work performed. Never invent details.','Do not quote verbatim. Return JSON with exactly these keys: '+JSON.stringify(labels),'Each key maps to an array of {"fact":"concise factual statement","paragraph":1}.','Use the supporting original paragraph number; [] means no documented facts.',paras.map(function(t,i){return '['+(i+1)+'] '+t;}).join('\n')].join('\n');
 var api=bc_sendPromptViaOpenAI(prompt,6500,MODEL_CHEAP);
 if(!api||!api.success)throw Error('Evidence extraction API failed');
 var data=JSON.parse(String(api.text||'').trim()),cells=[],facts=[],n=0;
 labels.forEach(function(name){
  if(!Array.isArray(data[name]))throw Error('Missing category '+name);
  var arr=data[name].map(function(item){
   var fact=String(item.fact||'').trim(),p=Number(item.paragraph);
   if(!fact||!Number.isInteger(p)||p<1||p>paras.length)throw Error('Invalid paragraph for '+name);
   facts.push({id:'F'+('000'+(++n)).slice(-4),category:name,fact:fact,sourceParagraph:p,sourceOrder:p});
   return {fact:fact,paragraph:p};
  });cells.push(JSON.stringify(arr));
 });
 facts.sort(function(a,b){return a.sourceParagraph-b.sourceParagraph;});
 facts.forEach(function(f,i){f.id='F'+('000'+(i+1)).slice(-4);});
 var base=bc_caseEvidenceBuildPackage_(id,html,digest,labels,labels.map(function(){return '[]';}));
 var pkg={version:'v10',postId:id,sourceDigest:digest,facts:facts,chronology:facts.filter(function(f){return f.category==='Restoration Methods';})};
 var mapped={};facts.forEach(function(f){mapped[f.sourceParagraph]=true;});
 var coverage={state:'SOURCE INDEXED',totalParagraphs:paras.length,paragraphsWithFacts:Object.keys(mapped).length,facts:facts.length};
 cells=cells.concat([JSON.stringify(pkg),JSON.stringify({links:[],status:'No inferred relationships'}),JSON.stringify(coverage),JSON.stringify({imagesAndVideos:base.media,links:base.links}),JSON.stringify({unknowns:base.unknowns}),bc_caseEvidenceStatus_(digest)]);
 cells.forEach(function(v){if(v.length>49000)throw Error('Evidence cell too long');});
 sh.getRange(row,first,1,16).setValues([cells]);
 bc_caseEvidenceArchive_(sh,row,id,'AFTER',run);
 return {success:true,skipped:false,cost:Number(api.cost||0),message:'v10 extracted and archived'};
}
function bc_caseEvidenceActiveRow() {
  return bc_caseEvidenceExtract_(SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getRow());
}

/**
 * Shared read-only evidence block for W1.5 and W2B.
 * Do not trust a prior extraction unless its original HTML fingerprint matches.
 */
function bc_caseEvidencePromptBlock_(row){
 var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts'),h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
 function c(n){var i=h.indexOf(n);if(i<0)throw Error('Missing '+n);return i+1;}
 if(String(sh.getRange(row,c('Article Type')).getValue())!=='Case Study')return '';
 var id=String(sh.getRange(row,c('Post ID')).getDisplayValue()).trim();
 var pkg=JSON.parse(String(sh.getRange(row,c('Case Study Evidence Package')).getValue()||'{}'));
 if(pkg.version!=='v10'||String(pkg.postId)!==id)throw Error('Extract v10 evidence before content generation');
 var labels=['Project Location & Property','Observed Conditions','Client Requirements & Restoration Objectives','Materials & Construction','Inspection & Tests','Restoration Methods','Products & Equipment','Technical Findings','Project Results','Maintenance Handover'];
 return '\nAUTHORITATIVE CASE STUDY FACTS (original article is not a source for new project claims)\n'+labels.map(function(name){
  var arr=JSON.parse(String(sh.getRange(row,c(name)).getValue()||'[]'));
  if(!Array.isArray(arr))throw Error('Bad evidence in '+name);
  return name+':\n'+(arr.length?arr.map(function(v){return '- '+v.fact+' [source paragraph '+v.paragraph+']';}).join('\n'):'(Not documented)');
 }).join('\n')+'\nUse only these facts and explicitly confirmed owner corrections for project-specific claims. Never invent missing details.\nEND CASE STUDY FACTS\n';
}
/** W0 filtered batch. */
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
