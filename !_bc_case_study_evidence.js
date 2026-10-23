/**
 * Content Intelligence: project facts extracted from original case studies.
 * Ten source-paragraph-referenced evidence categories feed downstream stages.
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
 * Media/link extractor retained for source provenance. Its legacy fact parser
 * is called with empty arrays; v10 facts are built from paragraph references.
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
/**
 * Append-only run history in the existing "Case Study Evidence Package" tab.
 * A:G identify the run and snapshot. H:W mirror posts' 16 evidence columns.
 * Never overwrite a prior history row; fail closed if the archive is unavailable.
 */
function bc_caseEvidenceArchive_(sh,row,postId,phase,runId,diagnostics) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var history=ss.getSheetByName('Case Study Evidence Package');
  if(!history)throw new Error('Case Study Evidence Package history sheet missing; source evidence not overwritten');
  var expected=['Project Location & Property','Observed Conditions','Client Requirements & Restoration Objectives',
    'Materials & Construction','Inspection & Tests','Restoration Methods','Products & Equipment',
    'Technical Findings','Project Results','Maintenance Handover','Case Study Evidence Package',
    'Evidence Relationships','Evidence Coverage','Evidence Media & Links','Evidence Gaps & Uncertainties',
    'Evidence Status','Post ID','Posts Row','Run ID','Snapshot Type','Snapshot Time','Evidence Version','Source Digest'];
  var historyHeadings=['Post ID','Posts Row','Run ID','Snapshot Type','Snapshot Time','Evidence Version','Source Digest'].concat(expected.slice(0,16));
  var diagnosticHeadings=['Prompt Fingerprint','AI Model','API Response ID','Token Usage','Response Fingerprint'];
  var existing=history.getRange(1,1,1,28).getDisplayValues()[0];
  historyHeadings.concat(diagnosticHeadings).forEach(function(h,i){if(existing[i]!==h)throw new Error('Archive header mismatch at '+(i+1)+': '+h);});
  var heads=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  var first=heads.indexOf(expected[0])+1;
  if(first<1||expected.slice(0,16).some(function(h,i){return heads[first-1+i]!==h;}))
    throw new Error('Evidence columns not contiguous in posts; archive aborted');
  var cells=sh.getRange(row,first,1,16).getValues()[0].map(function(v){return String(v==null?'':v);});
  var status=cells[15], match=status.match(/SOURCE-QUOTED-(v\d+)\s*\|\s*sha256:([a-f0-9]+)/i);
  var version=match?match[1]:(status?'legacy/unknown':'none');
  var digest=match?match[2]:'';
  var d=diagnostics||{};
  var diagnosticValues=[d.promptFingerprint||'',d.model||'',d.responseId||'',d.tokenUsage||'',d.responseFingerprint||''];
  var appended=[String(postId),row,String(runId),String(phase),new Date(),version,digest].concat(cells,diagnosticValues);
  var lock=LockService.getDocumentLock()||LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error('Could not lock evidence history for append');
  try {
    var next=Math.max(2,history.getLastRow()+1);
    if(next>history.getMaxRows())history.insertRowsAfter(history.getMaxRows(),Math.max(100,next-history.getMaxRows()));
    history.getRange(next,1,1,28).setValues([appended]);
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
 var prompt=['Extract documented facts from THIS floor restoration project in ONE pass. First move through the numbered source paragraphs in order, identifying each distinct project-specific fact; only then classify each fact.',
 'Do not skip factual inspections, negative findings, construction, conditions, client requests, restoration objectives, completed methods, equipment, results or maintenance simply because a paragraph also contains general explanation.',
 'A paragraph can have multiple distinct facts. Keep important operations separate, including stripping, repairs, cleaning, rinsing, extraction and sealing where performed.',
 'STRICT PROJECT-EVIDENCE RULE: only retain facts clearly connected to this specific job. Exclude generic material behaviour, hypothetical homeowner symptoms, possible contaminants, other projects and conditional treatments.',
 'Never turn advice, possibilities or options into observed findings or performed work. Distinguish client requests from contractor objectives and verified observations from possible causes.',
 'Read the article in paragraph order. For each paragraph ask whether it records any actual observation, inspection, negative finding, action or outcome about this project; include all distinct documented facts without repetition.',
 'Return one JSON object with exactly one key "facts", containing a single array ordered by ascending paragraph number.',
 'Each fact must be {"paragraph":1,"category":"one exact category name","fact":"concise factual statement"}.',
 'Use only these ten category names: '+JSON.stringify(labels),
 'Do not include quotations or commentary. An empty array is valid when no project facts exist.',
 paras.map(function(t,i){return '['+(i+1)+'] '+t;}).join('\n')].join('\n');
 var promptFingerprint=bc_caseEvidenceDigest_(prompt);
 var api=bc_sendPromptViaOpenAI(prompt,6500,MODEL_CHEAP);
 if(!api||!api.success)throw Error('Evidence extraction API failed');
 if(api.truncated||api.finishReason==='length')throw Error('Evidence extraction response truncated; existing evidence preserved');
 var data=JSON.parse(String(api.text||'').trim()),cells=[],facts=[],n=0;
 if(!data||!Array.isArray(data.facts))throw Error('Expected a single facts array in response');
 var byCategory={};labels.forEach(function(name){byCategory[name]=[];});
 var previousParagraph=0;
 data.facts.forEach(function(item){
   var fact=String(item&&item.fact||'').trim();
   var p=Number(item&&item.paragraph);
   var name=String(item&&item.category||'').trim();
   if(!fact||!Number.isInteger(p)||p<1||p>paras.length)throw Error('Invalid fact or paragraph reference');
   if(!Object.prototype.hasOwnProperty.call(byCategory,name))throw Error('Unknown evidence category '+name);
   if(p<previousParagraph)throw Error('AI facts not in source paragraph order');
   previousParagraph=p;
   facts.push({id:'F'+('000'+(++n)).slice(-4),category:name,fact:fact,sourceParagraph:p,sourceOrder:p});
   byCategory[name].push({fact:fact,paragraph:p});
 });
 labels.forEach(function(name){cells.push(JSON.stringify(byCategory[name]));});
 facts.sort(function(a,b){return a.sourceParagraph-b.sourceParagraph;});
 facts.forEach(function(f,i){f.id='F'+('000'+(i+1)).slice(-4);});
 var base=bc_caseEvidenceBuildPackage_(id,html,digest,labels,labels.map(function(){return '[]';}));
 var pkg={version:'v10',postId:id,sourceDigest:digest,facts:facts,chronology:facts.filter(function(f){return f.category==='Restoration Methods';})};
 var mapped={};facts.forEach(function(f){mapped[f.sourceParagraph]=true;});
 var coverage={state:'SOURCE INDEXED',totalParagraphs:paras.length,paragraphsWithFacts:Object.keys(mapped).length,facts:facts.length};
 cells=cells.concat([JSON.stringify(pkg),JSON.stringify({links:[],status:'No inferred relationships'}),JSON.stringify(coverage),JSON.stringify({imagesAndVideos:base.media,links:base.links}),JSON.stringify({unknowns:base.unknowns}),bc_caseEvidenceStatus_(digest)]);
 cells.forEach(function(v){if(v.length>49000)throw Error('Evidence cell too long');});
 sh.getRange(row,first,1,16).setValues([cells]);
 bc_caseEvidenceArchive_(sh,row,id,'AFTER',run,{
  promptFingerprint:promptFingerprint,
  model:String(api.model||MODEL_CHEAP),
  responseId:String(api.responseId||''),
  tokenUsage:JSON.stringify({promptTokens:Number(api.promptTokens||0),completionTokens:Number(api.completionTokens||0),finishReason:String(api.finishReason||'unknown')}),
  responseFingerprint:bc_caseEvidenceDigest_(String(api.text||''))
 });
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
