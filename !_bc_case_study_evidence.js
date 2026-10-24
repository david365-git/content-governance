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
function bc_caseEvidenceArchive_(sh,row,postId,phase,runId,diagnostics,snapshotCells) {
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
  var cells=(snapshotCells||sh.getRange(row,first,1,16).getValues()[0]).map(function(v){return String(v==null?'':v);});
  if(cells.length!==16)throw Error('Evidence archive snapshot must contain 16 cells');
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
function bc_caseEvidenceExtract_(row,modelOverride,twoPass,splitMode){
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
 'A paragraph can have multiple distinct facts. Extract each independently meaningful observation, inspection finding, work operation, equipment item and result as a SEPARATE record, even when several appear in one sentence.',
 'Do not compress a sequence of operations into a single fact: distinguish stripping, lifting or resetting, applying cleaner, agitation, rinsing, slurry extraction, sealing and buffing where actually documented.',
 'Do not combine different tools or products in one record: for example, a rotary machine and a wet vacuum are two separate equipment facts when both were used.',
 'Extract each distinct documented project fact separately, preserving meaningful details of the method, equipment, material, condition, inspection and outcome. Do not merge separate tools, actions or findings, or omit qualifying details merely to shorten the response. Avoid only genuine duplication and trivial restatement.',
 'STRICT PROJECT-EVIDENCE RULE: only retain facts clearly connected to this specific job. Exclude generic material behaviour, hypothetical homeowner symptoms, possible contaminants, other projects and conditional treatments.',
 'Never turn advice, possibilities or options into observed findings or performed work. Distinguish client requests from contractor objectives and verified observations from possible causes.',
 'Read the article in paragraph order. For each paragraph ask whether it records any actual observation, inspection, negative finding, action or outcome about this project; include all distinct documented facts without repetition.',
 'Return one JSON object with exactly one key "facts", containing a single array ordered by ascending paragraph number.',
 'Each fact must be {"paragraph":1,"category":"one exact category name","fact":"concise factual statement"}.',
 'Classify each extracted fact by what it actually describes, not by isolated keywords or the source section heading. Use the following category meanings consistently across all materials and projects:',
 'Project Location & Property: where the project took place, property type, rooms and scope of physical location.',
 'Observed Conditions: documented pre-restoration condition, symptoms, defects, existing circumstances and outcomes of earlier attempts to address them.',
 'Client Requirements & Restoration Objectives: stated client wishes, agreed scope and intended restoration goals; do not confuse goals with completed results.',
 'Materials & Construction: identified floor materials, construction, layout and installed components.',
 'Inspection & Tests: inspections, assessments, checks, tests and their findings actually made during professional project assessment or treatment; do not label a reported symptom or earlier cleaning attempt as a professional test.',
 'Restoration Methods: actions and treatment steps actually carried out on this project.',
 'Products & Equipment: chemicals, fillers, sealers, tools and machinery actually used.',
 'Technical Findings: project-specific explanations, constraints, confirmed causes and implications established from the evidence; do not promote general possibilities to confirmed findings.',
 'Project Results: documented changes and completed outcomes, not intended goals or generic expected benefits.',
 'Maintenance Handover: maintenance instructions or advice specifically given to the client; exclude standalone general reader guidance.',
 'Use only these ten category names: '+JSON.stringify(labels),
 'Do not include quotations or commentary. An empty array is valid when no project facts exist.',
 paras.map(function(t,i){return '['+(i+1)+'] '+t;}).join('\n')].join('\n');
 var splitAt=Math.ceil(paras.length/2);
 var fullParagraphBlock=paras.map(function(t,i){return '['+(i+1)+'] '+t;}).join('\n');
 var splitPromptA=prompt;
 var splitPromptB='';
 if(splitMode){
   var halfA=paras.slice(0,splitAt).map(function(t,i){return '['+(i+1)+'] '+t;}).join('\n');
   var halfB=paras.slice(splitAt).map(function(t,i){return '['+(splitAt+i+1)+'] '+t;}).join('\n');
   splitPromptA=prompt.slice(0,prompt.length-fullParagraphBlock.length)+halfA;
   splitPromptB=prompt.slice(0,prompt.length-fullParagraphBlock.length)+halfB;
   prompt=splitPromptA;
 }
 var promptFingerprint=bc_caseEvidenceDigest_(prompt);
 var chosenModel=modelOverride||MODEL_CHEAP;
 var api=bc_sendPromptViaOpenAI(prompt,6500,chosenModel);
 if(!api||!api.success)throw Error('Evidence extraction API failed');
 if(api.truncated||api.finishReason==='length')throw Error('Evidence extraction response truncated; existing evidence preserved');
 var data=JSON.parse(String(api.text||'').trim()),cells=[],facts=[],n=0;
 if(!data||!Array.isArray(data.facts))throw Error('Expected a single facts array in response');
 var byCategory={};labels.forEach(function(name){byCategory[name]=[];});

 data.facts.forEach(function(item){
   var fact=String(item&&item.fact||'').trim();
   var p=Number(item&&item.paragraph);
   var name=String(item&&item.category||'').trim();
   if(!fact||!Number.isInteger(p)||p<1||p>(splitMode?splitAt:paras.length))throw Error('Invalid fact or paragraph reference');
   if(!Object.prototype.hasOwnProperty.call(byCategory,name))throw Error('Unknown evidence category '+name);

   facts.push({id:'F'+('000'+(++n)).slice(-4),category:name,fact:fact,sourceParagraph:p,sourceOrder:p});
   byCategory[name].push({fact:fact,paragraph:p});
 });
 labels.forEach(function(name){cells.push(JSON.stringify(byCategory[name]));});
 if(splitMode){
   // Two independent Luna extractions: do not show either half the other's facts.
   // Preserve the first half in a distinct archive snapshot after validation.
   var firstHalfFacts=facts.slice();
   var firstHalfPromptFingerprint=promptFingerprint;
   var firstHalfApi=api;
   var second=bc_sendPromptViaOpenAI(splitPromptB,6500,MODEL_CHEAP);
   if(!second||!second.success||second.truncated||second.finishReason==='length')throw Error('Second half extraction failed; existing row unchanged');
   var secondData=JSON.parse(String(second.text||'').trim());
   if(!secondData||!Array.isArray(secondData.facts))throw Error('Second half did not return a facts array');
   secondData.facts.forEach(function(item){
     var p=Number(item&&item.paragraph),cat=String(item&&item.category||'').trim(),fact=String(item&&item.fact||'').trim();
     if(!fact||!Number.isInteger(p)||p<=splitAt||p>paras.length||!Object.prototype.hasOwnProperty.call(byCategory,cat))
       throw Error('Invalid second-half fact/category/paragraph');
     facts.push({category:cat,fact:fact,sourceParagraph:p,sourceOrder:p});
     byCategory[cat].push({fact:fact,paragraph:p});
   });
   // Keep the first and second halves independent and avoid cross-half duplicate filtering.
   cells=labels.map(function(name){return JSON.stringify(byCategory[name]);});
   api={model:MODEL_CHEAP,responseId:String(second.responseId||''),text:String(second.text||''),
     promptTokens:Number(firstHalfApi.promptTokens||0)+Number(second.promptTokens||0),
     completionTokens:Number(firstHalfApi.completionTokens||0)+Number(second.completionTokens||0),
     finishReason:String(second.finishReason||'unknown'),
     cost:Number(firstHalfApi.cost||0)+Number(second.cost||0)};
   promptFingerprint=bc_caseEvidenceDigest_(splitPromptA+'\n---SECOND HALF---\n'+splitPromptB);
   // First-half response metadata is held until both calls succeed.
   Logger.log('Split Luna first half: model='+String(firstHalfApi.model||MODEL_CHEAP)+
     ' responseId='+String(firstHalfApi.responseId||'')+
     ' promptFingerprint='+firstHalfPromptFingerprint+
     ' responseFingerprint='+bc_caseEvidenceDigest_(String(firstHalfApi.text||''))+
     ' facts='+firstHalfFacts.length);
 }
 if(splitMode){
   // Exact repetitions may be removed; semantic similarities are REVIEW ONLY.
   var seenSplit={},uniqueSplit=[],splitDuplicates=[];
   function norm_(v){return String(v||'').toLowerCase().replace(/[’‘]/g,"'").replace(/[^a-z0-9]+/g,' ').trim().replace(/ +/g,' ');}
   facts.forEach(function(f){
     var sig=f.category+'|'+norm_(f.fact),kept=seenSplit[sig];
     if(kept){
       if(!kept.alsoParagraphs)kept.alsoParagraphs=[];
       if(f.sourceParagraph!==kept.sourceParagraph&&kept.alsoParagraphs.indexOf(f.sourceParagraph)<0)kept.alsoParagraphs.push(f.sourceParagraph);
       splitDuplicates.push({keptParagraph:kept.sourceParagraph,removedParagraph:f.sourceParagraph,category:f.category,fact:f.fact});
     }else{seenSplit[sig]=f;uniqueSplit.push(f);}
   });
   var stop={the:1,and:1,was:1,were:1,with:1,from:1,into:1,that:1,this:1,for:1,had:1,been:1,after:1,before:1,used:1,using:1,use:1,applied:1,apply:1,each:1,over:1,across:1,then:1,also:1,are:1,has:1,have:1,its:1,their:1,they:1,which:1,while:1,where:1,when:1,by:1,to:1,of:1,in:1,on:1,at:1};
   function stem_(w){
     if(/^vacuum(ed|ing)?$/.test(w))return 'vacuum';
     if(/^repair(s|ed|ing)?$/.test(w))return 'repair';
     if(/^seal(ed|ing|er)?$/.test(w))return 'seal';
     if(w.length>5&&/ing$/.test(w))return w.slice(0,-3);
     if(w.length>4&&/ed$/.test(w))return w.slice(0,-2);
     if(w.length>4&&/s$/.test(w)&&!/ss$/.test(w))return w.slice(0,-1);
     return w;
   }
   function profile_(f){
     var raw=norm_(f.fact).split(' '),words={},nums={};
     raw.forEach(function(w){if(/^\\d+$/.test(w))nums[w]=true;else if(w.length>2&&!stop[w])words[stem_(w)]=true;});
     return {words:words,nums:nums,raw:raw};
   }
   var opp=[['wet','dry'],['acidic','alkaline'],['acid','alkaline'],['initial','final'],['coarse','fine'],['satin','gloss'],['satin','matt'],['matt','gloss']];
   // A mention of another grit stage is context, not evidence that the stages match.
   // Compare each fact's primary grit (first stated grit), not its whole number set.
   function primaryGrit_(f){
     var m=String(f.fact||'').match(/\b(\d[\d,]*)\s*(?:-\s*)?grit\b/i);
     return m?m[1].replace(/,/g,''):'';
   }
   function stage_(f){
     var t=String(f.fact||'').toLowerCase();
     if(/\b(?:ground|grinding|grind)\b/.test(t))return 'grinding';
     if(/\b(?:honed|honing|hone)\b/.test(t))return 'honing';
     if(/\b(?:polished|polishing|polish)\b/.test(t))return 'polishing';
     return '';
   }
   function isMaintenance_(f){
     return f.category==='Maintenance Handover' || /\b(?:recommended|advised)\b.{0,100}\b(?:maintain|maintenance|cleaner|cleaning)\b/i.test(String(f.fact||''));
   }
   var profiles=uniqueSplit.map(profile_),splitDuplicateCandidates=[];
   for(var di=0;di<uniqueSplit.length;di++)for(var dj=di+1;dj<uniqueSplit.length;dj++){
     var left=uniqueSplit[di],right=uniqueSplit[dj],pa=profiles[di],pb=profiles[dj];
     if(isMaintenance_(left)!==isMaintenance_(right))continue;
     var ga=primaryGrit_(left),gb=primaryGrit_(right);
     if(ga&&gb&&ga!==gb)continue;
     var sa=stage_(left),sb=stage_(right);
     if(sa&&sb&&sa!==sb)continue;
     var na=Object.keys(pa.nums),nb=Object.keys(pb.nums);
     if(na.length&&nb.length&&(na.length!==nb.length||na.some(function(x){return !pb.nums[x];})))continue;
     if(opp.some(function(pair){return (pa.raw.indexOf(pair[0])>=0&&pb.raw.indexOf(pair[1])>=0)||(pa.raw.indexOf(pair[1])>=0&&pb.raw.indexOf(pair[0])>=0);}))continue;
     var ka=Object.keys(pa.words),kb=Object.keys(pb.words),size=Math.min(ka.length,kb.length);
     if(size<3)continue;
     var shared=ka.filter(function(x){return pb.words[x];}).length;
     if(shared<3)continue;
     var overlap=shared/size,jaccard=shared/(ka.length+kb.length-shared),sameP=left.sourceParagraph===right.sourceParagraph;
     var verdict='';
     if(sameP){
       if(overlap>=0.75)verdict='likely-duplicate';
       else if(left.category!==right.category&&overlap>=0.5)verdict='complementary';
     }else if(shared>=4&&overlap>=0.8)verdict='likely-duplicate';
     if(verdict)splitDuplicateCandidates.push({verdict:verdict,paragraphA:left.sourceParagraph,paragraphB:right.sourceParagraph,categoryA:left.category,factA:left.fact,categoryB:right.category,factB:right.fact,overlap:Number(overlap.toFixed(2)),jaccard:Number(jaccard.toFixed(2))});
   }
   splitDuplicateCandidates.sort(function(a,b){return a.verdict===b.verdict?b.overlap-a.overlap:a.verdict==='likely-duplicate'?-1:1;});
   // Consolidate only strongly overlapping, same-paragraph METHOD / EQUIPMENT
   // accounts of the same operation. Preserve both category memberships and
   // the discarded wording as audit evidence. Other review candidates stay.
   var consolidatedSplit=[],inactiveSplit=[];
   splitDuplicateCandidates.forEach(function(c){
     if(c.verdict!=='likely-duplicate'||c.paragraphA!==c.paragraphB||c.overlap<0.8)return;
     var types=[c.categoryA,c.categoryB];
     if(types.indexOf('Restoration Methods')<0||types.indexOf('Products & Equipment')<0)return;
     var ia=uniqueSplit.findIndex(function(f){return f.sourceParagraph===c.paragraphA&&f.category===c.categoryA&&f.fact===c.factA;});
     var ib=uniqueSplit.findIndex(function(f){return f.sourceParagraph===c.paragraphB&&f.category===c.categoryB&&f.fact===c.factB;});
     if(ia<0||ib<0||inactiveSplit.indexOf(ia)>=0||inactiveSplit.indexOf(ib)>=0)return;
     var fa=uniqueSplit[ia],fb=uniqueSplit[ib];
     var ga=primaryGrit_(fa),gb=primaryGrit_(fb),sa=stage_(fa),sb=stage_(fb);
     if(ga&&gb&&ga!==gb||sa&&sb&&sa!==sb)return;
     var keepIdx=fa.fact.length>=fb.fact.length?ia:ib,dropIdx=keepIdx===ia?ib:ia;
     var kept=uniqueSplit[keepIdx],dropped=uniqueSplit[dropIdx];
     if(!kept.evidenceCategories)kept.evidenceCategories=[kept.category];
     if(kept.evidenceCategories.indexOf(dropped.category)<0)kept.evidenceCategories.push(dropped.category);
     if(!kept.alsoFacts)kept.alsoFacts=[];
     kept.alsoFacts.push({category:dropped.category,paragraph:dropped.sourceParagraph,fact:dropped.fact});
     if(!kept.alsoParagraphs)kept.alsoParagraphs=[];
     if(dropped.sourceParagraph!==kept.sourceParagraph&&kept.alsoParagraphs.indexOf(dropped.sourceParagraph)<0)kept.alsoParagraphs.push(dropped.sourceParagraph);
     inactiveSplit.push(dropIdx);
     consolidatedSplit.push({keptParagraph:kept.sourceParagraph,keptFact:kept.fact,mergedCategory:dropped.category,mergedFact:dropped.fact});
   });
   facts=uniqueSplit.filter(function(f,i){return inactiveSplit.indexOf(i)<0;});
   byCategory={};labels.forEach(function(name){byCategory[name]=[];});
   facts.forEach(function(f){
     (f.evidenceCategories||[f.category]).forEach(function(cat){
       byCategory[cat].push({fact:f.fact,paragraph:f.sourceParagraph,alsoParagraphs:f.alsoParagraphs||[],alsoFacts:f.alsoFacts||[]});
     });
   });
   cells=labels.map(function(name){return JSON.stringify(byCategory[name]);});
   Logger.log('Split Luna strongly overlapping method/equipment facts consolidated: '+consolidatedSplit.length+'; '+JSON.stringify(consolidatedSplit));

   facts.forEach(function(f){byCategory[f.category].push({fact:f.fact,paragraph:f.sourceParagraph,alsoParagraphs:f.alsoParagraphs||[]});});
   cells=labels.map(function(name){return JSON.stringify(byCategory[name]);});
   Logger.log('Split Luna exact removed: '+splitDuplicates.length+'; '+JSON.stringify(splitDuplicates));
   Logger.log('Split Luna review candidates: '+splitDuplicateCandidates.length+'; '+JSON.stringify(splitDuplicateCandidates.slice(0,40)));
 }
 facts.sort(function(a,b){return a.sourceParagraph-b.sourceParagraph;});
 facts.forEach(function(f,i){f.id='F'+('000'+(i+1)).slice(-4);});
 var base=bc_caseEvidenceBuildPackage_(id,html,digest,labels,labels.map(function(){return '[]';}));
 var pkg={version:'v10',postId:id,sourceDigest:digest,facts:facts,chronology:facts.filter(function(f){return f.category==='Restoration Methods';})};
 var mapped={};facts.forEach(function(f){mapped[f.sourceParagraph]=true;});
 var coverage={state:'SOURCE INDEXED',totalParagraphs:paras.length,paragraphsWithFacts:Object.keys(mapped).length,facts:facts.length};
 if(splitMode){coverage.exactDuplicatesRemoved=splitDuplicates.length;coverage.possibleDuplicatesForReview=splitDuplicateCandidates.length;coverage.duplicateReviewCandidates=splitDuplicateCandidates.slice(0,30);coverage.likelyDuplicateCount=splitDuplicateCandidates.filter(function(x){return x.verdict==='likely-duplicate';}).length;coverage.complementaryCount=splitDuplicateCandidates.filter(function(x){return x.verdict==='complementary';}).length;coverage.consolidatedMethodEquipmentFacts=consolidatedSplit.length;coverage.consolidatedFactsAudit=consolidatedSplit.slice(0,30);}
 if(splitMode){
   var firstHalfMapped={};firstHalfFacts.forEach(function(f){firstHalfMapped[f.sourceParagraph]=true;});
   var halfPkg={version:'v10',postId:id,sourceDigest:digest,facts:firstHalfFacts,
     chronology:firstHalfFacts.filter(function(f){return f.category==='Restoration Methods';})};
   var halfCells=labels.map(function(name){return JSON.stringify(firstHalfFacts.filter(function(f){return f.category===name;}).map(function(f){return {fact:f.fact,paragraph:f.sourceParagraph};}));});
   halfCells=halfCells.concat([JSON.stringify(halfPkg),JSON.stringify({links:[],status:'No inferred relationships'}),
     JSON.stringify({state:'SOURCE INDEXED',totalParagraphs:paras.length,paragraphsWithFacts:Object.keys(firstHalfMapped).length,facts:firstHalfFacts.length}),
     JSON.stringify({imagesAndVideos:base.media,links:base.links}),JSON.stringify({unknowns:base.unknowns}),bc_caseEvidenceStatus_(digest)]);
   halfCells.forEach(function(v){if(v.length>49000)throw Error('First-half archive cell too long');});
   cells.forEach(function(v){if(v.length>49000)throw Error('Merged evidence cell too long');});
   bc_caseEvidenceArchive_(sh,row,id,'FIRST_HALF',run,{
      promptFingerprint:firstHalfPromptFingerprint,model:String(firstHalfApi.model||MODEL_CHEAP),
      responseId:String(firstHalfApi.responseId||''),
      tokenUsage:JSON.stringify({promptTokens:Number(firstHalfApi.promptTokens||0),completionTokens:Number(firstHalfApi.completionTokens||0),finishReason:String(firstHalfApi.finishReason||'unknown')}),
      responseFingerprint:bc_caseEvidenceDigest_(String(firstHalfApi.text||''))
   },halfCells);
 }
 cells=cells.concat([JSON.stringify(pkg),JSON.stringify({links:[],status:'No inferred relationships'}),JSON.stringify(coverage),JSON.stringify({imagesAndVideos:base.media,links:base.links}),JSON.stringify({unknowns:base.unknowns}),bc_caseEvidenceStatus_(digest)]);
 if(twoPass){
   // Persist and archive the first Luna pass before supplementing it.
   cells.forEach(function(v){if(v.length>49000)throw Error('First-pass evidence cell too long');});
   sh.getRange(row,first,1,16).setValues([cells]);
   bc_caseEvidenceArchive_(sh,row,id,'FIRST_PASS',run,{
     promptFingerprint:promptFingerprint,model:String(api.model||chosenModel),
     responseId:String(api.responseId||''),
     tokenUsage:JSON.stringify({promptTokens:Number(api.promptTokens||0),completionTokens:Number(api.completionTokens||0),finishReason:String(api.finishReason||'unknown')}),
     responseFingerprint:bc_caseEvidenceDigest_(String(api.text||''))
   });
   var prior=facts.map(function(f){return {paragraph:f.sourceParagraph,category:f.category,fact:f.fact};});
   var missingPrompt=[
     'Find ONLY important project-specific facts missing from the existing extraction. This is a gap-finding pass, not a rewrite or independent new extraction.',
     'Read every numbered source paragraph against ALL facts already extracted. Do not repeat, paraphrase, split or reclassify an existing fact. Ignore general advice, hypotheticals, conditional treatments and other projects.',
     'Preserve meaningful omitted qualifiers, distinct equipment, technical methods, actual inspections and outcomes where explicitly stated in the source.',
     'Use these exact category names: '+JSON.stringify(labels),
     'Return JSON with exactly one key "facts": an array of {"paragraph":1,"category":"exact category","fact":"only genuinely missing documented fact"} in ascending paragraph order. Return [] if nothing material is missing.',
     'ALREADY EXTRACTED FACTS (DO NOT REPEAT):',JSON.stringify(prior),
     'ORIGINAL NUMBERED PARAGRAPHS:',paras.map(function(t,i){return '['+(i+1)+'] '+t;}).join('\n')
   ].join('\n');
   var extra=bc_sendPromptViaOpenAI(missingPrompt,6500,MODEL_CHEAP);
   if(!extra||!extra.success||extra.truncated||extra.finishReason==='length')throw Error('Second Luna pass failed or truncated; first pass preserved in archive');
   var extraData=JSON.parse(String(extra.text||'').trim());
   if(!extraData||!Array.isArray(extraData.facts))throw Error('Second Luna pass did not return facts array; first pass preserved');
   function wordSet_(v){return String(v).toLowerCase().replace(/[^a-z0-9 ]/g,' ').split(/\s+/).filter(function(w){return w.length>2;});}
   var additions=[];
   extraData.facts.forEach(function(item){
     var p=Number(item&&item.paragraph),category=String(item&&item.category||'').trim(),fact=String(item&&item.fact||'').trim();
     if(!Number.isInteger(p)||p<1||p>paras.length||!fact||!Object.prototype.hasOwnProperty.call(byCategory,category))throw Error('Invalid second-pass fact');
     var w=wordSet_(fact);
     var duplicate=facts.concat(additions).some(function(f){
       if(f.sourceParagraph!==p)return false;
       var other=wordSet_(f.fact),shared=w.filter(function(x){return other.indexOf(x)!==-1;}).length;
       return String(f.fact).toLowerCase()===fact.toLowerCase() || (w.length>0&&other.length>0&&shared/Math.min(w.length,other.length)>=0.96);
     });
     if(!duplicate)additions.push({category:category,fact:fact,sourceParagraph:p,sourceOrder:p});
   });
   facts=facts.concat(additions).sort(function(a,b){return a.sourceParagraph-b.sourceParagraph;});
   facts.forEach(function(f,i){f.id='F'+('000'+(i+1)).slice(-4);});
   byCategory={};labels.forEach(function(name){byCategory[name]=[];});
   facts.forEach(function(f){byCategory[f.category].push({fact:f.fact,paragraph:f.sourceParagraph});});
   pkg.facts=facts;pkg.chronology=facts.filter(function(f){return f.category==='Restoration Methods';});
   mapped={};facts.forEach(function(f){mapped[f.sourceParagraph]=true;});
   coverage.paragraphsWithFacts=Object.keys(mapped).length;coverage.facts=facts.length;
   cells=labels.map(function(name){return JSON.stringify(byCategory[name]);}).concat([
     JSON.stringify(pkg),JSON.stringify({links:[],status:'No inferred relationships'}),JSON.stringify(coverage),
     JSON.stringify({imagesAndVideos:base.media,links:base.links}),JSON.stringify({unknowns:base.unknowns}),bc_caseEvidenceStatus_(digest)
   ]);
   api={model:MODEL_CHEAP,responseId:String(extra.responseId||''),text:String(extra.text||''),
     promptTokens:Number(api.promptTokens||0)+Number(extra.promptTokens||0),
     completionTokens:Number(api.completionTokens||0)+Number(extra.completionTokens||0),
     finishReason:String(extra.finishReason||'unknown'),
     cost:Number(api.cost||0)+Number(extra.cost||0)};
   promptFingerprint=bc_caseEvidenceDigest_(prompt+'\nSECOND PASS:\n'+missingPrompt);
 }
 cells.forEach(function(v){if(v.length>49000)throw Error('Evidence cell too long');});
 sh.getRange(row,first,1,16).setValues([cells]);
 bc_caseEvidenceArchive_(sh,row,id,'AFTER',run,{
  promptFingerprint:promptFingerprint,
  model:String(api.model||chosenModel),
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
function bc_caseEvidenceActiveRowSplitLunaTest() {
  return bc_caseEvidenceExtract_(SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getRow(),MODEL_CHEAP,false,true);
}

function bc_caseEvidenceActiveRowTwoPassLunaTest() {
  return bc_caseEvidenceExtract_(SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getRow(),MODEL_CHEAP,true);
}

function bc_caseEvidenceActiveRowMidTest() {
  return bc_caseEvidenceExtract_(SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getRow(), MODEL_MID);
}

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
