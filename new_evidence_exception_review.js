/**
 * SHADOW TEST: AI evidence exceptions for selected Posts row.
 * Never changes Posts, source HTML, W2B or original evidence.
 * Stores only explicit owner answers in a separate ledger sheet.
 */
function new_eer_getSelected_() {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getActiveSheet();
  if (sh.getName()!=='posts') throw new Error('Select a row on the posts sheet first.');
  var row=sh.getActiveRange().getRow();
  if (row<2) throw new Error('Select a data row.');
  var heads=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  var data=sh.getRange(row,1,1,heads.length).getValues()[0];
  function col(name){var i=heads.indexOf(name);return i<0?'':String(data[i]||'').trim();}
  var id=col('Post ID')||col('ID');
  if (!id) throw new Error('Missing Post ID.');
  return {row:row,postId:id,title:col('Title'),articleType:col('Article Type')};
}
function new_eer_confirmations_(postId) {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Owner Confirmations');
  if(!sheet || sheet.getLastRow()<2)return [];
  var rows=sheet.getRange(2,1,sheet.getLastRow()-1,6).getValues();
  return rows.filter(function(x){return String(x[1]).trim()===String(postId)&&String(x[4]).trim();})
    .map(function(x){return {question:String(x[3]).trim(),answer:String(x[4]).trim()};});
}
function new_eer_reviewSelectedRow() {
  var p;
  try{p=new_eer_getSelected_();}catch(e){SpreadsheetApp.getUi().alert(String(e));return;}
  var html='<!doctype html><html><head><base target="_top"><style>body{font:14px Arial;padding:14px}button{padding:8px;margin:6px 0}textarea{width:96%;height:70px}#notice{white-space:pre-wrap} .item{border-top:1px solid #ccc;padding:12px 0}</style></head><body>'+
    '<b>Evidence exception test — Post '+p.postId+'</b><p>'+String(p.title).replace(/&/g,'&amp;').replace(/</g,'&lt;')+'</p>'+
    '<p>AI checks for important missing or contradictory facts. It cannot independently verify an undocumented fact.</p>'+
    '<button id="go" onclick="run()">Find exceptions (API cost)</button><div id="notice"></div><div id="list"></div>'+
    '<script>var PID='+JSON.stringify(p.postId)+', ROW='+p.row+';'+
    'function run(){document.getElementById("go").disabled=true;document.getElementById("notice").textContent="Checking source evidence...";google.script.run.withSuccessHandler(show).withFailureHandler(fail).new_eer_findExceptions(ROW,PID);}' +
    'function fail(e){document.getElementById("go").disabled=false;document.getElementById("notice").textContent=String(e);}' +
    'function show(r){document.getElementById("go").disabled=false;if(!r||!r.success){document.getElementById("notice").textContent=r&&r.message||"Failed";return;}document.getElementById("notice").textContent=r.questions.length+" material owner question(s); "+r.technicalExceptions.length+" technical/evidence issue(s); "+r.nonessential.length+" non-essential unknown(s). Status: "+r.status+". API cost $"+Number(r.cost||0).toFixed(4)+". Recheck after saving answers.";var d=document.getElementById("list");d.textContent="";r.technicalExceptions.forEach(function(q){var x=document.createElement("div");x.className="item";var p=document.createElement("p");p.textContent="Evidence/technical issue: "+q;x.appendChild(p);var t=document.createElement("textarea");t.placeholder="Your response or correction (optional)";x.appendChild(t);var sel=document.createElement("select");["Agree","Disagree","Needs correction"].forEach(function(v){var o=document.createElement("option");o.value=v;o.textContent=v;sel.appendChild(o);});x.appendChild(sel);var b=document.createElement("button");b.textContent="Save response";b.onclick=function(){b.disabled=true;google.script.run.withSuccessHandler(function(z){b.disabled=false;if(z&&z.success){b.textContent="Saved";}else{b.textContent=z&&z.message||"Save failed";}}).withFailureHandler(function(e){b.disabled=false;b.textContent=String(e);}).new_eer_saveReviewResponse(PID,q,sel.value,t.value);};x.appendChild(b);d.appendChild(x);});r.questions.forEach(function(q,i){var x=document.createElement("div");x.className="item";var h=document.createElement("p");h.textContent=q.question;x.appendChild(h);var context=document.createElement("p");context.textContent="Why ask: "+q.reason;x.appendChild(context);var t=document.createElement("textarea");t.placeholder="Your first-hand answer (leave blank to skip)";x.appendChild(t);var b=document.createElement("button");b.textContent="Save my confirmation";b.onclick=function(){var v=t.value.trim();if(!v)return; b.disabled=true;google.script.run.withSuccessHandler(function(z){b.disabled=false;context.textContent=z&&z.message||"Saved";}).withFailureHandler(function(e){b.disabled=false;context.textContent=String(e);}).new_eer_saveConfirmation(PID,q.question,v);};x.appendChild(b);d.appendChild(x);});}' +
    '</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(650).setHeight(620),'Evidence exception review — TEST');
}
function new_eer_findExceptions(row,expectedPostId) {
  try {
    var selected=new_eer_getSelected_();
    if (selected.row!==Number(row)||selected.postId!==String(expectedPostId))
      return {success:false,message:'Selection changed; re-open the review.'};
    var assembled=new_buildArticleWritingBrief(row);
    if (!assembled.success)return assembled;
    var confirmations=new_eer_confirmations_(expectedPostId);
    var prompt=[
      'You are an independent, cautious evidence exception checker for UK floor-restoration publishing.',
      'Audit ALL material evidence claims against source provenance, technical safeguards, project/owner evidence, media claims and article scope.',
      'Classify every MATERIAL issue, without an arbitrary maximum question count:',
      '(A) ownerQuestions: important facts needing firsthand owner confirmation,',
      '(B) technicalExceptions: material conflicts, unsupported claims or safety issues to fix from existing sources without asking owner,',
      '(C) nonessentialUnknowns: incidental facts that must not be invented but need no owner answer.',
      'Ask no question if evidence is sufficient. Do not turn ordinary missing metrics into questions.',
      'Do not repeat questions already answered in OWNER CONFIRMATIONS; incorporate these as owner testimony, not independent verification.',
      'If an existing answer directly resolves an issue, do not re-ask it; if it conflicts, flag the conflict as technicalException.',
      'A confirmed owner answer outranks AI inference, but cannot prove laboratory tests, engineering conclusions or photo provenance not directly observed.',
      'Return ONLY a JSON object with three arrays and a boolean reviewComplete. No markdown.',
      '{"ownerQuestions":[{"question":"direct question","reason":"material uncertainty"}],"technicalExceptions":["specific issue"],"nonessentialUnknowns":["unknown that should remain unclaimed"],"reviewComplete":true}',
      'reviewComplete must be false if you could not examine the complete supplied evidence. It does not mean externally verified or approved.',
      'Never invent facts or claim all exceptions are found with certainty.',
      'POST ID '+expectedPostId+'; ARTICLE TYPE '+selected.articleType,
      'OWNER CONFIRMATIONS FOR THIS POST ID: '+JSON.stringify(confirmations),
      'SOURCE AND EVIDENCE:',
      assembled.prompt
    ].join('\n');
    var model=typeof MODEL_CHEAP!=='undefined'?MODEL_CHEAP:null;
    if (!model||typeof bc_sendPromptViaOpenAI!=='function')throw new Error('OpenAI model or API sender unavailable.');
    var api=bc_sendPromptViaOpenAI(prompt,4000,model);
    if (!api||!api.success||api.truncated||api.finishReason==='length')
      return {success:false,message:'AI exception check incomplete: '+String(api&&api.message||api&&api.finishReason||'unknown')};
    var raw=String(api.text||'').trim().replace(/^```(?:json)?\s*/i,'').replace(/\s*```$/,'');
    var parsed=JSON.parse(raw);
    if(!parsed||!Array.isArray(parsed.ownerQuestions)||!Array.isArray(parsed.technicalExceptions)||
       !Array.isArray(parsed.nonessentialUnknowns)||parsed.reviewComplete!==true)
      throw new Error('AI did not complete a structured evidence review; no PASS issued.');
    var seen={};
    var questions=parsed.ownerQuestions.filter(function(q){
      if(!q||typeof q.question!=='string'||typeof q.reason!=='string')return false;
      var k=q.question.toLowerCase().replace(/[^a-z0-9]/g,'');
      if(!k||seen[k])return false;seen[k]=true;
      return !confirmations.some(function(c){
        return c.question.toLowerCase().replace(/[^a-z0-9]/g,'')===k;
      });
    }).map(function(q){return {question:q.question.slice(0,350),reason:q.reason.slice(0,500)};});
    var tech=parsed.technicalExceptions.filter(function(x){return typeof x==='string'&&x.trim();});
    var nonessential=parsed.nonessentialUnknowns.filter(function(x){return typeof x==='string'&&x.trim();});
    var status=questions.length||tech.length?'REVIEW REQUIRED':'NO MATERIAL EXCEPTIONS IDENTIFIED (AI CHECK ONLY)';
    return {success:true,questions:questions,technicalExceptions:tech,nonessential:nonessential,
      status:status,cost:Number(api.cost||0)};
  }catch(e){return {success:false,message:String(e)};}
}
function new_eer_saveConfirmation(postId,question,answer) {
  var lock=LockService.getDocumentLock();lock.waitLock(15000);
  try {
    var p=new_eer_getSelected_();
    if (p.postId!==String(postId))throw new Error('Selection changed; confirmation not saved.');
    question=String(question||'').trim();answer=String(answer||'').trim();
    if(!question||!answer)throw new Error('Question and answer are required.');
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getSheetByName('Owner Confirmations');
    if(!sheet){
      sheet=ss.insertSheet('Owner Confirmations');
      sheet.appendRow(['Timestamp','Post ID','Article Title','AI Question','Owner Answer','Status']);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(),p.postId,p.title,question,answer,'OWNER CONFIRMED — not yet integrated into W2B']);
    return {success:true,message:'Saved for Post '+p.postId+'. Not yet used by writing/fact-check stages.'};
  }catch(e){return {success:false,message:String(e)};}finally{lock.releaseLock();}
}

/**
 * Records owner's view on a technical/editorial finding.
 * Unlike firsthand project confirmations, this is NOT evidence that
 * a project event occurred. It is kept in a separate review ledger.
 */
function new_eer_saveReviewResponse(postId,finding,decision,reply) {
  var lock=LockService.getDocumentLock();
  if(!lock)return {success:false,message:'Document lock unavailable'};
  lock.waitLock(15000);
  try {
    var p=new_eer_getSelected_();
    if(String(p.postId)!==String(postId))throw new Error('Selection changed; nothing saved.');
    finding=String(finding||'').trim();
    decision=String(decision||'').trim();
    reply=String(reply||'').trim();
    if(!finding||['Agree','Disagree','Needs correction'].indexOf(decision)<0)
      throw new Error('Select a valid response.');
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getSheetByName('Evidence Review Responses');
    if(!sheet){
      sheet=ss.insertSheet('Evidence Review Responses');
      sheet.appendRow(['Timestamp','Post ID','Article Title','Finding','Decision','Owner Reply','Status']);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(),p.postId,p.title,finding,decision,reply,
      'REVIEW FEEDBACK — not yet integrated into W2B']);
    return {success:true,message:'Review response saved for Post '+postId};
  }catch(e){return {success:false,message:String(e)};}
  finally{lock.releaseLock();}
}
