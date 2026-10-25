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
    'function show(r){document.getElementById("go").disabled=false;if(!r||!r.success){document.getElementById("notice").textContent=r&&r.message||"Failed";return;}document.getElementById("notice").textContent=r.questions.length+" potential question(s); nothing saved. API cost $"+Number(r.cost||0).toFixed(4);var d=document.getElementById("list");d.textContent="";r.questions.forEach(function(q,i){var x=document.createElement("div");x.className="item";var h=document.createElement("p");h.textContent=q.question;x.appendChild(h);var context=document.createElement("p");context.textContent="Why ask: "+q.reason;x.appendChild(context);var t=document.createElement("textarea");t.placeholder="Your first-hand answer (leave blank to skip)";x.appendChild(t);var b=document.createElement("button");b.textContent="Save my confirmation";b.onclick=function(){var v=t.value.trim();if(!v)return; b.disabled=true;google.script.run.withSuccessHandler(function(z){b.disabled=false;context.textContent=z&&z.message||"Saved";}).withFailureHandler(function(e){b.disabled=false;context.textContent=String(e);}).new_eer_saveConfirmation(PID,q.question,v);};x.appendChild(b);d.appendChild(x);});}' +
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
    var prompt=[
      'You are an independent evidence exception checker for an English floor-restoration writing workflow.',
      'Use only the supplied source, project evidence and governance. Source statements and generic technical advice are not independent proof.',
      'Find at most TWO genuinely material, project-specific factual uncertainties or contradictions that require first-hand owner confirmation.',
      'A question must materially change the factual accuracy of the final article; do not ask for incidental missing measurements or details.',
      'Do NOT ask for facts already directly owner-confirmed in the supplied material. If no necessary exceptions, return an empty array.',
      'Return ONLY valid JSON: {"questions":[{"question":"one concise yes/no or factual question","reason":"why source cannot establish it"}]}',
      'Never answer your own question, invent owner confirmation or treat plausible construction as a verified fact.',
      'POST ID '+expectedPostId+'; ARTICLE TYPE '+selected.articleType,
      'SOURCE AND EVIDENCE:',
      assembled.prompt
    ].join('\n');
    var model=typeof MODEL_CHEAP!=='undefined'?MODEL_CHEAP:null;
    if (!model||typeof bc_sendPromptViaOpenAI!=='function')throw new Error('OpenAI model or API sender unavailable.');
    var api=bc_sendPromptViaOpenAI(prompt,1200,model);
    if (!api||!api.success||api.truncated||api.finishReason==='length')
      return {success:false,message:'AI exception check incomplete: '+String(api&&api.message||api&&api.finishReason||'unknown')};
    var raw=String(api.text||'').trim().replace(/^```(?:json)?\s*/i,'').replace(/\s*```$/,'');
    var parsed=JSON.parse(raw);
    if (!parsed||!Array.isArray(parsed.questions))throw new Error('AI returned invalid exception JSON.');
    var questions=parsed.questions.slice(0,2).filter(function(q){
      return q&&typeof q.question==='string'&&q.question.trim()&&typeof q.reason==='string';
    }).map(function(q){return {question:q.question.slice(0,350),reason:q.reason.slice(0,500)};});
    return {success:true,questions:questions,cost:Number(api.cost||0)};
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
