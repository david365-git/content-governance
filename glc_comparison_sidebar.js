/**
 * Google Love model comparison — isolated pilot, no production writes.
 * Open by running glc_openComparisonSidebar or using Row Analysis menu.
 */
function glc_openComparisonSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('glc_ComparisonSidebar')
      .setTitle('Google Love — Model Comparison')
      .setWidth(420)
  );
}
function glc_getArticle_(postId) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('posts');
  if(!sh) throw new Error('posts sheet not found.');
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  function ix(name){var n=headers.indexOf(name);if(n<0)throw new Error('Missing posts column: '+name);return n;}
  var idCol=ix('Post ID'), titleCol=ix('Title'), materialCol=ix('Stone Type'),typeCol=ix('Article Type'), urlCol=ix('URL');
  var rows=sh.getRange(2,1,Math.max(1,sh.getLastRow()-1),sh.getLastColumn()).getValues();
  var matches=rows.map(function(r,i){return {r:r,row:i+2};}).filter(function(x){return String(x.r[idCol]).trim()===String(postId).trim();});
  if(matches.length!==1)throw new Error('Expected one posts row for Post ID '+postId+'; found '+matches.length+'.');
  var rec=matches[0],original='',source='not available';
  var exp=ss.getSheetByName('site-export');
  if(exp && exp.getLastRow()>1) {
    var n=exp.getLastRow()-1,ids=exp.getRange(2,1,n,1).getDisplayValues();
    var hits=[];
    ids.forEach(function(x,i){if(String(x[0]).trim()===String(postId).trim())hits.push(i+2);});
    if(hits.length===1){original=String(exp.getRange(hits[0],15).getValue()||'');source='site-export!O'+hits[0];}
    else source='site-export match count: '+hits.length;
  }
  var technical=ss.getSheetByName('technical'),matrix='',technicalSource='not matched';
  if(technical){
    var mRows=technical.getLastRow(), mCols=technical.getLastColumn();
    // Inspect material labels in column A, never use a cached matrix output.
    var matVals=technical.getRange(2,1,Math.max(1,mRows-1),1).getDisplayValues();
    var found=[];
    matVals.forEach(function(x,i){if(String(x[0]).trim().toLowerCase()===String(rec.r[materialCol]).trim().toLowerCase())found.push(i+2);});
    if(found.length===1&&mCols>=19){
      matrix=technical.getRange(found[0],14,1,6).getDisplayValues()[0].join('\n\n--- APPROVED NEXT COLUMN ---\n\n');
      technicalSource='technical!N'+found[0]+':S'+found[0];
    } else technicalSource='technical match count: '+found.length;
  }
  var result={postId:String(postId),row:rec.row,title:String(rec.r[titleCol]||''),material:String(rec.r[materialCol]||''),type:String(rec.r[typeCol]||''),url:String(rec.r[urlCol]||''),original:original,originalSource:source,technical:matrix,technicalSource:technicalSource};
  return result;
}
function glc_loadArticle(postId){
  var a=glc_getArticle_(postId);
  return {postId:a.postId,row:a.row,title:a.title,material:a.material,type:a.type,url:a.url,originalSource:a.originalSource,originalCharacters:a.original.length,technicalSource:a.technicalSource,technicalCharacters:a.technical.length,images:(a.original.match(/<img\b/gi)||[]).length};
}
function glc_saveComparison(payload){
  payload=payload||{};
  var id=String(payload.postId||'').trim();
  if(!/^\d+$/.test(id))throw new Error('Enter a numeric Post ID.');
  var a=glc_getArticle_(id),ss=SpreadsheetApp.getActiveSpreadsheet();
  var sh=ss.getSheetByName('GL Model Comparison Tests');
  if(!sh){
    sh=ss.insertSheet('GL Model Comparison Tests');
    sh.appendRow(['Recorded UTC','Post ID','Title','Material','Type','Model / route','Prompt version','Run status','HTML or result','Notes','Original source','Original chars','Technical source','Technical chars','Image count','Source URL']);
    sh.setFrozenRows(1);
  }
  var output=String(payload.output||''),notes=String(payload.notes||''),route=String(payload.route||'');
  if(!['ChatGPT conversation','API — Luna','API — stronger model'].includes(route))throw new Error('Select a comparison route.');
  if(!output.trim())throw new Error('Paste a test result before saving.');
  if(output.length>49000)throw new Error('Result exceeds safe single-cell limit. Save a shortened result or store full HTML elsewhere.');
  var stamp=Utilities.formatDate(new Date(),'Etc/UTC',"yyyy-MM-dd'T'HH:mm:ss'Z'");
  sh.appendRow([stamp,a.postId,a.title,a.material,a.type,route,String(payload.promptVersion||''),'CAPTURED — NOT AUDITED',output,notes,a.originalSource,a.original.length,a.technicalSource,a.technical.length,(a.original.match(/<img\b/gi)||[]).length,a.url]);
  return {message:'Saved isolated comparison record to GL Model Comparison Tests, row '+sh.getLastRow(),row:sh.getLastRow()};
}
