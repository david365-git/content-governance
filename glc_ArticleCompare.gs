/** Google Love article comparison sidebar. Run glc_installCompareMenu once. */
function glc_addCompareMenu_() {
  SpreadsheetApp.getUi().createMenu('Article Comparison')
    .addItem('Open HTML Comparison Sidebar', 'glc_openCompareSidebar')
    .addToUi();
}
function glc_installCompareMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'glc_addCompareMenu_';
  }).forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('glc_addCompareMenu_').forSpreadsheet(ss).onOpen().create();
  glc_addCompareMenu_();
}
function glc_openCompareSidebar() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('glc_ArticleCompareSidebar')
    .setTitle('Google Love — Article Comparison'));
}
function glc_compareFindColumn_(sheet, heading) {
  var cols = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  var found=cols.indexOf(heading);
  if(found<0)throw new Error('Missing column: '+heading+' on '+sheet.getName());
  return found+1;
}
function glc_compareFindPostRow_(sh, col, postId) {
  var n=sh.getLastRow();
  if(n<2)throw new Error('No posts in '+sh.getName());
  var ids=sh.getRange(2,col,n-1,1).getDisplayValues();
  var row=0;
  for(var i=0;i<ids.length;i++){
    if(String(ids[i][0]).trim()===postId){if(row)throw new Error('Duplicate Post ID '+postId+' in '+sh.getName());row=i+2;}
  }
  if(!row)throw new Error('Post ID '+postId+' not found in '+sh.getName());
  return row;
}
function glc_compareLoad(postId){
  var id=String(postId||'').trim();
  if(!/^\d+$/.test(id))throw new Error('Enter a numeric Post ID.');
  var ss=SpreadsheetApp.getActiveSpreadsheet(),posts=ss.getSheetByName('posts'),site=ss.getSheetByName('site-export');
  if(!posts||!site)throw new Error('Both posts and site-export tabs are required.');
  var pId=glc_compareFindColumn_(posts,'Post ID'), pHtml=glc_compareFindColumn_(posts,'New HTML'),
      pImage=glc_compareFindColumn_(posts,'Post featured Image'), sId=glc_compareFindColumn_(site,'ID'),
      sHtml=glc_compareFindColumn_(site,'Full Post HTML');
  var pr=glc_compareFindPostRow_(posts,pId,id),sr=glc_compareFindPostRow_(site,sId,id);
  return {postId:id,postsRow:pr,sourceRow:sr,siteHtml:String(site.getRange(sr,sHtml).getValue()||''),
    newHtml:String(posts.getRange(pr,pHtml).getValue()||''),imageUrl:String(posts.getRange(pr,pImage).getValue()||'')};
}
function glc_compareSave(postId,field,value,original){
  var id=String(postId||'').trim();
  if(!/^\d+$/.test(id))throw new Error('Invalid Post ID.');
  var headings={newHtml:'New HTML',imageUrl:'Post featured Image'};
  if(!Object.prototype.hasOwnProperty.call(headings,field))throw new Error('This field cannot be changed.');
  if(typeof value!=='string'||typeof original!=='string')throw new Error('Invalid editor content.');
  if(field==='newHtml' && !value.trim())throw new Error('New HTML cannot be empty.');
  if(field==='imageUrl' && value.trim()&&!/^https:\/\//i.test(value.trim()))throw new Error('Featured image must be an HTTPS URL or empty.');
  var lock=LockService.getDocumentLock();
  if(!lock.tryLock(10000))throw new Error('Sheet busy. Please retry.');
  try{
    var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
    var r=glc_compareFindPostRow_(sh,glc_compareFindColumn_(sh,'Post ID'),id);
    var cell=sh.getRange(r,glc_compareFindColumn_(sh,headings[field]));
    if(String(cell.getValue()||'')!==original)throw new Error('This cell has changed since loading. Reload the Post ID before saving to avoid overwriting another edit.');
    cell.setValue(value);
    return {success:true,row:r,value:value,message:headings[field]+' saved for Post ID '+id+'.'};
  }finally{lock.releaseLock();}
}