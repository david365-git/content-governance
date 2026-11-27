/** Add to Content Engine project (NOT the Social Manager Apps Script project).
 * Explicitly sends the selected posts-sheet row to Social Content Manager.
 * Uses direct spreadsheet access; no public webhook or API key required.
 */
function ce_sendActiveArticleToSocialManager() {
  const source=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('posts');
  if(!source)throw new Error('The posts sheet was not found.');
  const selected=source.getActiveCell().getRow();
  if(selected<2)throw new Error('Select an article row first.');
  const headings=source.getRange(1,1,1,source.getLastColumn()).getValues()[0]
    .map(h=>String(h).trim().toLowerCase());
  const row=source.getRange(selected,1,1,source.getLastColumn()).getValues()[0];
  const byName=(...candidates)=>{
    for(const name of candidates){const index=headings.indexOf(name.toLowerCase());if(index>=0&&row[index]!==''&&row[index]!=null)return row[index];}
    return '';
  };
  const postId=String(byName('Post ID')||row[3]||'').trim();
  const html=String(byName('New HTML','Finished HTML')||row[97]||'');
  if(!/^\d+$/.test(postId)||!html.trim())throw new Error('Post ID or New HTML missing; nothing sent.');
  const target=SpreadsheetApp.openById('1XQbRXDfesWPEKY_EmD8YXj0M5CVtgP-C5OH2bAx3ucE');
  const s=target.getSheetByName('Article Library');
  if(!s)throw new Error('Social Manager Article Library sheet missing.');
  const headers=s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(String);
  const vals=s.getDataRange().getValues();
  let found=0;
  for(let i=1;i<vals.length;i++)if(String(vals[i][0]).trim()===postId){found=i+1;break;}
  const payload={
    'Article Post ID':postId,'Title':String(byName('Title','Post Title')||row[1]||''),
    'Article Type':String(byName('Article Type')||row[7]||''),
    'Stone Type':String(byName('Stone Type')||row[6]||''),
    'Locality':String(byName('Locality')||row[114]||''),
    'Parent Area':String(byName('Parent Area')||row[115]||''),
    'Article URL':String(byName('Article URL','Post URL','Permalink','URL')||''),
    'Finished HTML':html,
    'Source Version':'posts row '+selected,
    'Received UTC':new Date().toISOString(),
    'Status':'RECEIVED'
  };
  if(found){
    headers.forEach((h,i)=>{if(Object.prototype.hasOwnProperty.call(payload,h))s.getRange(found,i+1).setValue(payload[h]);});
  }else{s.appendRow(headers.map(h=>payload[h]===undefined?'':payload[h]));}
  SpreadsheetApp.getActive().toast('Post '+postId+' sent to Social Content Manager ('+(found?'updated':'new')+').');
  return {postId:postId,action:found?'UPDATED':'CREATED'};
}
