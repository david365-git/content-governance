/** Pinterest board mapping and evidence-based manual Pin captions. Never publishes automatically. */
const SCM_PINTEREST_BOARDS = [
 ['Limestone Floor Cleaning & Restoration','limestone'],
 ['Travertine Floor Cleaning & Restoration','travertine'],
 ['Marble Floor Polishing & Restoration','marble'],
 ['Terrazzo Floor Restoration','terrazzo'],
 ['Victorian & Encaustic Tile Restoration','victorian|encaustic|minton|edwardian'],
 ['Quarry Tile & Terracotta Restoration','quarry|terracotta'],
 ['Slate Floor Cleaning & Restoration','slate'],
 ['Porcelain & Ceramic Tile Cleaning','porcelain|ceramic'],
 ['Tile Floor Cleaning & Restoration','general|tile'],
 ['Stone Floor Care & Maintenance Tips','care'],
 ['Stone Restoration Videos','restoration video'],
 ['Stone Care Videos','care video']
];
function scm_pinterestSetupBoards() {
 const book=scm_book_();
 let sheet=book.getSheetByName('Pinterest Boards');
 if(!sheet) {
   sheet=book.insertSheet('Pinterest Boards');
   sheet.appendRow(['Board Name','Stone Type Keywords','Pinterest Board ID (optional)']);
   sheet.setFrozenRows(1);
   sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#e8f0e9');
   sheet.setColumnWidth(1,340);
   sheet.setColumnWidth(2,280);
 }
 const actual=sheet.getRange(1,1,1,3).getValues()[0].map(String);
 if(actual.join('|')!=='Board Name|Stone Type Keywords|Pinterest Board ID (optional)')
   throw new Error('Unexpected Pinterest Boards headers. Existing data preserved.');
 const existing=new Set(sheet.getDataRange().getValues().slice(1).map(r=>String(r[0]).trim().toLowerCase()));
 SCM_PINTEREST_BOARDS.forEach(row=>{if(!existing.has(row[0].toLowerCase()))sheet.appendRow([row[0],row[1],'']);});
 return {boards:sheet.getLastRow()-1};
}
function scm_pinterestArticle_(p) {
 const row=scm_rows_(SCM.articles).rows.find(r=>String(r.item['Article Post ID'])===String(p['Article Post ID']));
 return row?row.item:null;
}
function scm_pinterestImage_(p) {
 const row=scm_rows_(SCM.images).rows.find(r=>String(r.item['Opportunity ID'])===String(p['Opportunity ID']));
 return row?row.item:null;
}
function scm_pinterestStone_(article,image) {
 const fields=['Stone Type','Material','Stone','Stone / Tile Type','Primary Stone Type'];
 for(const field of fields) {
   const value=String(article?.[field]||'').trim();
   if(value)return value.toLowerCase();
 }
 const title=String(article?.Title||'');
 const img=String(image?.['Original Alt / Caption']||'');
 return (title+' '+img).toLowerCase();
}
function scm_pinterestBoardFor_(article,image) {
 const source=scm_pinterestStone_(article,image);
 const sheet=scm_book_().getSheetByName('Pinterest Boards');
 const rows=sheet?sheet.getDataRange().getValues().slice(1):SCM_PINTEREST_BOARDS.map(r=>[r[0],r[1]]);
 for(const row of rows) {
   const board=String(row[0]||'').trim();
   if(!board || board==='Tile Floor Cleaning & Restoration' ||
     /stone (restoration videos|care videos|floor care)/i.test(board))continue;
   const keywords=String(row[1]||'').toLowerCase().split('|').map(x=>x.trim()).filter(Boolean);
   if(keywords.some(word=>new RegExp('\\b'+word.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'\\b','i').test(source)))return board;
 }
 return 'Tile Floor Cleaning & Restoration';
}
function scm_pinterestCaption_(p,article,image) {
 const title=String(article?.Title||'Stone and tile floor restoration').trim();
 const angle=String(image?.['Distinct Post Angle']||image?.['Original Alt / Caption']||'').replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim();
 const detail=angle&&angle.toLowerCase()!==title.toLowerCase()?angle:title;
 const url=String(p['Article URL']||'').trim();
 return [detail.slice(0,220),'See the full Abbey Floor Care article and project photographs:',url].join('\n\n');
}

function scm_pinterestTitle_(article,image) {
 const angle=String(image?.['Distinct Post Angle']||image?.['Original Alt / Caption']||'').replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim();
 const title=angle||String(article?.Title||'Stone floor restoration');
 return title.slice(0,100);
}
