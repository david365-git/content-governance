/**
 * ce_CaseStudySampler.gs - EXPORT 5 CASE STUDIES PER STONE TYPE
 * Scans the "posts" sheet for rows where Article Type = Case Study,
 * groups by Stone Type (column G), takes up to 5 per stone type, and
 * writes their content to a single text file in Google Drive.
 * Falls back to "site-export" sheet's "Full Post HTML" (matched by Post ID)
 * if the row's New HTML column is empty.
 */

function exportCaseStudySamplesToFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  if (!sheet) { return { success: false, message: "posts sheet not found." }; }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function(h) { return String(h).trim(); });

  const colArticleType = headers.indexOf('Article Type');
  const colStoneType = headers.indexOf('Stone Type');
  const colNewHtml = headers.indexOf('New HTML');
  const colPostId = headers.indexOf('Post ID');
  const colTitle = 1; // column B

  if (colArticleType === -1 || colStoneType === -1 || colNewHtml === -1 || colPostId === -1) {
    return { success: false, message: "Could not find one of: Article Type, Stone Type, New HTML, Post ID columns." };
  }

  // Group rows by stone type, only Article Type = Case Study
  const byStoneType = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const articleType = String(row[colArticleType] || '').trim();
    if (articleType !== 'Case Study') continue;

    const stoneType = String(row[colStoneType] || '').trim();
    if (!stoneType) continue;

    if (!byStoneType[stoneType]) byStoneType[stoneType] = [];
    if (byStoneType[stoneType].length < 5) {
      byStoneType[stoneType].push({
        rowNum: i + 1,
        title: String(row[colTitle] || '').trim(),
        postId: String(row[colPostId] || '').trim().replace(/\.0$/, ''),
        newHtml: String(row[colNewHtml] || '').trim()
      });
    }
  }

  // Prepare site-export lookup (only if needed)
  let exportData = null;
  let exportIdCol = -1;
  let exportHtmlCol = -1;

  function getExportHtmlForPostId(postId) {
    if (!postId) return '';
    if (exportData === null) {
      const exportSheet = ss.getSheetByName('site-export');
      if (!exportSheet) { exportData = []; return ''; }
      exportData = exportSheet.getDataRange().getValues();
      const exportHeaders = exportData[0].map(function(h) { return String(h).trim(); });
      exportIdCol = exportHeaders.indexOf('ID');
      exportHtmlCol = exportHeaders.indexOf('Full Post HTML');
    }
    if (exportIdCol === -1 || exportHtmlCol === -1) return '';
    for (let j = 1; j < exportData.length; j++) {
      if (String(exportData[j][exportIdCol]).trim().replace(/\.0$/, '') === postId) {
        return String(exportData[j][exportHtmlCol] || '').trim();
      }
    }
    return '';
  }

  // Build the output text
  const lines = [];
  const stoneTypes = Object.keys(byStoneType).sort();

  stoneTypes.forEach(function(stoneType) {
    lines.push('='.repeat(70));
    lines.push('STONE TYPE: ' + stoneType);
    lines.push('='.repeat(70));

    byStoneType[stoneType].forEach(function(item, idx) {
      let html = item.newHtml;
      let source = 'New HTML';
      if (!html) {
        html = getExportHtmlForPostId(item.postId);
        source = 'site-export Full Post HTML (fallback — New HTML was empty)';
      }
      if (!html) {
        html = '[NO CONTENT FOUND — New HTML empty and no match in site-export for Post ID: ' + item.postId + ']';
      }

      lines.push('');
      lines.push('-'.repeat(70));
      lines.push('CASE STUDY ' + (idx + 1) + ' OF ' + byStoneType[stoneType].length + ' — ' + stoneType);
      lines.push('Row: ' + item.rowNum + ' | Post ID: ' + item.postId + ' | Title: ' + item.title);
      lines.push('Source: ' + source);
      lines.push('-'.repeat(70));
      lines.push(html);
    });

    lines.push('');
  });

  const output = lines.join('\n');

  const fileName = 'CaseStudySamples_' + new Date().toISOString().slice(0, 10) + '.txt';
  const file = DriveApp.createFile(fileName, output, MimeType.PLAIN_TEXT);

  return {
    success: true,
    message: 'Exported ' + stoneTypes.length + ' stone type(s), ' +
      Object.values(byStoneType).reduce(function(sum, arr) { return sum + arr.length; }, 0) +
      ' case stud(y/ies) total. File: ' + file.getName(),
    fileUrl: file.getUrl()
  };
}
