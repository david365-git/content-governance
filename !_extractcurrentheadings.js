function extractCurrentHeadingsToPosts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postsSheet = ss.getSheetByName('posts');
  const exportSheet = ss.getSheetByName('site-export');

  if (!postsSheet) throw new Error("Sheet 'posts' not found.");
  if (!exportSheet) throw new Error("Sheet 'site-export' not found.");

  const postsData = postsSheet.getDataRange().getValues();
  const exportData = exportSheet.getDataRange().getValues();

  if (postsData.length < 2) {
    throw new Error("The 'posts' sheet has no data rows.");
  }
  if (exportData.length < 2) {
    throw new Error("The 'site-export' sheet has no data rows.");
  }

  const postsHeaders = postsData[0];
  const exportHeaders = exportData[0];

  const postIdColPosts = postsHeaders.indexOf('Post ID');
  const currentHeadingsColPosts = postsHeaders.indexOf('Current Headings');
  const idColExport = exportHeaders.indexOf('ID');
  const fullPostHtmlColExport = exportHeaders.indexOf('Full Post HTML');

  if (postIdColPosts === -1) {
    throw new Error("Column 'Post ID' not found in 'posts' sheet.");
  }
  if (currentHeadingsColPosts === -1) {
    throw new Error("Column 'Current Headings' not found in 'posts' sheet.");
  }
  if (idColExport === -1) {
    throw new Error("Column 'ID' not found in 'site-export' sheet.");
  }
  if (fullPostHtmlColExport === -1) {
    throw new Error("Column 'Full Post HTML' not found in 'site-export' sheet.");
  }

  // Build lookup from site-export ID -> Full Post HTML
  const exportMap = new Map();
  for (let i = 1; i < exportData.length; i++) {
    const id = String(exportData[i][idColExport]).trim();
    const html = exportData[i][fullPostHtmlColExport];
    if (id) {
      exportMap.set(id, html);
    }
  }

  const output = [];

  for (let i = 1; i < postsData.length; i++) {
    const postId = String(postsData[i][postIdColPosts]).trim();

    if (!postId || !exportMap.has(postId)) {
      output.push(['']);
      continue;
    }

    const html = String(exportMap.get(postId) || '');
    const headings = extractHeadingsFromHtml(html);
    output.push([headings]);
  }

  postsSheet
    .getRange(2, currentHeadingsColPosts + 1, output.length, 1)
    .setValues(output);

  SpreadsheetApp.getUi().alert('Current headings extracted successfully.');
}

function extractHeadingsFromHtml(html) {
  if (!html) return '';

  const matches = html.match(/<h[234]\b[^>]*>[\s\S]*?<\/h[234]>/gi);
  if (!matches) return '';

  return matches
    .map(tag => cleanHeadingTag(tag))
    .join('\n');
}

function cleanHeadingTag(tag) {
  return tag
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/>\s+</g, '><')
    .trim();
}