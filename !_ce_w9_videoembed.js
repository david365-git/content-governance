function generateW9VideoEmbedPrompt(youtubeUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = (name) => headers.indexOf(name);

  const HTML_COL             = col('New HTML');
  const ARTICLE_TYPE_COL     = col('Article Type');
  const MATERIAL_ENTITY_COL  = col('Material Entity');
  const STONE_TYPE_COL       = col('Stone Type');
  const LOCALITY_COL         = col('Locality');
  const H1_COL               = col('New H1');
  const YOUTUBE_URL_COL      = col('YouTube URL');
  const YOUTUBE_ID_COL       = col('YouTube Video ID');

  const row    = sheet.getActiveRange().getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const existingHtml   = values[HTML_COL]             || '';
  const articleType    = values[ARTICLE_TYPE_COL]     || '';
  const materialEntity = values[MATERIAL_ENTITY_COL]  || '';
  const stoneType      = values[STONE_TYPE_COL]       || '';
  const locality       = values[LOCALITY_COL]         || '';
  const h1             = values[H1_COL]               || '';

  if (!youtubeUrl) {
    return 'ERROR: No YouTube URL provided.';
  }

  // Extract Video ID from YouTube URL
  let videoId = '';
  const match = youtubeUrl.match(/(?:v=|youtu\.be\/)([a-zA-Z0-9_-]{11})/);
  if (match) {
    videoId = match[1];
  } else {
    return 'ERROR: Could not extract Video ID from URL. Please check the URL and try again.';
  }

  const embedUrl = 'https://www.youtube.com/embed/' + videoId;

  // Write YouTube URL and Video ID back to posts sheet
  if (YOUTUBE_URL_COL > -1) {
    sheet.getRange(row, YOUTUBE_URL_COL + 1).setValue(youtubeUrl);
  }
  if (YOUTUBE_ID_COL > -1) {
    sheet.getRange(row, YOUTUBE_ID_COL + 1).setValue(videoId);
  }

  // Build the embed block instruction
  const embedBlock =
    '<figure class="video-embed" style="margin:24px 0;text-align:center">' +
    '<iframe width="560" height="315" src="' + embedUrl + '" ' +
    'title="' + h1 + '" frameborder="0" ' +
    'allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" ' +
    'allowfullscreen></iframe>' +
    '<figcaption>Video overview of the ' + (locality ? locality + ' ' : '') + materialEntity + ' restoration project.</figcaption>' +
    '</figure>';

  // Extract header block only — send only what is being changed
  var headerBlock = '';
  var headerMatch = existingHtml.match(/<header[\s\S]*?<\/header>/i);
  if (headerMatch) {
    headerBlock = headerMatch[0];
  } else {
    // Fallback — no header tag found, send first 500 chars as context
    headerBlock = existingHtml.substring(0, 500);
  }

  const prompt =
    'Return directly in chat. Do not open canvas or document mode.\n\n' +
    'You are updating the header block of an existing article. Insert the following video embed block immediately after the closing </p> tag of the first paragraph inside the <header> section.\n\n' +
    'VIDEO EMBED BLOCK TO INSERT:\n' + embedBlock + '\n\n' +
    'After inserting the embed block, add one short supporting sentence immediately after it, inside a <p> tag, that:\n' +
    '- Tells the reader the video gives a quick overview of the ' + (locality ? locality + ' ' : '') + materialEntity + ' restoration\n' +
    '- Encourages them to read the full article below for the complete detail\n' +
    '- Is written in plain, empathetic language suited to a homeowner audience\n' +
    '- Is no longer than 25 words\n\n' +
    'Return the updated <header> block only — nothing before it, nothing after it.\n' +
    'Your first token must be <header> and your last token must be </header>.\n' +
    'No explanation. No preamble. No markdown code fences. No surrounding HTML.\n\n' +
    'CRITICAL — DO NOT remove, alter, reorder, or rewrite any existing element inside the header. Only insert the video embed block and the one supporting sentence.\n\n' +
    'EXISTING HEADER BLOCK:\n' + headerBlock;

  return prompt;
}
function saveVideoEmbedHeaderToSheet(updatedHeader) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('posts');
    var row   = sheet.getActiveRange().getRow();
    if (row < 2) return { success: false, message: 'Select a data row first.' };

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var htmlIdx = headers.indexOf('New HTML');
    if (htmlIdx === -1) return { success: false, message: 'New HTML column not found.' };

    var existingHtml = String(sheet.getRange(row, htmlIdx + 1).getValue() || '').trim();
    if (!existingHtml) return { success: false, message: 'No existing HTML found in New HTML column for this row.' };

    // Validate the pasted block starts and ends with header tags
    var trimmed = updatedHeader.trim();
    if (!trimmed.match(/^<header/i)) return { success: false, message: 'Pasted block does not start with <header> — check the LLM output and try again.' };
    if (!trimmed.match(/<\/header>\s*$/i)) return { success: false, message: 'Pasted block does not end with </header> — check the LLM output and try again.' };

    // Replace existing header in full article with updated header
    var updatedHtml = existingHtml.replace(/<header[\s\S]*?<\/header>/i, trimmed);

    if (updatedHtml === existingHtml) {
      return { success: false, message: 'No header block found in existing HTML — replacement could not be made.' };
    }

    var cell = sheet.getRange(row, htmlIdx + 1);
    cell.setNumberFormat('@');
    cell.setValue(updatedHtml);

    return { success: true, message: 'Video embed header saved to row ' + row + '. Run W5C to rebuild schema.' };

  } catch(e) {
    return { success: false, message: 'saveVideoEmbedHeaderToSheet error: ' + e.message };
  }
}
