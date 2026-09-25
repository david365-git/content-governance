function generateW7VideoPrompt(videoFormat) {
  videoFormat = videoFormat || 'Brief';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = (name) => headers.indexOf(name);

  const META_DESC_COL = col('New Meta Description');
  const SCHEMA_COL = col('Schema (JSON-LD)');
  const ARTICLE_TYPE_COL = col('Article Type');
  const LOCALITY_COL = col('Locality');
  const STONE_TYPE_COL = col('Stone Type');
  const MATERIAL_ENTITY_COL = col('Material Entity');

  const row = sheet.getActiveRange().getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const articleType = values[ARTICLE_TYPE_COL] || '';
  const locality = values[LOCALITY_COL] || '';
  const stoneType = values[STONE_TYPE_COL] || '';
  const materialEntity = values[MATERIAL_ENTITY_COL] || '';
  const metaDesc = values[META_DESC_COL] || '';
  const schemaRaw = values[SCHEMA_COL] || '';

  // Extract Yoast Keyphrase from end of meta description
  let yoastKeyphrase = '';
  if (metaDesc.includes('Yoast Keyphrase:')) {
    yoastKeyphrase = metaDesc.split('Yoast Keyphrase:')[1].replace('.', '').trim();
  }

  // Extract mentions terms from schema JSON-LD
  let mentionsTerms = [];
  if (schemaRaw) {
    try {
      const cleanSchema = schemaRaw.replace(/<script[^>]*>/i, '').replace(/<\/script>/i, '').trim();
      const schemaObj = JSON.parse(cleanSchema);
      const graph = schemaObj['@graph'] || [];
      for (const node of graph) {
        if (node['@type'] === 'Article' && node.mentions) {
          mentionsTerms = node.mentions.map(m => m.name).filter(Boolean);
          break;
        }
      }
    } catch (e) {
      // Schema parse failed — fallback to keyphrase
    }
  }

  // Fallback if mentions empty
  if (mentionsTerms.length === 0 && yoastKeyphrase) {
    mentionsTerms = [yoastKeyphrase];
  }

  const m1 = mentionsTerms[0] || '';
  const m2 = mentionsTerms[1] || '';
  const m3 = mentionsTerms[2] || '';

  // ── STORE VIDEO FORMAT AND THUMBNAIL ENTITIES TO SHEET ──
  const VIDEO_FORMAT_COL      = col('Video Format');
  const THUMBNAIL_ENTITY_COL  = col('Thumbnail Entities');
  if (VIDEO_FORMAT_COL > -1) {
    sheet.getRange(row, VIDEO_FORMAT_COL + 1).setValue(videoFormat);
  }
  if (THUMBNAIL_ENTITY_COL > -1) {
    sheet.getRange(row, THUMBNAIL_ENTITY_COL + 1).setValue([m1, m2, m3].filter(Boolean).join(', '));
  }

  // ── FORMAT STYLE AND LENGTH ──
  const formatLengthMap = {
    'Brief':     'Target 30–60 seconds maximum. Hook the viewer in the first 10 seconds. No padding.',
    'Explainer': 'Target 3–5 minutes. Structure clearly with an opening problem, middle explanation, and closing outcome. Cover each key concept thoroughly.',
    'Cinematic': 'Target 5–10 minutes. Open with an immersive scene-setting narrative. Build the story through visual descriptions, emotional beats, and technical revelation. Close with a satisfying outcome.'
  };

  const styleMap = {
    'Brief':     'concise, direct, no padding',
    'Explainer': 'structured, comprehensive, connecting the dots between concepts',
    'Cinematic': 'immersive, story-led, rich with visual description and emotional narrative'
  };

  const formatStyle  = styleMap[videoFormat]  || styleMap['Brief'];
  const formatLength = formatLengthMap[videoFormat] || formatLengthMap['Brief'];

  let prompt = '';

  switch (articleType) {

    case 'Case Study':
      prompt = `Tell this restoration story as two senior stone conservators. Focus on why the ${locality} ${materialEntity} problem resisted ordinary cleaning due to ${m1} and ${m2}. Emphasise Abbey Floor Care's specialist expertise — run by David Allen with over 30 years of experience — as the reason the outcome was possible. Tone: ${formatStyle}. Do not give DIY advice.`;
      break;

    case 'Method Guide':
      prompt = `Discuss this as two specialist technicians walking through the professional method for ${yoastKeyphrase}. Focus on why ordinary approaches fail with ${materialEntity} due to ${m1}, ${m2} and ${m3}. Emphasise Abbey Floor Care's hands-on specialist experience. Tone: ${formatStyle}. Do not suggest DIY alternatives.`;
      break;

    case 'Diagnostic Guide':
      prompt = `Discuss this as two specialist diagnosticians helping a homeowner identify what is wrong with their ${materialEntity}. Focus on symptoms described in the article and what ${m1}, ${m2} and ${m3} indicate. Emphasise Abbey Floor Care's diagnostic expertise and depth of experience. Tone: ${formatStyle}. Do not give DIY advice.`;
      break;

    case 'Educational Guide':
      prompt = `Discuss this as two senior conservators explaining the science behind ${yoastKeyphrase} to a curious homeowner. Focus on why ${stoneType} behaves differently and what ${m1} and ${m2} mean in practice. Emphasise Abbey Floor Care's specialist knowledge built over decades of real project work. Tone: ${formatStyle}.`;
      break;

    case 'Hub Page':
      prompt = `Discuss this as two senior stone care specialists mapping the full ${stoneType} care landscape. Focus on why ${yoastKeyphrase} requires professional understanding and what ${m1} means for homeowners. Emphasise Abbey Floor Care as the authority behind this guide. Tone: ${formatStyle}.`;
      break;

    case 'Buyer Guide':
      prompt = `Discuss this as two specialist consultants advising a homeowner about to make a costly ${stoneType} decision. Focus on what goes wrong when buyers choose wrong for ${materialEntity} and why ${m1} matters. Emphasise Abbey Floor Care's expertise as the benchmark for getting this right. Tone: ${formatStyle}. Do not give DIY advice.`;
      break;

    case 'Service Page':
      prompt = `Discuss this as two specialist advisors confirming professional capability for ${yoastKeyphrase}. Focus on what separates a ${stoneType} specialist from a general contractor. Emphasise why homeowners trust Abbey Floor Care for this type of work. Tone: ${formatStyle}.`;
      break;

    case 'Geo Service Page':
      prompt = `Discuss this as two local specialist advisors confirming professional ${stoneType} capability in ${locality}. Focus on why local knowledge and Abbey Floor Care's depth of experience matters for ${yoastKeyphrase}. Tone: ${formatStyle}. Speak to homeowners in ${locality} ready to make a decision.`;
      break;

    default:
      prompt = `Discuss this as two senior stone care specialists covering ${yoastKeyphrase}. Focus on the specialist knowledge required for ${stoneType}. Emphasise Abbey Floor Care's expertise and experience. Tone: ${formatStyle}.`;
      break;
  }

  // Trim to 500 characters hard limit for Brief only
  if (videoFormat === 'Brief' && prompt.length > 500) {
    prompt = prompt.substring(0, 497) + '...';
  }

  const format = videoFormat;
  const length = formatLength;

  // Check for existing video in HTML or YouTube URL column
  const YOUTUBE_URL_COL = col('YouTube URL');
  const HTML_COL        = col('New HTML');
  const existingUrl     = YOUTUBE_URL_COL > -1 ? values[YOUTUBE_URL_COL] : '';
  const existingHtml    = HTML_COL > -1 ? values[HTML_COL] : '';
  const hasEmbedInHtml  = existingHtml.indexOf('youtube.com/embed') > -1 || existingHtml.indexOf('youtu.be') > -1;

  let videoWarning = '';
  if (existingUrl) {
    videoWarning = '⚠ WARNING: A video already exists for this row: ' + existingUrl + '\nYou may still use this prompt to replace it.\n\n';
  } else if (hasEmbedInHtml) {
    videoWarning = '⚠ WARNING: The existing HTML already contains a YouTube embed.\nYou may still use this prompt to add a new video.\n\n';
  }

  return videoWarning + 'VIDEO FORMAT: ' + format + '\n' + length + '\n\n' + prompt;
}

function generateW7VideoFilename() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = (name) => headers.indexOf(name);

  const ARTICLE_TYPE_COL    = col('Article Type');
  const LOCALITY_COL        = col('Locality');
  const MATERIAL_ENTITY_COL = col('Material Entity');

  const row    = sheet.getActiveRange().getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const articleType    = values[ARTICLE_TYPE_COL]    || '';
  const locality       = values[LOCALITY_COL]        || '';
  const materialEntity = values[MATERIAL_ENTITY_COL] || '';

  const slugify = (str) => str.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

  const articleTypeMap = {
    'Case Study':        'case-study',
    'Method Guide':      'method-guide',
    'Diagnostic Guide':  'diagnostic-guide',
    'Educational Guide': 'educational-guide',
    'Hub Page':          'hub-page',
    'Buyer Guide':       'buyer-guide',
    'Service Page':      'service-page',
    'Geo Service Page':  'geo-service'
  };

  const typeSlug     = articleTypeMap[articleType] || slugify(articleType);
  const localitySlug = locality ? slugify(locality) + '-' : '';
  const materialSlug = slugify(materialEntity);

  return 'abbey-' + typeSlug + '-' + localitySlug + materialSlug + '-video.mp4';
  }