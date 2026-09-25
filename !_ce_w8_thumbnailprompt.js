function generateW8ThumbnailPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = (name) => headers.indexOf(name);

  const ARTICLE_TYPE_COL = col('Article Type');
  const LOCALITY_COL     = col('Locality');
  const STONE_TYPE_COL   = col('Stone Type');
  const MATERIAL_ENTITY_COL = col('Material Entity');

  const row    = sheet.getActiveRange().getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const articleType    = values[ARTICLE_TYPE_COL] || '';
  const locality       = values[LOCALITY_COL]     || '';
  const stoneType      = values[STONE_TYPE_COL]   || '';
  const materialEntity = values[MATERIAL_ENTITY_COL] || '';

  // Overlay text by article type
  const overlayMap = {
    'Case Study':        'This ' + locality + ' ' + materialEntity + ' floor looked beyond saving...',
    'Method Guide':      'Why your ' + materialEntity + ' keeps looking dirty',
    'Diagnostic Guide':  'Is this what\'s wrong with your ' + stoneType + ' floor?',
    'Educational Guide': 'What most people get wrong about ' + stoneType,
    'Hub Page':          'Everything about ' + stoneType + ' care in one place',
    'Buyer Guide':       'Don\'t choose the wrong ' + stoneType + ' specialist',
    'Service Page':      'Abbey Floor Care — 30 years of specialist results',
    'Geo Service Page':  stoneType + ' specialist in ' + locality
  };

  const overlayText = overlayMap[articleType] || 'Don\'t struggle with your ' + stoneType + ' floor';

  // Emotional scene array — any scene works for any article type
  const scenes = [
    'A cartoon person with jaw dropped, eyes wide, hand over mouth, expression of pure shock. Electric yellow background.',
    'A cartoon person mid-jump punching the air with a huge grin, arms raised in triumph. Deep teal background.',
    'A cartoon person holding their head with both hands, comic-style sweat drops flying, expression of total overwhelm. Vivid orange background.',
    'A cartoon person pointing directly at the viewer with a knowing smile and one eyebrow raised. Purple background.',
    'A cartoon person frozen like a statue with eyes shut tight and hands over ears, refusing to deal with something. Bold red background.',
    'A cartoon person with a giant lightbulb above their head, eyes wide and lit up, finger pointing in the air in a eureka moment. Cobalt blue background.',
    'A cartoon person dramatically falling backwards off a chair in complete disbelief, arms flailing. Neon green background.',
    'A cartoon person standing with arms folded and a smug satisfied smile, one eyebrow raised in confidence. Hot pink background.',
    'A cartoon person sprinting away from a giant cartoon storm cloud chasing them, looking back in panic. Vivid yellow background.',
    'A cartoon person holding a giant magnifying glass zoomed in on their own confused face, utterly baffled. Bright orange background.',
    'A cartoon person holding a giant question mark above their head, shrugging with a completely baffled expression. Deep red background.',
    'A cartoon person standing in a dramatic spotlight with arms wide open and face beaming, pure relief and joy. Teal background.'
  ];

  // Pick a scene based on row number for variety across posts
  const scene = scenes[row % scenes.length];

  const prompt = 'Create a YouTube thumbnail image at 1280x720 pixels, 16:9 aspect ratio.\n\n' +
    'STYLE: Bold cartoon illustration with exaggerated expressions, ultra-vivid high contrast colours, thick black outlines, comic book energy. No photorealism. No floor imagery.\n\n' +
    'SCENE: ' + scene + '\n\n' +
    'TEXT OVERLAY: Add bold text in a large thick font reading: "' + overlayText + '" — text must be clearly readable at small sizes, high contrast white or yellow with thick black outline or drop shadow. Place text so it does not obscure the face.\n\n' +
    'TECHNICAL: 1280x720px, 16:9, no watermarks, no borders, no extra elements, clean and impactful.';

  return prompt;
}
function generateW8YouTubeDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = (name) => headers.indexOf(name);

  const ARTICLE_TYPE_COL     = col('Article Type');
  const LOCALITY_COL         = col('Locality');
  const STONE_TYPE_COL       = col('Stone Type');
  const MATERIAL_ENTITY_COL  = col('Material Entity');
  const H1_COL               = col('New H1');
  const URL_COL              = col('URL');
  const META_DESC_COL        = col('New Meta Description');

  const row    = sheet.getActiveRange().getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const articleType    = values[ARTICLE_TYPE_COL]    || '';
  const locality       = values[LOCALITY_COL]        || '';
  const stoneType      = values[STONE_TYPE_COL]      || '';
  const materialEntity = values[MATERIAL_ENTITY_COL] || '';
  const h1             = values[H1_COL]              || '';
  const url            = values[URL_COL]             || '';
  const metaDesc       = values[META_DESC_COL]       || '';

  // Extract Yoast Keyphrase
  let keyphrase = '';
  if (metaDesc.includes('Yoast Keyphrase:')) {
    keyphrase = metaDesc.split('Yoast Keyphrase:')[1].replace('.', '').trim();
  }

  // Video title — H1 trimmed to 100 characters
  const videoTitle = h1.length > 100 ? h1.substring(0, 97) + '...' : h1;

  // Description
  const descriptionMap = {
    'Case Study':        'See how this ' + locality + ' ' + materialEntity + ' floor was brought back from the brink. David Allen of Abbey Floor Care shares what was wrong, why ordinary cleaning failed, and how specialist restoration recovered the original surface.',
    'Method Guide':      'Discover why ordinary cleaning fails on ' + materialEntity + ' and what the professional method actually involves. David Allen of Abbey Floor Care explains the specialist approach behind lasting results.',
    'Diagnostic Guide':  'Not sure what is wrong with your ' + stoneType + ' floor? David Allen of Abbey Floor Care walks through the key symptoms and what they mean for your restoration options.',
    'Educational Guide': 'Understanding ' + stoneType + ' is the first step to protecting it. David Allen of Abbey Floor Care explains what most homeowners get wrong and why specialist knowledge makes the difference.',
    'Hub Page':          'Everything you need to know about ' + stoneType + ' care in one place. David Allen of Abbey Floor Care covers cleaning, restoration, sealing and long term protection.',
    'Buyer Guide':       'Choosing the right ' + stoneType + ' specialist could save you thousands. David Allen of Abbey Floor Care explains what to look for and what questions to ask before you commit.',
    'Service Page':      'Abbey Floor Care has been restoring ' + stoneType + ' floors for over 30 years. Find out what sets a genuine specialist apart from a general contractor.',
    'Geo Service Page':  'Looking for a ' + stoneType + ' specialist in ' + locality + '? David Allen of Abbey Floor Care covers what local homeowners need to know before booking a restoration.'
  };

  const description = (descriptionMap[articleType] || 'David Allen of Abbey Floor Care shares specialist knowledge on ' + stoneType + ' care and restoration.') +
    '\n\nAbbey Floor Care specialises in ' + stoneType + ' restoration, stone floor cleaning and period property floor care across the Midlands and nationally. David Allen has over 30 years of specialist experience.' +
    '\n\nRead the full article: ' + url +
    '\n\nContact Abbey Floor Care for a no-obligation assessment: https://www.abbeyfloorcare.co.uk/contact';

  // Hashtags
  const cleanMaterial = materialEntity.replace(/\s+/g, '');
  const cleanStone    = stoneType.replace(/\s+/g, '');
  const hashtags = '#' + cleanMaterial + ' #' + cleanStone + 'Restoration #AbbeyFloorCare';

  // YouTube tags
  const tags = [
    materialEntity,
    stoneType,
    stoneType + ' restoration',
    stoneType + ' cleaning',
    stoneType + ' specialist',
    materialEntity + ' restoration',
    keyphrase,
    locality ? locality + ' floor restoration' : '',
    locality ? stoneType + ' ' + locality : '',
    'Abbey Floor Care',
    'David Allen',
    'period property floors',
    'professional floor restoration',
    articleType.toLowerCase()
  ].filter(Boolean).filter(function(tag, index, self) {
    return self.indexOf(tag) === index;
  }).join(', ');

  const output =
    '── VIDEO TITLE ──\n' + videoTitle +
    '\n\n── DESCRIPTION ──\n' + description +
    '\n\n── HASHTAGS ──\n' + hashtags +
    '\n\n── YOUTUBE TAGS ──\n' + tags;

  return output;
}
function generateW8Filename() {
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

  return 'abbey-' + typeSlug + '-' + localitySlug + materialSlug + '-thumb.jpg';
}

function generateW8ThumbnailPromptV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('posts');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = (name) => headers.indexOf(name);

  const ARTICLE_TYPE_COL    = col('Article Type');
  const LOCALITY_COL        = col('Locality');
  const STONE_TYPE_COL      = col('Stone Type');
  const MATERIAL_ENTITY_COL = col('Material Entity');
  const H1_COL              = col('New H1');
  const W15D_COL            = col('W1.5D Final Plan');

  const row    = sheet.getActiveRange().getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const articleType    = values[ARTICLE_TYPE_COL]    || '';
  const locality       = values[LOCALITY_COL]        || '';
  const stoneType      = values[STONE_TYPE_COL]      || '';
  const materialEntity = values[MATERIAL_ENTITY_COL] || '';
  const h1             = values[H1_COL]              || '';
  const w15d           = values[W15D_COL]            || '';

  // ── HEADLINE MAP ──
  const headlineMap = {
    'Case Study': [
      'THIS ' + (locality || stoneType).toUpperCase() + ' FLOOR',
      materialEntity.toUpperCase(),
      'LOOKED BEYOND SAVING'
    ],
    'Method Guide': [
      'WHY YOUR',
      materialEntity.toUpperCase(),
      "WON'T COME CLEAN"
    ],
    'Diagnostic Guide': [
      'IS THIS',
      "WHAT'S WRONG",
      'WITH YOUR ' + stoneType.toUpperCase() + '?'
    ],
    'Educational Guide': [
      'WHAT MOST',
      'PEOPLE GET WRONG',
      'ABOUT ' + stoneType.toUpperCase()
    ],
    'Hub Page': [
      'EVERYTHING ABOUT',
      stoneType.toUpperCase(),
      'CARE EXPLAINED'
    ],
    'Buyer Guide': [
      "DON'T CHOOSE",
      'THE WRONG',
      stoneType.toUpperCase() + ' SPECIALIST'
    ],
    'Service Page': [
      'ABBEY FLOOR CARE',
      '30 YEARS OF',
      'SPECIALIST RESULTS'
    ],
    'Geo Service Page': [
      stoneType.toUpperCase(),
      'SPECIALIST IN',
      (locality || 'YOUR AREA').toUpperCase()
    ]
  };

  const headlineLines = headlineMap[articleType] || [
    "DON'T STRUGGLE",
    'WITH YOUR',
    stoneType.toUpperCase() + ' FLOOR'
  ];

  // ── EXTRACT SECTION 1 BRIEF FROM W1.5D ──
  let section1Brief = '';
  const sec1Match = w15d.match(/SECTION 1:.*?Content Brief:\s*(.*?)(?=Visual Pattern:|$)/s);
  if (sec1Match) section1Brief = sec1Match[1].trim();

  // ── CHARACTER EXPRESSION OVERRIDE ──
  let expressionOverride = '';
  if (section1Brief) {
    if (/beyond saving|beyond recovery|patchy|worn|unstable|failed|loose|deposits/i.test(section1Brief)) {
      expressionOverride = 'Standing back slightly, arms loosely raised at her sides, looking down at the floor with exhausted disbelief — someone who has tried everything and cannot understand why nothing has worked. Eyes wide, mouth slightly open. Thick black outlines. Exaggerated wide eyes. Comic-book energy. Dramatic rim lighting in orange and gold.';
    } else if (/damage|permanent|aggressive|unsafe|risk/i.test(section1Brief)) {
      expressionOverride = 'Stepping back in alarm, one hand raised as if to say stop, eyes wide with urgent warning. Expression of someone who has just realised something is about to go badly wrong. Thick black outlines. Exaggerated expression. Comic-book energy. Dramatic rim lighting in orange and gold.';
    } else if (/expect|realistic|outcome|another generation|life/i.test(section1Brief)) {
      expressionOverride = 'Standing with a cautious but relieved smile, one hand resting on her chin thoughtfully. Expression of someone pleasantly surprised by a better outcome than expected. Thick black outlines. Exaggerated expression. Comic-book energy. Dramatic rim lighting in orange and gold.';
    }
  }

  // ── CHARACTER MAP ──
  const characterMap = {
    'Case Study':        'A cartoon British female homeowner — middle-aged, neatly dressed in a practical jumper. Expression: jaw dropped, eyes wide in total disbelief at what has been revealed. One hand pointing down at the floor in amazement. Thick black outlines. Exaggerated wide eyes. Comic-book energy. Dramatic rim lighting in orange and gold.',
    'Method Guide':      'A cartoon British female homeowner — middle-aged, neatly dressed in a practical jumper. Expression: genuine bafflement mixed with growing frustration. Holding a mop that has clearly made no difference whatsoever. The mop head is visibly grey and worn. Thick black outlines. Exaggerated wide eyes. Comic-book energy. Dramatic rim lighting in orange and gold.',
    'Diagnostic Guide':  'A cartoon British female homeowner — middle-aged, wearing reading glasses pushed up on her forehead. Expression: squinting and puzzled, leaning forward with a giant magnifying glass examining something closely. Thick black outlines. Exaggerated expression. Comic-book energy. Dramatic rim lighting in orange and gold.',
    'Educational Guide': 'A cartoon British female homeowner — middle-aged, neatly dressed in a jumper. Expression: a giant lightbulb moment — eyes wide and lit up, finger pointing in the air in a eureka moment. Thick black outlines. Exaggerated expression. Comic-book energy. Dramatic rim lighting in orange and gold.',
    'Hub Page':          'A cartoon British female homeowner — middle-aged, standing with arms wide open gesturing at everything around them, expression of confident authority and welcome. Thick black outlines. Comic-book energy. Dramatic rim lighting in orange and gold.',
    'Buyer Guide':       'A cartoon British female homeowner — middle-aged, holding two competing products in each hand, looking between them with a deeply uncertain expression. One eyebrow raised. Thick black outlines. Comic-book energy. Dramatic rim lighting in orange and gold.',
    'Service Page':      'A cartoon British female homeowner — middle-aged, arms folded with a smug satisfied smile and one eyebrow raised in calm confidence. Expression of someone who made the right choice. Thick black outlines. Comic-book energy. Dramatic rim lighting in orange and gold.',
    'Geo Service Page':  'A cartoon British female homeowner — middle-aged, holding a phone to her ear with a relieved smile. Expression of someone who has just found the right local specialist. Thick black outlines. Comic-book energy. Dramatic rim lighting in orange and gold.'
  };

  const baseCharacter = characterMap[articleType] || characterMap['Method Guide'];
  const character = expressionOverride
    ? 'A cartoon British female homeowner — middle-aged, neatly dressed in a practical jumper. ' + expressionOverride
    : baseCharacter;

  // ── VISUAL CLUE ICONS MAP ──
  const visualClueMap = {
    'Victorian Tile': [
      { icon: 'a cartoon tile with white salt deposits blooming up through it', label: 'SALT BLOOM' },
      { icon: 'a cartoon tile with a loose edge lifting away from the floor', label: 'LOOSE TILES' },
      { icon: 'a cartoon tile surface with old bitumen adhesive marks visible', label: 'OLD RESIDUE' }
    ],
    'Travertine': [
      { icon: 'a cartoon stone tile with natural holes and voids filled with dark compacted dirt', label: 'CLOGGED VOIDS' },
      { icon: 'a cartoon tile surface with dull cloudy patches across it', label: 'ACID DAMAGE' },
      { icon: 'a cartoon tile with filler crumbling away from a surface hole', label: 'FILLER LOSS' }
    ],
    'Marble': [
      { icon: 'a cartoon marble tile with a dull cloudy etch mark where acid has dissolved the surface', label: 'ACID ETCH' },
      { icon: 'a cartoon marble floor with directional fine scratches dulling the polish', label: 'MICRO SCRATCHES' },
      { icon: 'a cartoon marble tile with a brown rust-coloured internal stain', label: 'IRON STAINING' }
    ],
    'Limestone': [
      { icon: 'a cartoon limestone tile with white efflorescence deposits on the surface', label: 'SALT DEPOSITS' },
      { icon: 'a cartoon limestone tile with a dull rough etch mark from acid contact', label: 'ACID ETCH' },
      { icon: 'a cartoon porous stone tile absorbing a liquid spill instantly', label: 'SEALER FAILED' }
    ],
    'Slate': [
      { icon: 'a cartoon slate tile with surface layers peeling and separating', label: 'DELAMINATION' },
      { icon: 'a cartoon slate tile with old acrylic sealer build-up at the edges', label: 'SEALER BUILD-UP' },
      { icon: 'a cartoon slate tile that has lost its colour depth and looks faded', label: 'COLOUR LOSS' }
    ],
    'Quarry Tile': [
      { icon: 'a cartoon quarry tile with white salt deposits on the surface', label: 'EFFLORESCENCE' },
      { icon: 'a cartoon quarry tile with a concave worn hollow in the centre from foot traffic', label: 'TRAFFIC WEAR' },
      { icon: 'a cartoon quarry tile with old polymerised linseed oil darkening the surface', label: 'OIL RESIDUE' }
    ],
    'Terracotta': [
      { icon: 'a cartoon terracotta tile with multiple layers of old acrylic coating peeling', label: 'COATING BUILD-UP' },
      { icon: 'a cartoon terracotta tile with white efflorescence deposits appearing through a failed sealer', label: 'SEALER FAILURE' },
      { icon: 'a cartoon terracotta tile darkened by old polymerised linseed oil', label: 'OIL STAINING' }
    ],
    'Sandstone': [
      { icon: 'a cartoon sandstone tile with white salt deposits on a riven surface', label: 'EFFLORESCENCE' },
      { icon: 'a cartoon riven sandstone surface with soil compacted deep into the texture', label: 'DEEP SOILING' },
      { icon: 'a cartoon sandstone tile with surface layers peeling from bedding plane separation', label: 'DELAMINATION' }
    ],
    'Terrazzo': [
      { icon: 'a cartoon terrazzo floor with the cement binder becoming porous and absorbing stains', label: 'BINDER DECAY' },
      { icon: 'a cartoon terrazzo surface with dull rough patches where acid has dissolved the cement', label: 'ACID DAMAGE' },
      { icon: 'a cartoon terrazzo tile with height variation between tiles causing a trip edge', label: 'LIPPAGE' }
    ],
    'Porcelain Tile': [
      { icon: 'a cartoon polished porcelain tile with gloss dulled by micro-abrasion from grit', label: 'GLOSS LOSS' },
      { icon: 'a cartoon grout line that has turned dark and resistant to cleaning', label: 'GROUT SOILING' },
      { icon: 'a cartoon porcelain tile edge chipped by DIY grout removal', label: 'EDGE CHIPPING' }
    ],
    'Ceramic Tile': [
      { icon: 'a cartoon ceramic tile grout line darkened with absorbed contamination', label: 'DARK GROUT' },
      { icon: 'a cartoon ceramic tile surface with polymer grout haze obscuring the glaze', label: 'GROUT HAZE' },
      { icon: 'a cartoon ceramic tile in a wet area with mould growth in the grout joints', label: 'MOULD GROWTH' }
    ],
    'Grout': [
      { icon: 'a cartoon grout joint with capillary absorption pulling contamination deep inside', label: 'DEEP SOILING' },
      { icon: 'a cartoon grout line with biological mould growth visible in the joint', label: 'MOULD' },
      { icon: 'a cartoon grout joint with fines eroded away making it porous and staining rapidly', label: 'SURFACE LOSS' }
    ]
  };

  const defaultClues = [
    { icon: 'a cartoon floor tile with a dull contaminated surface', label: 'SOILING' },
    { icon: 'a cartoon grout line darkened with absorbed contamination', label: 'GROUT FAILURE' },
    { icon: 'a cartoon tile with a failed sealer allowing liquid to absorb', label: 'SEALER FAILED' }
  ];

  // ── READ THUMBNAIL ENTITIES FROM SHEET ──
  const THUMBNAIL_ENTITY_COL = col('Thumbnail Entities');
  const thumbnailEntities = THUMBNAIL_ENTITY_COL > -1 ? values[THUMBNAIL_ENTITY_COL] : '';
  let dynamicClues = [];
  if (thumbnailEntities) {
    const entityList = thumbnailEntities.split(',').map(e => e.trim()).filter(Boolean);
    dynamicClues = entityList.slice(0, 3).map(e => ({ label: e.toUpperCase() }));
  }

  const visualClues = dynamicClues.length === 3
    ? dynamicClues
    : (visualClueMap[stoneType] || defaultClues);
// ── BACKGROUND COLOUR MAP ──
  const bgOptions = [
    'deep charcoal black',
    'deep forest green',
    'deep navy blue',
    'deep slate blue',
    'deep burgundy',
    'deep teal'
  ];

  const bgColour = bgOptions[Math.floor(Math.random() * bgOptions.length)];


  // ── STARBURST BADGE MAP ──
  const starburstMap = {
    'Case Study':        'SEE THE RESULT',
    'Method Guide':      'STOP SCRUBBING',
    'Diagnostic Guide':  'SPOT THE SIGNS',
    'Educational Guide': 'NOW IT MAKES SENSE',
    'Hub Page':          'FULL GUIDE',
    'Buyer Guide':       'READ THIS FIRST',
    'Service Page':      '30 YEARS EXPERIENCE',
    'Geo Service Page':  'LOCAL SPECIALIST'
  };

  const starburstText = starburstMap[articleType] || 'READ THIS FIRST';

  // ── BOTTOM CALLOUT MAP ──
  const calloutMap = {
    'Case Study':        'THE RESTORATION THAT CHANGED EVERYTHING',
    'Method Guide':      'THE METHOD THAT ACTUALLY WORKS',
    'Diagnostic Guide':  'IDENTIFY IT BEFORE YOU TREAT IT',
    'Educational Guide': 'UNDERSTAND YOUR FLOOR BEFORE YOU CLEAN IT',
    'Hub Page':          'EVERYTHING YOU NEED IN ONE PLACE',
    'Buyer Guide':       'WHAT TO ASK BEFORE YOU COMMIT',
    'Service Page':      'REAL WORK ON REAL FLOORS',
    'Geo Service Page':  'ASSESSED AND RESTORED LOCALLY'
  };

  // ── EXTRACT SECTION 1 HEADING FROM W1.5D ──
  let dynamicCallout = '';
  const sec1HeadingMatch = w15d.match(/SECTION 1:\s*(?:Heading\s+H\d:\s*)?([^\n]+)/i);
  if (sec1HeadingMatch) {
    dynamicCallout = sec1HeadingMatch[1].trim().toUpperCase();
    // Trim to a punchy callout if over 60 characters
    if (dynamicCallout.length > 60) {
      dynamicCallout = dynamicCallout.substring(0, dynamicCallout.lastIndexOf(' ', 60)).trim();
    }
  }

  const bottomCallout = dynamicCallout || calloutMap[articleType] || 'THE PROFESSIONAL APPROACH';

  // ── BUILD PROMPT ──
  const prompt =
    'Create a high-CTR viral YouTube thumbnail at 1280x720 pixels, 16:9 aspect ratio.\n\n' +

    'THUMBNAIL PURPOSE:\n' +
    'This thumbnail is for a ' + articleType + ' video about ' + materialEntity + '.\n' +
    (locality ? 'Location context: ' + locality + '.\n' : '') +
    'The viewer is a UK homeowner with a ' + stoneType + ' floor problem.\n' +
    'The thumbnail must communicate instantly: the video solves a real problem this homeowner recognises.\n\n' +

    'VISUAL STYLE:\n' +
    'Ultra high-CTR YouTube thumbnail with hyper-stylised comic-book marketing energy.\n' +
    'Thick black outlines. Neon glow accents. Aggressive contrast. Stacked typography.\n' +
    'Dramatic rim lighting. Exaggerated facial expressions. Viral SEO thumbnail aesthetic.\n' +
    'Layered sticker graphics. Vector-cartoon hybrid illustration. Mobile-first readability.\n' +
    'YouTube growth hacking thumbnail style. High-density visual layout. Click-focused composition.\n\n' +
    'No photorealism. No cinematic realism. No soft muted colours. No minimalism.\n' +
    'No flat design. No realistic textures. No empty background. No corporate design.\n' +
    'No thin typography. No elegant luxury styling.\n\n' +

    'COMPOSITION:\n' +
    'Left two thirds — bold stacked typography dominating the space.\n' +
    'Right third — cartoon character as the emotional anchor.\n' +
    'Bottom strip — three small visual clue icons in a row.\n' +
    'Background — ' + bgColour + ' with explosive orange and gold radial glow behind the character.\n\n' +

    'CHARACTER:\n' +
    character + '\n\n' +

    'VISUAL CLUE TAGS (bottom strip — three text blocks in a row):\n' +
    'Tag 1 — "' + visualClues[0].label + '"\n' +
    'Tag 2 — "' + visualClues[1].label + '"\n' +
    'Tag 3 — "' + visualClues[2].label + '"\n' +
    'Each tag is a bold sticker-style rectangular badge. Ultra bold condensed caps. White text on gold or orange fill. Thick black outline. No icons. Text only. Maximum visibility at thumbnail scale.\n\n' +

    'TYPOGRAPHY:\n' +
    'Main headline (left side, stacked, maximum size):\n' +
    'Line 1: "' + headlineLines[0] + '"\n' +
    'Line 2: "' + headlineLines[1] + '"\n' +
    'Line 3: "' + headlineLines[2] + '"\n' +
    'Font style: ultra bold condensed sans-serif. White text with thick black stroke outline.\n' +
    'Line 2 in bright gold or yellow for emphasis.\n\n' +
    'Secondary badge (top right corner — circular sticker burst):\n' +
    '"' + starburstText + '"\n' +
    'Bold white text on red background. Thick black outline. Starburst shape.\n\n' +
    'Bottom callout strip (full width under the icons):\n' +
    '"' + bottomCallout + '"\n' +
    'Bold caps. White on dark charcoal. High contrast.\n\n' +

    'COLOUR PALETTE:\n' +
    'Background: ' + bgColour + '\n' +
    'Radial glow: explosive orange and warm gold behind character\n' +
    'Typography: white with black stroke, gold or yellow accent on Line 2\n' +
    'Icon badges: gold or orange fill with black outline\n' +
    'Starburst badge: red and white\n' +
    'Overall energy: high contrast, aggressive, visually loud\n\n' +

    'MARKETING ELEMENTS:\n' +
    'Starburst callout badge top right.\n' +
    'Three visual clue icon badges bottom strip.\n' +
    'Layered typography with size hierarchy.\n' +
    'Glow border effect around character silhouette.\n' +
    'Thin gold divider line separating headline from icon strip.\n\n' +

    'MOBILE READABILITY:\n' +
    'All text must be clearly readable at 120x68 pixels.\n' +
    'Character expression must read clearly at thumbnail scale.\n' +
    'Icon labels must be legible at small size — maximum 2 words each.\n' +
    'No fine detail that disappears at small scale.\n\n' +

    'TECHNICAL OUTPUT:\n' +
    '1280x720 pixels. 16:9 aspect ratio.\n' +
    'No watermarks. No borders. No extra elements beyond the described composition.\n' +
    'Clean impactful design. Print-ready quality.';

  return prompt;
}