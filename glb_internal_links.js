/**
 * Google Love governed internal-link audit — DEVELOPMENT ONLY.
 * Pure functions: no spreadsheet writes, no external requests, no publishing.
 * The existing W1.5D final plan (posts column CX) is the approved authority.
 * This file is intentionally not yet called by the live Google Love runner.
 */
function glb_linkCanonical_(value) {
  var candidate = String(value || '').trim().replace(/&amp;/gi, '&');
  if (candidate.charAt(0) === '/' && candidate.charAt(1) !== '/') {
    candidate = 'https://www.abbeyfloorcare.co.uk' + candidate;
  }
  var match = /^https?:\/\/(?:www\.)?abbeyfloorcare\.co\.uk(\/[^?#]*)?(?:[?#].*)?$/i.exec(candidate);
  if (!match) return '';
  var path = String(match[1] || '/').replace(/\/{2,}/g, '/').replace(/\/+$/, '') || '/';
  return path.toLowerCase();
}

function glb_linkPlanFromCx_(cx, feedsHub, articleType, ownUrl) {
  var content = String(cx || '').trim();
  var type = String(articleType || '').trim();
  var hub = glb_linkCanonical_(feedsHub);
  var own = glb_linkCanonical_(ownUrl);
  var isHub = type === 'Hub Page';
  if (!/Stage 1\.5D complete\.\s*$/i.test(content)) {
    return {ready:false,reason:'No completed W1.5D final plan in CX',required:[],hub:''};
  }
  var result = [], section = 0, sectionHeading = '', seen = {};
  var lines = content.split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var heading = /^\s*SECTION\s+(\d+):/i.exec(lines[i]);
    if (heading) {
      section = Number(heading[1]);
      var headingText = /Heading H2:\s*(.+?)\s*$/i.exec(lines[i]);
      sectionHeading = headingText ? headingText[1].trim() : '';
      continue;
    }
    var assignment = /^\s*Internal link:\s*(.+?)\s*$/i.exec(lines[i]);
    if (!assignment || /^None\s*$/i.test(assignment[1])) continue;
    if (!section) return {ready:false,reason:'CX link appears outside a numbered section',required:[],hub:''};
    var tokens = assignment[1].split(/\s*;\s*/);
    for (var j = 0; j < tokens.length; j++) {
      var url = tokens[j].trim();
      var path = glb_linkCanonical_(url);
      if (!path || path === own || path === '/contact') {
        return {ready:false,reason:'CX contains an invalid, self, or contact destination',required:[],hub:''};
      }
      if (seen[path]) {
        return {ready:false,reason:'CX assigns the same internal destination more than once',required:[],hub:''};
      }
      seen[path] = true;
      result.push({url:url,path:path,section:section,heading:sectionHeading,kind:hub && path===hub ? 'hub' : 'lateral'});
    }
  }
  if (!isHub) {
    if (!hub || hub === own) {
      return {ready:false,reason:'Feeds Hub is missing, invalid or points to the current article',required:[],hub:''};
    }
    if (!seen[hub]) {
      return {ready:false,reason:'The completed CX plan does not assign Feeds Hub',required:[],hub:hub};
    }
  } else if (hub && seen[hub]) {
    return {ready:false,reason:'Hub page must not assign itself as a destination',required:[],hub:''};
  }
  return {ready:true,required:result,hub:hub,reason:''};
}

function glb_linkHrefPaths_(html) {
  var safe = String(html || '')
    .replace(/<script\b[^>]*>[\s\S]*?<\/script\s*>/gi, '')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style\s*>/gi, '')
    // Only meaningful editorial-body anchors count, not the fixed CTA/footer.
    .replace(/<!--[^]*?-->/g, '')
    .replace(/<footer\b[^>]*>[\s\S]*?<\/footer\s*>/gi, '');
  var paths = {};
  var matcher = /<a\b[^>]*\bhref\s*=\s*(?:"([^"]+)"|'([^']+)')/gi;
  var match;
  while ((match = matcher.exec(safe)) !== null) {
    var path = glb_linkCanonical_(match[1] || match[2]);
    if (path) paths[path] = true;
  }
  return paths;
}

function glb_linkAuditHtml_(html, plan) {
  if (!plan || !plan.ready) {
    return {pass:false,missing:[],reason:plan && plan.reason || 'No approved link plan'};
  }
  var anchors = glb_linkHrefPaths_(html), missing = [];
  plan.required.forEach(function(link) {
    if (!anchors[link.path]) missing.push(link);
  });
  return {pass:missing.length===0,missing:missing,reason:missing.length
    ? 'Generated HTML is missing approved internal anchors' : ''};
}


/**
 * Detect whether CX was built from the current CW source plan.
 * W1.5D is permitted to alter link lines, contact routing and its completion
 * marker, but not the governed section structure or technical instructions.
 */
function glb_linkPlanIsCurrent_(cw, cx) {
  if (!String(cw || '').trim() || !String(cx || '').trim()) return false;
  function baseline(value) {
    return String(value).split(/\r?\n/)
      .filter(function(line) {
        return !/^\s*(?:Internal link:|Contact link:|Stage 1\.5[CD] complete\.)/i.test(line);
      }).join('\n').replace(/\s+/g, ' ').trim();
  }
  return baseline(cw) === baseline(cx);
}

/**
 * Read-only per-post preflight using the existing W1.5D approval and silo rules.
 * Nothing is automatically selected or approved when CX is missing.
 */
function glb_linkPreflight_(article) {
  var sh = glb_posts_();
  var heads = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
    .map(function(v) { return String(v || '').trim(); });
  function value(header) {
    var col = heads.indexOf(header) + 1;
    if (!col) throw Error('Required posts column missing: ' + header);
    return String(sh.getRange(article.row, col).getValue() || '').trim();
  }
  if (String(value('Post ID')) !== String(article.id))
    throw Error('Posts row changed during Google Love link preflight.');
  var type = value('Article Type');
  var url = value('URL');
  var hub = value('Feeds Hub');
  // The existing W1.5D writer stores its governed final plan in CX (102).
  var cx = String(sh.getRange(article.row, 102).getValue() || '').trim();
  var cw = String(sh.getRange(article.row, 101).getValue() || '').trim();
  if (!cw || !cx)
    throw Error('Complete Content Generation stages W1.5C and W1.5D to populate CW and CX before running Google Love.');
  if (!glb_linkPlanIsCurrent_(cw, cx))
    throw Error('Google Love links need W1.5D review: CX no longer matches its CW source plan.');
  var plan = glb_linkPlanFromCx_(cx, hub, type, url);
  if (!plan.ready) throw Error('Google Love links need W1.5D review: ' + plan.reason);
  var limit = getMaxLateralLinks(type);
  if (limit === null || !Number.isInteger(limit) || limit < 0)
    throw Error('Google Love cannot verify the article-type lateral link budget.');
  var lateral = plan.required.filter(function(item) { return item.kind === 'lateral'; });
  if (type !== 'Hub Page' && lateral.length > limit)
    throw Error('Approved CX contains more lateral links than article-type rules allow.');
  var candidates = buildEnhancedSiloMap(value('Stone Type'), article.row, url);
  var eligible = {};
  candidates.forEach(function(page) {
    var canonical = glb_linkCanonical_(page.url);
    if (canonical) eligible[canonical] = true;
  });
  lateral.forEach(function(item) {
    if (!eligible[item.path]) throw Error('CX lateral link no longer eligible: ' + item.url);
  });
  return plan;
}

/** Give Sol only the exact approved URLs. No new destinations may be inferred. */
function glb_linkPrompt_(plan) {
  if (!plan || !plan.ready) throw Error('No verified internal link plan.');
  var lines = [
    'GOVERNED INTERNAL LINKS — MANDATORY:',
    'The exact listed URLs were approved by the W1.5D governance plan, not invented by the writer.',
    'Include every listed URL once as a real <a href="exact URL">natural contextual anchor</a> in the article body.',
    'The section numbers and H2 topics identify where W1.5D assigned each destination. Place each anchor in a substantive paragraph or list item explaining that topic; do not append a generic block of links at the end.',
    'Do not put the links only in a contact footer, script, data block, author biography or navigation comment.',
    'Do not invent, substitute, shorten or redirect destinations. Avoid claims not supported by the evidence.',
    'Do not create a new link when W1.5D has approved none. Contact CTA is separate.'
  ];
  plan.required.forEach(function(item) {
    lines.push('Approved ' + item.kind + ' link for planned SECTION ' +
      item.section + (item.heading ? ' (topic: ' + item.heading + ')' : '') + ': ' + item.url);
  });
  return lines.join('\n');
}
