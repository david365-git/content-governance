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
  var result = [], section = 0, seen = {};
  var lines = content.split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var heading = /^\s*SECTION\s+(\d+):/i.exec(lines[i]);
    if (heading) {section = Number(heading[1]);continue;}
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
      result.push({url:url,path:path,section:section,kind:hub && path===hub ? 'hub' : 'lateral'});
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
    .replace(/<style\b[^>]*>[\s\S]*?<\/style\s*>/gi, '');
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
