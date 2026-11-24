/**
 * W1.5D_Sol selects governed links for the ACTUAL finished Sol article H2s.
 * The original W1.5D script and Content Generation CU:CX remain unchanged.
 */
function glb_solCanonical_(url){
 var text=String(url||'').trim().replace(/&amp;/gi,'&');
 if(/^\/(?!\/)/.test(text))text='https://www.abbeyfloorcare.co.uk'+text;
 var m=/^https?:\/\/(?:www\.)?abbeyfloorcare\.co\.uk(\/[^?#]*)?(?:[?#].*)?$/i.exec(text);
 if(!m)return '';
 return (m[1]||'/').replace(/\/{2,}/g,'/').replace(/\/+$/,'').toLowerCase()||'/';
}
/**
 * Geographical educational pages are not general-purpose destinations.
 * Case studies are explicitly exempt: projects can illustrate a material
 * or method outside their own locality, without implying local service.
 * Deterministic matching prevents model-only geographic exceptions.
 */
function glb_solGeoEducationAllowed_(page,currentLocality){
 var type=String(page.articleType||'').trim();
 if(type==='Case Study')return true;
 if(type!=='Educational Guide'&&type!=='Diagnostic Guide'&&
    type!=='Method Guide'&&type!=='Buyer Guide'&&type!=='FAQ Spoke')return true;
 var own=String(currentLocality||'').trim().toLowerCase();
 var target=String(page.locality||'').trim().toLowerCase();
 if(target)return !!own&&own===target;
 // Some geographically targeted educational pages have no locality metadata.
 // Detect explicit UK location targeting in the title/URL, but never suppress
 // case studies or make geographic assumptions from an unrecognised name.
 var title=String(page.title||'').toLowerCase();
 var url=String(page.url||'').toLowerCase().replace(/[-_]/g,' ');
 var markers=['glasgow','edinburgh','london','birmingham','manchester',
  'liverpool','bristol','leeds','sheffield','newcastle','nottingham',
  'derby','leicester','coventry','oxford','cambridge','york','walsall',
  'stirling','dundee','aberdeen','inverness','perth','fife',
  'croydon','surrey','staffordshire','somerset','scotland','wales'];
 var found=markers.filter(function(place){
  var pattern=new RegExp('\\b'+place+'\\b','i');
  return pattern.test(title)||pattern.test(url);
 });
 if(!found.length)return true;
 // Require an explicit matching current locality for geographically focused
 // educational pages. Blank locality on a national guide never grants access.
 return !!own&&found.some(function(place){
  return own===place||own.indexOf(place+',')===0||own.indexOf(', '+place)>=0;
 });
}
function glb_solSections_(html){
 var source=String(html||''),head=/<h2\b[^>]*>[\s\S]*?<\/h2\s*>/gi,m,list=[];
 while((m=head.exec(source)))list.push({start:m.index,after:head.lastIndex,title:m[0].replace(/<[^>]*>/g,' ').trim()});
 if(!list.length)throw Error('W1.5D_Sol: no H2 sections in finished article.');
 return list.map(function(h,i){
  var end=i+1<list.length?list[i+1].start:source.length;
  var body=source.slice(h.after,end).split(/<footer\b|<div\b[^>]*class=["'][^"']*abbey-bio-box/i)[0];
  var p=/<p\b[^>]*>([\s\S]*?)<\/p\s*>/gi,match,paras=[];
  while((match=p.exec(body))) {
   // Hide text that cannot be linked: old tags were stripped, misleading the model.
   var selectable=match[1]
    .replace(/<a\b[^>]*>[\s\S]*?<\/a\s*>/gi,' [existing link unavailable] ')
    .replace(/<(?:strong|em|span|code|script|style)\b[^>]*>[\s\S]*?<\/(?:strong|em|span|code|script|style)\s*>/gi,' [formatted text unavailable] ');
   var text=selectable.replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim();
   if(text.length>45)paras.push(text.slice(0,350));
  }
  return {number:i+1,title:h.title,start:h.after,end:end,paras:paras.slice(0,5)};
 });
}
function glb_solSelectLinks_(article,html){
 var sh=glb_posts_(),heads=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
 function read(h){var i=heads.indexOf(h);if(i<0)throw Error('Missing posts heading '+h);
  return String(sh.getRange(article.row,i+1).getValue()||'').trim();}
 if(read('Post ID')!==String(article.id))throw Error('W1.5D_Sol post row moved.');
 var type=read('Article Type'),material=read('Stone Type'),url=read('URL'),hub=read('Feeds Hub');
 var own=glb_solCanonical_(url),hubPath=glb_solCanonical_(hub),isHub=type==='Hub Page';
 var currentLocality=read('Locality');
 if(!isHub&&(!hubPath||hubPath===own))throw Error('W1.5D_Sol invalid Feeds Hub.');
 var limit=getMaxLateralLinks(type);
 if(limit===null||!Number.isInteger(limit)||limit<0)throw Error('W1.5D_Sol invalid lateral-link budget.');
 var used={},pages=buildEnhancedSiloMap(material,article.row,url).filter(function(p){
  var path=glb_solCanonical_(p.url);
  if(!path||path===own||path===hubPath||path==='/contact'||used[path])return false;
  if(!glb_solGeoEducationAllowed_(p,currentLocality))return false;
  used[path]=true;return true;
 });
 if(pages.length>150)throw Error('Too many eligible pages for safe selection.');
 var sections=glb_solSections_(html);
 var prompt=[
  'W1.5D_Sol: select contextual internal links for FINISHED article; do NOT rewrite it.',
  'JSON only: {"assignments":[{"section":1,"page":0,"anchorText":"exact phrase","readerBenefit":"specific distinct reason to visit that page"}]}',
  'PAGE 0 is the material hub, mandatory once for non-hub articles; forbidden for Hub Pages.',
  'Choose at most '+(isHub?pages.length:limit)+' other pages, zero if none relevant.',
  'Never repeat page numbers. Respect same-material, locality and cannibalisation scope.',
  'Geographically targeted educational destinations are eligible ONLY when the current article shares their explicit locality; do not force local guides into national advice.',
  'Case studies are different: allow nonlocal projects only when their documented methods, conditions or outcomes directly help the reader; do not imply they happened locally.',
  'For EACH selected page provide readerBenefit: a concise, specific explanation of what the reader learns from that destination BEYOND what this section already explains.',
  'Do not pick a destination merely because it shares the material, repeats the warning or uses similar keywords.',
  'Choose anchors that accurately signal the destination topic, not generic statements of fact or misleading location/service claims.',
  'Anchor text must be an exact, single text-node substring from a paragraph in that H2 section.',
  'Paragraph excerpts mark existing links and protected formatting as unavailable. Do not select words from these unavailable spans as anchor text.',
  'Choose a descriptive, naturally occurring phrase 3 to 75 characters long.',
  'No headings, existing links, formatted words, image captions, footers or invented text.',
  'ARTICLE '+JSON.stringify({id:article.id,title:article.title,type:type,material:material}),
  'HUB PAGE 0 '+JSON.stringify(isHub?null:{url:hub}),
  'APPROVED SILO PAGES '+JSON.stringify(pages.map(function(p,i){return {
   page:i+1,url:p.url,title:p.title,type:p.articleType,locality:p.locality,
   sameLocality:p.sameLocality,primaryEntity:p.primaryEntity,
   supportingEntities:p.supportingEntities,safeHandoff:p.safeHandoff,
   cannibalGuard:p.cannibalGuard,caseEvidence:String(p.caseEvidence||'').slice(0,250)
  };})),
  'REAL SOL H2 SECTIONS '+JSON.stringify(sections.map(function(s){return {
   section:s.number,heading:s.title,paragraphs:s.paras
  };}))
 ].join('\n');
 if(prompt.length>110000)throw Error('W1.5D_Sol prompt exceeds safe size.');
 var api=bc_sendPromptViaOpenAI(prompt,2500,MODEL_CHEAP);
 if(!api||!api.success||!String(api.text||'').trim())
  throw Error('W1.5D_Sol API selection failed: '+String(api&&api.message||'no response'));
 var obj;try{obj=JSON.parse(String(api.text||'').trim());}
 catch(e){throw Error('W1.5D_Sol invalid JSON; Sol draft preserved.');}
 if(!obj||!Array.isArray(obj.assignments))throw Error('W1.5D_Sol missing assignments.');
 var seen={},hubCount=0,lateral=0,links=[];
 obj.assignments.forEach(function(x){
  var section=Number(x.section),page=Number(x.page),anchor=String(x.anchorText||'').trim();
  if(!Number.isInteger(section)||section<1||section>sections.length||
   !Number.isInteger(page)||page<0||page>pages.length)throw Error('Invalid page/section selection.');
  if(seen[page])throw Error('Duplicate destination selected.');
  seen[page]=true;
  if(page===0)hubCount++;else lateral++;
  if(anchor.length<3||anchor.length>75||/[<>]/.test(anchor))throw Error('Invalid anchor phrase.');
  var benefit=String(x.readerBenefit||'').trim();
  if(benefit.length<25||benefit.length>350)throw Error('W1.5D_Sol requires a distinct reader benefit for every selected link.');
  var target=page===0?hub:pages[page-1].url;
  if(!glb_solCanonical_(target)||glb_solCanonical_(target)===own)throw Error('Invalid destination.');
  links.push({section:section,url:target,anchor:anchor,readerBenefit:benefit});
 });
 if((!isHub&&hubCount!==1)||(isHub&&hubCount!==0)||lateral>(isHub?pages.length:limit))
  throw Error('Hub quota or lateral budget not met.');
 return {links:links,cost:Number(api.cost||0)};
}

/**
 * Identify each exact linkable occurrence, ignoring existing links, formatted
 * spans and non-paragraph content. The same text-node rules apply to
 * candidate validation and insertion.
 */
function glb_solAnchorPositions_(html,section,anchor){
 var sections=glb_solSections_(html),sec=sections[section-1];
 if(!sec)return [];
 var within=html.slice(sec.start,sec.end);
 var editable=within.split(/<footer\b|<div\b[^>]*class=["'][^"']*abbey-bio-box/i)[0];
 var p=/<p\b[^>]*>[\s\S]*?<\/p\s*>/gi,m,places=[];
 while((m=p.exec(editable))){
  var tokens=m[0].split(/(<[^>]*>)/g),inAnchor=false,forbiddenDepth=0;
  for(var i=0;i<tokens.length;i++){
   var token=tokens[i];
   if(token.charAt(0)==='<'){
    if(/^<a\b/i.test(token))inAnchor=true;
    if(/^<\/a\s*>/i.test(token))inAnchor=false;
    if(/^<(?:script|style|code|strong|em|span|svg)\b/i.test(token))forbiddenDepth++;
    if(/^<\/(?:script|style|code|strong|em|span|svg)\s*>/i.test(token))forbiddenDepth=Math.max(0,forbiddenDepth-1);
   }else if(!inAnchor&&!forbiddenDepth){
    var from=0,at;
    while((at=token.indexOf(anchor,from))>=0){
     places.push({sec:sec,offset:m.index,raw:m[0],tokens:tokens,token:i,at:at});
     from=at+anchor.length;
    }
   }
  }
 }
 return places;
}
/**
 * Use the chosen section whenever its anchor occurs uniquely there.
 * If missing there, correct the section ONLY when the exact phrase appears
 * once in a different eligible paragraph in the whole article.
 * No text is rewritten or guessed.
 */
function glb_solResolveAnchor_(html,assignment){
 var sections=glb_solSections_(html);
 var intended=glb_solAnchorPositions_(html,assignment.section,assignment.anchor);
 if(intended.length===1)return assignment;
 if(intended.length>1)throw Error('W1.5D_Sol ambiguous anchor in H2 section '+assignment.section+': '+assignment.anchor);
 var candidates=[];
 for(var n=1;n<=sections.length;n++){
  var matches=glb_solAnchorPositions_(html,n,assignment.anchor);
  for(var i=0;i<matches.length;i++)candidates.push(n);
 }
 if(candidates.length!==1)
  throw Error('W1.5D_Sol anchor not uniquely linkable in any H2 paragraph: '+assignment.anchor+
   ' ('+candidates.length+' matches); saved Sol draft retained.');
 var corrected={};
 Object.keys(assignment).forEach(function(key){corrected[key]=assignment[key];});
 corrected.section=candidates[0];
 return corrected;
}
/** Insert one validated, unique existing phrase into an eligible paragraph. */
function glb_solInsertOne_(html,assignment){
 var resolved=glb_solResolveAnchor_(html,assignment);
 var sec=glb_solSections_(html)[resolved.section-1];
 var places=glb_solAnchorPositions_(html,resolved.section,resolved.anchor);
 if(places.length!==1)throw Error('W1.5D_Sol linkable anchor changed unexpectedly: '+resolved.anchor);
 var found=places[0],text=found.tokens[found.token];
 var href=resolved.url.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
 found.tokens[found.token]=text.slice(0,found.at)+'<a href="'+href+'">'+resolved.anchor+'</a>'+text.slice(found.at+resolved.anchor.length);
 var within=html.slice(sec.start,sec.end);
 var editable=within.split(/<footer\b|<div\b[^>]*class=["'][^"']*abbey-bio-box/i)[0];
 var modified=editable.slice(0,found.offset)+found.tokens.join('')+editable.slice(found.offset+found.raw.length)+within.slice(editable.length);
 return {html:html.slice(0,sec.start)+modified+html.slice(sec.end),assignment:resolved};
}
function glb_solApplyLinks_(original,selection){
 var html=String(original||''),links=selection.links.slice(),placed=[];
 // Resolve and validate against the current HTML, including any earlier links.
 links.forEach(function(link){
  var insertion=glb_solInsertOne_(html,link);
  html=insertion.html;
  placed.push(insertion.assignment);
 });
 // Undo only the inserted links. Article text, attributes and images must be identical.
 var revert=html;
 placed.forEach(function(link){
  var href=link.url.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
  var snippet='<a href="'+href+'">'+link.anchor+'</a>';
  if(revert.split(snippet).length!==2)throw Error('W1.5D_Sol inserted anchor is not unique.');
  revert=revert.replace(snippet,link.anchor);
 });
 if(revert!==original)throw Error('W1.5D_Sol altered the original article outside approved anchors.');
 var sections=glb_solSections_(html);
 placed.forEach(function(link){
  var sec=sections[link.section-1];
  var href=link.url.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
  if(!sec||html.slice(sec.start,sec.end).indexOf('<a href="'+href+'">'+link.anchor+'</a>')<0)
   throw Error('W1.5D_Sol placement validation failed.');
 });
 return {html:html,links:placed};
}
function glb_runW15DSol_(article,html){
 var selected=glb_solSelectLinks_(article,html);
 var placed=glb_solApplyLinks_(html,selected);
 return {html:placed.html,links:placed.links,cost:selected.cost};
}
