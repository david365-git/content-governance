/**
 * W1.5D_Sol selects governed links for the ACTUAL finished Sol article H2s.
 * The original W1.5D script and Content Generation CU:CX remain unchanged.
 */
function glb_solSections_(html){
 var source=String(html||''),head=/<h2\b[^>]*>[\s\S]*?<\/h2\s*>/gi,m,list=[];
 while((m=head.exec(source)))list.push({start:m.index,after:head.lastIndex,title:m[0].replace(/<[^>]*>/g,' ').trim()});
 if(!list.length)throw Error('W1.5D_Sol: no H2 sections in finished article.');
 return list.map(function(h,i){
  var end=i+1<list.length?list[i+1].start:source.length;
  var body=source.slice(h.after,end).split(/<footer\b|<div\b[^>]*class=["'][^"']*abbey-bio-box/i)[0];
  var p=/<p\b[^>]*>([\s\S]*?)<\/p\s*>/gi,match,paras=[];
  while((match=p.exec(body))) {
   var text=match[1].replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim();
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
 var own=glb_linkCanonical_(url),hubPath=glb_linkCanonical_(hub),isHub=type==='Hub Page';
 if(!isHub&&(!hubPath||hubPath===own))throw Error('W1.5D_Sol invalid Feeds Hub.');
 var limit=getMaxLateralLinks(type);
 if(limit===null||!Number.isInteger(limit)||limit<0)throw Error('W1.5D_Sol invalid lateral-link budget.');
 var used={},pages=buildEnhancedSiloMap(material,article.row,url).filter(function(p){
  var path=glb_linkCanonical_(p.url);
  if(!path||path===own||path===hubPath||path==='/contact'||used[path])return false;
  used[path]=true;return true;
 });
 if(pages.length>150)throw Error('Too many eligible pages for safe selection.');
 var sections=glb_solSections_(html);
 var prompt=[
  'W1.5D_Sol: select contextual internal links for FINISHED article; do NOT rewrite it.',
  'JSON only: {"assignments":[{"section":1,"page":0,"anchorText":"exact phrase"}]}',
  'PAGE 0 is the material hub, mandatory once for non-hub articles; forbidden for Hub Pages.',
  'Choose at most '+(isHub?pages.length:limit)+' other pages, zero if none relevant.',
  'Never repeat page numbers. Respect same-material, locality and cannibalisation scope.',
  'Prefer relevant same-locality pages. Nonlocal case studies must not be represented as local.',
  'Anchor text must be an exact, single text-node substring from a paragraph in that H2 section.',
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
  var target=page===0?hub:pages[page-1].url;
  if(!glb_linkCanonical_(target)||glb_linkCanonical_(target)===own)throw Error('Invalid destination.');
  links.push({section:section,url:target,anchor:anchor});
 });
 if((!isHub&&hubCount!==1)||(isHub&&hubCount!==0)||lateral>(isHub?pages.length:limit))
  throw Error('Hub quota or lateral budget not met.');
 return {links:links,cost:Number(api.cost||0)};
}
