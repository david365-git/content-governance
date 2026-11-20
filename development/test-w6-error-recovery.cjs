/**
 * W6 offline recovery regression test. No Apps Script, WordPress or sheet writes.
 * Run from repository root: node development/test-w6-error-recovery.cjs
 */
const fs = require('fs');
const vm = require('vm');
const assert = require('node:assert/strict');
const code = fs.readFileSync('ce_W6_WordPress_Direct.js', 'utf8');
const sidebar = fs.readFileSync('ce_Sidebar_W6.html', 'utf8');
const sidebarJs = sidebar.match(/<script>([\s\S]*?)<\/script>/)[1];
new vm.Script(code);
new vm.Script(sidebarJs);
const context = vm.createContext({});
vm.runInContext(code, context);
const intended = {
  newH1: 'New title', newHtml: '<p>content</p>',
  newMetaTitle: 'Expected SEO', newMetaDesc: 'Expected description',
  yoastKeyphrase: 'steam mop terrazzo',
  featuredImageUrl: 'https://www.abbeyfloorcare.co.uk/uploads/feature.webp'
};
const before = {status:'future',publishedDate:'2027-01-01T09:00:00'};
let reads = [];
context.ce_w6WpRequest_ = (method,path) => {
  assert.equal(method,'get'); reads.push(path);
  if(path.includes('/media/')) return {source_url:intended.featuredImageUrl};
  return {
    id:48148, status:'publish', date:'2026-11-20T18:21:08',
    modified:'2026-11-20T18:23:05', featured_media:48354,
    title:{raw:intended.newH1}, content:{raw:intended.newHtml},
    meta:{_yoast_wpseo_title:'old title',_yoast_wpseo_metadesc:'old description',
      _yoast_wpseo_focuskw:intended.yoastKeyphrase}
  };
};
const result = context.ce_w6ReadBackAfterError_(48148,'republish',before,intended,Error('HTTP 500'));
assert.equal(result.success,false);
assert.equal(result.retryBlocked,true);
assert.equal(result.published,true);
assert.equal(result.checks.titleMatches,true);
assert.equal(result.checks.htmlMatches,true);
assert.equal(result.checks.metaTitleMatches,false);
assert.equal(result.checks.metaDescriptionMatches,false);
assert.equal(result.checks.focusKeyphraseMatches,true);
assert.equal(result.checks.featuredImageMatches,true);
assert.equal(result.checks.schemaSaved,null); // Not exposed by WP REST.
assert.equal(reads.length,2);
context.ce_w6WpRequest_ = () => {throw Error('GET failed')};
const missing = context.ce_w6ReadBackAfterError_(48148,'update',before,intended,Error('HTTP 500'));
assert.equal(missing.retryBlocked,true);
assert.equal(missing.recoveryCheck,'unavailable');
let logWrites=0;
context.ce_getW6Data=()=>({success:true,postId:48148});
context.ce_w6WordPressPublish=()=>({success:false,retryBlocked:true,message:'partial'});
context.ce_updatePublicationLog=()=>{logWrites++;return {success:true}};
context.ce_updateSiteExport=()=>{logWrites++;return {success:true}};
context.ce_w6SaveConfirmedSchema_=()=>{logWrites++;return {success:true}};
const fail = context.ce_w6DirectPublishAndRecord({postId:48148,mode:'republish'});
assert.equal(fail.retryBlocked,true);
assert.equal(logWrites,0);
console.log('PASS: W6 recovered partial publication without retry or sheet writes; sidebar syntax valid');
