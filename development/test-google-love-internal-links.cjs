/**
 * Offline regression tests for Google Love W1.5D-link integration.
 * Run locally: node development/test-google-love-internal-links.cjs
 * No Apps Script calls, network requests, or spreadsheet writes.
 */
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');
const helper = fs.readFileSync('glb_internal_links.js', 'utf8');
const runner = fs.readFileSync('glb_batch_orchestrator.js', 'utf8');
const context = vm.createContext({console});
vm.runInContext(helper, context);
new vm.Script(runner); // syntax-only: never executes production Apps Script
const hub = 'https://www.abbeyfloorcare.co.uk/home-garden/terrazzo-flooring/';
const clean = 'https://www.abbeyfloorcare.co.uk/home-garden/tile-care/is-it-safe-to-use-off-the-shelf-tile-cleaners-on-terrazzo-a-guide/';
const polish = 'https://www.abbeyfloorcare.co.uk/home-garden/tile-care/terrazzo-tile-polishing-services-transform-your-floors/';
const own = 'https://www.abbeyfloorcare.co.uk/home-garden/floor-care/can-you-use-a-steam-mop-on-terrazzo/';
const cx = [
  'SECTION 1: Origin','Internal link: None',
  'SECTION 6: Cleaning options','Internal link: '+clean,
  'SECTION 9: Professional correction','Internal link: '+polish,
  'SECTION 11: Next steps','Internal link: '+hub,
  'Stage 1.5D complete.'
].join('\n');
const cw = cx.replace('Stage 1.5D complete.', 'Stage 1.5C complete. Waiting for Stage 1.5D.');
assert.strictEqual(context.glb_linkPlanIsCurrent_(cw,cx),true);
assert.strictEqual(context.glb_linkPlanIsCurrent_(cw+' changed',cx),false);
const plan = context.glb_linkPlanFromCx_(cx,hub,'Educational Guide',own);
assert.strictEqual(plan.ready,true);
assert.strictEqual(plan.required.length,3);
assert.deepStrictEqual(Array.from(plan.required,x=>x.section),[6,9,11]);
assert.strictEqual(context.glb_linkAuditHtml_('<a href="/contact">Contact</a>',plan).pass,false);
assert.strictEqual(context.glb_linkAuditHtml_(clean+' '+polish+' '+hub,plan).pass,false);
assert.strictEqual(context.glb_linkAuditHtml_('<script><a href="'+clean+'">X</a></script>',plan).pass,false);
assert.strictEqual(context.glb_linkAuditHtml_('<!-- <a href="'+clean+'">X</a> -->',plan).pass,false);
assert.strictEqual(context.glb_linkAuditHtml_('<footer><a href="'+hub+'">Hub</a></footer>',plan).pass,false);
assert.strictEqual(context.glb_linkAuditHtml_(
  '<p><a href="'+clean+'">Cleaner guide</a></p><p><a href="'+polish+'">Polishing guide</a></p><p><a href="'+hub+'">Terrazzo hub</a></p>',plan).pass,true);
assert.strictEqual(context.glb_linkAuditHtml_(
  '<a href="/home-garden/tile-care/is-it-safe-to-use-off-the-shelf-tile-cleaners-on-terrazzo-a-guide">Cleaner</a>'+
  '<a href="'+polish+'">Polishing</a><a href="'+hub+'">Hub</a>',plan).pass,true);
assert.strictEqual(context.glb_linkPlanFromCx_('',hub,'Educational Guide',own).ready,false);
assert.strictEqual(context.glb_linkPlanFromCx_(cx.replace(hub,'None'),hub,'Educational Guide',own).ready,false);
assert.strictEqual(context.glb_linkPlanFromCx_(cx.replace(clean,own),hub,'Educational Guide',own).ready,false);
assert.strictEqual(context.glb_linkPlanFromCx_(cx.replace(clean,'https://example.com/other'),hub,'Educational Guide',own).ready,false);
assert.strictEqual(context.glb_linkPlanFromCx_(cx.replace(clean,polish),hub,'Educational Guide',own).ready,false);
assert.strictEqual(context.glb_linkPrompt_(plan).includes(clean),true);
assert.strictEqual(context.glb_linkPrompt_(plan).includes(hub),true);
// Mock row/candidate/quota preflight without making any live spreadsheet requests.
const headers = ['Post ID','Article Type','URL','Feeds Hub','Stone Type'];
const lookup = {'Post ID':'48148','Article Type':'Educational Guide','URL':own,
  'Feeds Hub':hub,'Stone Type':'Terrazzo'};
const worksheet = {getLastColumn:()=>headers.length,
  getRange:(row,col)=>({
    getDisplayValues:()=>[headers],
    getValue:()=>row===283 && col===102 ? cx : row===283 && col===101 ? cw : lookup[headers[col-1]]
  })};
context.glb_posts_ = () => worksheet;
context.getMaxLateralLinks = () => 2;
context.buildEnhancedSiloMap = () => [{url:clean},{url:polish}];
const approved = context.glb_linkPreflight_({id:'48148',row:283});
assert.strictEqual(approved.required.length,3);
context.getMaxLateralLinks = () => 1;
assert.throws(() => context.glb_linkPreflight_({id:'48148',row:283}),/more lateral links/);
context.getMaxLateralLinks = () => 2;
context.buildEnhancedSiloMap = () => [{url:clean}];
assert.throws(() => context.glb_linkPreflight_({id:'48148',row:283}),/no longer eligible/);
console.log('PASS: Google Love link governance, prompt, HTML and preflight regression checks');
