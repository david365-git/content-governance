// No real WordPress, Sheets, Apps Script or paid API calls.
// Run: node control/test_google_love_stop.cjs
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const page = fs.readFileSync(path.join(__dirname, '..', 'glb_BatchSidebar.html'), 'utf8');
const js = page.match(/<script>([\s\S]*?)<\/script>/);
assert.ok(js, 'sidebar script missing');
assert.match(page, /Stop after current article/);
const ids = ['100', '200'];
function runCase(stopPhase) {
  const nodes = new Map();
  function elem(id) {
    if (!nodes.has(id)) nodes.set(id, {value: '', textContent: '', disabled: false,
       style: {}, scrollHeight: 0, scrollTop: 0, className: '',
       appendChild() {}, children: []});
    return nodes.get(id);
  }
  elem('ids').value = ids.join('\n');
  const requests = [];
  let successes = 0;
  let ctx;
  const google = {script: {run: {
    withSuccessHandler(fn) {
      const chain = {
        withFailureHandler() { return chain; },
        glb_progress() { fn([]); },
        glb_prepare() { fn({ids}); },
        glb_runDirectPost(id, phase) { execute(id, phase, fn); },
        glb_runSeoStage(id, phase) { execute(id, phase, fn); },
      };
      return chain;
    }
  }}};
  function execute(id, phase, fn) {
    requests.push(id + ':' + phase);
    if (id === '100' && phase === stopPhase) ctx.stopRun();
    successes++;
    fn({success:true});
  }
  ctx = {
    document: {
      getElementById: elem,
      createElement(tag) { return {tagName: tag, textContent:'', className:'',
          appendChild() {}}; }
    },
    google,
    setInterval() {},
    console,
  };
  vm.createContext(ctx);
  vm.runInContext(js[1], ctx, {timeout:1000});
  ctx.runPosts();
  return {requests, status: elem('status').textContent, successes};
}
for (const phase of ['CI', 'SOL', 'W5', 'W5C']) {
  const result = runCase(phase);
  assert.deepEqual(result.requests, ['100:CI', '100:SOL', '100:W5', '100:W5C'],
    'stop at ' + phase + ' must finish current article and not start second');
  assert.match(result.status, /PAUSED/);
  console.log('PASS stop at ' + phase + ': completed first article, did not start second');
}
console.log('All simulated article-boundary tests passed. No external calls made.');
