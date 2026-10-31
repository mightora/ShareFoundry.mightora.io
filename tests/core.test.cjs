const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '..');
const sandbox = {module:{exports:{}},console};
vm.runInNewContext(fs.readFileSync(path.join(root,'src/core.js'),'utf8'),sandbox);
const core = sandbox.module.exports;

test('creates supported Graph columns with built-in Title not duplicated',()=>{
  const payload=core.graphPayload({name:'Projects',description:'For delivery',columns:[
    {displayName:'Status',type:'choice',choices:'Open, Done',indexed:true},
    {displayName:'Due date',type:'dateTime'},
    {displayName:'Effort',type:'number'}
  ]});
  assert.equal(payload.displayName,'Projects');
  assert.equal(payload.list.template,'genericList');
  assert.equal(payload.columns[0].name,'Status');
  assert.equal(payload.columns[0].choice.choices.join(','),'Open,Done');
  assert.equal(payload.columns[1].dateTime.format,'dateOnly');
  assert.equal(payload.columns.length,3);
});

test('prevents collisions, unsupported types and invalid choices',()=>{
  assert.throws(()=>core.validateProject({name:'A',columns:[{displayName:'A B',type:'text'},{displayName:'AB',type:'text'}]}),/Duplicate/);
  assert.throws(()=>core.validateProject({name:'A',columns:[{displayName:'Bad',type:'script'}]}),/Unsupported/);
  assert.throws(()=>core.validateProject({name:'A',columns:[{displayName:'Status',type:'choice',choices:' '}]}),/Choice/);
});

test('CSV quotes embedded values',()=>{
  const csv=core.csvSchema({name:'Projects',columns:[{displayName:'Notes "Internal"',type:'text'}]});
  assert.ok(csv.includes('"Notes ""Internal"""'));
});

test('formatting patterns use the correct schema and safe accent set',()=>{
  for(const key of ['status','priority','progress','link']){
    const payload=core.formattingPayload(key,'green');
    assert.match(payload.$schema,/column-formatting\.schema\.json/);
    assert.ok(payload.elmType);
    assert.ok(JSON.stringify(payload).length>100);
  }
  assert.throws(()=>core.formattingPayload('nonsense','green'),/Unknown/);
});

test('built editions are self-contained and offline has no remote code requests',()=>{
  const online=fs.readFileSync(path.join(root,'dist/index.html'),'utf8');
  const offline=fs.readFileSync(path.join(root,'dist/offline/sharepoint-maker-toolkit.html'),'utf8');
  assert.ok(online.includes('mightora-header'));
  assert.ok(online.includes('shared-ui@main'));
  assert.ok(online.includes('class="author-section-overlay"'));
  assert.ok(online.includes('class="author-photos"'));
  assert.ok(online.includes('class="author-avatar"'));
  assert.ok(!online.includes('class="author-fallback content"'));
  assert.ok(offline.includes('connect-src \'none\''));
  assert.ok(offline.includes('class="author-avatar author-avatar-placeholder"'));
  assert.ok(offline.includes('class="author-blog-logo author-brand-placeholder"'));
  assert.ok(!offline.includes('<img src="https://cdn.jsdelivr.net/gh/mightora/shared-ui@main/'));
  assert.ok(!offline.includes('<script src='));
  assert.ok(!offline.includes('<link rel="stylesheet" href="https:'));
  assert.ok(offline.includes('data:image/png;base64,'));
  for(const id of ['designer','formatting','exports','privacy'])assert.ok(offline.includes('id="'+id+'"'));
  assert.ok(fs.statSync(path.join(root,'dist/offline/sharepoint-maker-toolkit.zip')).size>1000);
});
