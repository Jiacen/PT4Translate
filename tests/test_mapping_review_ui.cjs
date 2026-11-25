const assert = require('node:assert/strict');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const source = readFileSync(join(__dirname, '../sitrain_app/static/mapping-review.js'), 'utf8');

function setup() {
  let html = '';
  const listeners = {}, calls = [];
  const elements = new Map();
  const element = selector => {
    if (!elements.has(selector)) elements.set(selector, {textContent: '', disabled: false,
      classList: {add(){},remove(){},toggle(){}}, querySelectorAll:()=>[],
      addEventListener(event, fn) { listeners[event] = fn; },
      close() { listeners.close?.(); }});
    return elements.get(selector);
  };
  const data = (id='one', completed=0) => ({initialized: true, progress: {initial_pending:2,completed,remaining:2-completed,automatic:5},
    mappings: id ? [{id,media_part:'word/media/image10.emf',word_pages:[9],word_locations:[],candidates:[{slide_number:3}],ppt_file_id:'ppt'}] : [],
    word_name:'chapter.docx',ppt_name:'chapter.pptx',ppt_file_id:'ppt',ppt_version_id:'v',ppt_pages:12,page_index_status:'ready'});
  const ctx = vm.createContext({$:element, form: text=>{html=text;}, esc:String, toast(){}, activeProject:null,
    document:{querySelector:()=>null,querySelectorAll:()=>[]},
    api: async (url,options)=>{calls.push({url,options});return data();},refreshAll:async()=>{},openSyncReview:async()=>{}});
  vm.runInContext(source,ctx);
  return {ctx,data,calls,html:()=>html,state:()=>vm.runInContext('mappingReview',ctx)};
}

test('one pending image, real Word page and explicit confirmation before saving',async()=>{
  const ui=setup();
  await ui.ctx.showPendingMappings('chapter',ui.data());
  assert.match(ui.html(),/Word 第 9 页/);
  assert.match(ui.html(),/初始待确认/);
  assert.match(ui.html(),/id="mapping-confirm"[^>]*disabled/);
  ui.ctx.chooseMappingPage(3);
  assert.equal(ui.calls.length,0);
  let release;
  const pending=new Promise(resolve=>{release=resolve;});
  ui.ctx.api=async (url,options)=>{ui.calls.push({url,options});return url.includes('/decision')?pending:ui.data('two',1);};
  const save=ui.ctx.confirmMappingChoice();
  await ui.ctx.confirmMappingChoice();
  assert.equal(ui.calls.length,1);
  release({}); await save;
  assert.equal(ui.calls.length,2);
  assert.ok(ui.calls.every(call=>!call.url.endsWith('/apply')));
  assert.equal(ui.state().data.mappings[0].id,'two');
  assert.equal(ui.state().selected,null);
  assert.equal(ui.state().data.progress.completed,1);
});

test('failed save keeps current image and choice for retry',async()=>{
  const ui=setup(); await ui.ctx.showPendingMappings('chapter',ui.data());
  ui.ctx.chooseMappingPage('keep');
  ui.ctx.api=async()=>{throw new Error('网络中断');};
  await ui.ctx.confirmMappingChoice();
  assert.equal(ui.state().data.mappings[0].id,'one');
  assert.equal(ui.state().selected,'keep');
  assert.equal(ui.state().busy,false);
  assert.match(ui.html(),/网络中断/);
});

test('page-index response enriches current image without clearing its selection',async()=>{
  const ui=setup(); await ui.ctx.showPendingMappings('chapter',ui.data());
  ui.ctx.chooseMappingPage(3);
  ui.ctx.api=async()=>({status:'ready',pages:{'word/media/image10.emf':[8]},locations:{}});
  await ui.ctx.loadMappingWordPages();
  assert.equal(ui.state().selected,3);
  assert.match(ui.html(),/Word 第 8 页/);
});
