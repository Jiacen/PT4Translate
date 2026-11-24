const assert = require('node:assert/strict');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

// Exercise the real handler without a browser or a live translation service.
const source = readFileSync(join(__dirname, '../sitrain_app/static/app.js'), 'utf8');
const handler = source.slice(source.indexOf('async function submitPackageImport()'), source.indexOf('function newChapter()'));

function setup() {
  let resolve, reject;
  const response = new Promise((yes, no) => { resolve = yes; reject = no; });
  const button = {disabled: false, textContent: '扫描并导入', isConnected: true};
  const calls = [], notices = [], opened = [];
  let closed = 0;
  const elements = {
    '#package-import-submit': button,
    '#package-source-lang': {value: 'auto'},
    '#package-target-lang': {value: 'Chinese'},
    '#form-dialog': {close() { closed++; }},
  };
  const context = vm.createContext({
    activeProject: {id: 'project-1'},
    packageSelection: {word: ['word-file'], ppt: ['ppt-file']},
    packageImportInFlight: false,
    FormData: class { append() {} },
    $: selector => elements[selector],
    api: (url, options) => { calls.push({url, options}); return response; },
    toast: (message, bad) => notices.push({message, bad}),
    refreshAll: async () => {},
    openProject: async id => opened.push(id),
  });
  vm.runInContext(handler, context);
  return {context, button, calls, notices, opened, resolve, reject, closed: () => closed};
}

test('repeated clicks send one request and keep the button disabled until completion', async () => {
  const ui = setup();
  const pending = ui.context.submitPackageImport();
  await ui.context.submitPackageImport();
  assert.equal(ui.calls.length, 1);
  assert.equal(ui.button.disabled, true);
  assert.match(ui.button.textContent, /请稍候/);
  ui.resolve({files: [{id: 'word'}, {id: 'ppt'}], skipped_files: [], paired_chapters: 1});
  await pending;
  assert.equal(ui.context.packageImportInFlight, false);
  assert.equal(ui.button.disabled, false);
  assert.equal(ui.closed(), 1);
  assert.deepEqual(ui.opened, ['project-1']);
});

test('an upload error unlocks the button for a deliberate retry', async () => {
  const ui = setup();
  const pending = ui.context.submitPackageImport();
  ui.reject(new Error('上传失败'));
  await pending;
  assert.equal(ui.context.packageImportInFlight, false);
  assert.equal(ui.button.disabled, false);
  assert.equal(ui.notices[0].bad, true);
  ui.context.api = async () => { ui.calls.push({retry: true}); return {files: [], skipped_files: [{id: 'word'}, {id: 'ppt'}]}; };
  await ui.context.submitPackageImport();
  assert.equal(ui.calls.length, 2);
  assert.match(ui.notices.at(-1).message, /已跳过重复导入/);
});

test('finishing a duplicate import preserves navigation and reports the skipped files', async () => {
  const ui = setup();
  const pending = ui.context.submitPackageImport();
  ui.context.activeProject = null;
  ui.button.isConnected = false;
  ui.resolve({files: [], skipped_files: [{id: 'word'}, {id: 'ppt'}]});
  await pending;
  assert.equal(ui.closed(), 0);
  assert.deepEqual(ui.opened, []);
  assert.match(ui.notices[0].message, /所选 2 个文件已在项目中/);
  assert.equal(ui.context.packageImportInFlight, false);
});
