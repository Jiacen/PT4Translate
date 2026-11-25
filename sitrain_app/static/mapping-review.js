let mappingReview = null;

function mappingDialog(html) {
  form(html);
  $('#form-dialog').classList.add('mapping-dialog');
}
function closeMappingReview() { $('#form-dialog').close(); }
$('#form-dialog').addEventListener('close', () => {
  mappingReview = null;
  $('#form-dialog').classList.remove('mapping-dialog');
  document.querySelector('#mapping-image-dialog')?.close();
});
$('#form-dialog').addEventListener('cancel', event => { if (mappingReview?.busy) event.preventDefault(); });

async function syncMap(cid) {
  try {
    const data = await api(`/api/chapters/${cid}/sync/pending?limit=1`);
    if (data.initialized) return showPendingMappings(cid, data);
    form(`<div class="dialog-head"><h2>确认图文对应</h2><button type="button" onclick="closeMappingReview()" aria-label="关闭">×</button></div><p class="notice">系统将比对本章节 Word 图片与 PPT 页面。无法可靠识别的图片由您逐项确认，确认完毕后统一生成同步 Word。</p><button type="button" class="primary full" onclick="startAutomaticSync('${cid}')">开始识别图片</button>`);
  } catch (error) { toast(error.message, true); }
}

async function startAutomaticSync(cid) {
  mappingDialog('<div class="dialog-head"><h2>正在比对图片</h2></div><p class="notice">正在识别 Word 中的幻灯片图片并比对 PPT 页面，请稍候…</p>');
  try {
    await api(`/api/chapters/${cid}/sync/map`, {method: 'POST'});
    await showPendingMappings(cid);
  } catch (error) {
    mappingDialog(`<div class="dialog-head"><h2>暂未完成图片识别</h2><button type="button" onclick="closeMappingReview()">×</button></div><p class="notice">${esc(error.message)}</p><button type="button" onclick="startAutomaticSync('${cid}')">重试</button>`);
  }
}

async function showPendingMappings(cid, existing) {
  try {
    const data = existing || await api(`/api/chapters/${cid}/sync/pending?limit=1`);
    if (!data.initialized) return syncMap(cid);
    mappingReview = {cid, data, selected: null, busy: false, loadingPages: false, error: ''};
    renderMappingReview();
    if (data.progress.remaining && data.page_index_status === 'not_loaded') loadMappingWordPages();
  } catch (error) { toast(error.message, true); }
}

function mappingWordPageLabel(mapping) {
  const pages = mapping.word_pages || [];
  return pages.length ? `Word 第 ${pages.join('、')} 页` : 'Word 页码暂未获取';
}
function mappingPptUrl(page) {
  const data = mappingReview.data;
  return `/api/files/${data.ppt_file_id}/preview/${page}?version_id=${encodeURIComponent(data.ppt_version_id)}`;
}
function mappingChoice(page) {
  const selected = mappingReview.selected === page;
  return `<article class="mapping-candidate ${selected ? 'selected' : ''}"><button type="button" class="mapping-image" onclick="enlargeMappingImage(${page})" aria-label="放大 PPT 第 ${page} 页"><img src="${mappingPptUrl(page)}" alt="PPT 第 ${page} 页预览"><span>点击放大</span></button><label class="mapping-radio"><input type="radio" name="mapping-choice" value="${page}" ${selected ? 'checked' : ''} onchange="chooseMappingPage(${page})">PPT 第 ${page} 页</label></article>`;
}

function renderMappingReview() {
  const state = mappingReview;
  if (!state) return;
  const {data} = state, p = data.progress, mapping = data.mappings[0];
  const heading = `<div class="dialog-head"><h2>${mapping ? '逐项确认图片对应' : '图片对应已确认'}</h2><button type="button" onclick="closeMappingReview()" aria-label="保存进度并关闭" ${state.busy ? 'disabled' : ''}>×</button></div>`;
  const progress = `<div class="mapping-progress" aria-label="确认进度"><div><b>${p.initial_pending}</b><span>初始待确认</span></div><div><b>${p.completed}</b><span>已完成</span></div><div><b>${p.remaining}</b><span>剩余待确认</span></div></div><progress class="mapping-progress-bar" max="${Math.max(p.initial_pending, 1)}" value="${p.initial_pending ? p.completed : 1}"></progress><p class="mapping-meta">另有 ${p.automatic} 张图片已自动匹配。确认进度自动保存。</p>`;
  const error = `<p id="mapping-error" class="mapping-error" role="alert">${esc(state.error)}</p>`;
  if (!mapping) {
    mappingDialog(`${heading}${progress}<p class="notice">所有待确认图片已处理。接下来生成同步 Word，再检查更新后的页面。</p>${error}<div class="mapping-footer"><button type="button" onclick="closeMappingReview()" ${state.busy ? 'disabled' : ''}>稍后继续</button><button type="button" class="primary" onclick="finishMappingReview()" ${state.busy ? 'disabled' : ''}>${state.busy ? '正在生成同步 Word…' : data.needs_apply ? '生成同步 Word' : '查看同步后 Word'}</button></div>`);
    return;
  }
  const candidates = [...new Set((mapping.candidates || []).map(c => c.slide_number))];
  if (Number.isInteger(state.selected) && !candidates.includes(state.selected)) candidates.push(state.selected);
  const pageInfo = mapping.word_locations?.filter(l => l.printed_page !== l.page).map(l => `实际第 ${l.page} 页 / 文档标注 ${l.printed_page}`).join('；');
  const wordImage = `/api/mappings/${mapping.id}/word-image`;
  mappingDialog(`${heading}${progress}<p class="mapping-instruction">先放大对照，再选择对应 PPT 页或保留 Word 原图，最后点击“确认并下一项”。</p><div class="mapping-columns"><section><h3>${esc(mappingWordPageLabel(mapping))}</h3><p class="mapping-meta">${esc(data.word_name)}<br>同步前校对版${pageInfo ? ` · ${esc(pageInfo)}` : ''}</p><button type="button" class="mapping-image mapping-word-image" onclick="enlargeMappingImage(null)" aria-label="放大 Word 原图"><img src="${wordImage}" alt="${esc(mappingWordPageLabel(mapping))}的原图"><span>点击放大 Word 原图</span></button>${!mapping.word_pages?.length ? `<p class="mapping-meta">${state.loadingPages ? '正在读取 Word 实际页码…' : '无法确定页码时不会用图片序号代替。'}</p><button type="button" onclick="loadMappingWordPages()" ${state.loadingPages ? 'disabled' : ''}>${state.loadingPages ? '正在获取页码…' : '重新获取 Word 页码'}</button>` : ''}</section><section><h3>选择对应 PPT 页面</h3><p class="mapping-meta">${esc(data.ppt_name)}</p><div class="mapping-candidates">${candidates.map(mappingChoice).join('')}</div><div class="mapping-manual"><label>其他 PPT 页<input id="mapping-custom-page" type="number" min="1" max="${data.ppt_pages}" placeholder="1–${data.ppt_pages}"></label><button type="button" onclick="chooseOtherMappingPage()">查看并选择</button></div><label class="mapping-keep ${state.selected === 'keep' ? 'selected' : ''}"><input type="radio" name="mapping-choice" ${state.selected === 'keep' ? 'checked' : ''} onchange="chooseMappingPage('keep')">保留 Word 原图</label></section></div>${error}<div class="mapping-footer"><span id="mapping-selection" aria-live="polite">${mappingSelectionLabel()}</span><button type="button" class="primary" id="mapping-confirm" onclick="confirmMappingChoice()" ${state.selected === null || state.busy ? 'disabled' : ''}>${state.busy ? '正在保存…' : p.remaining === 1 ? '确认，完成最后一项' : '确认并下一项'}</button></div>`);
  if (state.busy) $('#dynamic-form').querySelectorAll('button,input').forEach(control => { control.disabled = true; });
}

function mappingSelectionLabel() {
  const selected = mappingReview?.selected;
  return selected === null ? '尚未选择' : selected === 'keep' ? '已选择：保留 Word 原图' : `已选择：PPT 第 ${selected} 页`;
}
function chooseMappingPage(page) {
  if (!mappingReview || mappingReview.busy) return;
  mappingReview.selected = page;
  mappingReview.error = '';
  document.querySelectorAll('.mapping-candidate').forEach(card => card.classList.toggle('selected', Number(card.querySelector('input').value) === page));
  document.querySelector('.mapping-keep')?.classList.toggle('selected', page === 'keep');
  $('#mapping-selection').textContent = mappingSelectionLabel();
  $('#mapping-confirm').disabled = false;
  $('#mapping-error').textContent = '';
}
function chooseOtherMappingPage() {
  const page = Number($('#mapping-custom-page').value);
  if (!Number.isInteger(page) || page < 1 || page > mappingReview.data.ppt_pages) {
    $('#mapping-error').textContent = `请输入 1–${mappingReview.data.ppt_pages} 范围内的 PPT 页码`;
    return;
  }
  mappingReview.selected = page;
  renderMappingReview();
  enlargeMappingImage(page);
}
async function loadMappingWordPages() {
  const state = mappingReview;
  if (!state || state.loadingPages || state.busy) return;
  state.loadingPages = true;
  renderMappingReview();
  try {
    const index = await api(`/api/chapters/${state.cid}/sync/word-pages`, {method: 'POST'});
    if (mappingReview !== state) return;
    state.data.page_index_status = index.status;
    for (const mapping of state.data.mappings) {
      mapping.word_pages = index.pages?.[mapping.media_part] || [];
      mapping.word_locations = index.locations?.[mapping.media_part] || [];
    }
  } catch (error) { state.error = error.message; }
  finally {
    state.loadingPages = false;
    if (mappingReview === state && !state.busy) renderMappingReview();
  }
}
async function confirmMappingChoice() {
  const state = mappingReview;
  if (!state || state.busy || state.selected === null) return;
  const mapping = state.data.mappings[0], selected = state.selected;
  state.busy = true; state.error = ''; renderMappingReview();
  try {
    await api(`/api/mappings/${mapping.id}/decision`, {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(selected === 'keep' ? {keep_original: true} : {slide_number: selected})});
    const data = await api(`/api/chapters/${state.cid}/sync/pending?limit=1`);
    if (mappingReview !== state) return;
    state.data = data; state.selected = null;
  } catch (error) { state.error = `未能完成确认：${error.message}。请重试。`; }
  finally {
    state.busy = false;
    if (mappingReview === state) renderMappingReview();
  }
}
async function finishMappingReview() {
  const state = mappingReview;
  if (!state || state.busy || state.data.progress.remaining) return;
  state.busy = true; state.error = ''; renderMappingReview();
  try {
    const result = await api(`/api/chapters/${state.cid}/sync/apply`, {method: 'POST'});
    closeMappingReview();
    await refreshAll();
    if (activeProject) await openProject(activeProject.id);
    await openSyncReview(result.word_file_id);
  } catch (error) {
    state.error = error.message; state.busy = false;
    if (mappingReview === state) renderMappingReview();
  }
}

let mappingImageZoom = 1;
function enlargeMappingImage(page) {
  const state = mappingReview;
  if (!state || state.busy) return;
  const mapping = state.data.mappings[0];
  let dialog = document.querySelector('#mapping-image-dialog');
  if (!dialog) {
    dialog = document.createElement('dialog'); dialog.id = 'mapping-image-dialog';
    document.body.append(dialog);
  }
  const word = {title: `${mappingWordPageLabel(mapping)} · 原图`, src: `/api/mappings/${mapping.id}/word-image`};
  const items = page === null ? [word] : [word, {title: `PPT 第 ${page} 页 · 已发布版`, src: mappingPptUrl(page)}];
  mappingImageZoom = 1;
  dialog.innerHTML = `<div class="dialog-head"><h2>${page === null ? '放大 Word 原图' : 'Word / PPT 放大对照'}</h2><button type="button" onclick="$('#mapping-image-dialog').close()" aria-label="关闭放大图">×</button></div><div class="mapping-zoom-tools"><button type="button" onclick="zoomMappingImage(-0.25)" aria-label="缩小">−</button><output id="mapping-zoom-value">100%</output><button type="button" onclick="zoomMappingImage(0.25)" aria-label="放大">＋</button><button type="button" onclick="zoomMappingImage(0,true)">适合窗口</button><span>滚轮缩放 · 放大后可滚动查看</span></div><div class="mapping-large-grid ${items.length === 1 ? 'single' : ''}">${items.map(item => `<section><h3>${esc(item.title)}</h3><div class="mapping-large-scroll"><img src="${item.src}" alt="${esc(item.title)}"></div></section>`).join('')}</div><div class="mapping-footer"><button type="button" onclick="$('#mapping-image-dialog').close()">返回选择</button></div>`;
  dialog.querySelectorAll('.mapping-large-scroll').forEach(pane => pane.addEventListener('wheel', event => {
    event.preventDefault(); zoomMappingImage(event.deltaY < 0 ? 0.1 : -0.1);
  }, {passive: false}));
  dialog.showModal();
  dialog.querySelectorAll('img').forEach(img => {
    img.onload = () => fitMappingImages();
  });
  fitMappingImages();
}
function fitMappingImages() {
  document.querySelectorAll('#mapping-image-dialog .mapping-large-scroll img').forEach(img => {
    const pane = img.parentElement;
    if (!img.naturalWidth || !pane.clientHeight) return;
    const fitWidth = Math.min(pane.clientWidth, pane.clientHeight * img.naturalWidth / img.naturalHeight);
    img.style.width = `${Math.max(1, fitWidth * mappingImageZoom)}px`;
    img.style.marginInline = mappingImageZoom <= 1 ? 'auto' : '0';
  });
}
function zoomMappingImage(delta, reset = false) {
  mappingImageZoom = reset ? 1 : Math.max(0.5, Math.min(4, Math.round((mappingImageZoom + delta) * 100) / 100));
  fitMappingImages();
  $('#mapping-zoom-value').textContent = `${Math.round(mappingImageZoom * 100)}%`;
}
