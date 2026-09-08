let reviewData=null, reviewSelection=null, reviewDirty=false, reviewSaving=false, reviewRequest=0, reviewWheelDelta=0, reviewWheelLock=false;
let wordReviewUnits=[], wordSelectedUnitId=null;

async function continueReview(){
  if(!leaveReviewEdit())return;
  const id=activeFile,page=activePage;
  try{
    await api(`/api/files/${id}/reopen-review`,{method:'POST'});
    await refreshAll();await selectFile(id,page);
    toast('已建立可编辑副本，之前的发布文件保留');
  }catch(error){toast(error.message,true)}
}

function leaveReviewEdit(){
  if(reviewSaving){toast('正在保存，请稍候');return false}
  if(reviewDirty&&!window.confirm('当前修改尚未保存，确定放弃修改？'))return false;
  reviewDirty=false;return true;
}
const previousSelectFile=selectFile, previousReviewPage=showReviewPage, previousSwitchView=switchView;
selectFile=async function(...args){if(leaveReviewEdit())return previousSelectFile(...args)};
showReviewPage=function(...args){if(leaveReviewEdit())return previousReviewPage(...args)};
switchView=function(name){if(name!=='review'&&!leaveReviewEdit())return;return previousSwitchView(name)};
window.addEventListener('beforeunload',event=>{if(reviewDirty||reviewSaving){event.preventDefault();event.returnValue=''}});

async function showEditablePreview(id,page){
  const ticket=++reviewRequest;
  reviewData=null;reviewSelection=null;
  const preview=$('#preview');
  syncReviewPageNavigation();
  if(activeFileDetail?.kind==='word'){
    return showWordReview(id,page,ticket);
  }
  if(activeFileDetail?.kind!=='powerpoint'){
    preview.innerHTML=`<img src="/api/files/${id}/preview/${page}" alt="文件预览"><p>此文件类型当前仅支持预览。</p>`;
    return;
  }
  try{
    const data=await api(`/api/files/${id}/pages/${page}/objects`);
    if(ticket!==reviewRequest)return;
    const imageUrl=`/api/files/${id}/preview/${page}?version_id=${encodeURIComponent(data.version_id)}`;
    // Do not remove the current slide until the next rendered slide is decoded.
    // This keeps both cached and first-time page switches visually continuous.
    await preloadReviewImage(imageUrl);
    if(ticket!==reviewRequest)return;
    reviewData={...data,fileId:id,page};
    const editable=activeFileDetail.status==='review';
    preview.classList.add('preview-switching');
    preview.innerHTML=`<div class="ppt-review-controls"><span>${editable?'点击文字或图片进行修改；也可从对象列表选择。':'已发布版本 · 只读预览'}</span><select id="ppt-object-picker" aria-label="选择页面对象"><option value="">选择文字或图片…</option>${data.objects.map((o,i)=>`<option value="${i}">${esc(o.kind==='image'?'图片':'文字')} · ${esc(o.name)}</option>`).join('')}</select><button id="ppt-export" class="primary">发布并下载 PPTX</button></div><div class="ppt-canvas" style="aspect-ratio:${data.width}/${data.height}"><img id="ppt-page-image" alt="第 ${page} 页校对版" src="${imageUrl}"><div class="ppt-hit-layer" id="ppt-hit-layer" hidden>${editable?data.objects.map((o,i)=>`<button type="button" class="ppt-hit ${o.kind}" data-index="${i}" title="${esc(o.name)}" aria-label="编辑${esc(o.name)}" style="left:${o.box[0]*100}%;top:${o.box[1]*100}%;width:${o.box[2]*100}%;height:${o.box[3]*100}%;transform:rotate(${o.rotation}deg)"></button>`).join(''):''}</div></div><div id="ppt-render-status" role="status">预览与下载使用同一份 PPTX。修改后请检查文字是否溢出。</div><section id="ppt-editor" class="ppt-editor" hidden></section>`;
    const img=$('#ppt-page-image');
    img.onload=()=>{if(ticket!==reviewRequest)return;$('#ppt-hit-layer').hidden=false;window.requestAnimationFrame(()=>preview.classList.remove('preview-switching'))};
    img.onerror=()=>{if(ticket!==reviewRequest)return;$('#ppt-render-status').textContent='预览生成失败；已保存的 PPTX 仍保留。可重新选择本页重试。'};
    if(img.complete&&img.naturalWidth)img.onload();
    $$('.ppt-hit').forEach(button=>button.onclick=()=>selectPptObject(Number(button.dataset.index)));
    $('#ppt-object-picker').disabled=!editable;
    $('#ppt-object-picker').onchange=event=>{if(event.target.value!=='')selectPptObject(Number(event.target.value))};
    $('#ppt-export').onclick=()=>{if(reviewDirty){toast('请先保存或取消当前修改',true);return}if(!reviewSaving)doPublish(id)};
    installPreviewWheelNavigation(preview);
  }catch(error){if(ticket===reviewRequest)preview.textContent=error.message}
}

async function showWordReview(id,page,ticket){
  const preview=$('#preview');
  preview.innerHTML=`<div class="word-review-layout" style="display:grid;grid-template-columns:minmax(360px,1.15fr) minmax(330px,.85fr);gap:18px;align-items:start"><section id="word-preview-stage">正在加载可编辑 Word 页面…</section><section id="word-review-panel" class="ppt-editor" aria-live="polite">正在加载 Word 校对单元…</section></div>`;
  try{
    const [data,objects]=await Promise.all([api(`/api/files/${id}/units`),api(`/api/files/${id}/pages/${page}/word-objects`)]);
    if(ticket!==reviewRequest)return;
    wordReviewUnits=(data.units||[]).filter(unit=>isMeaningfulWordUnit(unit));
    wordSelectedUnitId=wordReviewUnits.some(unit=>unit.id===wordSelectedUnitId)?wordSelectedUnitId:null;
    const imageUrl=`/api/files/${id}/preview/${page}?version_id=${encodeURIComponent(objects.version_id)}`;
    const stage=$('#word-preview-stage');
    stage.innerHTML=`<div class="ppt-canvas" style="aspect-ratio:${objects.width}/${objects.height}"><img id="word-preview-image" alt="Word 第 ${page} 页预览" src="${imageUrl}"><div class="ppt-hit-layer" id="word-hit-layer" hidden>${objects.objects.map(object=>`<button type="button" class="ppt-hit text" data-unit-id="${object.unit_id}" title="编辑：${esc(object.name)}" aria-label="编辑 Word 文本：${esc(object.name)}" style="left:${object.box[0]*100}%;top:${object.box[1]*100}%;width:${object.box[2]*100}%;height:${object.box[3]*100}%"></button>`).join('')}</div></div><p>点击页面中的文字块即可编辑。嵌入的幻灯片图片将在配套 PPT 校对并完成同步后更新。</p>`;
    const image=$('#word-preview-image');
    image.onload=()=>{$('#word-hit-layer').hidden=false};
    if(image.complete&&image.naturalWidth)image.onload();
    $$('#word-hit-layer .ppt-hit').forEach(button=>button.onclick=()=>{selectWordUnit(button.dataset.unitId);$('#word-review-panel').scrollIntoView({block:'nearest',behavior:'smooth'})});
    // Bind to this page stage, which is recreated on every page change, so one
    // wheel gesture can advance only one Word page.
    installPreviewWheelNavigation(stage);
    renderWordReviewPanel();
  }catch(error){if(ticket===reviewRequest)$('#word-review-panel').textContent=`无法加载 Word 校对内容：${error.message}`}
}

function isMeaningfulWordUnit(unit){
  const source=(unit.source_text||'').replace(/[\s\u240a\u240d]/g,'');
  return source.length>=2&&/[A-Za-z\u4e00-\u9fff]/.test(source);
}

function wordSelectedUnit(){return wordReviewUnits.find(unit=>unit.id===wordSelectedUnitId)||null}

function renderWordReviewPanel(){
  const panel=$('#word-review-panel');if(!panel)return;
  const selected=wordSelectedUnit();
  if(!selected){panel.innerHTML=`<div class="ppt-editor-heading"><b>Word 文本校对</b><span>${wordReviewUnits.length} 条可编辑内容</span></div><p>请先点击左侧 Word 页面中的正文文字。这里会显示对应原文、译文和术语操作。</p>`;return}
  panel.innerHTML=`<div class="ppt-editor-heading"><b>Word 文本校对</b><span>${esc(selected.status)}</span></div><label>原文<div class="source" style="max-height:180px;overflow:auto">${esc(selected.source_text)}</div></label><label>译文<textarea id="word-unit-text" rows="12" maxlength="20000">${esc(selected.human_text??selected.machine_text??'')}</textarea></label><div class="ppt-assist-actions"><button id="word-unit-save">保存修改</button><button id="word-unit-confirm" class="primary">确认并写入翻译记忆</button><button id="word-add-glossary">加入全局术语库</button></div><div id="word-glossary-form"></div><small>术语请填写词或短语；整段内容应保存为译文或确认到翻译记忆。</small>`;
  $('#word-unit-save').onclick=()=>saveWordUnit(false);
  $('#word-unit-confirm').onclick=()=>saveWordUnit(true);
  $('#word-add-glossary').onclick=()=>openWordGlossary(selected);
}

function selectWordUnit(unitId){wordSelectedUnitId=unitId;renderWordReviewPanel()}

async function saveWordUnit(confirm){
  const selected=wordSelectedUnit(), field=$('#word-unit-text');
  if(!selected||!field)return;
  const text=field.value.trim();
  if(!text)return toast('译文不能为空',true);
  try{
    const saved=await api(`/api/units/${selected.id}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({human_text:text,status:confirm?'confirmed':'modified',revision:selected.revision})});
    wordReviewUnits=wordReviewUnits.map(unit=>unit.id===saved.id?saved:unit);
    toast(confirm?'已确认并写入翻译记忆，正在更新 Word 预览':'已保存，正在更新 Word 预览');
    renderWordReviewPanel();
    window.setTimeout(()=>showWordReview(activeFile,activePage,++reviewRequest),800);
  }catch(error){toast(error.message,true)}
}

function openWordGlossary(unit){
  const form=$('#word-glossary-form');if(!form)return;
  form.innerHTML=`<label>英文术语<input id="word-glossary-source" maxlength="300" value="${esc(unit.source_text||'')}"></label><label>中文标准译法<input id="word-glossary-target" maxlength="300" value="${esc(unit.human_text??unit.machine_text??'')}"></label><div class="ppt-assist-actions"><button id="word-glossary-save" class="primary">保存术语</button><button id="word-glossary-cancel">取消</button></div>`;
  $('#word-glossary-cancel').onclick=()=>{form.innerHTML=''};
  $('#word-glossary-save').onclick=async()=>{
    const source=$('#word-glossary-source').value.trim(),target=$('#word-glossary-target').value.trim();
    if(!source||!target)return toast('请填写英文术语和中文标准译法',true);
    try{await api('/api/glossary',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({source,target,mode:'required'})});toast('已加入全局术语库');form.innerHTML=''}catch(error){toast(error.message,true)}
  };
}

function preloadReviewImage(src){
  return new Promise((resolve,reject)=>{
    const image=new Image();
    image.onload=()=>resolve();
    image.onerror=()=>reject(new Error('新页面预览生成失败，请稍后重试'));
    image.src=src;
  });
}

function syncReviewPageNavigation(){
  const current=$('#slide-list .active');
  if(!current)return;
  window.requestAnimationFrame(()=>current.scrollIntoView({block:'nearest',inline:'nearest',behavior:'smooth'}));
}

function installPreviewWheelNavigation(preview){
  // A small accumulated threshold prevents touchpads from turning one gesture
  // into several page changes. Editing controls keep their native scroll/input.
  preview.addEventListener('wheel',event=>{
    if(reviewSaving || event.ctrlKey || event.target.closest('.ppt-editor, textarea, input, select, button'))return;
    event.preventDefault();
    reviewWheelDelta += event.deltaY;
    if(reviewWheelLock || Math.abs(reviewWheelDelta)<70)return;
    const direction=reviewWheelDelta>0?1:-1;
    reviewWheelDelta=0;
    const pages=Math.max(1,Number(activeFileDetail?.preflight?.pages||1));
    const next=Math.min(pages,Math.max(1,activePage+direction));
    if(next===activePage)return;
    reviewWheelLock=true;
    showReviewPage(next);
    window.setTimeout(()=>{reviewWheelLock=false},280);
  },{passive:false});
}

function selectPptObject(index){
  if(!leaveReviewEdit())return;
  const object=reviewData?.objects[index];if(!object)return;
  reviewSelection=object;
  $('#ppt-object-picker').value=String(index);
  $$('.ppt-hit').forEach(el=>el.classList.toggle('selected',Number(el.dataset.index)===index));
  const editor=$('#ppt-editor');editor.hidden=false;
  editor.innerHTML=`<div class="ppt-editor-heading"><b>${esc(object.name)}</b><button id="ppt-cancel">取消</button></div>${object.kind==='text'?`<p>编辑下方文字，保留原段落格式。段内可换行。</p>${object.paragraphs.map((text,i)=>`<label>段落 ${i+1}<textarea class="ppt-paragraph" rows="${Math.min(5,Math.max(2,text.split('\n').length))}" maxlength="20000">${esc(text)}</textarea></label>`).join('')}<section id="ppt-review-assist" class="ppt-review-assist" aria-live="polite">正在检查可用的校对操作…</section>`:'<p>替换当前图片，保留其位置、尺寸和原有裁剪。建议使用相同宽高比的图片。</p><input id="ppt-replacement" type="file" accept="image/png,image/jpeg,image/webp"><small>PNG / JPEG / WebP，最大 10 MB</small>'}<div class="ppt-editor-actions"><span id="ppt-save-status" role="status">尚未修改</span><button id="ppt-save" class="primary">保存到 PPTX</button></div>`;
  editor.querySelectorAll('textarea,input').forEach(input=>input.oninput=()=>{reviewDirty=true;$('#ppt-save-status').textContent='尚未保存'});
  $('#ppt-cancel').onclick=()=>{if(leaveReviewEdit()){editor.hidden=true;reviewSelection=null;$$('.ppt-hit').forEach(el=>el.classList.remove('selected'))}};
  $('#ppt-save').onclick=savePptObject;
  if(object.kind==='text')loadReviewAssist(object);
  editor.scrollIntoView({block:'nearest',behavior:'smooth'});
}

async function loadReviewAssist(object){
  const area=$('#ppt-review-assist');if(!area||!reviewData)return;
  try{
    const context=await api(`/api/files/${reviewData.fileId}/pages/${reviewData.page}/review-context?object_id=${encodeURIComponent(object.id)}`);
    if(reviewSelection!==object||!$('#ppt-review-assist'))return;
    if(!context.trusted){
      area.innerHTML=`<b>校对辅助</b><p>${esc(context.reason)} 如需沉淀为术语，可手工填写英中词对。</p><button id="ppt-add-glossary">加入全局术语库</button>`;
      $('#ppt-add-glossary').onclick=()=>openReviewGlossary({source:'',target:''});
      return;
    }
    const memoryLabel=context.memory_scope==='project'?'加入项目翻译记忆':'加入本地翻译记忆';
    area.innerHTML=`<b>校对辅助</b><p>已可靠关联原文：<code>${esc(context.source)}</code></p><div class="ppt-assist-actions"><button id="ppt-apply-source">应用到本文件的同源文本</button><button id="ppt-save-memory">${memoryLabel}</button><button id="ppt-add-glossary">加入全局术语库</button></div><small>批量应用只匹配原文完全一致且可定位的文本，不会影响其他文件。</small>`;
    $('#ppt-apply-source').onclick=()=>applySameSource(object);
    $('#ppt-save-memory').onclick=()=>saveReviewMemory(object);
    $('#ppt-add-glossary').onclick=()=>openReviewGlossary(context);
  }catch(error){area.textContent=`校对辅助暂不可用：${error.message}`}
}

function reviewActionReady(){
  if(reviewDirty){toast('请先保存当前修改后再执行校对辅助操作',true);return false}
  if(reviewSaving){toast('正在保存，请稍候');return false}
  return Boolean(reviewData&&reviewSelection);
}

async function applySameSource(object){
  if(!reviewActionReady())return;
  try{
    const result=await api(`/api/files/${reviewData.fileId}/pages/${reviewData.page}/apply-same-source`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({base_version:reviewData.version_id,object_id:object.id})});
    toast(`已应用到 ${result.applied} 处同源文本${result.skipped?`；跳过 ${result.skipped} 处无法可靠定位的文本`:''}`);
    await refreshAll();await showEditablePreview(reviewData.fileId,reviewData.page);
  }catch(error){toast(error.message,true)}
}

async function saveReviewMemory(object){
  if(!reviewActionReady())return;
  try{
    const result=await api(`/api/files/${reviewData.fileId}/pages/${reviewData.page}/translation-memory`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({base_version:reviewData.version_id,object_id:object.id})});
    toast(`已加入${result.scope==='project'?'项目':'本地'}翻译记忆`);
  }catch(error){toast(error.message,true)}
}

function openReviewGlossary(context){
  const area=$('#ppt-review-assist');if(!area)return;
  area.innerHTML=`<b>加入全局术语库</b><label>英文术语<input id="review-glossary-source" maxlength="300" value="${esc(context.source||'')}"></label><label>中文标准译法<input id="review-glossary-target" maxlength="300" value="${esc(context.target||'')}"></label><div class="ppt-assist-actions"><button id="review-glossary-save" class="primary">保存术语</button><button id="review-glossary-cancel">取消</button></div><small>请填写术语级词对；整句或段落不建议加入术语库。</small>`;
  $('#review-glossary-cancel').onclick=()=>loadReviewAssist(reviewSelection);
  $('#review-glossary-save').onclick=async()=>{
    const source=$('#review-glossary-source').value.trim(),target=$('#review-glossary-target').value.trim();
    if(!source||!target)return toast('请填写英文术语和中文标准译法',true);
    try{await api('/api/glossary',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({source,target,mode:'required'})});toast('已加入全局术语库');loadReviewAssist(reviewSelection)}catch(error){toast(error.message,true)}
  };
}

async function savePptObject(){
  if(reviewSaving||!reviewSelection||!reviewData)return;
  const {fileId,page,version_id}=reviewData, object=reviewSelection;
  let options,url;
  if(object.kind==='text'){
    options={method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({base_version:version_id,object_id:object.id,paragraphs:$$('.ppt-paragraph').map(el=>el.value)})};
    url=`/api/files/${fileId}/pages/${page}/text`;
  }else{
    const file=$('#ppt-replacement').files[0];if(!file)return toast('请选择替换图片',true);
    if(file.size>10*1024*1024)return toast('图片不能超过 10 MB',true);
    const body=new FormData();body.append('upload',file);body.append('base_version',version_id);body.append('object_id',object.id);
    options={method:'POST',body};url=`/api/files/${fileId}/pages/${page}/image`;
  }
  reviewSaving=true;$('#ppt-save').disabled=true;$('#ppt-save-status').textContent='正在保存…';
  try{
    await api(url,options);reviewDirty=false;
    toast('已保存到 PPTX，正在更新本页预览');
    await showEditablePreview(fileId,page);
  }catch(error){toast(error.message,true);if($('#ppt-save-status'))$('#ppt-save-status').textContent='保存失败，修改仍保留，可重试'}
  finally{reviewSaving=false;if($('#ppt-save'))$('#ppt-save').disabled=false}
}
