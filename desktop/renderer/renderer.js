const $ = selector => document.querySelector(selector);
const page = $('#page');
let state, catalog, libraryItems = [], libraryConfig = { categories: [] }, catalogMetadata = {}, currentPage = 'dbn';
let pdfjsLibPromise = null, activePdf = null;
let viewHistory = [], viewHistoryIndex = -1;
let pendingDatabaseManifest = null;
let pendingProgramUpdate = null;
let selectedArchicadProjectId = null;
let selectedEdessbProjectId = null;
let edessbConfig = { portalLinks: [], documentTypes: [] };
let edessbInstructions = [];
let selectedEdessbInstructionCategory = null;
let edessbInstructionSearch = '';
let selectedMaterialCategory = 'all';
let materialSearch = '';
const expandedMaterialIds = new Set();
let selectedLibraryCategory = 'dstu';
const escape = value => String(value || '').replace(/[&<>"]/g, char => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' }[char]));
const icon = (name, className = '') => `<svg class="icon ${className}" aria-hidden="true"><use href="icons.svg#${name}"></use></svg>`;
const saveState = async () => window.ekp.saveState(state);
const recent = async (key, value) => { state[key] = [value, ...(state[key] || []).filter(item => item !== value)].slice(0, 8); await saveState(); renderSidebar(); };

function attachmentOpenLabel(filename) {
  const extension = String(filename || '').split('.').pop().toLowerCase();
  if (['doc','docx','rtf'].includes(extension)) return 'Відкрити у Word';
  if (['xls','xlsx','xlsm','csv'].includes(extension)) return 'Відкрити в Excel';
  if (['ppt','pptx'].includes(extension)) return 'Відкрити у PowerPoint';
  if (['zip','rar','7z'].includes(extension)) return 'Відкрити архів';
  return 'Відкрити файл';
}
function confirmDelete(message, title = 'Підтвердження видалення') {
  return new Promise(resolve => {
    const dialog = $('#delete-confirm');
    $('#delete-confirm-title').textContent = title;
    $('#delete-confirm-message').textContent = message;
    let completed = false;
    const finish = answer => {
      if (completed) return;
      completed = true;
      dialog.close();
      resolve(answer);
    };
    $('#delete-confirm-cancel').onclick = () => finish(false);
    $('#delete-confirm-accept').onclick = () => finish(true);
    dialog.oncancel = event => { event.preventDefault(); finish(false); };
    dialog.showModal();
  });
}
function confirmUserDataImport(labels = []) {
  return new Promise(resolve=>{
    const dialog=$('#user-data-import-confirm');let completed=false;
    $('#user-data-import-confirm-message').textContent=`Будуть замінені лише: ${labels.join(', ')}. Всі інші дані залишаться без змін. Рекомендуємо спочатку виконати експорт цих категорій.`;
    const finish=answer=>{if(completed)return;completed=true;dialog.close();resolve(answer);};
    $('#user-data-import-cancel').onclick=()=>finish(false);
    $('#user-data-import-accept').onclick=()=>finish(true);
    dialog.oncancel=event=>{event.preventDefault();finish(false);};
    dialog.showModal();
  });
}

function renderSidebar() {
  document.documentElement.dataset.theme = state.theme || 'light';
  $('#settings-theme-label').textContent = state.theme === 'dark' ? 'Світла тема' : 'Темна тема';
  $('#settings-theme-icon').setAttribute('href', state.theme === 'dark' ? 'icons.svg#sun' : 'icons.svg#moon');
  const activePage = currentPage === 'edessb-instructions' ? 'edessb' : currentPage;
  document.querySelectorAll('.primary-nav button').forEach(button => button.classList.toggle('active', button.dataset.page === activePage));
  const pinnedDbn = catalog.filter(item => (state.pinned_dbn || []).includes(item.path)).map(item => ({ label:`${item.number} — ${item.title}`, kind:'book-open', action: () => openDbn(item), unpin: async () => { state.pinned_dbn=(state.pinned_dbn||[]).filter(path=>path!==item.path); await saveState(); renderSidebar(); } }));
  const dbnMap = new Map(catalog.map(item => [item.path, item]));
  const recently = (state.recent_dbn || []).map(id => dbnMap.get(id)).filter(Boolean).map(item => ({label:item.number,kind:'book-open',action:()=>openDbn(item)}));
  const list = (title, items) => !items.length ? '' : `<div class="side-heading">${title}</div><div class="sidebar-list">${items.slice(0,6).map((item,i) => `<div class="sidebar-list-item"><button data-list="${title}" data-index="${i}" title="${escape(item.label)}">${icon(item.kind, 'small')}<span>${escape(item.label)}</span></button>${item.unpin?`<button class="sidebar-unpin" data-unpin="${i}" title="Прибрати із закріплених">${icon('x','small')}</button>`:''}</div>`).join('')}</div>`;
  const pinned = pinnedDbn; $('#side-lists').innerHTML = list('Закріплені', pinned) + list('Нещодавні', recently);
  document.querySelectorAll('[data-list]').forEach(button => button.onclick = () => (button.dataset.list === 'Закріплені' ? pinned : recently)[Number(button.dataset.index)].action());
  document.querySelectorAll('[data-unpin]').forEach(button => button.onclick = event => { event.stopPropagation(); pinned[Number(button.dataset.unpin)]?.unpin?.(); });
}

function formatCatalogDate(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return 'не вказана';
  const months = ['січня','лютого','березня','квітня','травня','червня','липня','серпня','вересня','жовтня','листопада','грудня'];
  return `${Number(match[3])} ${months[Number(match[2]) - 1]} ${match[1]} року`;
}
function readableError(error, fallback = 'Сталася невідома помилка.') {
  let message = String(error?.message || error || '').trim();
  message = message.replace(/^Error invoking remote method ['"][^'"]+['"]:\s*/i, '');
  message = message.replace(/^Error:\s*/i, '');
  return message || fallback;
}
function readableReleaseNotes(value) {
  let text = String(value || '').trim();
  if (!text) return '';
  for (let pass = 0; pass < 3; pass++) {
    if (!/<\/?[a-z][^>]*>|&(?:lt|gt|amp|quot|#\d+|#x[\da-f]+);/i.test(text)) break;
    const parsed = new DOMParser().parseFromString(text, 'text/html');
    parsed.querySelectorAll('br').forEach(node => node.replaceWith('\n'));
    parsed.querySelectorAll('li').forEach(node => { node.prepend('• '); node.append('\n'); });
    parsed.querySelectorAll('h1,h2,h3,h4,h5,h6,p,div,ul,ol').forEach(node => node.append('\n'));
    const decoded = String(parsed.body.textContent || '').trim();
    if (!decoded || decoded === text) break;
    text = decoded;
  }
  return text
    .replace(/^\s{0,3}#{1,6}\s+/gm, '')
    .replace(/^\s*[-*+]\s+/gm, '• ')
    .replace(/\*\*([^*\n]+)\*\*/g, '$1')
    .replace(/`([^`\n]+)`/g, '$1')
    .replace(/<\/?[a-z][^>]*>/gi, '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
function renderDatabaseUpdate() {
  const update = $('#database-update');
  update.innerHTML = '<span>Дата оновлення бази будівельних норм</span><strong></strong>';
  update.querySelector('strong').textContent = formatCatalogDate(catalogMetadata.updated_at);
}
function showDatabaseUpdateDialog(title, message, canInstall = false) {
  const dialog = $('#database-update-dialog');
  $('#database-update-title').textContent = title;
  $('#database-update-message').textContent = message;
  $('#database-update-install').hidden = !canInstall;
  $('#database-update-cancel').textContent = canInstall ? 'Скасувати' : 'Закрити';
  if (!dialog.open) dialog.showModal();
}
async function checkDatabaseUpdate() {
  const button = $('#check-database-update'); const spinner = button.querySelector('.icon');
  button.disabled = true; spinner.classList.add('spinning');
  try {
    const result = await window.ekp.checkDatabaseUpdate();
    pendingDatabaseManifest = result.manifest;
    if (result.updateAvailable) showDatabaseUpdateDialog('Доступне оновлення', `Встановлена версія: ${result.currentVersion}\nНова версія: ${result.latestVersion}\nДата бази: ${formatCatalogDate(result.databaseDate)}`, true);
    else showDatabaseUpdateDialog('База актуальна', `Установлена найновіша версія ${result.currentVersion}.`);
  } catch (error) { pendingDatabaseManifest = null; showDatabaseUpdateDialog('Не вдалося перевірити базу', readableError(error)); }
  finally { button.disabled = false; spinner.classList.remove('spinning'); }
}
async function installDatabaseUpdate() {
  if (!pendingDatabaseManifest) return;
  const install = $('#database-update-install'); const cancel = $('#database-update-cancel');
  install.disabled = true; cancel.disabled = true;
  const stopProgress = window.ekp.onDatabaseUpdateProgress(value => { $('#database-update-message').textContent = value.message; });
  try {
    const result = await window.ekp.installDatabaseUpdate(pendingDatabaseManifest);
    [catalog, catalogMetadata] = await Promise.all([window.ekp.catalog(), window.ekp.catalogMetadata()]);
    renderDatabaseUpdate(); renderSidebar();
    showDatabaseUpdateDialog('Оновлення завершено', `Базу оновлено до версії ${result.version}.`);
    pendingDatabaseManifest = null;
  } catch (error) { showDatabaseUpdateDialog('Оновлення не встановлено', `${readableError(error)}\n\nЛокальна база залишилася без змін.`); }
  finally { stopProgress(); install.disabled = false; cancel.disabled = false; }
}
function showProgramUpdateDialog(title, message, canInstall = false) {
  const dialog = $('#program-update-dialog');
  $('#program-update-title').textContent = title;
  $('#program-update-message').textContent = message;
  $('#program-update-install').hidden = !canInstall;
  $('#program-update-cancel').textContent = canInstall ? 'Пізніше' : 'Закрити';
  $('#program-update-progress').hidden = true;
  $('#program-update-progress span').style.width = '0%';
  if (!dialog.open) dialog.showModal();
}
function showAvailableProgramUpdate(result) {
  const notes = readableReleaseNotes(result.releaseNotes);
  showProgramUpdateDialog('Доступна нова версія CoDA', '', true);
  const noteLines = notes.split('\n').map(line=>line.trim()).filter(Boolean);
  $('#program-update-message').innerHTML = `<div class="program-update-versions"><div><span>Встановлена версія</span><strong>${escape(result.currentVersion)}</strong></div><div><span>Нова версія</span><strong>${escape(result.latestVersion)}</strong></div></div>${noteLines.length?`<div class="program-release-notes">${noteLines.map(line=>/^(Що нового|Встановлення|Приватність):?$/i.test(line)?`<h3>${escape(line.replace(/:$/,''))}</h3>`:line.startsWith('•')?`<div class="program-release-note">${escape(line)}</div>`:`<p>${escape(line)}</p>`).join('')}</div>`:''}`;
}
async function checkProgramUpdate(manual = false) {
  if (manual) showProgramUpdateDialog('Перевірка оновлень…', 'CoDA перевіряє наявність нової версії.');
  try {
    const result = await window.ekp.checkProgramUpdate();
    if (result.updateAvailable) {
      pendingProgramUpdate = result;
      showAvailableProgramUpdate(result);
    } else if (manual) {
      pendingProgramUpdate = null;
      showProgramUpdateDialog(result.supported ? 'CoDA оновлена' : 'Перевірка у встановленій програмі', result.message || `У вас встановлена найновіша версія ${result.currentVersion}.`);
    }
  } catch (error) {
    pendingProgramUpdate = null;
    if (manual) showProgramUpdateDialog('Не вдалося перевірити оновлення', readableError(error));
  }
}
async function installProgramUpdate() {
  if (!pendingProgramUpdate) return;
  const install = $('#program-update-install');
  const cancel = $('#program-update-cancel');
  const close = $('#program-update-close-x');
  const progress = $('#program-update-progress');
  install.disabled = true; cancel.disabled = true; close.disabled = true;
  install.hidden = true; progress.hidden = false;
  $('#program-update-title').textContent = 'Завантаження оновлення';
  $('#program-update-message').textContent = 'Підготовка завантаження…';
  const stopProgress = window.ekp.onProgramUpdateProgress(value => {
    $('#program-update-message').textContent = value.message;
    $('#program-update-progress span').style.width = `${value.percent || 0}%`;
  });
  try { await window.ekp.installProgramUpdate(); }
  catch (error) {
    showProgramUpdateDialog('Оновлення не встановлено', readableError(error), true);
    install.disabled = false; cancel.disabled = false; close.disabled = false;
  } finally { stopProgress(); }
}
function updateHistoryButtons() {
  $('#nav-back').disabled = viewHistoryIndex <= 0;
  $('#nav-forward').disabled = viewHistoryIndex >= viewHistory.length - 1;
}
function rememberView(view) {
  const previous = viewHistory[viewHistoryIndex];
  if (previous && JSON.stringify(previous) === JSON.stringify(view)) return;
  viewHistory = viewHistory.slice(0, viewHistoryIndex + 1);
  viewHistory.push(view);
  viewHistoryIndex = viewHistory.length - 1;
  updateHistoryButtons();
}
async function showHistoryView(view) {
  if (!['pdf', 'library-pdf', 'edessb-instruction-pdf'].includes(view.type)) disposePdfViewer();
  if (view.type === 'library-pdf') {
    const item = libraryItems.find(entry => entry.id === view.id);
    if (item) { currentPage = 'library'; await renderPdfViewer({ number: item.name, id: item.id, attachment: true, library: true }); }
  } else if (view.type === 'edessb-instruction-pdf') {
    const item = edessbInstructions.find(entry => entry.filename === view.id);
    if (item) { currentPage = 'edessb-instructions'; await renderPdfViewer({ ...item, number: item.title, edessbInstruction: true }); }
  } else if (view.type === 'pdf') {
    const item = catalog.find(entry => entry.path === view.id);
    if (item) { currentPage = 'dbn'; await renderPdfViewer(item); }
  } else {
    currentPage = view.id;
    const renderer = { dbn: renderDbn, library: renderLibrary, excel: renderExcel, edessb: renderEdessb, 'edessb-instructions': renderEdessbInstructions, materials: renderMaterials }[view.id];
    if (renderer) await renderer();
  }
  renderSidebar(); updateHistoryButtons();
}
async function moveHistory(offset) {
  const next = viewHistoryIndex + offset;
  if (next < 0 || next >= viewHistory.length) return;
  viewHistoryIndex = next;
  await showHistoryView(viewHistory[next]);
}
function disposePdfViewer() {
  if (!activePdf) return;
  activePdf.observer?.disconnect();
  activePdf.renderTasks.forEach(task => task?.cancel?.());
  activePdf.textLayers.forEach(layer => layer.cancel());
  if (activePdf.selectionFrame) cancelAnimationFrame(activePdf.selectionFrame);
  activePdf = null;
  document.onselectionchange = null;
  window.onkeydown = null; window.onkeyup = null; window.onblur = null;
}

function head(title) { return `<div class="page-head"><span class="folder">${icon('folder')}</span><span class="page-title">${title}</span></div>`; }
function renderDbn() {
  const categoryOf = item => String(item.category || '').trim() || 'Без категорії';
  const categoryCounts = new Map();
  catalog.forEach(item => categoryCounts.set(categoryOf(item), (categoryCounts.get(categoryOf(item)) || 0) + 1));
  const categories = [...categoryCounts.keys()].sort((a, b) => a.localeCompare(b, 'uk'));
  const categoryButtons = [`<button class="active" data-category-index="-1"><span>Усі категорії</span><strong>${catalog.length}</strong></button>`, ...categories.map((category, index) => `<button data-category-index="${index}" title="${escape(category)}"><span>${escape(category)}</span><strong>${categoryCounts.get(category)}</strong></button>`)].join('');
  const rows = catalog.map(item => `<tr data-path="${escape(item.path)}" data-category="${escape(categoryOf(item))}"><td class="pin">${(state.pinned_dbn||[]).includes(item.path)?icon('star','small star'):''}</td><td>${escape(item.number)}</td><td>${escape(item.title)}</td></tr>`).join('');
  page.innerHTML = head('Державні будівельні норми') + `<div class="page-content dbn-page"><h1 class="hero">Державні будівельні норми</h1><p class="subtitle">Локальна бібліотека • ${catalog.length} документів</p><div class="dbn-layout"><section class="dbn-catalog"><table class="table"><thead><tr><th></th><th>Номер</th><th>Назва</th></tr></thead><tbody>${rows}</tbody></table></section><aside class="dbn-categories"><div class="dbn-categories-title">Категорії</div><div class="dbn-category-list">${categoryButtons}</div></aside></div><div class="dbn-floating-controls" role="search"><input id="dbn-search" class="input" placeholder="Пошук за номером або назвою"><button id="pin-dbn" class="button">${icon('pin')}<span>Закріпити / відкріпити</span></button></div></div>`;
  let selectedRow = null; document.querySelectorAll('.table tbody tr').forEach(row => { row.onclick=()=>{ document.querySelectorAll('.table tbody tr').forEach(item=>item.style.background=''); row.style.background='var(--hover)'; selectedRow=row; }; row.ondblclick=()=>openDbn(catalog.find(item=>item.path===row.dataset.path)); });
  let selectedCategory = null;
  const applyDbnFilters = () => { const query=$('#dbn-search').value.trim().toLocaleLowerCase('uk'); document.querySelectorAll('.table tbody tr').forEach(row => { const searchMatches=!query||row.innerText.toLocaleLowerCase('uk').includes(query); const categoryMatches=!selectedCategory||row.dataset.category===selectedCategory; row.hidden=!(searchMatches&&categoryMatches); }); };
  $('#dbn-search').oninput = applyDbnFilters;
  document.querySelectorAll('[data-category-index]').forEach(button => button.onclick = () => { const index=Number(button.dataset.categoryIndex); selectedCategory=index<0?null:categories[index]; document.querySelectorAll('[data-category-index]').forEach(item=>item.classList.toggle('active',item===button)); applyDbnFilters(); });
  $('#pin-dbn').onclick = async () => { if(!selectedRow) return; const path=selectedRow.dataset.path; const set=new Set(state.pinned_dbn||[]); set.has(path)?set.delete(path):set.add(path); state.pinned_dbn=[...set]; await saveState(); renderDbn(); renderSidebar(); };
}
async function loadPdfJs() {
  if (!pdfjsLibPromise) {
    pdfjsLibPromise = import('../../node_modules/pdfjs-dist/build/pdf.mjs').then(module => {
      module.GlobalWorkerOptions.workerSrc = new URL('../../node_modules/pdfjs-dist/build/pdf.worker.mjs', window.location.href).href;
      return module;
    });
  }
  return pdfjsLibPromise;
}

function pdfDomTextMap(layer) {
  const nodes = [], walker = document.createTreeWalker(layer, NodeFilter.SHOW_TEXT);
  let text = '', node;
  while ((node = walker.nextNode())) { const start = text.length; text += node.nodeValue || ''; nodes.push({ node, start, end:text.length }); }
  return { text, nodes };
}
function pdfTextMap(layer) {
  const page=Number(layer?.closest('[data-pdf-page]')?.dataset.pdfPage),mapped=activePdf?.pageTextMaps?.get(page);
  if(!mapped)return pdfDomTextMap(layer);
  const nodes=[];
  mapped.items.forEach(item=>{let cursor=item.start,walker=document.createTreeWalker(item.element,NodeFilter.SHOW_TEXT),node;while((node=walker.nextNode())){const start=cursor;cursor+=node.nodeValue?.length||0;nodes.push({node,element:item.element,start,end:cursor});}});
  return{text:mapped.text,nodes,items:mapped.items};
}
function pdfTextLayerMap(textLayer) {
  const items=[];let text='';
  textLayer.textContentItemsStr.forEach((value,index)=>{const start=text.length;text+=value||'';const element=textLayer.textDivs[index];if(element)items.push({element,text:value||'',start,end:text.length});});
  return{text,items};
}
function pdfBoundaryOffset(map, container, offset) {
  if (container.nodeType === Node.TEXT_NODE) { const item = map.nodes.find(entry => entry.node === container); return item ? item.start + Math.min(offset, item.end-item.start) : -1; }
  const first = container.childNodes?.[offset];
  if (first) { const item = map.nodes.find(entry => entry.node === first || first.contains?.(entry.node)); if (item) return item.start; }
  const previous = container.childNodes?.[offset-1];
  if (previous) { const items = map.nodes.filter(entry => entry.node === previous || previous.contains?.(entry.node)); if (items.length) return items.at(-1).end; }
  return -1;
}
function normalizePdfRects(rects) {
  const lines=[];
  rects.filter(rect=>rect.width>.5&&rect.height>.5).sort((a,b)=>a.top-b.top||a.left-b.left).forEach(rect=>{
    const center=(rect.top+rect.bottom)/2;
    let line=lines.find(item=>Math.abs(center-item.center)<=Math.max(1.5,Math.min(rect.height,item.height)*.35));
    if(!line){line={top:rect.top,bottom:rect.bottom,center,height:rect.height,segments:[]};lines.push(line);}
    else{line.top=Math.min(line.top,rect.top);line.bottom=Math.max(line.bottom,rect.bottom);line.center=(line.top+line.bottom)/2;line.height=line.bottom-line.top;}
    line.segments.push({left:rect.left,right:rect.right});
  });
  return lines.flatMap(line=>{
    const gap=Math.max(1.5,line.height*.28),merged=[];
    line.segments.sort((a,b)=>a.left-b.left).forEach(segment=>{const current=merged.at(-1);if(current&&segment.left<=current.right+gap)current.right=Math.max(current.right,segment.right);else merged.push({...segment});});
    return merged.map(segment=>({left:segment.left,top:line.top,width:segment.right-segment.left,height:line.bottom-line.top}));
  });
}
function pdfRectsFromOffsets(shell,layer,start,end) {
  if(!shell||!layer||end<=start)return[];
  const map=pdfTextMap(layer),parent=shell.getBoundingClientRect(),rects=[];
  map.nodes.forEach(item=>{
    let from=Math.max(start,item.start),to=Math.min(end,item.end);if(to<=from)return;
    const selectedText=map.text.slice(from,to),leading=selectedText.length-selectedText.trimStart().length,trailing=selectedText.length-selectedText.trimEnd().length;from+=leading;to-=trailing;if(to<=from)return;
    const range=document.createRange();range.setStart(item.node,from-item.start);range.setEnd(item.node,to-item.start);const itemBounds=(item.element||item.node.parentElement).getBoundingClientRect();
    [...range.getClientRects()].forEach(rect=>{const left=Math.max(0,rect.left-parent.left,itemBounds.left-parent.left),top=Math.max(0,rect.top-parent.top,itemBounds.top-parent.top),right=Math.min(parent.width,rect.right-parent.left,itemBounds.right-parent.left),bottom=Math.min(parent.height,rect.bottom-parent.top,itemBounds.bottom-parent.top);if(right-left>.5&&bottom-top>.5)rects.push({left,top,right,bottom,width:right-left,height:bottom-top});});
  });
  return normalizePdfRects(rects);
}
function pdfNormalizeGeometry(shell,rects){const width=shell.clientWidth||1,height=shell.clientHeight||1;return rects.map(rect=>({x:rect.left/width,y:rect.top/height,width:rect.width/width,height:rect.height/height}));}
function pdfRectsFromGeometry(shell,geometry){const width=shell.clientWidth||1,height=shell.clientHeight||1;return Array.isArray(geometry)?geometry.map(rect=>({left:rect.x*width,top:rect.y*height,width:rect.width*width,height:rect.height*height})):[];}
function pdfDrawRects(shell,rects,className,bookmarkId='') {
  const overlay=shell?.querySelector('.pdf-markup-layer');if(!overlay)return;
  rects.forEach(rect=>{if(rect.width<.5||rect.height<.5)return;const mark=document.createElement('span');mark.className=`pdf-mark ${className}`;if(bookmarkId)mark.dataset.bookmarkId=bookmarkId;Object.assign(mark.style,{left:`${rect.left}px`,top:`${rect.top}px`,width:`${rect.width}px`,height:`${rect.height}px`});overlay.append(mark);});
}
function currentPdfSelections() {
  if(!activePdf)return[];const selection=window.getSelection();if(!selection||selection.isCollapsed||selection.rangeCount!==1)return[];const range=selection.getRangeAt(0);
  const startLayer=(range.startContainer.nodeType===Node.ELEMENT_NODE?range.startContainer:range.startContainer.parentElement)?.closest?.('.textLayer'),endLayer=(range.endContainer.nodeType===Node.ELEMENT_NODE?range.endContainer:range.endContainer.parentElement)?.closest?.('.textLayer');if(!startLayer||!endLayer)return[];
  const startPage=Number(startLayer.closest('[data-pdf-page]')?.dataset.pdfPage),endPage=Number(endLayer.closest('[data-pdf-page]')?.dataset.pdfPage);if(!startPage||!endPage||endPage<startPage)return[];
  return[...document.querySelectorAll('.textLayer')].map(layer=>{const shell=layer.closest('[data-pdf-page]'),page=Number(shell?.dataset.pdfPage);if(!shell||page<startPage||page>endPage)return null;const map=pdfTextMap(layer),start=layer===startLayer?pdfBoundaryOffset(map,range.startContainer,range.startOffset):0,end=layer===endLayer?pdfBoundaryOffset(map,range.endContainer,range.endOffset):map.text.length;return start>=0&&end>start?{selection,range,layer,shell,map,start,end,page}:null;}).filter(Boolean);
}
function renderPdfNativeSelection() {
  document.querySelectorAll('.pdf-mark-selection').forEach(mark=>mark.remove());if(!activePdf)return;
  const selected=currentPdfSelections();activePdf.nativeSelection=selected.map(item=>({page:item.page,start:item.start,end:item.end}));
  selected.forEach(item=>pdfDrawRects(item.shell,pdfRectsFromOffsets(item.shell,item.layer,item.start,item.end),'pdf-mark-selection'));
}
function queuePdfNativeSelectionRender(){if(!activePdf||activePdf.selectionFrame)return;activePdf.selectionFrame=requestAnimationFrame(()=>{if(activePdf)activePdf.selectionFrame=0;renderPdfNativeSelection();});}
function pdfAppendInlineText(element,text,className='') {
  if(!className){element.append(document.createTextNode(text));return;}
  const leading=text.length-text.trimStart().length,trailing=text.length-text.trimEnd().length,end=text.length-trailing;
  if(leading)element.append(document.createTextNode(text.slice(0,leading)));
  if(end>leading){const mark=document.createElement('mark');mark.className=`pdf-inline-mark ${className}`;mark.textContent=text.slice(leading,end);element.append(mark);}
  if(trailing)element.append(document.createTextNode(text.slice(end)));
}
function renderPdfInlineMarks(pageNumber) {
  const map=activePdf?.pageTextMaps?.get(pageNumber);if(!map)return;
  const marks=[...activePdf.bookmarks.flatMap((item,index)=>pdfBookmarkSegments(item).filter(segment=>segment.page===pageNumber).map(segment=>({start:segment.start,end:segment.end,className:`pdf-inline-${item.color}`,priority:10-index/10000}))),...activePdf.searchMatches.filter(item=>item.page===pageNumber).map(item=>{const index=activePdf.searchMatches.indexOf(item);return{start:item.start,end:item.end,className:index===activePdf.searchIndex?'pdf-inline-search-current':'pdf-inline-search',priority:index===activePdf.searchIndex?30:20};})];
  map.items.forEach(item=>{
    const local=marks.filter(mark=>mark.end>item.start&&mark.start<item.end),boundaries=new Set([0,item.text.length]);local.forEach(mark=>{boundaries.add(Math.max(0,mark.start-item.start));boundaries.add(Math.min(item.text.length,mark.end-item.start));});const points=[...boundaries].sort((a,b)=>a-b);item.element.replaceChildren();
    for(let index=0;index<points.length-1;index++){const from=points[index],to=points[index+1],winner=local.filter(mark=>mark.start<item.start+to&&mark.end>item.start+from).sort((a,b)=>b.priority-a.priority)[0];pdfAppendInlineText(item.element,item.text.slice(from,to),winner?.className||'');}
  });
}
function renderPdfPageMarks(pageNumber) {
  if(!activePdf)return; const shell=document.querySelector(`[data-pdf-page="${pageNumber}"]`), layer=shell?.querySelector('.textLayer'), overlay=shell?.querySelector('.pdf-markup-layer'); if(!layer||!overlay)return;
  overlay.replaceChildren();
  renderPdfInlineMarks(pageNumber);
  if(activePdf.nativeSelection?.some(item=>item.page===pageNumber)){const selected=currentPdfSelections().find(item=>item.page===pageNumber);if(selected)pdfDrawRects(shell,pdfRectsFromOffsets(shell,layer,selected.start,selected.end),'pdf-mark-selection');}
}
const pdfBookmarkColor=color=>({yellow:'#d7b900',green:'#39a85d',blue:'#278fdc',pink:'#d65491',orange:'#df7e20'})[color]||'#d7b900';
function pdfBookmarkSegments(item){const segments=Array.isArray(item?.segments)&&item.segments.length?item.segments:[item];return segments.map(segment=>({page:Number(segment?.page),start:Number(segment?.start),end:Number(segment?.end),geometry:Array.isArray(segment?.geometry)?segment.geometry:[]})).filter(segment=>Number.isInteger(segment.page)&&segment.page>0&&Number.isInteger(segment.start)&&segment.start>=0&&Number.isInteger(segment.end)&&segment.end>segment.start).sort((a,b)=>a.page-b.page);}
function pdfBookmarkPageLabel(item){const segments=pdfBookmarkSegments(item),first=segments[0]?.page||Number(item?.page)||1,last=segments.at(-1)?.page||first;return first===last?`Сторінка ${first}`:`Сторінки ${first}–${last}`;}
function renderPdfBookmarksPanel() {
  if(!activePdf||!$('#pdf-bookmarks-list'))return;
  $('#pdf-bookmarks-list').innerHTML=activePdf.bookmarks.length?activePdf.bookmarks.map(item=>`<button class="pdf-bookmark-item${activePdf.activeBookmarkId===item.id?' active':''}" data-pdf-bookmark="${escape(item.id)}" style="--bookmark-color:${pdfBookmarkColor(item.color)}" title="${escape(item.text)}"><strong>${escape(item.label)}</strong><span>${pdfBookmarkPageLabel(item)} · ${escape(item.text)}</span><span class="pdf-bookmark-delete" data-delete-pdf-bookmark="${escape(item.id)}" title="Видалити">${icon('trash')}</span></button>`).join(''):'<div class="pdf-bookmark-empty">Виділіть текст у документі, оберіть колір і збережіть текстову закладку.</div>';
  document.querySelectorAll('[data-pdf-bookmark]').forEach(button=>button.onclick=async()=>{const item=activePdf?.bookmarks.find(entry=>entry.id===button.dataset.pdfBookmark),segments=pdfBookmarkSegments(item);if(!item||!segments.length)return;activePdf.activeBookmarkId=item.id;await scrollToPdfPage(segments[0].page);segments.forEach(segment=>renderPdfPageMarks(segment.page));renderPdfBookmarksPanel();});
  document.querySelectorAll('[data-delete-pdf-bookmark]').forEach(button=>button.onclick=async event=>{event.stopPropagation();const id=button.dataset.deletePdfBookmark;await window.ekp.deletePdfBookmark({documentId:activePdf.documentId,id});activePdf.bookmarks=activePdf.bookmarks.filter(item=>item.id!==id);document.querySelectorAll('[data-pdf-page]').forEach(shell=>renderPdfPageMarks(Number(shell.dataset.pdfPage)));renderPdfBookmarksPanel();});
}
function pdfSearchResultWord(count){const mod100=count%100,mod10=count%10;return mod10===1&&mod100!==11?'збіг':mod10>=2&&mod10<=4&&(mod100<12||mod100>14)?'збіги':'збігів';}
function renderPdfSearchPanel() {
  if(!activePdf||!$('#pdf-search-results'))return;
  const query=activePdf.searchQuery||'',matches=activePdf.searchMatches;
  $('#pdf-search-summary').textContent=query?`${matches.length} ${pdfSearchResultWord(matches.length)} · «${query}»`:'';
  $('#pdf-search-results').innerHTML=matches.length?matches.map((item,index)=>`<button class="pdf-search-result${index===activePdf.searchIndex?' active':''}" data-pdf-search-result="${index}"><span class="pdf-search-result-page">Сторінка ${item.page}</span><span class="pdf-search-result-context">${escape(item.before)}<strong>${escape(item.term)}</strong>${escape(item.after)}</span></button>`).join(''):'<div class="pdf-bookmark-empty">За цим запитом збігів не знайдено.</div>';
  document.querySelectorAll('[data-pdf-search-result]').forEach(button=>button.onclick=()=>goToPdfSearchMatch(Number(button.dataset.pdfSearchResult)));
}
function showPdfSearchPanel() {
  $('#pdf-bookmarks-view').hidden=true;$('#pdf-search-view').hidden=false;renderPdfSearchPanel();
}
function closePdfSearch() {
  if(!activePdf)return;activePdf.searchToken++;activePdf.searchMatches=[];activePdf.searchIndex=-1;activePdf.searchQuery='';$('#pdf-search-input').value='';$('#pdf-search-count').textContent='';$('#pdf-search-view').hidden=true;$('#pdf-bookmarks-view').hidden=false;document.querySelectorAll('[data-pdf-page]').forEach(shell=>renderPdfPageMarks(Number(shell.dataset.pdfPage)));renderPdfBookmarksPanel();
}
function clearPdfSelectionEditor(){if(activePdf)activePdf.selection=null;if($('#pdf-selection-editor'))$('#pdf-selection-editor').hidden=true;}
function capturePdfSelection(){
  const selected=currentPdfSelections(),segments=[],texts=[];if(!selected.length)return;
  selected.forEach(item=>{let{start,end}=item;const raw=item.map.text.slice(start,end),leading=raw.length-raw.trimStart().length,trailing=raw.length-raw.trimEnd().length;start+=leading;end-=trailing;const text=item.map.text.slice(start,end).trim();if(!text)return;const geometry=pdfNormalizeGeometry(item.shell,pdfRectsFromOffsets(item.shell,item.layer,start,end));if(!geometry.length)return;segments.push({page:item.page,start,end,geometry});texts.push(text);});
  if(!segments.length)return;const text=texts.join('\n\n').slice(0,20000),first=segments[0],last=segments.at(-1);activePdf.selection={page:first.page,endPage:last.page,start:first.start,end:first.end,geometry:first.geometry,segments,text,color:'yellow'};$('#pdf-selection-caption').textContent=segments.length===1?'Виділений текст':`Виділений текст · сторінки ${first.page}–${last.page}`;$('#pdf-selection-preview').textContent=text;$('#pdf-bookmark-label').value='';$('#pdf-selection-editor').hidden=false;document.querySelectorAll('.pdf-color').forEach(button=>button.classList.toggle('active',button.dataset.color==='yellow'));
}
async function savePdfSelectionBookmark(){if(!activePdf?.selection)return;const bookmark=await window.ekp.addPdfBookmark({documentId:activePdf.documentId,...activePdf.selection,label:$('#pdf-bookmark-label').value});activePdf.bookmarks.unshift(bookmark);activePdf.activeBookmarkId=bookmark.id;clearPdfSelectionEditor();window.getSelection()?.removeAllRanges();pdfBookmarkSegments(bookmark).forEach(segment=>renderPdfPageMarks(segment.page));renderPdfBookmarksPanel();}
async function indexPdfText(){if(!activePdf?.document)return;const session=activePdf;for(let page=1;page<=session.document.numPages;page++){if(session.pageSearchIndexes.has(page))continue;const pdfPage=await session.document.getPage(page),content=await pdfPage.getTextContent();if(activePdf!==session)return;const pageIndex=PdfSearchIndex.buildPageIndex(content);session.pageTexts.set(page,pageIndex.rawText);session.pageSearchIndexes.set(page,pageIndex);}}
async function goToPdfSearchMatch(index){if(!activePdf?.searchMatches.length)return;activePdf.searchIndex=(index+activePdf.searchMatches.length)%activePdf.searchMatches.length;const match=activePdf.searchMatches[activePdf.searchIndex];await scrollToPdfPage(match.page);$('#pdf-search-count').textContent=`${activePdf.searchIndex+1} / ${activePdf.searchMatches.length}`;document.querySelectorAll('[data-pdf-page]').forEach(shell=>renderPdfPageMarks(Number(shell.dataset.pdfPage)));renderPdfSearchPanel();requestAnimationFrame(()=>document.querySelector(`[data-pdf-page="${match.page}"] .pdf-inline-search-current`)?.scrollIntoView({behavior:'smooth',block:'center',inline:'center'}));}
async function performPdfSearch(){if(!activePdf?.document)return;const query=PdfSearchIndex.normalizeQuery($('#pdf-search-input').value);if(!query){$('#pdf-search-input').focus();return;}const token=++activePdf.searchToken;activePdf.searchMatches=[];activePdf.searchIndex=-1;activePdf.searchQuery=query;$('#pdf-search-count').textContent='Пошук…';showPdfSearchPanel();$('#pdf-search-summary').textContent=`Пошук «${query}»…`;$('#pdf-search-results').innerHTML='<div class="pdf-bookmark-empty">Переглядаємо текст документа…</div>';document.querySelectorAll('[data-pdf-page]').forEach(shell=>renderPdfPageMarks(Number(shell.dataset.pdfPage)));await indexPdfText();if(!activePdf||token!==activePdf.searchToken)return;for(const[page,pageIndex]of activePdf.pageSearchIndexes){const remaining=5000-activePdf.searchMatches.length;if(remaining<=0)break;activePdf.searchMatches.push(...PdfSearchIndex.findMatches(pageIndex,query,remaining).map(match=>({page,...match})));}if(!activePdf.searchMatches.length){$('#pdf-search-count').textContent='0 збігів';renderPdfSearchPanel();return;}await goToPdfSearchMatch(0);}

async function renderPdfPage(pageNumber) {
  if (!activePdf?.document || activePdf.rendered.has(pageNumber)) return;
  const shell = document.querySelector(`[data-pdf-page="${pageNumber}"]`);
  if (!shell || activePdf.renderTasks.has(pageNumber)) return;
  const session = activePdf;
  session.renderTasks.set(pageNumber, null);
  try {
    const pdfPage = await session.document.getPage(pageNumber);
    if (activePdf !== session || !document.body.contains(shell)) return;
    const viewport = pdfPage.getViewport({ scale: session.scale });
    const canvas = shell.querySelector('canvas');
    const context = canvas.getContext('2d');
    const pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
    canvas.width = Math.floor(viewport.width * pixelRatio);
    canvas.height = Math.floor(viewport.height * pixelRatio);
    canvas.style.width = `${Math.floor(viewport.width)}px`;
    canvas.style.height = `${Math.floor(viewport.height)}px`;
    shell.style.width = `${Math.floor(viewport.width)}px`;
    shell.style.height = `${Math.floor(viewport.height)}px`;
    shell.style.minHeight = `${Math.floor(viewport.height)}px`;
    shell.style.setProperty('--scale-factor', String(viewport.scale));
    shell.style.setProperty('--user-unit', String(viewport.userUnit));
    const task = pdfPage.render({ canvasContext: context, viewport, transform: pixelRatio === 1 ? null : [pixelRatio, 0, 0, pixelRatio, 0, 0] });
    session.renderTasks.set(pageNumber, task);
    await task.promise;
    if (activePdf !== session) return;
    const textLayerElement = shell.querySelector('.textLayer');
    const textContent = await pdfPage.getTextContent();
    if (activePdf !== session || !document.body.contains(textLayerElement)) return;
    const textLayer = new session.pdfjs.TextLayer({ textContentSource: textContent, container: textLayerElement, viewport });
    session.textLayers.set(pageNumber, textLayer);
    await textLayer.render();
    const textMap=pdfTextLayerMap(textLayer),pageIndex=PdfSearchIndex.buildPageIndex(textContent);session.pageTextMaps.set(pageNumber,textMap);session.pageTexts.set(pageNumber,pageIndex.rawText);session.pageSearchIndexes.set(pageNumber,pageIndex);
    renderPdfPageMarks(pageNumber);
    if (activePdf === session) session.rendered.add(pageNumber);
  } catch (error) {
    if (error?.name !== 'RenderingCancelledException') throw error;
  } finally { session.renderTasks.delete(pageNumber); }
}

function updateCurrentPdfPage() {
  if (!activePdf?.document) return;
  const stage = $('#pdf-stage');
  const center = stage.getBoundingClientRect().top + stage.clientHeight / 2;
  let nearest = activePdf.pageNumber, distance = Infinity;
  document.querySelectorAll('[data-pdf-page]').forEach(shell => {
    const rect = shell.getBoundingClientRect();
    const candidate = Math.abs(rect.top + rect.height / 2 - center);
    if (candidate < distance) { distance = candidate; nearest = Number(shell.dataset.pdfPage); }
  });
  activePdf.pageNumber = nearest;
  $('#pdf-page-number').value = nearest;
}

async function scrollToPdfPage(pageNumber, behavior = 'smooth') {
  if (!activePdf?.document) return;
  activePdf.pageNumber = Math.max(1, Math.min(activePdf.document.numPages, pageNumber));
  await renderPdfPage(activePdf.pageNumber);
  document.querySelector(`[data-pdf-page="${activePdf.pageNumber}"]`)?.scrollIntoView({ behavior, block: 'start' });
  $('#pdf-page-number').value = activePdf.pageNumber;
}

function buildPdfPages() {
  activePdf.textLayers.forEach(layer => layer.cancel());
  activePdf.textLayers.clear();
  activePdf.pageTextMaps.clear();
  const placeholderWidth=Math.floor((activePdf.defaultPageSize?.width||612)*activePdf.scale),placeholderHeight=Math.floor((activePdf.defaultPageSize?.height||792)*activePdf.scale);
  const pages = Array.from({ length: activePdf.document.numPages }, (_, index) => `<section class="pdf-page-shell" data-pdf-page="${index + 1}" style="width:${placeholderWidth}px;height:${placeholderHeight}px"><canvas></canvas><div class="pdf-markup-layer" aria-hidden="true"></div><div class="textLayer" aria-label="Текст сторінки ${index + 1}"></div><span class="pdf-page-label">${index + 1}</span></section>`).join('');
  $('#pdf-stage').innerHTML = `<div class="pdf-pages${activePdf.layout === 'spread' ? ' two-page' : ''}">${pages}</div>`;
  activePdf.observer?.disconnect();
  activePdf.observer = new IntersectionObserver(entries => entries.forEach(entry => { if (entry.isIntersecting) renderPdfPage(Number(entry.target.dataset.pdfPage)); }), { root: $('#pdf-stage'), rootMargin: '700px 0px', threshold: .01 });
  document.querySelectorAll('[data-pdf-page]').forEach(shell => activePdf.observer.observe(shell));
  $('#pdf-stage').onscroll = updateCurrentPdfPage;
}

async function setPdfLayout(layout) {
  if (!activePdf?.document) return;
  if (activePdf.layout === layout) return;
  const pageNumber = activePdf.pageNumber;
  if (layout === 'spread') {
    activePdf.singleScale = activePdf.scale;
    activePdf.layout = 'spread';
    const firstPage = await activePdf.document.getPage(1);
    const baseWidth = firstPage.getViewport({ scale: 1 }).width;
    const availableWidth = Math.max(500, $('#pdf-stage').clientWidth - 84);
    const fitScale = (availableWidth - 24) / (baseWidth * 2);
    setPdfScale(Math.max(.5, Math.min(1.25, fitScale)));
  } else {
    activePdf.layout = 'single';
    setPdfScale(activePdf.singleScale || 1.25);
  }
  $('#pdf-layout-single').classList.toggle('active', activePdf.layout === 'single');
  $('#pdf-layout-spread').classList.toggle('active', activePdf.layout === 'spread');
  $('#pdf-layout-single').setAttribute('aria-pressed', String(activePdf.layout === 'single'));
  $('#pdf-layout-spread').setAttribute('aria-pressed', String(activePdf.layout === 'spread'));
  requestAnimationFrame(() => scrollToPdfPage(pageNumber, 'auto'));
}

function setPdfScale(scale) {
  if (!activePdf?.document) return;
  activePdf.scale = Math.max(.5, Math.min(3, Math.round(scale * 100) / 100));
  activePdf.renderTasks.forEach(task => task?.cancel?.());
  activePdf.renderTasks.clear();
  activePdf.textLayers.forEach(layer => layer.cancel());
  activePdf.textLayers.clear();
  activePdf.pageTextMaps.clear();
  activePdf.rendered.clear();
  $('#pdf-zoom-value').textContent = `${Math.round(activePdf.scale * 100)}%`;
  const pageNumber = activePdf.pageNumber;
  buildPdfPages();
  requestAnimationFrame(() => scrollToPdfPage(pageNumber, 'auto'));
}

async function renderPdfViewer(item) {
  disposePdfViewer();
  const documentId = item.library ? `library:${item.id}` : item.edessbInstruction ? `edessb-instruction:${item.filename}` : `dbn:${item.path}`;
  activePdf = { item, documentId, document: null, defaultPageSize: null, pdfjs: null, pageNumber: 1, scale: 1.25, singleScale: 1.25, layout: 'single', rendered: new Set(), renderTasks: new Map(), textLayers: new Map(), pageTexts: new Map(), pageTextMaps: new Map(), pageSearchIndexes: new Map(), bookmarks: [], selection: null, nativeSelection: [], selectionFrame: 0, activeBookmarkId: '', searchMatches: [], searchIndex: -1, searchQuery: '', searchToken: 0, observer: null, rightAlt: false, zoomWheelLocked: false };
  const backLabel = item.library ? 'До бібліотеки' : item.edessbInstruction ? 'До інструкцій' : 'До каталогу';
  page.innerHTML = head(item.number) + `<div class="pdf-viewer"><div class="pdf-toolbar"><button id="pdf-back" class="button">${icon('chevron-left')}<span>${backLabel}</span></button><div class="pdf-search"><div class="pdf-search-box">${icon('search')}<input id="pdf-search-input" class="pdf-search-input" placeholder="Пошук у документі"></div><button id="pdf-search-submit" class="button primary">Пошук</button><span id="pdf-search-count" class="pdf-search-count"></span></div><div class="pdf-toolbar-group"><button id="pdf-prev" class="pdf-icon-button" title="Попередня сторінка">${icon('chevron-left')}</button><input id="pdf-page-number" class="pdf-page-input" type="number" min="1" value="1"><span id="pdf-page-count">/ —</span><button id="pdf-next" class="pdf-icon-button" title="Наступна сторінка">${icon('chevron-right')}</button></div><div class="pdf-layout-switch" aria-label="Режим відображення сторінок"><button id="pdf-layout-single" class="pdf-icon-button active" title="Одна сторінка" aria-pressed="true">${icon('page-single')}</button><button id="pdf-layout-spread" class="pdf-icon-button" title="Дві сторінки поруч" aria-pressed="false">${icon('pages-two')}</button></div><div class="pdf-toolbar-group"><button id="pdf-zoom-out" class="pdf-icon-button" title="Зменшити">${icon('minus')}</button><span id="pdf-zoom-value">125%</span><button id="pdf-zoom-in" class="pdf-icon-button" title="Збільшити">${icon('plus')}</button></div><button id="pdf-open-system" class="button">${icon('external-link')}<span>Відкрити окремо</span></button></div><div class="pdf-workspace"><div id="pdf-stage" class="pdf-stage"><div id="pdf-status" class="pdf-status">Завантаження документа…</div></div><aside class="pdf-bookmarks-panel"><div id="pdf-bookmarks-view" class="pdf-side-view"><div class="pdf-bookmarks-head">${icon('bookmark')}<span>Текстові закладки</span></div><div id="pdf-selection-editor" class="pdf-selection-editor" hidden><div id="pdf-selection-caption" class="pdf-selection-caption">Виділений текст</div><div id="pdf-selection-preview" class="pdf-selection-preview"></div><input id="pdf-bookmark-label" class="pdf-bookmark-label" maxlength="120" placeholder="Ключове слово"><div class="pdf-color-row"><button class="pdf-color active" data-color="yellow" title="Жовтий"></button><button class="pdf-color" data-color="green" title="Зелений"></button><button class="pdf-color" data-color="blue" title="Блакитний"></button><button class="pdf-color" data-color="pink" title="Рожевий"></button><button class="pdf-color" data-color="orange" title="Помаранчевий"></button></div><div class="pdf-selection-actions"><button id="pdf-bookmark-cancel" class="button">Скасувати</button><button id="pdf-bookmark-save" class="button primary">Зберегти</button></div></div><div id="pdf-bookmarks-list" class="pdf-bookmarks-list"></div></div><div id="pdf-search-view" class="pdf-side-view" hidden><div class="pdf-bookmarks-head">${icon('search')}<span>Результати пошуку</span></div><div id="pdf-search-summary" class="pdf-search-summary"></div><div id="pdf-search-results" class="pdf-search-results"></div><div class="pdf-search-footer"><button id="pdf-search-close" class="button">Закрити пошук</button></div></div></aside></div></div>`;
  $('#pdf-back').onclick = () => item.library ? moveHistory(-1) : navigate(item.edessbInstruction ? 'edessb-instructions' : 'dbn');
  $('#pdf-open-system').onclick = () => item.library ? window.ekp.openLibraryItem(item.id) : item.edessbInstruction ? window.ekp.openEdessbInstruction(item.filename) : window.ekp.openDbn(item.path);
  if(!item.library&&!item.edessbInstruction){const button=document.createElement('button');button.id='pdf-edessb';button.className='button pdf-edessb-link';button.disabled=!item.edessb_url;button.title=item.edessb_url?'Відкрити сторінку цього ДБН у ЄДЕССБ':'Для цього документа посилання ЄДЕССБ відсутнє';button.innerHTML=`${icon('external-link')}<span>ЄДЕССБ</span>`;$('#pdf-open-system').insertAdjacentElement('afterend',button);button.onclick=()=>{if(item.edessb_url)window.ekp.openExternal(item.edessb_url);};}
  $('#pdf-prev').onclick = () => scrollToPdfPage(activePdf.pageNumber - 1);
  $('#pdf-next').onclick = () => scrollToPdfPage(activePdf.pageNumber + 1);
  $('#pdf-layout-single').onclick = () => setPdfLayout('single');
  $('#pdf-layout-spread').onclick = () => setPdfLayout('spread');
  $('#pdf-zoom-out').onclick = () => setPdfScale(activePdf.scale - .25);
  $('#pdf-zoom-in').onclick = () => setPdfScale(activePdf.scale + .25);
  $('#pdf-page-number').onchange = event => scrollToPdfPage(Number(event.target.value) || 1);
  $('#pdf-search-submit').onclick = performPdfSearch;
  $('#pdf-search-input').onkeydown = event => { if (event.key === 'Enter') event.preventDefault(); };
  $('#pdf-search-close').onclick = closePdfSearch;
  $('#pdf-stage').addEventListener('pointerup', () => setTimeout(capturePdfSelection));
  document.onselectionchange = queuePdfNativeSelectionRender;
  document.querySelectorAll('.pdf-color').forEach(button => button.onclick = () => { if (!activePdf?.selection) return; activePdf.selection.color = button.dataset.color; document.querySelectorAll('.pdf-color').forEach(item => item.classList.toggle('active', item === button)); });
  $('#pdf-bookmark-cancel').onclick = () => { clearPdfSelectionEditor(); window.getSelection()?.removeAllRanges(); };
  $('#pdf-bookmark-save').onclick = savePdfSelectionBookmark;
  window.onkeydown = event => { if (event.code === 'AltRight' && activePdf) activePdf.rightAlt = true; };
  window.onkeyup = event => { if (event.code === 'AltRight' && activePdf) activePdf.rightAlt = false; };
  window.onblur = () => { if (activePdf) activePdf.rightAlt = false; };
  $('#pdf-stage').onwheel = event => {
    if (!activePdf?.rightAlt) return;
    event.preventDefault();
    if (activePdf.zoomWheelLocked) return;
    activePdf.zoomWheelLocked = true;
    setPdfScale(activePdf.scale + (event.deltaY < 0 ? .1 : -.1));
    setTimeout(() => { if (activePdf) activePdf.zoomWheelLocked = false; }, 90);
  };
  try {
    const [pdfjs, data] = await Promise.all([loadPdfJs(), item.library ? window.ekp.readLibraryPdf(item.id) : item.edessbInstruction ? window.ekp.readEdessbInstructionPdf(item.filename) : window.ekp.readPdf(item.path)]);
    activePdf.pdfjs = pdfjs;
    activePdf.document = await pdfjs.getDocument({ data: new Uint8Array(data) }).promise;
    const defaultPage=await activePdf.document.getPage(1),defaultViewport=defaultPage.getViewport({scale:1});activePdf.defaultPageSize={width:defaultViewport.width,height:defaultViewport.height};
    activePdf.bookmarks = await window.ekp.pdfBookmarks(activePdf.documentId);
    $('#pdf-page-count').textContent = `/ ${activePdf.document.numPages}`;
    renderPdfBookmarksPanel();
    buildPdfPages();
    await renderPdfPage(1);
  } catch (error) {
    const status = $('#pdf-status');
    if (status) status.textContent = `Не вдалося відкрити PDF: ${error.message || error}`;
    else console.error(error);
  }
}

async function openDbn(item) { await recent('recent_dbn', item.path); currentPage = 'dbn'; rememberView({ type: 'pdf', id: item.path }); await renderPdfViewer(item); renderSidebar(); }
async function openEdessbInstruction(item) {
  if (!item) return;
  currentPage = 'edessb-instructions';
  rememberView({ type: 'edessb-instruction-pdf', id: item.filename });
  await renderPdfViewer({ ...item, number: item.title, edessbInstruction: true });
  renderSidebar();
}
async function renderEdessbInstructions() {
  page.innerHTML = head('Інструкції ЄДЕССБ') + `<div class="page-content dbn-page edessb-instructions-page"><div class="edessb-loading">Завантаження каталогу інструкцій…</div></div>`;
  try { edessbInstructions = await window.ekp.edessbInstructions(); }
  catch (error) {
    page.innerHTML = head('Інструкції ЄДЕССБ') + `<div class="page-content dbn-page edessb-instructions-page"><div class="edessb-empty">${icon('info')}<h2>Не вдалося прочитати каталог</h2><p>${escape(readableError(error))}</p><button id="back-to-edessb" class="button">${icon('chevron-left')}<span>Повернутися до ЄДЕССБ</span></button></div></div>`;
    $('#back-to-edessb').onclick = () => navigate('edessb');
    return;
  }
  if (currentPage !== 'edessb-instructions') return;
  const categoryOf = item => String(item.category || '').trim() || 'Без категорії';
  const categoryCounts = new Map();
  edessbInstructions.forEach(item => categoryCounts.set(categoryOf(item), (categoryCounts.get(categoryOf(item)) || 0) + 1));
  const categories = [...categoryCounts.keys()].sort((a, b) => a.localeCompare(b, 'uk'));
  if (selectedEdessbInstructionCategory && !categoryCounts.has(selectedEdessbInstructionCategory)) selectedEdessbInstructionCategory = null;
  const favoriteInstructions = new Set(state.favorite_edessb_instructions || []);
  const orderedInstructions = [...edessbInstructions].sort((a, b) => {
    const favoriteDifference = Number(favoriteInstructions.has(b.filename)) - Number(favoriteInstructions.has(a.filename));
    return favoriteDifference || a.order - b.order || a.title.localeCompare(b.title, 'uk');
  });
  const categoryButtons = [`<button class="${selectedEdessbInstructionCategory ? '' : 'active'}" data-edessb-instruction-category-index="-1"><span>Усі категорії</span><strong>${edessbInstructions.length}</strong></button>`, ...categories.map((category, index) => `<button class="${selectedEdessbInstructionCategory === category ? 'active' : ''}" data-edessb-instruction-category-index="${index}" title="${escape(category)}"><span>${escape(category)}</span><strong>${categoryCounts.get(category)}</strong></button>`)].join('');
  const rows = orderedInstructions.map(item => {
    const favorite = favoriteInstructions.has(item.filename);
    const favoriteLabel = favorite ? 'Прибрати з обраних' : 'Додати до обраних';
    return `<tr data-instruction-file="${escape(item.filename)}" data-category="${escape(categoryOf(item))}"><td><div class="edessb-instruction-title"><button type="button" class="edessb-instruction-favorite${favorite ? ' active' : ''}" data-favorite-instruction="${escape(item.filename)}" title="${favoriteLabel}" aria-label="${favoriteLabel}" aria-pressed="${favorite}">${icon('star')}</button><span>${escape(item.title)}</span></div></td><td>${escape(item.version || '—')}</td><td>${item.pages || '—'}</td></tr>`;
  }).join('');
  page.innerHTML = head('Інструкції ЄДЕССБ') + `<div class="page-content dbn-page edessb-instructions-page"><button id="back-to-edessb" class="archicad-back">${icon('chevron-left')}<span>До ЄДЕССБ</span></button><h1 class="hero">Інструкції ЄДЕССБ</h1><p class="subtitle">Вбудований каталог • ${edessbInstructions.length} документів • оновлюється разом із CoDA</p><div class="dbn-layout"><section class="dbn-catalog"><table class="table edessb-instructions-table"><thead><tr><th>Найменування</th><th>Версія</th><th>Сторінок</th></tr></thead><tbody>${rows}</tbody></table></section><aside class="dbn-categories"><div class="dbn-categories-title">Категорії</div><div class="dbn-category-list">${categoryButtons}</div></aside></div><div class="dbn-floating-controls" role="search"><input id="edessb-instructions-search" class="input" value="${escape(edessbInstructionSearch)}" placeholder="Пошук за назвою інструкції"><button id="open-edessb-instruction" class="button primary edessb-instructions-open" disabled>${icon('book-open')}<span>Відкрити інструкцію</span></button></div></div>`;
  $('#back-to-edessb').onclick = () => navigate('edessb');
  let selectedRow = null;
  const selectRow = row => {
    document.querySelectorAll('.edessb-instructions-table tbody tr').forEach(item => item.classList.toggle('selected', item === row));
    selectedRow = row;
    $('#open-edessb-instruction').disabled = !row;
  };
  document.querySelectorAll('.edessb-instructions-table tbody tr').forEach(row => {
    row.onclick = () => selectRow(row);
    row.ondblclick = () => openEdessbInstruction(edessbInstructions.find(item => item.filename === row.dataset.instructionFile));
  });
  document.querySelectorAll('[data-favorite-instruction]').forEach(button => {
    button.onclick = async event => {
      event.stopPropagation();
      const filename = button.dataset.favoriteInstruction;
      const favorites = new Set(state.favorite_edessb_instructions || []);
      favorites.has(filename) ? favorites.delete(filename) : favorites.add(filename);
      state.favorite_edessb_instructions = [...favorites];
      await saveState();
      await renderEdessbInstructions();
    };
    button.ondblclick = event => event.stopPropagation();
  });
  const applyFilters = () => {
    const query = edessbInstructionSearch.trim().toLocaleLowerCase('uk');
    document.querySelectorAll('.edessb-instructions-table tbody tr').forEach(row => {
      const searchMatches = !query || row.innerText.toLocaleLowerCase('uk').includes(query);
      const categoryMatches = !selectedEdessbInstructionCategory || row.dataset.category === selectedEdessbInstructionCategory;
      row.hidden = !(searchMatches && categoryMatches);
      if (row.hidden && row === selectedRow) selectRow(null);
    });
  };
  $('#edessb-instructions-search').oninput = event => { edessbInstructionSearch = event.target.value; applyFilters(); };
  document.querySelectorAll('[data-edessb-instruction-category-index]').forEach(button => button.onclick = () => {
    const index = Number(button.dataset.edessbInstructionCategoryIndex);
    selectedEdessbInstructionCategory = index < 0 ? null : categories[index];
    document.querySelectorAll('[data-edessb-instruction-category-index]').forEach(item => item.classList.toggle('active', item === button));
    applyFilters();
  });
  $('#open-edessb-instruction').onclick = () => selectedRow && openEdessbInstruction(edessbInstructions.find(item => item.filename === selectedRow.dataset.instructionFile));
  applyFilters();
}
function libraryFileSize(value) {
  const bytes=Number(value)||0;
  if(bytes<1024)return `${bytes} Б`;
  if(bytes<1024*1024)return `${(bytes/1024).toFixed(bytes<10240?1:0)} КБ`;
  return `${(bytes/1024/1024).toFixed(bytes<10*1024*1024?1:0)} МБ`;
}
function requestLibraryItem() {
  return new Promise(resolve=>{
    const dialog=$('#library-add-dialog'),category=$('#library-item-category'),name=$('#library-item-name'),fileName=$('#library-item-file-name'),error=$('#library-item-error');
    let done=false,sourcePath='';
    const finish=value=>{if(done)return;done=true;if(dialog.open)dialog.close();resolve(value);};
    category.innerHTML=libraryConfig.categories.map(item=>`<option value="${escape(item.id)}">${escape(item.label)}</option>`).join('');
    category.value=selectedLibraryCategory||libraryConfig.categories[0]?.id||'dstu';name.value='';fileName.textContent='Файл не обрано';error.hidden=true;
    $('#library-item-select-file').onclick=async()=>{try{const selected=await window.ekp.selectLibraryFile();if(selected.canceled)return;sourcePath=selected.sourcePath;fileName.textContent=selected.name;if(!name.value.trim())name.value=selected.name.replace(/\.[^.]+$/,'');error.hidden=true;}catch(fileError){error.textContent=readableError(fileError);error.hidden=false;}};
    $('#library-item-cancel').onclick=()=>finish(null);
    $('#library-item-accept').onclick=()=>{const payload={category:category.value,name:name.value.trim(),sourcePath};if(!payload.name){error.textContent='Вкажіть найменування документа.';error.hidden=false;name.focus();return;}if(!payload.sourcePath){error.textContent='Оберіть файл документа.';error.hidden=false;return;}finish(payload);};
    dialog.oncancel=event=>{event.preventDefault();finish(null);};dialog.onclose=()=>finish(null);dialog.showModal();requestAnimationFrame(()=>name.focus());
  });
}
function requestLibraryCategory() {
  return new Promise(resolve=>{
    const dialog=$('#library-category-dialog'),name=$('#library-category-name'),error=$('#library-category-error');
    let done=false;
    const finish=value=>{if(done)return;done=true;if(dialog.open)dialog.close();resolve(value);};
    name.value='';error.hidden=true;
    $('#library-category-cancel').onclick=()=>finish(null);
    $('#library-category-accept').onclick=()=>{const value=name.value.trim();if(!value){error.textContent='Вкажіть назву категорії.';error.hidden=false;name.focus();return;}finish(value);};
    name.oninput=()=>{error.hidden=true;};
    name.onkeydown=event=>{if(event.key==='Enter'){event.preventDefault();$('#library-category-accept').click();}};
    dialog.oncancel=event=>{event.preventDefault();finish(null);};dialog.onclose=()=>finish(null);dialog.showModal();requestAnimationFrame(()=>name.focus());
  });
}
async function openLibraryItem(item) {
  if(!item?.exists){showArchicadMessage('Файл не знайдено','Файл документа відсутній у локальному сховищі Бібліотеки.');return;}
  if(item.isPdf){currentPage='library';rememberView({type:'library-pdf',id:item.id});await renderPdfViewer({number:item.name,id:item.id,library:true});renderSidebar();return;}
  try{await window.ekp.openLibraryItem(item.id);}catch(error){showArchicadMessage('Не вдалося відкрити файл',readableError(error));}
}
async function renderLibrary() {
  [libraryItems,libraryConfig]=await Promise.all([window.ekp.libraryItems(),window.ekp.libraryConfig()]);
  const categories=libraryConfig.categories;
  if(!categories.some(item=>item.id===selectedLibraryCategory))selectedLibraryCategory=categories[0]?.id||'dstu';
  const counts=new Map(categories.map(category=>[category.id,libraryItems.filter(item=>item.category===category.id).length]));
  const categoryButtons=categories.map(category=>`<button class="${category.id===selectedLibraryCategory?'active':''}" data-library-category="${escape(category.id)}"><span>${escape(category.label)}</span><strong>${counts.get(category.id)||0}</strong></button>`).join('');
  const visible=libraryItems.filter(item=>item.category===selectedLibraryCategory);
  const cards=visible.map(item=>`<article class="library-document${item.exists?'':' missing'}"><div class="library-document-icon">${icon('file-text')}</div><button class="library-document-copy" data-library-open="${escape(item.id)}" ${item.exists?'':'disabled'}><strong>${escape(item.name)}</strong><span>${escape(item.originalName)} · ${item.extension?escape(item.extension.toUpperCase()):'Файл'} · ${libraryFileSize(item.size)}</span></button><div class="library-document-actions"><button class="button library-open" data-library-open="${escape(item.id)}" ${item.exists?'':'disabled'}>${icon(item.isPdf?'book-open':'external-link')}<span>${item.isPdf?'Переглянути PDF':attachmentOpenLabel(item.originalName)}</span></button><button class="button library-delete" data-library-delete="${escape(item.id)}" title="Видалити документ" aria-label="Видалити документ">${icon('trash')}</button></div></article>`).join('');
  const activeCategory=categories.find(item=>item.id===selectedLibraryCategory);
  page.innerHTML=head('Бібліотека')+`<div class="page-content library-page"><div class="library-hero"><div><h1 class="hero">Бібліотека</h1><p class="subtitle">Власне локальне сховище стандартів і нормативних документів</p></div><div class="library-toolbar"><button id="open-library-folder" class="button">${icon('folder')}<span>Відкрити папку</span></button><button id="add-library-item" class="button primary">${icon('plus')}<span>Додати документ</span></button></div></div><div class="library-layout"><aside class="library-categories"><div class="library-categories-title"><span>Категорії</span><button id="add-library-category" title="Створити категорію" aria-label="Створити категорію">${icon('plus')}</button></div>${categoryButtons}</aside><section class="library-documents"><div class="library-section-head"><div><h2>${escape(activeCategory?.label||'Документи')}</h2><p>${visible.length} ${visible.length===1?'документ':'документів'}</p></div></div><div class="library-document-list">${cards||`<div class="library-empty">${icon('library')}<h3>У цій категорії ще немає документів</h3><p>Додайте PDF для перегляду всередині CoDA або файл іншого формату для відкриття у системній програмі.</p><button id="add-library-item-empty" class="button primary">${icon('plus')}<span>Додати документ</span></button></div>`}</div></section></div></div>`;
  document.querySelectorAll('[data-library-category]').forEach(button=>button.onclick=()=>{selectedLibraryCategory=button.dataset.libraryCategory;renderLibrary();});
  const add=async()=>{const payload=await requestLibraryItem();if(!payload)return;try{const saved=await window.ekp.addLibraryItem(payload);selectedLibraryCategory=saved.category;await renderLibrary();}catch(error){showArchicadMessage('Не вдалося додати документ',readableError(error));}};
  $('#add-library-item').onclick=add;if($('#add-library-item-empty'))$('#add-library-item-empty').onclick=add;
  $('#add-library-category').onclick=async()=>{const label=await requestLibraryCategory();if(!label)return;try{const category=await window.ekp.addLibraryCategory(label);selectedLibraryCategory=category.id;await renderLibrary();}catch(error){showArchicadMessage('Не вдалося створити категорію',readableError(error));}};
  $('#open-library-folder').onclick=async()=>{try{await window.ekp.openLibraryFolder();}catch(error){showArchicadMessage('Не вдалося відкрити папку',readableError(error));}};
  document.querySelectorAll('[data-library-open]').forEach(button=>button.onclick=()=>openLibraryItem(libraryItems.find(item=>item.id===button.dataset.libraryOpen)));
  document.querySelectorAll('[data-library-delete]').forEach(button=>button.onclick=async()=>{const item=libraryItems.find(entry=>entry.id===button.dataset.libraryDelete);if(!item||!await confirmDelete(`Документ «${item.name}» і його локальна копія будуть видалені з Бібліотеки. Початковий файл залишиться без змін.`,'Видалити документ?'))return;try{await window.ekp.deleteLibraryItem(item.id);await renderLibrary();}catch(error){showArchicadMessage('Не вдалося видалити документ',readableError(error));}});
}
function showArchicadMessage(title, message) {
  const dialog=$('#archicad-message-dialog');
  $('#archicad-message-title').textContent=title;
  $('#archicad-message-text').textContent=message;
  dialog.onclose=()=>window.ekp.window.releaseAlwaysOnTop();
  dialog.showModal();
}
function confirmArchicadImport(project) {
  return new Promise(resolve=>{
    const dialog=$('#archicad-import-dialog');
    $('#archicad-import-message').textContent=`Збережені в Excel зміни ID та класифікацій буде записано у відкритий проєкт Archicad. Проєкт CoDA: «${project.title}».`;
    let done=false;
    const finish=value=>{if(done)return;done=true;dialog.close();resolve(value);};
    $('#archicad-import-cancel').onclick=()=>finish(false);
    $('#archicad-import-accept').onclick=()=>finish(true);
    dialog.oncancel=event=>{event.preventDefault();finish(false);};
    dialog.showModal();
  });
}
function confirmArchicadWindowSelection() {
  return new Promise(resolve=>{
    const dialog=$('#archicad-window-dialog');
    let done=false;
    const finish=value=>{if(done)return;done=true;if(dialog.open)dialog.close();resolve(value);};
    $('#archicad-window-cancel').onclick=()=>finish(false);
    $('#archicad-window-start').onclick=()=>finish(true);
    dialog.oncancel=event=>{event.preventDefault();finish(false);};
    dialog.onclose=()=>finish(false);
    dialog.showModal();
  });
}
function requestArchicadProjectName() {
  return new Promise(resolve=>{
    const dialog=$('#archicad-project-create-dialog');
    const input=$('#archicad-project-name');
    const error=$('#archicad-project-create-error');
    let done=false;
    const finish=value=>{if(done)return;done=true;if(dialog.open)dialog.close();resolve(value);};
    input.value='';error.hidden=true;
    $('#archicad-project-create-cancel').onclick=()=>finish(null);
    $('#archicad-project-create-accept').onclick=()=>{const value=input.value.trim();if(!value){error.textContent='Вкажіть назву проєкту.';error.hidden=false;input.focus();return;}finish(value);};
    input.onkeydown=event=>{if(event.key==='Enter'){$('#archicad-project-create-accept').click();}};
    dialog.oncancel=event=>{event.preventDefault();finish(null);};
    dialog.onclose=()=>finish(null);
    dialog.showModal();
    requestAnimationFrame(()=>input.focus());
  });
}
function archicadConnectionHtml(status) {
  return status.connected
    ? `<div class="archicad-connection connected">${icon('sheet')}<span>Підключено: Archicad ${escape(status.version)}, build ${escape(status.buildNumber)}${status.port?` · порт ${escape(status.port)}`:''}</span></div>`
    : `<div class="archicad-connection disconnected">${icon('info')}<span>${escape(status.message||'Archicad не підключено')}</span></div>`;
}
function archicadConnectionsHtml(connections) {
  if(!connections.length)return archicadConnectionHtml({connected:false,message:'Відкритих проєктів Archicad не знайдено'});
  if(connections.length===1)return archicadConnectionHtml({connected:true,...connections[0]});
  return `<div class="archicad-connection connected">${icon('sheet')}<span>Знайдено відкритих проєктів Archicad: ${connections.length}. Підключення обирається всередині кожного проєкту CoDA.</span></div>`;
}
function archicadDevelopmentNotice(){return `<div class="archicad-development-notice">${icon('info')}<div><strong>Функціонал перебуває в розробці</strong><span>Розділ Archicad / Excel ще не доопрацьований. Перед імпортом змін рекомендуємо зберігати резервну копію проєкту Archicad.</span></div></div>`;}
async function renderExcel(){
  page.innerHTML=head('Archicad / Excel')+`<div class="page-content archicad-page"><div class="archicad-empty">Завантаження проєктів…</div></div>`;
  const [connections,projects]=await Promise.all([window.ekp.archicadConnections(),window.ekp.archicadProjects()]);
  if(currentPage!=='excel')return;
  let selected=projects.find(project=>project.id===selectedArchicadProjectId);
  if(selectedArchicadProjectId&&!selected)selectedArchicadProjectId=null;

  if(!selected){
    page.innerHTML=head('Archicad / Excel')+`<div class="page-content archicad-page">${archicadDevelopmentNotice()}<div class="archicad-page-heading"><div><h1 class="hero">Проєкти Archicad</h1><p class="subtitle">Кожен проєкт — окрема папка зі своїми Excel-файлами та інструментами обміну</p></div><button id="create-archicad-project" class="button primary">${icon('plus')}<span>Створити проєкт</span></button></div>${archicadConnectionsHtml(connections)}<div class="archicad-project-browser">${projects.length?projects.map(project=>`<div class="archicad-folder-row"><button class="archicad-folder-card" data-ac-project="${project.id}"><span class="archicad-folder-icon">${icon('folder')}</span><span class="archicad-folder-copy"><strong>${escape(project.title)}</strong><small>${project.files.length} ${project.files.length===1?'Excel-файл':'Excel-файлів'} · створено ${escape(new Date(project.createdAt).toLocaleString('uk-UA'))}</small></span>${icon('chevron-right','folder-arrow')}</button><button class="archicad-folder-delete" data-ac-delete-project="${project.id}" title="Видалити проєкт" aria-label="Видалити проєкт ${escape(project.title)}">${icon('trash')}</button></div>`).join(''):`<div class="archicad-empty"><div>${icon('folder')}</div><h2>Проєктів ще немає</h2><p>Створіть папку проєкту, а потім сформуйте всередині потрібні Excel-файли.</p></div>`}</div></div>`;
    $('#create-archicad-project').onclick=async()=>{const name=await requestArchicadProjectName();if(!name)return;try{const project=await window.ekp.createArchicadProject(name);selectedArchicadProjectId=project.id;await renderExcel();}catch(error){showArchicadMessage('Не вдалося створити проєкт',error?.message||String(error));}};
    document.querySelectorAll('[data-ac-project]').forEach(button=>button.onclick=()=>{selectedArchicadProjectId=button.dataset.acProject;renderExcel();});
    document.querySelectorAll('[data-ac-delete-project]').forEach(button=>button.onclick=async()=>{const project=projects.find(item=>item.id===button.dataset.acDeleteProject);if(!project||!await confirmDelete(`Проєкт «${project.title}» та всі Excel-файли всередині нього буде видалено без можливості відновлення. Копії, які ви раніше зберегли на комп’ютері, залишаться без змін.`,'Видалити проєкт?'))return;try{await window.ekp.deleteArchicadProject(project.id);if(selectedArchicadProjectId===project.id)selectedArchicadProjectId=null;await renderExcel();}catch(error){showArchicadMessage('Не вдалося видалити проєкт',error?.message||String(error));}});
    return;
  }

  if(!selected.connection&&connections.length===1){
    try{selected=await window.ekp.setArchicadProjectConnection({projectId:selected.id,port:connections[0].port});}
    catch(error){showArchicadMessage('Не вдалося підключити Archicad',error?.message||String(error));}
  }
  const elementsFile=selected.files.find(file=>file.id==='all-elements');
  const selectedPort=Number(selected.connection?.port)||0;
  const activeConnection=connections.find(connection=>connection.port===selectedPort);
  const connectionMessage=selectedPort&&!activeConnection?`Раніше вибране підключення (порт ${selectedPort}) зараз недоступне. Оберіть інший відкритий проєкт Archicad.`:connections.length?'Оберіть, з якого відкритого проєкту Archicad читати дані та куди імпортувати зміни.':'Відкрийте потрібний проєкт в Archicad, а потім оновіть список.';
  const connectionOptions=connections.length?`<option value="">Оберіть проєкт Archicad…</option>${connections.map(connection=>`<option value="${connection.port}" ${connection.port===selectedPort?'selected':''}>${escape(connection.label)}</option>`).join('')}`:'<option value="">Відкритих проєктів Archicad не знайдено</option>';
  page.innerHTML=head(selected.title)+`<div class="page-content archicad-page">${archicadDevelopmentNotice()}<button id="back-to-archicad-projects" class="archicad-back">${icon('chevron-left')}<span>Усі проєкти</span></button><div class="archicad-page-heading project-heading"><div><h1 class="hero">${escape(selected.title)}</h1><p class="subtitle">Excel-файли та інструменти цього проєкту</p></div><button id="delete-current-archicad-project" class="button archicad-danger">${icon('trash')}<span>Видалити проєкт</span></button></div><section class="archicad-project-connection ${activeConnection?'connected':selectedPort?'warning':''}"><div class="archicad-connection-copy"><strong>${icon('sheet')}Підключення до Archicad</strong><span>${escape(connectionMessage)}</span></div><div class="archicad-connection-controls"><select id="archicad-connection-select" ${connections.length?'':'disabled'}>${connectionOptions}</select><button id="detect-active-archicad" class="button primary" title="Визначити підключення за активним вікном Archicad">${icon('external-link')}<span>Обрати через вікно</span></button><button id="refresh-archicad-connections" class="button" title="Оновити список підключень">${icon('refresh-cw')}</button></div></section><div class="archicad-tools"><article class="archicad-tool-card"><header><div class="archicad-tool-icon">${icon('sheet')}</div><div><h2>Усі елементи</h2><p>Унікальний ID, ID елемента, слой і класифікація всіх елементів вибраного файлу Archicad.</p></div><button id="export-all-elements" class="button primary">${icon(elementsFile?'refresh-cw':'plus')}<span>${elementsFile?'Оновити Excel':'Створити Excel'}</span></button></header>${elementsFile?`<div class="archicad-file"><div class="archicad-file-icon">${icon('sheet')}</div><div class="archicad-file-copy"><strong>${escape(elementsFile.name)}</strong><span>${elementsFile.rowCount||0} елементів · оновлено ${escape(new Date(elementsFile.updatedAt).toLocaleString('uk-UA'))}</span></div><div class="archicad-file-actions"><button class="button" data-ac-open="${elementsFile.id}">${icon('external-link')}<span>Відкрити Excel</span></button><button class="button" data-ac-save="${elementsFile.id}">${icon('save')}<span>Зберегти</span></button><button class="button primary" data-ac-import="${elementsFile.id}">${icon('upload')}<span>Імпортувати в Archicad</span></button><button class="button archicad-file-delete" data-ac-delete-file="${elementsFile.id}" title="Видалити Excel">${icon('trash')}</button></div></div>`:'<div class="archicad-tool-empty">Excel ще не створено.</div>'}</article><article class="archicad-future-card"><div>${icon('plus')}</div><div><strong>Наступні Excel-інструменти</strong><span>Тут згодом можна додати відомості дверей, вікон, приміщень та інші специфікації.</span></div></article></div></div>`;
  $('#back-to-archicad-projects').onclick=()=>{selectedArchicadProjectId=null;renderExcel();};
  $('#delete-current-archicad-project').onclick=async()=>{if(!await confirmDelete(`Проєкт «${selected.title}» та всі Excel-файли всередині нього буде видалено без можливості відновлення. Копії, які ви раніше зберегли на комп’ютері, залишаться без змін.`,'Видалити проєкт?'))return;try{await window.ekp.deleteArchicadProject(selected.id);selectedArchicadProjectId=null;await renderExcel();}catch(error){showArchicadMessage('Не вдалося видалити проєкт',error?.message||String(error));}};
  $('#archicad-connection-select').onchange=async event=>{const port=Number(event.target.value);if(!port)return;event.target.disabled=true;try{await window.ekp.setArchicadProjectConnection({projectId:selected.id,port});await renderExcel();}catch(error){showArchicadMessage('Не вдалося перепідключити проєкт',error?.message||String(error));event.target.disabled=false;}};
  $('#detect-active-archicad').onclick=async()=>{if(!await confirmArchicadWindowSelection())return;const button=$('#detect-active-archicad');button.disabled=true;button.innerHTML=`${icon('refresh-cw','spinning')}<span>Перейдіть в Archicad…</span>`;try{const connection=await window.ekp.activeArchicadConnection();await window.ekp.setArchicadProjectConnection({projectId:selected.id,port:connection.port});await renderExcel();showArchicadMessage('Archicad підключено',`Проєкт CoDA «${selected.title}» підключено до вікна Archicad на порті ${connection.port}.`);}catch(error){showArchicadMessage('Не вдалося визначити Archicad',error?.message||String(error));button.disabled=false;button.innerHTML=`${icon('external-link')}<span>Обрати через вікно</span>`;}};
  $('#refresh-archicad-connections').onclick=()=>renderExcel();
  $('#export-all-elements').onclick=async()=>{const button=$('#export-all-elements');if(elementsFile&&!window.confirm('Оновлення замінить внутрішній Excel актуальними даними з Archicad. Незбережені зміни в таблиці буде втрачено. Продовжити?'))return;button.disabled=true;button.innerHTML=`${icon('refresh-cw','spinning')}<span>Читання елементів…</span>`;try{await window.ekp.exportArchicadElements(selected.id);await renderExcel();}catch(error){showArchicadMessage('Не вдалося створити Excel',error?.message||String(error));button.disabled=false;}};
  document.querySelectorAll('[data-ac-open]').forEach(button=>button.onclick=async()=>{try{await window.ekp.openArchicadProjectFile({projectId:selected.id,fileId:button.dataset.acOpen});}catch(error){showArchicadMessage('Не вдалося відкрити Excel',error?.message||String(error));}});
  document.querySelectorAll('[data-ac-save]').forEach(button=>button.onclick=async()=>{button.disabled=true;try{const result=await window.ekp.saveArchicadProjectFile({projectId:selected.id,fileId:button.dataset.acSave});if(!result.canceled)showArchicadMessage('Excel-файл збережено',`Копію створено: ${result.filePath}`);}catch(error){showArchicadMessage('Не вдалося зберегти файл',error?.message||String(error));}finally{button.disabled=false;}});
  document.querySelectorAll('[data-ac-delete-file]').forEach(button=>button.onclick=async()=>{if(!elementsFile||!await confirmDelete(`Excel «${elementsFile.name}» буде видалено з проєкту «${selected.title}» без можливості відновлення. Копії, які ви зберегли в інших папках, залишаться без змін.`,'Видалити Excel-файл?'))return;try{await window.ekp.deleteArchicadProjectFile({projectId:selected.id,fileId:button.dataset.acDeleteFile});await renderExcel();}catch(error){showArchicadMessage('Не вдалося видалити Excel',error?.message||String(error));}});
  document.querySelectorAll('[data-ac-import]').forEach(button=>button.onclick=async()=>{if(!await confirmArchicadImport(selected))return;button.disabled=true;button.innerHTML=`${icon('refresh-cw','spinning')}<span>Імпорт…</span>`;try{const result=await window.ekp.importArchicadProject({projectId:selected.id,fileId:button.dataset.acImport});const layerWarning=result.layerChangesSkipped?`\n\n${result.layerChangesSkipped} змін слоя не застосовано: офіційний JSON API Archicad надає назву слоя лише для читання.`:'';showArchicadMessage('Імпорт завершено',`Оновлено ID: ${result.elementIdsUpdated}. Оновлено класифікацій: ${result.classificationsUpdated}. Без змін: ${result.unchanged}.${layerWarning}`);}catch(error){showArchicadMessage('Не вдалося імпортувати зміни',error?.message||String(error));}finally{button.disabled=false;button.innerHTML=`${icon('upload')}<span>Імпортувати в Archicad</span>`;}});
}
function edessbPortalSection() {
  return `<section class="edessb-portal"><div class="edessb-section-heading"><div><h1 class="hero">ЄДЕССБ</h1><p class="subtitle">Швидкий перехід до основних розділів Єдиної державної електронної системи у сфері будівництва</p></div><button id="open-edessb-instructions" class="button primary">${icon('book-open')}<span>Інструкція</span></button></div><div class="edessb-portal-grid">${edessbConfig.portalLinks.map(link=>`<button class="edessb-portal-link" data-edessb-portal="${escape(link.url)}">${icon('external-link')}<span>${escape(link.label)}</span></button>`).join('')}</div><div class="edessb-privacy">${icon('shield')}<p>CoDA не збирає і не передає жодної інформації, яка була введена в браузері і не має жодного доступу до нього.</p></div></section>`;
}
function requestEdessbProject() {
  return new Promise(resolve=>{
    const dialog=$('#edessb-project-create-dialog'),identifier=$('#edessb-project-identifier'),title=$('#edessb-project-title'),organization=$('#edessb-project-organization'),error=$('#edessb-project-create-error');let done=false;
    const finish=value=>{if(done)return;done=true;if(dialog.open)dialog.close();resolve(value);};
    identifier.value='';title.value='';organization.value='';error.hidden=true;
    $('#edessb-project-create-cancel').onclick=()=>finish(null);
    $('#edessb-project-create-accept').onclick=()=>{const payload={identifier:identifier.value.trim(),title:title.value.trim(),organization:organization.value.trim()};const invalid=!/^\d+$/.test(payload.identifier)?identifier:!payload.title?title:!payload.organization?organization:null;if(invalid){error.textContent=invalid===identifier?'Вкажіть цифровий ідентифікаційний номер проєкту.':invalid===title?'Вкажіть назву проєкту.':'Вкажіть проєктну організацію.';error.hidden=false;invalid.focus();return;}finish(payload);};
    dialog.oncancel=event=>{event.preventDefault();finish(null);};dialog.onclose=()=>finish(null);dialog.showModal();requestAnimationFrame(()=>identifier.focus());
  });
}
function requestUserDataGroups(groups, mode) {
  return new Promise(resolve=>{
    const dialog=$('#user-data-groups-dialog'),list=$('#user-data-groups-list'),error=$('#user-data-groups-error');let completed=false;
    $('#user-data-groups-title').textContent=mode==='export'?'Що зберегти?':'Що замінити?';
    $('#user-data-groups-description').textContent=mode==='export'?'Оберіть одну або кілька незалежних категорій для файла .codasaves.':'Оберіть категорії з резервної копії. Лише вони будуть замінені.';
    list.innerHTML=groups.map(group=>`<label class="user-data-group"><input type="checkbox" value="${escape(group.id)}" checked><span><strong>${escape(group.label)}</strong><small>${escape(group.description)}</small></span></label>`).join('');error.hidden=true;
    const finish=value=>{if(completed)return;completed=true;if(dialog.open)dialog.close();resolve(value);};
    $('#user-data-groups-cancel').onclick=()=>finish(null);
    $('#user-data-groups-accept').onclick=()=>{const ids=[...list.querySelectorAll('input:checked')].map(input=>input.value);if(!ids.length){error.textContent='Оберіть хоча б одну категорію.';error.hidden=false;return;}finish(ids);};
    dialog.oncancel=event=>{event.preventDefault();finish(null);};dialog.onclose=()=>finish(null);dialog.showModal();
  });
}
function requestEdessbRevision(typeId, revisionNumber, existing = null) {
  return new Promise(resolve=>{
    const dialog=$('#edessb-document-dialog'),type=$('#edessb-document-type'),url=$('#edessb-document-url'),error=$('#edessb-document-error'),pdfName=$('#edessb-document-pdf-name'),clearPdf=$('#edessb-document-clear-pdf'),accept=$('#edessb-document-accept');let done=false,sourcePath='',removePdf=false;
    const finish=value=>{if(done)return;done=true;if(dialog.open)dialog.close();resolve(value);};
    type.innerHTML=edessbConfig.documentTypes.map(item=>`<option value="${escape(item.id)}">${escape(item.label)} (${escape(item.short)})</option>`).join('');type.value=typeId;url.value=existing?.url||'';error.hidden=true;pdfName.textContent=existing?.storedName||'Файл не обрано';clearPdf.hidden=!existing?.storedName;$('#edessb-document-dialog-title').textContent=existing?`Редагувати редакцію №${revisionNumber}`:`Редакція №${revisionNumber}`;$('#edessb-document-dialog-subtitle').textContent=existing?'Змініть посилання або додайте чи замініть PDF-файл':'Достатньо посилання, PDF-файла або обох';accept.textContent=existing?'Зберегти зміни':'Зберегти редакцію';
    $('#edessb-document-select-pdf').onclick=async()=>{try{const selected=await window.ekp.selectEdessbPdf();if(selected.canceled)return;sourcePath=selected.sourcePath;removePdf=false;pdfName.textContent=selected.name;clearPdf.hidden=false;error.hidden=true;}catch(reason){error.textContent=readableError(reason);error.hidden=false;}};
    clearPdf.onclick=()=>{sourcePath='';removePdf=Boolean(existing?.storedName);pdfName.textContent='Файл не обрано';clearPdf.hidden=true;};
    $('#edessb-document-cancel').onclick=()=>finish(null);
    accept.onclick=()=>{const keepExistingPdf=Boolean(existing?.storedName&&!removePdf&&!sourcePath),payload={type:type.value,number:revisionNumber,url:url.value.trim(),sourcePath,removePdf};if(!payload.url&&!payload.sourcePath&&!keepExistingPdf){error.textContent='Додайте посилання, PDF-файл або обидва.';error.hidden=false;return;}if(payload.url){let parsed;try{parsed=new URL(payload.url);}catch{error.textContent='Вкажіть правильне посилання на документ.';error.hidden=false;url.focus();return;}if(parsed.protocol!=='https:'||!(parsed.hostname==='e-construction.gov.ua'||parsed.hostname.endsWith('.e-construction.gov.ua'))){error.textContent='Посилання повинно вести на захищену сторінку e-construction.gov.ua.';error.hidden=false;url.focus();return;}}finish(payload);};
    dialog.oncancel=event=>{event.preventDefault();finish(null);};dialog.onclose=()=>finish(null);dialog.showModal();requestAnimationFrame(()=>url.focus());
  });
}
function bindEdessbPortalLinks() {
  document.querySelectorAll('[data-edessb-portal]').forEach(button=>button.onclick=()=>window.ekp.openExternal(button.dataset.edessbPortal));
  $('#open-edessb-instructions').onclick=()=>navigate('edessb-instructions');
}
async function renderEdessb() {
  page.innerHTML=head('ЄДЕССБ')+`<div class="page-content edessb-page"><div class="edessb-loading">Завантаження проєктів…</div></div>`;
  const [config,projects]=await Promise.all([window.ekp.edessbConfig(),window.ekp.edessbProjects()]);
  if(currentPage!=='edessb')return;
  edessbConfig=config;
  let selected=projects.find(project=>project.id===selectedEdessbProjectId);
  if(selectedEdessbProjectId&&!selected)selectedEdessbProjectId=null;
  if(!selected){
    page.innerHTML=head('ЄДЕССБ')+`<div class="page-content edessb-page">${edessbPortalSection()}<div class="edessb-storage-note">${icon('folder')}<div><strong>PDF зберігаються окремо від програми</strong><p>CoDA копіює й перейменовує файли у папці «edessb_documents». Їх можна відкрити та скопіювати без запуску CoDA.</p></div><button id="open-edessb-root" class="button">Відкрити папку PDF</button></div><section class="edessb-projects-section"><div class="edessb-section-heading"><div><h2>Мої проєкти</h2><p>Проєкти впорядковано за ідентифікаційним номером</p></div><button id="create-edessb-project" class="button primary">${icon('plus')}<span>Створити проєкт</span></button></div><div class="edessb-project-list">${projects.length?projects.map(project=>`<article class="edessb-project-row"><button class="edessb-project-main" data-edessb-project="${escape(project.id)}"><span class="edessb-project-number">№${escape(project.identifier||'—')}</span><span class="edessb-project-copy"><strong>${escape(project.title)}</strong><small>${escape(project.organization)}</small></span>${icon('chevron-right','folder-arrow')}</button><button class="edessb-project-delete" data-edessb-delete-project="${escape(project.id)}" title="Видалити проєкт">${icon('trash')}</button></article>`).join(''):`<div class="edessb-empty">${icon('folder')}<h2>Проєктів ще немає</h2><p>Створіть перший проєкт і додайте посилання або PDF-документи.</p></div>`}</div></section></div>`;
    bindEdessbPortalLinks();
    $('#open-edessb-root').onclick=()=>window.ekp.openEdessbDocumentsFolder();
    $('#create-edessb-project').onclick=async()=>{const payload=await requestEdessbProject();if(!payload)return;try{const project=await window.ekp.createEdessbProject(payload);selectedEdessbProjectId=project.id;await renderEdessb();}catch(error){showArchicadMessage('Не вдалося створити проєкт',readableError(error));}};
    document.querySelectorAll('[data-edessb-project]').forEach(button=>button.onclick=()=>{selectedEdessbProjectId=button.dataset.edessbProject;renderEdessb();});
    document.querySelectorAll('[data-edessb-delete-project]').forEach(button=>button.onclick=async()=>{const project=projects.find(item=>item.id===button.dataset.edessbDeleteProject);if(!project||!await confirmDelete(`Проєкт «${project.title}» та всі PDF-документи всередині нього буде видалено без можливості відновлення.`,'Видалити проєкт ЄДЕССБ?'))return;try{await window.ekp.deleteEdessbProject(project.id);await renderEdessb();}catch(error){showArchicadMessage('Не вдалося видалити проєкт',readableError(error));}});
    return;
  }
  const documentMap=new Map((selected.documents||[]).map(document=>[document.type,document]));
  page.innerHTML=head(selected.title)+`<div class="page-content edessb-page">${edessbPortalSection()}<section class="edessb-project-detail"><button id="back-to-edessb-projects" class="archicad-back">${icon('chevron-left')}<span>Усі проєкти</span></button><div class="edessb-section-heading edessb-project-heading"><div><div class="edessb-project-title-line"><span class="edessb-project-number">№${escape(selected.identifier)}</span><h2>${escape(selected.title)}</h2></div><p>${escape(selected.organization)}</p></div><div class="edessb-heading-actions"><button id="open-current-edessb-folder" class="button">${icon('folder')}<span>Папка PDF</span></button><button id="delete-current-edessb-project" class="button archicad-danger">${icon('trash')}<span>Видалити проєкт</span></button></div></div><div class="edessb-document-list">${edessbConfig.documentTypes.map(type=>{const document=documentMap.get(type.id),revisions=document?.revisions||[],next=Math.max(0,...revisions.map(item=>item.number))+1;return `<article class="edessb-document-card ${revisions.length?'ready':''}"><div class="edessb-document-header"><div class="edessb-document-icon">${icon('file-text')}</div><div class="edessb-document-copy"><div><h3>${escape(type.label)}</h3><span class="edessb-document-code">${escape(type.short)}</span></div><p>${revisions.length?`${revisions.length} ред.`:'Редакцій ще немає'}</p></div><button class="button ${revisions.length?'':'primary'}" data-edessb-add-revision="${escape(type.id)}" data-next-revision="${next}">${icon('plus')}<span>${revisions.length?'Додати наступну редакцію':'Додати документ'}</span></button></div>${revisions.length?`<div class="edessb-revision-list">${revisions.map(revision=>`<div class="edessb-revision-row"><div><strong>Редакція №${revision.number}</strong><small>${revision.storedName?escape(revision.storedName):'Без PDF'}${revision.url?' · є посилання':''}</small></div><div class="edessb-revision-actions">${revision.url?`<button class="button" data-edessb-open-link="${escape(revision.url)}">${icon('external-link')}<span>Посилання</span></button>`:''}${revision.storedName?`<button class="button" data-edessb-open-revision="${escape(type.id)}" data-revision-number="${revision.number}" ${revision.fileExists?'':'disabled'}>${icon('file-text')}<span>PDF</span></button>`:''}<button class="button" data-edessb-edit-revision="${escape(type.id)}" data-revision-number="${revision.number}">${icon('pencil')}<span>Редагувати</span></button><button class="button edessb-document-delete" data-edessb-delete-revision="${escape(type.id)}" data-revision-number="${revision.number}" title="Видалити редакцію">${icon('trash')}</button></div></div>`).join('')}</div>`:''}</article>`;}).join('')}</div></section></div>`;
  bindEdessbPortalLinks();
  $('#back-to-edessb-projects').onclick=()=>{selectedEdessbProjectId=null;renderEdessb();};
  $('#open-current-edessb-folder').onclick=()=>window.ekp.openEdessbDocumentsFolder(selected.identifier);
  $('#delete-current-edessb-project').onclick=async()=>{if(!await confirmDelete(`Проєкт «${selected.title}» та всі PDF-документи всередині нього буде видалено без можливості відновлення.`,'Видалити проєкт ЄДЕССБ?'))return;try{await window.ekp.deleteEdessbProject(selected.id);selectedEdessbProjectId=null;await renderEdessb();}catch(error){showArchicadMessage('Не вдалося видалити проєкт',readableError(error));}};
  document.querySelectorAll('[data-edessb-open-link]').forEach(button=>button.onclick=()=>window.ekp.openExternal(button.dataset.edessbOpenLink));
  document.querySelectorAll('[data-edessb-open-revision]').forEach(button=>button.onclick=async()=>{try{await window.ekp.openEdessbRevision({projectId:selected.id,type:button.dataset.edessbOpenRevision,number:Number(button.dataset.revisionNumber)});}catch(error){showArchicadMessage('Не вдалося відкрити PDF',readableError(error));}});
  document.querySelectorAll('[data-edessb-add-revision]').forEach(button=>button.onclick=async()=>{const payload=await requestEdessbRevision(button.dataset.edessbAddRevision,Number(button.dataset.nextRevision));if(!payload)return;button.disabled=true;try{await window.ekp.addEdessbRevision({projectId:selected.id,...payload});await renderEdessb();}catch(error){showArchicadMessage('Не вдалося додати редакцію',readableError(error));}finally{button.disabled=false;}});
  document.querySelectorAll('[data-edessb-edit-revision]').forEach(button=>button.onclick=async()=>{const typeId=button.dataset.edessbEditRevision,number=Number(button.dataset.revisionNumber),revision=documentMap.get(typeId)?.revisions.find(item=>item.number===number);if(!revision)return;const payload=await requestEdessbRevision(typeId,number,revision);if(!payload)return;button.disabled=true;try{await window.ekp.updateEdessbRevision({projectId:selected.id,...payload});await renderEdessb();}catch(error){showArchicadMessage('Не вдалося змінити редакцію',readableError(error));}finally{button.disabled=false;}});
  document.querySelectorAll('[data-edessb-delete-revision]').forEach(button=>button.onclick=async()=>{const type=edessbConfig.documentTypes.find(item=>item.id===button.dataset.edessbDeleteRevision),number=Number(button.dataset.revisionNumber);if(!type||!await confirmDelete(`Редакцію №${number} документа «${type.label}» буде видалено разом із її локальним PDF.`,'Видалити редакцію?'))return;try{await window.ekp.deleteEdessbRevision({projectId:selected.id,type:type.id,number});await renderEdessb();}catch(error){showArchicadMessage('Не вдалося видалити редакцію',readableError(error));}});
}
function requestMaterial(library, existing = null) {
  return new Promise(resolve=>{
    const categories=library.categories,dialog=$('#material-dialog'),source=$('#material-source'),warning=$('#material-builtin-warning'),category=$('#material-category'),name=$('#material-name'),description=$('#material-description'),installation=$('#material-installation'),filesList=$('#material-dialog-files-list'),error=$('#material-dialog-error'),hint=$('#material-dialog-code-hint'),accept=$('#material-dialog-accept');let completed=false,attachments=(existing?.attachments||[]).map(({exists,...item})=>({...item}));
    const finish=value=>{if(completed)return;completed=true;if(dialog.open)dialog.close();resolve(value);};
    $('#material-dialog-title').textContent=existing?'Редагувати матеріал':'Новий матеріал';
    source.value=existing?.source||'user';
    category.innerHTML=categories.map(item=>`<option value="${escape(item.code)}">${escape(item.name)} (${escape(item.code)})</option>`).join('');
    const preferredCategory=existing?.categoryCode||(selectedMaterialCategory!=='all'&&categories.some(item=>item.code===selectedMaterialCategory)?selectedMaterialCategory:'');
    category.value=preferredCategory||categories[0]?.code||'';
    name.value=existing?.name||'';description.value=existing?.description||'';installation.value=existing?.installationRules||'';error.hidden=true;
    const renderFiles=()=>{filesList.innerHTML=attachments.length?attachments.map((file,index)=>`<div class="material-dialog-file"><div>${icon('file-text')}<span title="${escape(file.name)}">${escape(file.name)}</span>${file.sourcePath?'<small>новий</small>':''}</div><button class="button material-dialog-file-remove" type="button" data-material-dialog-file-remove="${index}" title="Прибрати файл" aria-label="Прибрати файл">${icon('x')}</button></div>`).join(''):'<div class="material-dialog-files-empty">Файли ще не прикріплено</div>';filesList.querySelectorAll('[data-material-dialog-file-remove]').forEach(button=>button.onclick=()=>{attachments.splice(Number(button.dataset.materialDialogFileRemove),1);renderFiles();});};
    const updateHint=()=>{const prefix=category.value,marker=source.value==='user'?'К':'';warning.hidden=source.value!=='builtin';accept.textContent=source.value==='builtin'?'Зберегти як базовий':'Зберегти як користувацький';hint.textContent=existing&&existing.categoryCode===prefix&&existing.source===source.value?`Ідентифікаційний код: ${existing.code}`:`Новий код матиме формат ${prefix}${marker}0001 і буде визначений автоматично.`;};
    category.onchange=updateHint;source.onchange=updateHint;updateHint();
    renderFiles();
    $('#material-dialog-add-files').onclick=async()=>{try{const selected=await window.ekp.selectMaterialFiles();const known=new Set(attachments.filter(item=>item.sourcePath).map(item=>item.sourcePath));attachments.push(...selected.filter(item=>!known.has(item.sourcePath)));renderFiles();}catch(fileError){error.textContent=readableError(fileError);error.hidden=false;}};
    $('#material-dialog-cancel').onclick=()=>finish(null);
    $('#material-dialog-accept').onclick=()=>{const payload={id:existing?.id||'',source:source.value,categoryCode:category.value,name:name.value.trim(),description:description.value.trim(),installationRules:installation.value.trim(),attachments};if(!payload.name){error.textContent='Вкажіть назву матеріалу.';error.hidden=false;name.focus();return;}finish(payload);};
    dialog.oncancel=event=>{event.preventDefault();finish(null);};dialog.onclose=()=>finish(null);dialog.showModal();requestAnimationFrame(()=>name.focus());
  });
}
function materialMatches(material) {
  const query=materialSearch.trim().toLocaleLowerCase('uk-UA');
  if(selectedMaterialCategory!=='all'&&material.categoryCode!==selectedMaterialCategory)return false;
  return !query||[material.code,material.name,material.description,material.installationRules,material.category,...(material.attachments||[]).map(file=>file.name)].some(value=>String(value||'').toLocaleLowerCase('uk-UA').includes(query));
}
function materialResultsHtml(library) {
  const materials=library.materials.filter(materialMatches),category=library.categories.find(item=>item.code===selectedMaterialCategory);
  return `<div class="materials-results-head"><div><h2>${escape(category?.name||'Усі матеріали')}</h2><p>Знайдено: ${materials.length}</p></div></div><div class="materials-list">${materials.length?materials.map(material=>{const expanded=expandedMaterialIds.has(material.id),attachments=material.attachments||[];return `<article class="material-card ${expanded?'expanded':''}"><div class="material-card-header"><span class="material-code">${escape(material.code)}</span><div class="material-card-title"><h3>${escape(material.name)}</h3></div>${attachments.length?`<span class="material-attachment-count" title="Прикріплених файлів: ${attachments.length}">${icon('file-text')}<span>${attachments.length}</span></span>`:''}<span class="material-source ${material.source==='user'?'user':''}">${material.source==='user'?'Користувацький':'Базовий'}</span><div class="material-card-actions"><button class="button material-expand" data-material-toggle="${escape(material.id)}" aria-expanded="${expanded}"><span>${expanded?'Згорнути':'Розкрити'}</span>${icon('chevron-down')}</button><button class="button material-icon-action" data-material-edit="${escape(material.id)}" title="Редагувати матеріал" aria-label="Редагувати матеріал">${icon('pencil')}</button><button class="button material-icon-action danger" data-material-delete="${escape(material.id)}" title="Видалити матеріал" aria-label="Видалити матеріал">${icon('trash')}</button></div></div>${expanded?`<div class="material-sections"><div class="material-section"><strong>Опис</strong><p>${escape(material.description||'Опис не вказано')}</p></div><div class="material-section"><strong>Правила монтажу</strong><p>${escape(material.installationRules||'Правила монтажу не вказано')}</p></div></div><div class="material-files-section"><div class="material-files-title"><strong>Прикріплені файли</strong><span>${attachments.length}</span></div>${attachments.length?`<div class="material-files-list">${attachments.map(file=>`<button class="button material-file-button" data-material-open-file="${escape(material.id)}" data-material-attachment="${escape(file.id)}" ${file.exists?'':'disabled'}>${icon('file-text')}<span title="${escape(file.name)}">${escape(file.name)}</span>${icon('external-link')}</button>`).join('')}</div>`:'<p>Файлів не прикріплено. Додати їх можна через редагування матеріалу.</p>'}</div>`:''}</article>`;}).join(''):`<div class="material-empty">${icon('search')}<h3>Матеріалів не знайдено</h3><p>Змініть категорію або пошуковий запит чи додайте власний матеріал.</p></div>`}</div>`;
}
function bindMaterialResultActions(library, refreshResults) {
  document.querySelectorAll('[data-material-toggle]').forEach(button=>button.onclick=()=>{const id=button.dataset.materialToggle;if(expandedMaterialIds.has(id))expandedMaterialIds.delete(id);else expandedMaterialIds.add(id);refreshResults();});
  document.querySelectorAll('[data-material-edit]').forEach(button=>button.onclick=async()=>{const material=library.materials.find(item=>item.id===button.dataset.materialEdit);if(!material)return;const payload=await requestMaterial(library,material);if(!payload)return;try{await window.ekp.saveMaterial(payload);await renderMaterials();}catch(error){showArchicadMessage('Не вдалося зберегти матеріал',readableError(error));}});
  document.querySelectorAll('[data-material-delete]').forEach(button=>button.onclick=async()=>{const material=library.materials.find(item=>item.id===button.dataset.materialDelete);if(!material||!await confirmDelete(`Матеріал «${material.name}» (${material.code}) буде видалено без можливості відновлення.`,'Видалити матеріал?'))return;try{await window.ekp.deleteMaterial(material.id);expandedMaterialIds.delete(material.id);await renderMaterials();}catch(error){showArchicadMessage('Не вдалося видалити матеріал',readableError(error));}});
  document.querySelectorAll('[data-material-open-file]').forEach(button=>button.onclick=async()=>{const material=library.materials.find(item=>item.id===button.dataset.materialOpenFile),attachment=material?.attachments?.find(item=>item.id===button.dataset.materialAttachment);if(!attachment)return;try{await window.ekp.openMaterialAttachment(attachment.storedName);}catch(error){showArchicadMessage('Не вдалося відкрити файл',readableError(error));}});
}
async function renderMaterials() {
  page.innerHTML=head('Будівельні матеріали')+`<div class="page-content materials-page"><div class="materials-placeholder">Завантаження бази матеріалів…</div></div>`;
  let library;try{library=await window.ekp.materialsLibrary();}catch(error){page.innerHTML=head('Будівельні матеріали')+`<div class="page-content materials-page"><div class="material-empty">${icon('info')}<h3>Не вдалося прочитати базу</h3><p>${escape(readableError(error))}</p></div></div>`;return;}
  if(currentPage!=='materials')return;
  if(selectedMaterialCategory!=='all'&&!library.categories.some(item=>item.code===selectedMaterialCategory))selectedMaterialCategory='all';
  const allCount=library.materials.length;
  page.innerHTML=head('Будівельні матеріали')+`<div class="page-content materials-page"><div class="materials-page-heading"><div><h1 class="hero">Будівельні матеріали</h1><p class="subtitle">Каталог матеріалів із незмінними унікальними кодами</p></div><div class="materials-heading-actions"><button id="open-materials-folder" class="button">${icon('folder')}<span>Папка файлів</span></button><button id="open-materials-excel" class="button">${icon('sheet')}<span>Відкрити Excel</span></button><button id="create-user-material" class="button primary">${icon('plus')}<span>Додати матеріал</span></button></div></div><div class="materials-info">${icon('info')}<div><strong>Базові й користувацькі матеріали зберігаються в одному Excel, а вкладення — в окремій папці.</strong> Користувацькі записи з літерою «К» у коді та їхні файли зберігаються під час оновлення. Базові записи й базові файли, додані через встановлену програму, призначені лише для підготовки бази розробником і можуть бути замінені наступною версією CoDA. Версія бази: ${escape(library.databaseVersion||'не вказана')}.</div></div><div class="materials-controls"><input id="materials-search" class="input" value="${escape(materialSearch)}" placeholder="Пошук за кодом, назвою, описом, правилами монтажу або вкладенням"></div><div class="materials-layout"><aside class="materials-categories"><h2>Категорії</h2><div class="materials-category-list"><button data-material-category="all" class="${selectedMaterialCategory==='all'?'active':''}"><span>Усі матеріали</span><strong>${allCount}</strong></button>${library.categories.map(category=>`<button data-material-category="${escape(category.code)}" class="${selectedMaterialCategory===category.code?'active':''}" title="${escape(category.description)}"><span>${escape(category.name)}</span><strong>${category.count}</strong></button>`).join('')}</div></aside><section id="materials-results">${materialResultsHtml(library)}</section></div></div>`;
  const refreshResults=()=>{const results=$('#materials-results');results.innerHTML=materialResultsHtml(library);bindMaterialResultActions(library,refreshResults);document.querySelectorAll('[data-material-category]').forEach(button=>button.classList.toggle('active',button.dataset.materialCategory===selectedMaterialCategory));};
  $('#materials-search').oninput=event=>{materialSearch=event.target.value;refreshResults();};
  document.querySelectorAll('[data-material-category]').forEach(button=>button.onclick=()=>{selectedMaterialCategory=button.dataset.materialCategory;refreshResults();});
  $('#open-materials-excel').onclick=async()=>{try{await window.ekp.openMaterialsDatabase();}catch(error){showArchicadMessage('Не вдалося відкрити Excel',readableError(error));}};
  $('#open-materials-folder').onclick=async()=>{try{await window.ekp.openMaterialAttachmentsFolder();}catch(error){showArchicadMessage('Не вдалося відкрити папку файлів',readableError(error));}};
  $('#create-user-material').onclick=async()=>{const payload=await requestMaterial(library);if(!payload)return;try{const saved=await window.ekp.saveMaterial(payload);selectedMaterialCategory=saved.categoryCode;materialSearch='';await renderMaterials();}catch(error){showArchicadMessage('Не вдалося додати матеріал',readableError(error));}};
  bindMaterialResultActions(library,refreshResults);
}
function navigate(to){rememberView({type:'page',id:to});showHistoryView({type:'page',id:to});}
document.querySelectorAll('.primary-nav button').forEach(button=>button.onclick=()=>navigate(button.dataset.page));
$('#toggle-sidebar').onclick=()=>$('#sidebar').classList.toggle('collapsed');
$('#nav-back').onclick=()=>moveHistory(-1);
$('#nav-forward').onclick=()=>moveHistory(1);
const closeSettings=()=>{ $('#settings-dropdown').hidden=true; $('#settings-toggle').setAttribute('aria-expanded','false'); };
$('#settings-toggle').onclick=event=>{ event.stopPropagation(); const willOpen=$('#settings-dropdown').hidden; $('#settings-dropdown').hidden=!willOpen; $('#settings-toggle').setAttribute('aria-expanded',String(willOpen)); };
$('#settings-theme').onclick=async()=>{state.theme=state.theme==='dark'?'light':'dark';await saveState();renderSidebar();closeSettings();};
$('#settings-program-update').onclick=()=>{closeSettings();checkProgramUpdate(true);};
$('#settings-about').onclick=()=>{closeSettings();$('#about').showModal();};
$('#settings-updates').onclick=async()=>{closeSettings();const entries=await window.ekp.updateHistory();$('#updates-content').innerHTML=entries.length?`<div class="updates-list">${entries.map(entry=>`<article class="update-entry"><div class="update-version">${escape(entry.version)}</div><div class="update-description">${escape(entry.description)}</div></article>`).join('')}</div>`:'<p class="updates-empty">Історія поки порожня. Заповніть файл update_history.xlsx у папці data.</p>';$('#updates').showModal();};
$('#settings-user-data').onclick=()=>{closeSettings();$('#user-data-status').textContent='';$('#user-data-dialog').showModal();};
document.addEventListener('click',event=>{if(!$('#settings-menu').contains(event.target))closeSettings();});
document.querySelectorAll('.dialog-close').forEach(button=>button.onclick=()=>button.closest('dialog').close());
$('#check-database-update').onclick=checkDatabaseUpdate;
$('#database-update-cancel').onclick=()=>$('#database-update-dialog').close();
$('#database-update-install').onclick=installDatabaseUpdate;
$('#program-update-cancel').onclick=()=>$('#program-update-dialog').close();
$('#program-update-close-x').onclick=()=>$('#program-update-dialog').close();
$('#program-update-install').onclick=installProgramUpdate;
$('#archicad-message-ok').onclick=()=>$('#archicad-message-dialog').close();
$('#user-data-export').onclick=async()=>{const button=$('#user-data-export'),status=$('#user-data-status');button.disabled=true;try{const groups=await window.ekp.userDataGroups(),ids=await requestUserDataGroups(groups,'export');if(!ids)return;status.textContent='Створення резервної копії…';const result=await window.ekp.exportUserData(ids);if(result.canceled){status.textContent='Експорт скасовано.';return;}$('#user-data-dialog').close();showArchicadMessage('Резервну копію створено',`Вибрані дані збережено у файл:\n${result.packagePath}`);}catch(error){status.textContent=`Не вдалося створити копію: ${readableError(error)}`;}finally{button.disabled=false;}};
$('#user-data-import').onclick=async()=>{const button=$('#user-data-import'),status=$('#user-data-status');button.disabled=true;try{status.textContent='Перевірка резервної копії…';const prepared=await window.ekp.prepareUserDataImport();if(prepared.canceled){status.textContent='Імпорт скасовано.';return;}const ids=await requestUserDataGroups(prepared.groups,'import');if(!ids){status.textContent='Імпорт скасовано.';return;}const selectedGroups=prepared.groups.filter(group=>ids.includes(group.id));if(!await confirmUserDataImport(selectedGroups.map(group=>group.label))){status.textContent='Імпорт скасовано.';return;}status.textContent='Відновлення вибраних даних…';const result=await window.ekp.importUserData(ids);[state,libraryItems]=await Promise.all([window.ekp.state(),window.ekp.libraryItems()]);selectedLibraryCategory='dstu';selectedArchicadProjectId=null;selectedEdessbProjectId=null;viewHistory=[];viewHistoryIndex=-1;$('#user-data-dialog').close();renderSidebar();renderDatabaseUpdate();navigate('dbn');showArchicadMessage('Вибрані дані відновлено',`Замінено: ${selectedGroups.map(group=>group.label).join(', ')}. Інші категорії не змінено. Імпортовано файлів: ${result.fileCount}.`);}catch(error){status.textContent=`Не вдалося імпортувати дані: ${readableError(error)}`;}finally{button.disabled=false;}};
$('#minimize').onclick=()=>window.ekp.window.minimize();$('#maximize').onclick=()=>window.ekp.window.maximize();$('#close').onclick=()=>window.ekp.window.close();
(async()=>{let appVersion;[state,catalog,libraryItems,libraryConfig,catalogMetadata,appVersion]=await Promise.all([window.ekp.state(),window.ekp.catalog(),window.ekp.libraryItems(),window.ekp.libraryConfig(),window.ekp.catalogMetadata(),window.ekp.appVersion()]);$('#about-version').textContent=appVersion;renderSidebar();renderDatabaseUpdate();navigate('dbn');setTimeout(()=>checkProgramUpdate(false),1200);})();
