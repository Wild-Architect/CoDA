const $ = selector => document.querySelector(selector);
const page = $('#page');
let state, catalog, notes, importedNotes = [], catalogMetadata = {}, selectedNote = null, currentPage = 'dbn';
let pdfjsLibPromise = null, activePdf = null;
let viewHistory = [], viewHistoryIndex = -1;
let attachmentPreviewUrls = [];
let pendingDatabaseManifest = null;
let pendingProgramUpdate = null;
let selectedArchicadProjectId = null;
const allNotes = () => [...notes, ...importedNotes];
const escape = value => String(value || '').replace(/[&<>"]/g, char => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' }[char]));
const icon = (name, className = '') => `<svg class="icon ${className}" aria-hidden="true"><use href="icons.svg#${name}"></use></svg>`;
const saveState = async () => window.ekp.saveState(state);
const recent = async (key, value) => { state[key] = [value, ...(state[key] || []).filter(item => item !== value)].slice(0, 8); await saveState(); renderSidebar(); };

function clearAttachmentPreviews() {
  attachmentPreviewUrls.forEach(url => URL.revokeObjectURL(url));
  attachmentPreviewUrls = [];
}
function scrollChatToEnd() {
  requestAnimationFrame(() => {
    const surface = $('.surface');
    const thread = $('.chat-thread');
    if (surface) surface.scrollTop = 0;
    if (thread) thread.scrollTop = thread.scrollHeight;
  });
}
function attachmentOpenLabel(filename) {
  const extension = String(filename || '').split('.').pop().toLowerCase();
  if (['doc','docx','rtf'].includes(extension)) return 'Відкрити у Word';
  if (['xls','xlsx','xlsm','csv'].includes(extension)) return 'Відкрити в Excel';
  if (['ppt','pptx'].includes(extension)) return 'Відкрити у PowerPoint';
  if (['zip','rar','7z'].includes(extension)) return 'Відкрити архів';
  return 'Відкрити файл';
}
async function loadAttachmentPreviews(note) {
  const imageTypes = new Set(['jpg','jpeg','png','gif','webp','bmp','svg']);
  for (const preview of document.querySelectorAll('[data-preview-file]')) {
    const index = Number(preview.dataset.previewFile);
    const file = note.attachments[index];
    const extension = String(file?.name || '').split('.').pop().toLowerCase();
    if (file && extension === 'pdf') {
      preview.classList.add('pdf-launch');
      preview.innerHTML = `<div class="pdf-launch-icon">${icon('file-text')}</div><div class="attachment-copy"><strong title="${escape(file.name)}">${escape(file.name)}</strong><span>PDF-документ</span></div><button class="button primary">Переглянути PDF</button>`;
      preview.querySelector('button').onclick = () => openAttachmentPdf(note, file);
      continue;
    }
    if (!file || !imageTypes.has(extension)) {
      preview.classList.add('generic');
      preview.innerHTML = `<div class="generic-file-icon">${icon('file-text')}</div><div class="attachment-copy"><strong title="${escape(file?.name)}">${escape(file?.name)}</strong><span>${extension ? extension.toUpperCase() + '-файл' : 'Файл'}</span></div><button class="button">${attachmentOpenLabel(file?.name)}</button>`;
      preview.querySelector('button').onclick = () => window.ekp.openFile(file.path);
      continue;
    }
    try {
      const result = await window.ekp.readAttachment(file.path);
      const url = URL.createObjectURL(new Blob([new Uint8Array(result.data)], { type: result.mime }));
      if (!document.body.contains(preview) || selectedNote !== note) { URL.revokeObjectURL(url); continue; }
      attachmentPreviewUrls.push(url);
      preview.classList.add('image-preview');
      preview.innerHTML = `<img src="${url}" alt="${escape(file.name)}" title="Відкрити у стандартній програмі">`;
      preview.querySelector('img').onclick = () => window.ekp.openFile(file.path);
      preview.querySelector('img').onload = scrollChatToEnd;
    } catch {
      preview.classList.add('generic');
      preview.textContent = 'Не вдалося завантажити попередній перегляд';
    }
  }
  scrollChatToEnd();
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

function renderSidebar() {
  document.documentElement.dataset.theme = state.theme || 'light';
  $('#settings-theme-label').textContent = state.theme === 'dark' ? 'Світла тема' : 'Темна тема';
  $('#settings-theme-icon').setAttribute('href', state.theme === 'dark' ? 'icons.svg#sun' : 'icons.svg#moon');
  document.querySelectorAll('.primary-nav button').forEach(button => button.classList.toggle('active', button.dataset.page === currentPage));
  const pinnedNotes = notes.filter(note => note.pinned).map(note => ({ label: note.title, kind:'notebook-pen', action: () => openNote(note), unpin: async () => { note.pinned=false; await window.ekp.saveNote(note); notes=await window.ekp.notes(); if(selectedNote?.id===note.id) selectedNote=notes.find(item=>item.id===note.id)||selectedNote; renderSidebar(); } }));
  const pinnedDbn = catalog.filter(item => (state.pinned_dbn || []).includes(item.path)).map(item => ({ label:item.number, kind:'book-open', action: () => openDbn(item), unpin: async () => { state.pinned_dbn=(state.pinned_dbn||[]).filter(path=>path!==item.path); await saveState(); renderSidebar(); } }));
  const noteMap = new Map(notes.map(note => [note.id, note])); const dbnMap = new Map(catalog.map(item => [item.path, item]));
  const recently = [...(state.recent_notes || []).map(id => noteMap.get(id)).filter(Boolean).map(note => ({label:note.title,kind:'notebook-pen',action:()=>openNote(note)})), ...(state.recent_dbn || []).map(id => dbnMap.get(id)).filter(Boolean).map(item => ({label:item.number,kind:'book-open',action:()=>openDbn(item)}))];
  const list = (title, items) => !items.length ? '' : `<div class="side-heading">${title}</div><div class="sidebar-list">${items.slice(0,6).map((item,i) => `<div class="sidebar-list-item"><button data-list="${title}" data-index="${i}">${icon(item.kind, 'small')}<span>${escape(item.label)}</span></button>${item.unpin?`<button class="sidebar-unpin" data-unpin="${i}" title="Прибрати із закріплених">${icon('x','small')}</button>`:''}</div>`).join('')}</div>`;
  const pinned = [...pinnedNotes, ...pinnedDbn]; $('#side-lists').innerHTML = list('Закріплені', pinned) + list('Нещодавні', recently);
  document.querySelectorAll('[data-list]').forEach(button => button.onclick = () => (button.dataset.list === 'Закріплені' ? pinned : recently)[Number(button.dataset.index)].action());
  document.querySelectorAll('[data-unpin]').forEach(button => button.onclick = event => { event.stopPropagation(); pinned[Number(button.dataset.unpin)]?.unpin?.(); });
}

function formatCatalogDate(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return 'не вказана';
  const months = ['січня','лютого','березня','квітня','травня','червня','липня','серпня','вересня','жовтня','листопада','грудня'];
  return `${Number(match[3])} ${months[Number(match[2]) - 1]} ${match[1]} року`;
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
  } catch (error) { pendingDatabaseManifest = null; showDatabaseUpdateDialog('Не вдалося перевірити базу', error?.message || String(error)); }
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
  } catch (error) { showDatabaseUpdateDialog('Оновлення не встановлено', `${error?.message || String(error)}\n\nЛокальна база залишилася без змін.`); }
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
async function checkProgramUpdate(manual = false) {
  if (manual) showProgramUpdateDialog('Перевірка оновлень…', 'CoDA перевіряє наявність нової версії.');
  try {
    const result = await window.ekp.checkProgramUpdate();
    if (result.updateAvailable) {
      pendingProgramUpdate = result;
      const notes = result.releaseNotes ? `\n\n${result.releaseNotes}` : '';
      showProgramUpdateDialog('Доступна нова версія CoDA', `Встановлена версія: ${result.currentVersion}\nНова версія: ${result.latestVersion}${notes}`, true);
    } else if (manual) {
      pendingProgramUpdate = null;
      showProgramUpdateDialog(result.supported ? 'CoDA оновлена' : 'Перевірка у встановленій програмі', result.message || `У вас встановлена найновіша версія ${result.currentVersion}.`);
    }
  } catch (error) {
    pendingProgramUpdate = null;
    if (manual) showProgramUpdateDialog('Не вдалося перевірити оновлення', error?.message || String(error));
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
    showProgramUpdateDialog('Оновлення не встановлено', error?.message || String(error), true);
    install.disabled = false; cancel.disabled = false; close.disabled = false;
  } finally { stopProgress(); }
}
function openExportDialog() {
  $('#export-notes-list').innerHTML=notes.length?notes.map(note=>`<label class="admin-note-option"><input type="checkbox" value="${escape(note.id)}"><span>${escape(note.title)}</span></label>`).join(''):'<p class="updates-empty">Немає власних нотаток для експорту.</p>';
  $('#export-status').textContent='';
  $('#export-dialog').showModal();
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
  if (view.type !== 'pdf') disposePdfViewer();
  if (view.type === 'attachment-pdf') {
    const note = (view.imported ? importedNotes : notes).find(entry => entry.id === view.noteId);
    const file = note?.attachments?.find(entry => entry.path === view.id);
    if (note && file) { selectedNote = note; currentPage = 'notes'; await renderPdfViewer({ number: file.name, path: file.path, attachment: true }); }
  } else if (view.type === 'pdf') {
    const item = catalog.find(entry => entry.path === view.id);
    if (item) { currentPage = 'dbn'; await renderPdfViewer(item); }
  } else if (view.type === 'note') {
    selectedNote = (view.imported ? importedNotes : notes).find(note => note.id === view.id) || selectedNote;
    currentPage = 'notes'; renderNotes();
  } else {
    currentPage = view.id;
    await ({ dbn: renderDbn, notes: renderNotes, excel: renderExcel }[view.id])();
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
  activePdf = null;
  window.onkeydown = null; window.onkeyup = null; window.onblur = null;
}

function head(title) { return `<div class="page-head"><span class="folder">${icon('folder')}</span><span class="page-title">${title}</span></div>`; }
function renderDbn() {
  $('.surface').classList.remove('notes-mode');
  const categoryOf = item => String(item.category || '').trim() || 'Без категорії';
  const categoryCounts = new Map();
  catalog.forEach(item => categoryCounts.set(categoryOf(item), (categoryCounts.get(categoryOf(item)) || 0) + 1));
  const categories = [...categoryCounts.keys()].sort((a, b) => a.localeCompare(b, 'uk'));
  const categoryButtons = [`<button class="active" data-category-index="-1"><span>Усі категорії</span><strong>${catalog.length}</strong></button>`, ...categories.map((category, index) => `<button data-category-index="${index}" title="${escape(category)}"><span>${escape(category)}</span><strong>${categoryCounts.get(category)}</strong></button>`)].join('');
  const rows = catalog.map(item => `<tr data-path="${escape(item.path)}" data-category="${escape(categoryOf(item))}"><td class="pin">${(state.pinned_dbn||[]).includes(item.path)?icon('star','small star'):''}</td><td>${escape(item.number)}</td><td>${escape(item.title)}</td><td class="edessb-cell"><button class="edessb-button${item.edessb_url?'':' is-empty'}" data-edessb="${escape(item.edessb_url)}" title="${item.edessb_url?'Відкрити сторінку ДБН у ЄДЕССБ':'Додайте посилання в Excel'}">ЄДЕССБ</button></td></tr>`).join('');
  page.innerHTML = head('Державні будівельні норми') + `<div class="page-content dbn-page"><h1 class="hero">Державні будівельні норми</h1><p class="subtitle">Локальна бібліотека • ${catalog.length} документів</p><div class="dbn-layout"><section class="dbn-catalog"><table class="table"><thead><tr><th></th><th>Номер</th><th>Назва</th><th class="edessb-heading">Посилання</th></tr></thead><tbody>${rows}</tbody></table></section><aside class="dbn-categories"><div class="dbn-categories-title">Категорії</div><div class="dbn-category-list">${categoryButtons}</div></aside></div><div class="dbn-floating-controls" role="search"><input id="dbn-search" class="input" placeholder="Пошук за номером або назвою"><button id="pin-dbn" class="button">${icon('pin')}<span>Закріпити / відкріпити</span></button></div></div>`;
  let selectedRow = null; document.querySelectorAll('.table tbody tr').forEach(row => { row.onclick=()=>{ document.querySelectorAll('.table tbody tr').forEach(item=>item.style.background=''); row.style.background='var(--hover)'; selectedRow=row; }; row.ondblclick=()=>openDbn(catalog.find(item=>item.path===row.dataset.path)); });
  document.querySelectorAll('[data-edessb]').forEach(button => button.onclick = event => { event.stopPropagation(); if (button.dataset.edessb) window.ekp.openExternal(button.dataset.edessb); });
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

function scrollToPdfPage(pageNumber, behavior = 'smooth') {
  if (!activePdf?.document) return;
  activePdf.pageNumber = Math.max(1, Math.min(activePdf.document.numPages, pageNumber));
  document.querySelector(`[data-pdf-page="${activePdf.pageNumber}"]`)?.scrollIntoView({ behavior, block: 'start' });
  $('#pdf-page-number').value = activePdf.pageNumber;
  renderPdfPage(activePdf.pageNumber);
}

function buildPdfPages() {
  activePdf.textLayers.forEach(layer => layer.cancel());
  activePdf.textLayers.clear();
  const pages = Array.from({ length: activePdf.document.numPages }, (_, index) => `<section class="pdf-page-shell" data-pdf-page="${index + 1}"><canvas></canvas><div class="textLayer" aria-label="Текст сторінки ${index + 1}"></div><span class="pdf-page-label">${index + 1}</span></section>`).join('');
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
  activePdf.scale = Math.max(.5, Math.min(3, Math.round(scale * 10) / 10));
  activePdf.renderTasks.forEach(task => task?.cancel?.());
  activePdf.renderTasks.clear();
  activePdf.textLayers.forEach(layer => layer.cancel());
  activePdf.textLayers.clear();
  activePdf.rendered.clear();
  $('#pdf-zoom-value').textContent = `${Math.round(activePdf.scale * 100)}%`;
  const pageNumber = activePdf.pageNumber;
  buildPdfPages();
  requestAnimationFrame(() => scrollToPdfPage(pageNumber, 'auto'));
}

async function renderPdfViewer(item) {
  $('.surface').classList.remove('notes-mode');
  disposePdfViewer();
  activePdf = { item, document: null, pdfjs: null, pageNumber: 1, scale: 1.25, singleScale: 1.25, layout: 'single', rendered: new Set(), renderTasks: new Map(), textLayers: new Map(), observer: null, rightAlt: false, zoomWheelLocked: false };
  const backLabel = item.attachment ? 'До нотатки' : 'До каталогу';
  page.innerHTML = head(item.number) + `<div class="pdf-viewer"><div class="pdf-toolbar"><button id="pdf-back" class="button">${icon('chevron-left')}<span>${backLabel}</span></button><div class="pdf-toolbar-group"><button id="pdf-prev" class="pdf-icon-button" title="Попередня сторінка">${icon('chevron-left')}</button><input id="pdf-page-number" class="pdf-page-input" type="number" min="1" value="1"><span id="pdf-page-count">/ —</span><button id="pdf-next" class="pdf-icon-button" title="Наступна сторінка">${icon('chevron-right')}</button></div><div class="pdf-layout-switch" aria-label="Режим відображення сторінок"><button id="pdf-layout-single" class="pdf-icon-button active" title="Одна сторінка" aria-pressed="true">${icon('page-single')}</button><button id="pdf-layout-spread" class="pdf-icon-button" title="Дві сторінки поруч" aria-pressed="false">${icon('pages-two')}</button></div><div class="pdf-toolbar-group"><button id="pdf-zoom-out" class="pdf-icon-button" title="Зменшити">${icon('minus')}</button><span id="pdf-zoom-value">125%</span><button id="pdf-zoom-in" class="pdf-icon-button" title="Збільшити">${icon('plus')}</button></div><span class="pdf-zoom-hint">Правий Alt + колесо — масштаб</span><button id="pdf-open-system" class="button">${icon('external-link')}<span>Відкрити окремо</span></button></div><div id="pdf-stage" class="pdf-stage"><div id="pdf-status" class="pdf-status">Завантаження документа…</div></div></div>`;
  $('#pdf-back').onclick = () => item.attachment ? moveHistory(-1) : navigate('dbn');
  $('#pdf-open-system').onclick = () => item.attachment ? window.ekp.openFile(item.path) : window.ekp.openDbn(item.path);
  $('#pdf-prev').onclick = () => scrollToPdfPage(activePdf.pageNumber - 1);
  $('#pdf-next').onclick = () => scrollToPdfPage(activePdf.pageNumber + 1);
  $('#pdf-layout-single').onclick = () => setPdfLayout('single');
  $('#pdf-layout-spread').onclick = () => setPdfLayout('spread');
  $('#pdf-zoom-out').onclick = () => setPdfScale(activePdf.scale - .25);
  $('#pdf-zoom-in').onclick = () => setPdfScale(activePdf.scale + .25);
  $('#pdf-page-number').onchange = event => scrollToPdfPage(Number(event.target.value) || 1);
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
    const [pdfjs, source] = await Promise.all([loadPdfJs(), item.attachment ? window.ekp.readAttachment(item.path) : window.ekp.readPdf(item.path)]);
    const data = item.attachment ? source.data : source;
    activePdf.pdfjs = pdfjs;
    activePdf.document = await pdfjs.getDocument({ data: new Uint8Array(data) }).promise;
    $('#pdf-page-count').textContent = `/ ${activePdf.document.numPages}`;
    buildPdfPages();
    await renderPdfPage(1);
  } catch (error) {
    const status = $('#pdf-status');
    if (status) status.textContent = `Не вдалося відкрити PDF: ${error.message || error}`;
    else console.error(error);
  }
}

async function openDbn(item) { await recent('recent_dbn', item.path); currentPage = 'dbn'; rememberView({ type: 'pdf', id: item.path }); await renderPdfViewer(item); renderSidebar(); }
async function openAttachmentPdf(note, file) { selectedNote = note; currentPage = 'notes'; rememberView({ type: 'attachment-pdf', id: file.path, noteId: note.id, imported:Boolean(note.imported) }); await renderPdfViewer({ number: file.name, path: file.path, attachment: true }); renderSidebar(); }
function renderNotes() {
  $('.surface').classList.add('notes-mode');
  clearAttachmentPreviews();
  if (selectedNote && !allNotes().includes(selectedNote) && selectedNote.id) selectedNote = (selectedNote.imported ? importedNotes : notes).find(note=>note.id===selectedNote.id) || null;
  selectedNote ||= notes[0] || importedNotes[0] || { title:'', body:'', messages:[], attachments:[], pinned:false };
  const imported = Boolean(selectedNote.imported);
  selectedNote.attachments ||= [];
  if (!Array.isArray(selectedNote.messages) || (!selectedNote.messages.length && selectedNote.body?.trim())) selectedNote.messages = selectedNote.body?.trim() ? [{ id:`legacy-${selectedNote.id||'new'}`, text:selectedNote.body, created_at:selectedNote.created_at||'' }] : [];
  const privateButtons=notes.map(note=>`<button data-note="${note.id}" data-imported="false" class="${note===selectedNote?'active':''}">${note.pinned?icon('star','small star'):icon('file-text','small')}${escape(note.title)}</button>`).join('');
  const importedButtons=importedNotes.map(note=>`<button data-note="${note.id}" data-imported="true" class="${note===selectedNote?'active':''}">${icon('library','small')}${escape(note.title)}</button>`).join('');
  const attachmentCard=(file,index)=>`<article class="chat-attachment"><div class="attachment-preview" data-preview-file="${index}"><span>Завантаження перегляду…</span></div>${imported?'':`<button class="attachment-delete attachment-card-delete" data-delete-file="${index}" title="Видалити вкладення">${icon('trash','small')}</button>`}</article>`;
  const latestMessageTime=Math.max(Date.parse(selectedNote.created_at||'')||0,...selectedNote.messages.map(message=>Date.parse(message.created_at||'')||0));
  selectedNote.attachments.forEach((file,index)=>{if(!file.created_at)file.created_at=new Date(latestMessageTime+index+1).toISOString();});
  const timeline=[...selectedNote.messages.map((message,index)=>({kind:'message',index,time:Date.parse(message.created_at||'')||index})),...selectedNote.attachments.map((file,index)=>({kind:'attachment',index,time:Date.parse(file.created_at||'')||latestMessageTime+index+1}))].sort((a,b)=>a.time-b.time||a.index-b.index);
  const timelineHtml=timeline.map(item=>item.kind==='message'
    ? `<div class="chat-message"><article class="chat-bubble">${imported?'':`<button class="message-delete" data-delete-message="${item.index}" title="Видалити повідомлення">${icon('trash','small')}</button>`}<p class="chat-bubble-text">${escape(selectedNote.messages[item.index].text)}</p></article></div>`
    : `<div class="chat-message"><article class="chat-bubble attachment-bubble"><div class="chat-attachments">${attachmentCard(selectedNote.attachments[item.index],item.index)}</div></article></div>`).join('');
  const header=imported
    ? `<header class="chat-header"><div class="chat-title">${escape(selectedNote.title)}</div><button id="delete-imported-note" class="button imported-note-delete" title="Видалити імпортовану нотатку">${icon('trash')}<span>Видалити</span></button><span class="readonly-badge">${icon('library','small')}Імпортована</span></header><div class="public-note-meta">Добірка: ${escape(selectedNote.source_library||'не вказана')} · Автор: ${escape(selectedNote.source_author||selectedNote.author||'не вказаний')}</div>`
    : `<header class="chat-header"><input id="note-title" class="chat-title" placeholder="Назва нотатки" value="${escape(selectedNote.title)}"><button id="save-note-title" class="button title-save" title="Зберегти назву">${icon('save')}<span>Зберегти назву</span></button><button id="pin-note" class="button" title="${selectedNote.pinned?'Відкріпити':'Закріпити'}">${icon(selectedNote.pinned?'star':'pin',selectedNote.pinned?'small star':'')}</button><button id="delete-note" class="button" title="Видалити нотатку">${icon('trash')}</button></header>`;
  const composer=`<div class="chat-composer${imported?' imported-composer':''}">${imported?'':`<button id="attach" class="composer-attach" title="Прикріпити файл">${icon('plus')}</button>`}<textarea id="note-body" placeholder="${imported?'Додайте повідомлення — нотатка стане вашою…':'Напишіть повідомлення…'}"></textarea><button id="save-note" class="composer-send"><span>Відправити</span>${icon('arrow-up')}</button></div>`;
  page.innerHTML=head('Нотатки')+`<div class="notes-page"><div class="notes-layout"><aside class="notes-list"><div class="notes-transfer-actions"><button id="new-note" class="notes-menu-action">${icon('plus')}<span>Нова нотатка</span></button><button id="import-notes" class="notes-menu-action">${icon('download')}<span>Імпорт нотаток</span></button><button id="export-notes" class="notes-menu-action">${icon('upload')}<span>Експорт нотаток</span></button></div><div class="notes-list-heading">Мої нотатки</div>${privateButtons||'<div class="notes-empty-small">Поки немає нотаток</div>'}<div class="notes-list-heading">Імпортовані нотатки</div>${importedButtons||'<div class="notes-empty-small">Поки немає імпортованих нотаток</div>'}</aside><section class="chat-note">${header}<div class="chat-thread">${timelineHtml||`<div class="chat-empty"><span>${imported?'Додайте повідомлення, щоб перенести нотатку до «Моїх нотаток»':'Напишіть текст або прикріпіть фото чи PDF'}</span></div>`}</div>${composer}</section></div></div>`;
  document.querySelectorAll('[data-note]').forEach(button=>button.onclick=()=>openNote((button.dataset.imported==='true'?importedNotes:notes).find(note=>note.id===button.dataset.note)));
  $('#export-notes').onclick=openExportDialog;
  $('#import-notes').onclick=async()=>{try{const result=await window.ekp.importNotes();if(result.canceled)return;importedNotes=await window.ekp.importedNotes();selectedNote=importedNotes[0]||notes[0]||null;renderNotes();renderSidebar();}catch(error){window.alert(`Не вдалося імпортувати нотатки: ${error?.message||error}`);}};
  $('#new-note').onclick=()=>{selectedNote={id:null,title:'',body:'',messages:[],attachments:[],pinned:false};renderNotes();};
  if (!imported) {
    document.querySelectorAll('[data-delete-file]').forEach(button=>button.onclick=async()=>{const index=Number(button.dataset.deleteFile);const file=selectedNote.attachments[index];if(!file||!await confirmDelete(`Вкладення «${file.name}» буде видалено зі сховища CoDA. Оригінальний файл залишиться без змін.`, 'Видалити вкладення?'))return;await window.ekp.deleteAttachment(file.path);selectedNote.attachments.splice(index,1);selectedNote=await window.ekp.saveNote(selectedNote);notes=await window.ekp.notes();renderNotes();renderSidebar();});
    document.querySelectorAll('[data-delete-message]').forEach(button=>button.onclick=async()=>{const index=Number(button.dataset.deleteMessage);if(!selectedNote.messages[index]||!await confirmDelete('Це повідомлення буде видалено без можливості відновлення.','Видалити повідомлення?'))return;selectedNote.messages.splice(index,1);selectedNote.body=selectedNote.messages.map(message=>message.text).join('\n\n');selectedNote=await window.ekp.saveNote(selectedNote);notes=await window.ekp.notes();renderNotes();renderSidebar();});
    $('#attach').onclick=async()=>{const id=selectedNote.id||crypto.randomUUID().replaceAll('-','');selectedNote.id=id;selectedNote.title=$('#note-title').value;const added=await window.ekp.addAttachments(id);if(!added.length)return;const text=$('#note-body').value.trim();let sequence=Date.now();if(text)selectedNote.messages.push({id:crypto.randomUUID(),text,created_at:new Date(sequence++).toISOString()});added.forEach(file=>file.created_at=new Date(sequence++).toISOString());selectedNote.body=selectedNote.messages.map(message=>message.text).join('\n\n');selectedNote.attachments.push(...added);selectedNote=await window.ekp.saveNote(selectedNote);notes=await window.ekp.notes();renderNotes();renderSidebar();};
    $('#save-note-title').onclick=async()=>{selectedNote.title=$('#note-title').value;selectedNote=await window.ekp.saveNote(selectedNote);notes=await window.ekp.notes();renderNotes();renderSidebar();};
    $('#pin-note').onclick=async()=>{selectedNote.pinned=!selectedNote.pinned;selectedNote.title=$('#note-title').value;selectedNote=await window.ekp.saveNote(selectedNote);notes=await window.ekp.notes();renderNotes();renderSidebar();};
    $('#save-note').onclick=async()=>{const text=$('#note-body').value.trim();selectedNote.title=$('#note-title').value;if(!text&&!selectedNote.id)return;if(text)selectedNote.messages.push({id:crypto.randomUUID(),text,created_at:new Date().toISOString()});selectedNote.body=selectedNote.messages.map(message=>message.text).join('\n\n');selectedNote=await window.ekp.saveNote(selectedNote);notes=await window.ekp.notes();await recent('recent_notes',selectedNote.id);renderNotes();};
    $('#note-body').onkeydown=event=>{if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();$('#save-note').click();}};
    $('#delete-note').onclick=async()=>{if(selectedNote.id&&await confirmDelete('Нотатку та всі її вкладення буде видалено без можливості відновлення.','Видалити нотатку?')){await window.ekp.deleteNote(selectedNote.id);notes=await window.ekp.notes();selectedNote=notes[0]||importedNotes[0]||null;renderNotes();renderSidebar();}};
  } else {
    $('#delete-imported-note').onclick=async()=>{if(!await confirmDelete(`Імпортовану нотатку «${selectedNote.title}» та її локальні вкладення буде видалено з CoDA. Початковий файл .codanotes залишиться без змін.`,'Видалити імпортовану нотатку?'))return;try{await window.ekp.deleteImportedNote(selectedNote.id);importedNotes=await window.ekp.importedNotes();selectedNote=importedNotes[0]||notes[0]||null;renderNotes();renderSidebar();}catch(error){showArchicadMessage('Не вдалося видалити нотатку',error?.message||String(error));}};
    $('#save-note').onclick=async()=>{const text=$('#note-body').value.trim();if(!text)return;selectedNote.messages.push({id:crypto.randomUUID(),text,created_at:new Date().toISOString()});selectedNote.body=selectedNote.messages.map(message=>message.text).join('\n\n');selectedNote=await window.ekp.promoteImportedNote(selectedNote);[notes,importedNotes]=await Promise.all([window.ekp.notes(),window.ekp.importedNotes()]);await recent('recent_notes',selectedNote.id);renderNotes();renderSidebar();};
  }
  $('#note-body').onkeydown=event=>{if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();$('#save-note').click();}};
  loadAttachmentPreviews(selectedNote); scrollChatToEnd();
}
async function openNote(note) { if(!note)return; selectedNote=note; if(!note.imported)await recent('recent_notes',note.id);currentPage='notes';rememberView({type:'note',id:note.id,imported:Boolean(note.imported)});renderNotes();renderSidebar(); }
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
async function renderExcel(){
  $('.surface').classList.remove('notes-mode');
  page.innerHTML=head('Archicad / Excel')+`<div class="page-content archicad-page"><div class="archicad-empty">Завантаження проєктів…</div></div>`;
  const [connections,projects]=await Promise.all([window.ekp.archicadConnections(),window.ekp.archicadProjects()]);
  if(currentPage!=='excel')return;
  let selected=projects.find(project=>project.id===selectedArchicadProjectId);
  if(selectedArchicadProjectId&&!selected)selectedArchicadProjectId=null;

  if(!selected){
    page.innerHTML=head('Archicad / Excel')+`<div class="page-content archicad-page"><div class="archicad-page-heading"><div><h1 class="hero">Проєкти Archicad</h1><p class="subtitle">Кожен проєкт — окрема папка зі своїми Excel-файлами та інструментами обміну</p></div><button id="create-archicad-project" class="button primary">${icon('plus')}<span>Створити проєкт</span></button></div>${archicadConnectionsHtml(connections)}<div class="archicad-project-browser">${projects.length?projects.map(project=>`<div class="archicad-folder-row"><button class="archicad-folder-card" data-ac-project="${project.id}"><span class="archicad-folder-icon">${icon('folder')}</span><span class="archicad-folder-copy"><strong>${escape(project.title)}</strong><small>${project.files.length} ${project.files.length===1?'Excel-файл':'Excel-файлів'} · створено ${escape(new Date(project.createdAt).toLocaleString('uk-UA'))}</small></span>${icon('chevron-right','folder-arrow')}</button><button class="archicad-folder-delete" data-ac-delete-project="${project.id}" title="Видалити проєкт" aria-label="Видалити проєкт ${escape(project.title)}">${icon('trash')}</button></div>`).join(''):`<div class="archicad-empty"><div>${icon('folder')}</div><h2>Проєктів ще немає</h2><p>Створіть папку проєкту, а потім сформуйте всередині потрібні Excel-файли.</p></div>`}</div></div>`;
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
  page.innerHTML=head(selected.title)+`<div class="page-content archicad-page"><button id="back-to-archicad-projects" class="archicad-back">${icon('chevron-left')}<span>Усі проєкти</span></button><div class="archicad-page-heading project-heading"><div><h1 class="hero">${escape(selected.title)}</h1><p class="subtitle">Excel-файли та інструменти цього проєкту</p></div><button id="delete-current-archicad-project" class="button archicad-danger">${icon('trash')}<span>Видалити проєкт</span></button></div><section class="archicad-project-connection ${activeConnection?'connected':selectedPort?'warning':''}"><div class="archicad-connection-copy"><strong>${icon('sheet')}Підключення до Archicad</strong><span>${escape(connectionMessage)}</span></div><div class="archicad-connection-controls"><select id="archicad-connection-select" ${connections.length?'':'disabled'}>${connectionOptions}</select><button id="detect-active-archicad" class="button primary" title="Визначити підключення за активним вікном Archicad">${icon('external-link')}<span>Обрати через вікно</span></button><button id="refresh-archicad-connections" class="button" title="Оновити список підключень">${icon('refresh-cw')}</button></div></section><div class="archicad-tools"><article class="archicad-tool-card"><header><div class="archicad-tool-icon">${icon('sheet')}</div><div><h2>Усі елементи</h2><p>Унікальний ID, ID елемента, слой і класифікація всіх елементів вибраного файлу Archicad.</p></div><button id="export-all-elements" class="button primary">${icon(elementsFile?'refresh-cw':'plus')}<span>${elementsFile?'Оновити Excel':'Створити Excel'}</span></button></header>${elementsFile?`<div class="archicad-file"><div class="archicad-file-icon">${icon('sheet')}</div><div class="archicad-file-copy"><strong>${escape(elementsFile.name)}</strong><span>${elementsFile.rowCount||0} елементів · оновлено ${escape(new Date(elementsFile.updatedAt).toLocaleString('uk-UA'))}</span></div><div class="archicad-file-actions"><button class="button" data-ac-open="${elementsFile.id}">${icon('external-link')}<span>Відкрити Excel</span></button><button class="button" data-ac-save="${elementsFile.id}">${icon('save')}<span>Зберегти</span></button><button class="button primary" data-ac-import="${elementsFile.id}">${icon('upload')}<span>Імпортувати в Archicad</span></button><button class="button archicad-file-delete" data-ac-delete-file="${elementsFile.id}" title="Видалити Excel">${icon('trash')}</button></div></div>`:'<div class="archicad-tool-empty">Excel ще не створено.</div>'}</article><article class="archicad-future-card"><div>${icon('plus')}</div><div><strong>Наступні Excel-інструменти</strong><span>Тут згодом можна додати відомості дверей, вікон, приміщень та інші специфікації.</span></div></article></div></div>`;
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
document.addEventListener('click',event=>{if(!$('#settings-menu').contains(event.target))closeSettings();});
document.querySelectorAll('.dialog-close').forEach(button=>button.onclick=()=>button.closest('dialog').close());
$('#check-database-update').onclick=checkDatabaseUpdate;
$('#database-update-cancel').onclick=()=>$('#database-update-dialog').close();
$('#database-update-install').onclick=installDatabaseUpdate;
$('#program-update-cancel').onclick=()=>$('#program-update-dialog').close();
$('#program-update-close-x').onclick=()=>$('#program-update-dialog').close();
$('#program-update-install').onclick=installProgramUpdate;
$('#export-close').onclick=()=>$('#export-dialog').close();
$('#export-success-ok').onclick=()=>$('#export-success-dialog').close();
$('#archicad-message-ok').onclick=()=>$('#archicad-message-dialog').close();
$('#export-select-all').onclick=()=>{const boxes=[...document.querySelectorAll('#export-notes-list input[type="checkbox"]')];const select=boxes.some(box=>!box.checked);boxes.forEach(box=>box.checked=select);$('#export-select-all').textContent=select?'Зняти вибір':'Вибрати всі';};
$('#export-create').onclick=async()=>{const button=$('#export-create');const status=$('#export-status');const noteIds=[...document.querySelectorAll('#export-notes-list input:checked')].map(input=>input.value);if(!noteIds.length){status.textContent='Оберіть хоча б одну нотатку.';return;}button.disabled=true;status.textContent='Підготовка файла…';try{const result=await window.ekp.exportNotes({noteIds,libraryName:$('#library-name').value,author:$('#library-author').value});if(result.canceled){status.textContent='Створення файла скасовано.';return;}$('#export-dialog').close();$('#export-success-message').textContent=`Експортовано ${result.count} нотаток у файл ${result.packagePath}`;$('#export-success-dialog').showModal();}catch(error){status.textContent=`Помилка: ${error?.message||error}`;}finally{button.disabled=false;}};
$('#minimize').onclick=()=>window.ekp.window.minimize();$('#maximize').onclick=()=>window.ekp.window.maximize();$('#close').onclick=()=>window.ekp.window.close();
(async()=>{[state,catalog,notes,importedNotes,catalogMetadata]=await Promise.all([window.ekp.state(),window.ekp.catalog(),window.ekp.notes(),window.ekp.importedNotes(),window.ekp.catalogMetadata()]);renderSidebar();renderDatabaseUpdate();navigate('dbn');setTimeout(()=>checkProgramUpdate(false),1200);})();
