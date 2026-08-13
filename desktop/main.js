const { app, BrowserWindow, dialog, ipcMain, shell, session } = require('electron');
const fs = require('fs/promises');
const fsSync = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFile } = require('child_process');
const { promisify } = require('util');
const XLSX = require('xlsx');
const { autoUpdater } = require('electron-updater');
const archicad = require('./archicad-service');
const archicadProjects = require('./archicad-projects');

const execFileAsync = promisify(execFile);
const MANIFEST_URL = 'https://raw.githubusercontent.com/Wild-Architect/CoDA-Database/main/manifest.json';
const BUILTIN_DATABASE_VERSION = '2026.08.09.2';
app.setAppUserModelId('ua.coda.desktop');
autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = true;
autoUpdater.allowPrerelease = false;

const root = path.resolve(__dirname, '..');
const resourcesRoot = app.isPackaged ? process.resourcesPath : root;
let databaseRoot = resourcesRoot;
let catalogDataDir = path.join(databaseRoot, 'data');
let writableDataDir = path.join(root, 'data');
let notesDir = path.join(writableDataDir, 'notes');
let attachmentsDir = path.join(writableDataDir, 'attachments');
let stateFile = path.join(writableDataDir, 'ui_state.json');
let updateStateFile = path.join(writableDataDir, 'database_update_state.json');
let importedNotesDir = path.join(writableDataDir, 'imported_notes');
let pdfBookmarksFile = path.join(writableDataDir, 'pdf_bookmarks.json');

function normalizeRelative(value) {
  const normalized = String(value || '').replaceAll('\\', '/').replace(/^\/+/, '');
  if (!normalized || normalized.split('/').some(part => !part || part === '.' || part === '..')) throw new Error('Маніфест містить небезпечний шлях до файла.');
  return normalized;
}
function resolveInside(base, relative) {
  const target = path.resolve(base, normalizeRelative(relative));
  if (!target.startsWith(`${path.resolve(base)}${path.sep}`)) throw new Error('Маніфест містить небезпечний шлях до файла.');
  return target;
}
function directDownloadCandidates(value) {
  const url = new URL(value);
  if (url.protocol !== 'https:') throw new Error('Посилання для оновлення повинно використовувати HTTPS.');
  const candidates = [url.toString()];
  if (url.hostname === '1drv.ms' || url.hostname.endsWith('.sharepoint.com') || url.hostname.includes('onedrive.live.com')) {
    const download = new URL(url); download.searchParams.set('download', '1'); candidates.unshift(download.toString());
  }
  return [...new Set(candidates)];
}
async function fetchDownload(value) {
  let lastError;
  for (const candidate of directDownloadCandidates(value)) {
    try {
      const response = await fetch(candidate, { redirect: 'follow', headers: { 'User-Agent': 'CoDA/1.0' } });
      if (!response.ok) throw new Error(`сервер повернув ${response.status}`);
      return response;
    } catch (error) { lastError = error; }
  }
  throw new Error(`Не вдалося завантажити файл з OneDrive: ${lastError?.message || 'невідома помилка'}`);
}
async function downloadOneDriveWithBrowser(value, destination) {
  const downloadSession = session.fromPartition('coda-onedrive-download');
  const win = new BrowserWindow({ show: false, webPreferences: { session: downloadSession, contextIsolation: true, nodeIntegration: false, sandbox: true } });
  win.webContents.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36');
  let timer;
  let completedCleanup;
  try {
    const completed = new Promise((resolve, reject) => {
      timer = setTimeout(() => reject(new Error('OneDrive не розпочав завантаження впродовж 90 секунд.')), 90000);
      const onDownload = (_event, item, contents) => {
        if (contents.id !== win.webContents.id) return;
        item.setSavePath(destination);
        item.once('done', (_doneEvent, status) => status === 'completed' ? resolve() : reject(new Error(`OneDrive завершив завантаження зі станом ${status}.`)));
      };
      downloadSession.on('will-download', onDownload);
      completedCleanup = () => downloadSession.removeListener('will-download', onDownload);
    });
    await win.loadURL(value);
    await win.webContents.executeJavaScript(`new Promise((resolve, reject) => {
      let attempts = 0;
      const timer = setInterval(() => {
        const button = document.querySelector('[data-automationid="download"]');
        if (button) { clearInterval(timer); button.click(); resolve(true); }
        else if (++attempts > 240) { clearInterval(timer); reject(new Error('Кнопку завантаження не знайдено.')); }
      }, 250);
    })`);
    await completed;
  } finally {
    clearTimeout(timer);
    if (typeof completedCleanup === 'function') completedCleanup();
    if (!win.isDestroyed()) win.destroy();
  }
}
async function readOneDriveTextWithBrowser(value) {
  const previewSession = session.fromPartition('coda-onedrive-preview');
  const win = new BrowserWindow({ show: false, webPreferences: { session: previewSession, contextIsolation: true, nodeIntegration: false, sandbox: true } });
  win.webContents.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36');
  try {
    await win.loadURL(value);
    return await win.webContents.executeJavaScript(`new Promise((resolve, reject) => {
      let attempts = 0;
      const timer = setInterval(() => {
        const model = window.monaco?.editor?.getModels?.()[0];
        const value = model?.getValue?.();
        if (value && value.trim().startsWith('{') && value.trim().endsWith('}')) { clearInterval(timer); resolve(value); }
        else if (++attempts > 240) { clearInterval(timer); reject(new Error('Не вдалося прочитати текст manifest.json у OneDrive.')); }
      }, 250);
    })`);
  } finally { if (!win.isDestroyed()) win.destroy(); }
}
async function downloadToFile(value, destination) {
  try {
    const response = await fetchDownload(value);
    await fs.writeFile(destination, Buffer.from(await response.arrayBuffer()));
  } catch (httpError) {
    const host = new URL(value).hostname;
    if (host !== '1drv.ms' && !host.endsWith('.sharepoint.com') && !host.includes('onedrive.live.com')) throw httpError;
    await downloadOneDriveWithBrowser(value, destination);
  }
}
async function fetchManifest() {
  let manifest;
  try {
    const text = await (await fetchDownload(MANIFEST_URL)).text();
    manifest = JSON.parse(text.replace(/^\uFEFF/, ''));
  } catch {
    try { manifest = JSON.parse((await readOneDriveTextWithBrowser(MANIFEST_URL)).replace(/^\uFEFF/, '')); }
    catch { throw new Error('Не вдалося прочитати опублікований manifest.json через OneDrive.'); }
  }
  if (manifest.schemaVersion !== 1 || !manifest.latestVersion || !manifest.fullPackage) throw new Error('Формат manifest.json не підтримується.');
  return manifest;
}
async function sha256File(file) {
  const hash = crypto.createHash('sha256');
  const handle = await fs.open(file, 'r');
  try { for await (const chunk of handle.createReadStream()) hash.update(chunk); }
  finally { await handle.close(); }
  return hash.digest('hex');
}
async function downloadPackage(info, folder) {
  if (!info?.url) throw new Error(`У manifest.json не заповнене посилання для пакета ${info?.file || ''}.`);
  const file = path.join(folder, path.basename(info.file || 'update.zip'));
  await downloadToFile(info.url, file);
  const size = (await fs.stat(file)).size;
  if (Number(info.size) && size !== Number(info.size)) throw new Error(`Розмір пакета ${info.file} не збігається з маніфестом.`);
  const hash = await sha256File(file);
  if (hash !== String(info.sha256 || '').toLowerCase()) throw new Error(`Контрольна сума пакета ${info.file} не збігається.`);
  return file;
}
async function expandZip(zipFile, destination) {
  await fs.mkdir(destination, { recursive: true });
  const command = `Expand-Archive -LiteralPath '${zipFile.replaceAll("'", "''")}' -DestinationPath '${destination.replaceAll("'", "''")}' -Force`;
  await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], { windowsHide: true, maxBuffer: 1024 * 1024 });
}
async function copyBundledDatabase(destination) {
  await fs.mkdir(destination, { recursive: true });
  await fs.copyFile(path.join(resourcesRoot, 'data', 'dbn_catalog.xlsx'), path.join(destination, 'dbn_catalog.xlsx'));
  await fs.cp(path.join(resourcesRoot, 'building_codes'), path.join(destination, 'building_codes'), { recursive: true });
}
async function verifyDatabase(folder, manifest) {
  for (const entry of manifest.files || []) {
    const file = resolveInside(folder, entry.path);
    const stat = await fs.stat(file).catch(() => null);
    if (!stat?.isFile() || stat.size !== Number(entry.size) || await sha256File(file) !== String(entry.sha256).toLowerCase()) throw new Error(`Не пройшла перевірка файла ${entry.path}.`);
  }
}
function updateChainFrom(manifest, startingVersion) {
  const chain = []; let version = startingVersion;
  while (version !== manifest.latestVersion) {
    const next = (manifest.updates || []).find(item => item.from === version);
    if (!next || chain.length > 100) return null;
    chain.push(next); version = next.to;
  }
  return chain;
}
function selectPackages(manifest, currentVersion) {
  if (currentVersion) {
    const incremental = updateChainFrom(manifest, currentVersion);
    if (incremental) return incremental;
  }
  const afterFull = updateChainFrom(manifest, manifest.fullPackage.version);
  if (!afterFull) throw new Error('Маніфест не містить повного ланцюжка оновлень від початкової до найновішої бази.');
  return [manifest.fullPackage, ...afterFull];
}
async function installDatabaseUpdate(manifest, sender) {
  const installedDatabase = path.join(writableDataDir, 'dbn_database');
  const hasInstalledDatabase = Boolean(await fs.stat(path.join(installedDatabase, 'dbn_catalog.xlsx')).catch(() => null));
  const state = await readJson(updateStateFile, { version: BUILTIN_DATABASE_VERSION });
  const currentVersion = hasInstalledDatabase ? (state.version || BUILTIN_DATABASE_VERSION) : BUILTIN_DATABASE_VERSION;
  const packages = selectPackages(manifest, currentVersion);
  if (!packages.length) return { updated: false, version: manifest.latestVersion, databaseDate: manifest.databaseDate };
  await fs.mkdir(writableDataDir, { recursive: true });
  const tempRoot = await fs.mkdtemp(path.join(writableDataDir, '.coda-update-'));
  const nextDatabase = path.join(tempRoot, 'database');
  try {
    sender.send('database:update-progress', { message: 'Підготовка локальної бази…' });
    const fullInstall = packages[0] === manifest.fullPackage;
    if (fullInstall) await fs.mkdir(nextDatabase, { recursive: true });
    else if (hasInstalledDatabase) await fs.cp(installedDatabase, nextDatabase, { recursive: true });
    else await copyBundledDatabase(nextDatabase);
    for (let index = 0; index < packages.length; index += 1) {
      const info = packages[index];
      sender.send('database:update-progress', { message: `Завантаження пакета ${index + 1} з ${packages.length}…` });
      const zip = await downloadPackage(info, tempRoot);
      const unpacked = path.join(tempRoot, `unpacked-${index}`);
      await expandZip(zip, unpacked);
      const instructions = await readJson(path.join(unpacked, 'update.json'), {});
      await fs.rm(path.join(unpacked, 'update.json'), { force: true });
      for (const relative of instructions.delete || []) await fs.rm(resolveInside(nextDatabase, relative), { force: true, recursive: true });
      await fs.cp(unpacked, nextDatabase, { recursive: true, force: true });
    }
    sender.send('database:update-progress', { message: 'Перевірка завантажених файлів…' });
    await verifyDatabase(nextDatabase, manifest);
    const backup = path.join(writableDataDir, 'dbn_database.backup');
    await fs.rm(backup, { recursive: true, force: true });
    if (await fs.stat(installedDatabase).catch(() => null)) await fs.rename(installedDatabase, backup);
    try { await fs.rename(nextDatabase, installedDatabase); }
    catch (error) { if (await fs.stat(backup).catch(() => null)) await fs.rename(backup, installedDatabase); throw error; }
    await fs.rm(backup, { recursive: true, force: true });
    await writeJson(updateStateFile, { version: manifest.latestVersion, databaseDate: manifest.databaseDate, installedAt: new Date().toISOString() });
    databaseRoot = installedDatabase; catalogDataDir = databaseRoot;
    return { updated: true, version: manifest.latestVersion, databaseDate: manifest.databaseDate };
  } finally { await fs.rm(tempRoot, { recursive: true, force: true }); }
}

async function listImportedNotes() {
  await ensureDataDirs();
  const files = await fs.readdir(importedNotesDir);
  const imported = await Promise.all(files.filter(file => file.endsWith('.json')).map(file => readJson(path.join(importedNotesDir, file), null)));
  return imported.filter(Boolean).map(note => ({ ...note, imported: true, pinned: false })).sort((a, b) => String(b.updated_at || '').localeCompare(String(a.updated_at || '')));
}

async function deleteImportedNote(id) {
  const safeId = safeNoteId(id);
  const noteFile = path.join(importedNotesDir, `${safeId}.json`);
  if (!(await fs.stat(noteFile).catch(() => null))?.isFile()) throw new Error('Імпортована нотатка більше не існує.');
  await Promise.all([
    fs.rm(noteFile, { force: true }),
    fs.rm(path.join(attachmentsDir, safeId), { recursive: true, force: true }),
  ]);
  return true;
}

async function exportNotesLibrary(noteIds, libraryName, author) {
  const chosenIds = [...new Set((noteIds || []).map(safeNoteId))];
  if (!chosenIds.length) throw new Error('Виберіть хоча б одну нотатку.');
  const available = new Map((await listNotes()).map(note => [note.id, note]));
  const selected = chosenIds.map(id => available.get(id));
  if (selected.some(note => !note)) throw new Error('Одна з вибраних нотаток більше не існує.');
  const destination = await dialog.showOpenDialog({ title: 'Виберіть папку для пакета бібліотеки', properties: ['openDirectory', 'createDirectory'] });
  if (destination.canceled || !destination.filePaths[0]) return { canceled: true };
  const tempRoot = await fs.mkdtemp(path.join(writableDataDir, '.coda-notes-export-'));
  const stage = path.join(tempRoot, 'library');
  try {
    await fs.mkdir(stage, { recursive: true });
    const exportedNotes = [];
    for (const note of selected) {
      const exported = { ...note, author: String(author || '').trim() || 'Не вказано', pinned: false, attachments: [] };
      for (const file of note.attachments || []) {
        const source = resolveAttachment(file.path);
        const relative = `attachments/${note.id}/${path.basename(source)}`;
        const target = resolveInside(stage, relative);
        await fs.mkdir(path.dirname(target), { recursive: true });
        await fs.copyFile(source, target);
        exported.attachments.push({ ...file, path: relative });
      }
      exportedNotes.push(exported);
    }
    const name = String(libraryName || '').trim() || 'Бібліотека нотаток';
    await writeJson(path.join(stage, 'library.json'), { schemaVersion: 1, name, author: String(author || '').trim(), exportedAt: new Date().toISOString(), notes: exportedNotes });
    const safeName = name.replace(/[<>:"/\\|?*]+/g, '-').replace(/\s+/g, ' ').slice(0, 80) || 'Нотатки-CoDA';
    const packageName = `${safeName}.codanotes`;
    const packagePath = path.join(destination.filePaths[0], packageName);
    const temporaryZip = path.join(tempRoot, 'notes-package.zip');
    await fs.rm(packagePath, { force: true });
    const command = `Compress-Archive -Path '${path.join(stage, '*').replaceAll("'", "''")}' -DestinationPath '${temporaryZip.replaceAll("'", "''")}' -CompressionLevel Optimal -Force`;
    await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], { windowsHide: true, maxBuffer: 1024 * 1024 });
    await fs.copyFile(temporaryZip, packagePath);
    return { canceled: false, packagePath, count: exportedNotes.length, name };
  } finally { await fs.rm(tempRoot, { recursive: true, force: true }); }
}

async function importNotesLibrary() {
  const chosen = await dialog.showOpenDialog({ title: 'Імпортувати нотатки CoDA', properties: ['openFile'], filters: [{ name: 'Файл нотаток CoDA', extensions: ['codanotes'] }] });
  if (chosen.canceled || !chosen.filePaths[0]) return { canceled: true };
  const tempRoot = await fs.mkdtemp(path.join(writableDataDir, '.coda-notes-import-'));
  const createdIds = [];
  try {
    const unpacked = path.join(tempRoot, 'library');
    const archive = path.join(tempRoot, 'notes-package.zip');
    await fs.copyFile(chosen.filePaths[0], archive);
    await expandZip(archive, unpacked);
    const library = await readJson(path.join(unpacked, 'library.json'), null);
    if (library?.schemaVersion !== 1 || !Array.isArray(library.notes)) throw new Error('Файл не містить коректної бібліотеки нотаток CoDA.');
    for (const sourceNote of library.notes) {
      const id = crypto.randomUUID().replaceAll('-', '');
      createdIds.push(id);
      const attachments = [];
      for (const file of sourceNote.attachments || []) {
        const source = resolveInside(unpacked, file.path);
        if (!(await fs.stat(source).catch(() => null))?.isFile()) throw new Error(`У пакеті відсутнє вкладення ${file.name || file.path}.`);
        const folder = path.join(attachmentsDir, id);
        await fs.mkdir(folder, { recursive: true });
        const target = path.join(folder, `${crypto.randomUUID().slice(0, 8)}_${path.basename(source)}`);
        await fs.copyFile(source, target);
        attachments.push({ ...file, path: path.relative(writableDataDir, target) });
      }
      const note = { ...sourceNote, id, source_library: library.name || 'Імпортована бібліотека', source_author: sourceNote.author || library.author || '', imported: true, pinned: false, attachments };
      await writeJson(path.join(importedNotesDir, `${id}.json`), note);
    }
    return { canceled: false, count: library.notes.length, name: library.name || 'Імпортована бібліотека' };
  } catch (error) {
    await Promise.all(createdIds.flatMap(id => [
      fs.rm(path.join(importedNotesDir, `${id}.json`), { force: true }),
      fs.rm(path.join(attachmentsDir, id), { recursive: true, force: true }),
    ]));
    throw error;
  } finally { await fs.rm(tempRoot, { recursive: true, force: true }); }
}

async function promoteImportedNote(note) {
  const id = safeNoteId(note.id);
  const imported = await readJson(path.join(importedNotesDir, `${id}.json`), null);
  if (!imported) throw new Error('Імпортована нотатка більше не існує.');
  const saved = await saveNote({ ...imported, ...note, imported: false, pinned: false });
  await fs.rm(path.join(importedNotesDir, `${id}.json`), { force: true });
  return saved;
}

async function ensureDataDirs() { await Promise.all([fs.mkdir(notesDir, { recursive: true }), fs.mkdir(attachmentsDir, { recursive: true }), fs.mkdir(importedNotesDir, { recursive: true })]); }
async function readJson(file, fallback) { try { return JSON.parse(await fs.readFile(file, 'utf8')); } catch { return fallback; } }
async function writeJson(file, value) { await fs.writeFile(file, JSON.stringify(value, null, 2), 'utf8'); }
const USER_DATA_ARCHIVE_KIND = 'CoDA User Data';
const USER_DATA_ENTRIES = ['ui_state.json', 'pdf_bookmarks.json', 'notes', 'imported_notes', 'attachments', 'archicad_projects'];
async function copyUserDataEntry(source, target) {
  const stat = await fs.lstat(source).catch(() => null);
  if (!stat) return false;
  if (stat.isSymbolicLink()) throw new Error(`Символічні посилання не підтримуються: ${source}`);
  if (stat.isDirectory()) {
    await fs.mkdir(target, { recursive: true });
    for (const entry of await fs.readdir(source)) await copyUserDataEntry(path.join(source, entry), path.join(target, entry));
    return true;
  }
  if (!stat.isFile()) return false;
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.copyFile(source, target);
  return true;
}
async function archiveFileList(folder, relative = '') {
  const result = [];
  for (const entry of await fs.readdir(folder, { withFileTypes: true }).catch(() => [])) {
    if (entry.isSymbolicLink()) throw new Error(`Символічні посилання не підтримуються: ${entry.name}`);
    const itemRelative = path.posix.join(relative.replaceAll('\\', '/'), entry.name);
    const itemPath = path.join(folder, entry.name);
    if (entry.isDirectory()) result.push(...await archiveFileList(itemPath, itemRelative));
    else if (entry.isFile()) {
      const stat = await fs.stat(itemPath);
      result.push({ path: itemRelative, size: stat.size, sha256: await sha256File(itemPath) });
    }
  }
  return result;
}
function archiveEntryAllowed(relative) {
  const normalized = String(relative || '').replaceAll('\\', '/');
  return normalized && !normalized.startsWith('/') && !normalized.split('/').includes('..') && USER_DATA_ENTRIES.includes(normalized.split('/')[0]);
}
async function exportUserDataArchive() {
  const date = new Date().toISOString().slice(0, 10);
  const selected = await dialog.showSaveDialog({ title: 'Зберегти дані користувача CoDA', defaultPath: `CoDA-${date}.codasaves`, filters: [{ name: 'Резервна копія CoDA', extensions: ['codasaves'] }] });
  if (selected.canceled || !selected.filePath) return { canceled: true };
  const packagePath = selected.filePath.toLowerCase().endsWith('.codasaves') ? selected.filePath : `${selected.filePath}.codasaves`;
  const tempRoot = await fs.mkdtemp(path.join(writableDataDir, '.coda-saves-export-'));
  try {
    const stage = path.join(tempRoot, 'archive'), dataRoot = path.join(stage, 'data');
    await fs.mkdir(dataRoot, { recursive: true });
    const entries = [];
    for (const entry of USER_DATA_ENTRIES) if (await copyUserDataEntry(path.join(writableDataDir, entry), path.join(dataRoot, entry))) entries.push(entry);
    const files = await archiveFileList(dataRoot);
    await writeJson(path.join(stage, 'manifest.json'), { schemaVersion: 1, kind: USER_DATA_ARCHIVE_KIND, appVersion: app.getVersion(), exportedAt: new Date().toISOString(), entries, files });
    const temporaryZip = path.join(tempRoot, 'user-data.zip');
    const command = `Compress-Archive -Path '${path.join(stage, '*').replaceAll("'", "''")}' -DestinationPath '${temporaryZip.replaceAll("'", "''")}' -CompressionLevel Optimal -Force`;
    await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], { windowsHide: true, maxBuffer: 1024 * 1024 });
    await fs.copyFile(temporaryZip, packagePath);
    return { canceled: false, packagePath, fileCount: files.length, exportedAt: new Date().toISOString() };
  } finally { await fs.rm(tempRoot, { recursive: true, force: true }); }
}
async function validateUserDataArchive(folder) {
  const manifest = await readJson(path.join(folder, 'manifest.json'), null);
  if (manifest?.schemaVersion !== 1 || manifest.kind !== USER_DATA_ARCHIVE_KIND || !Array.isArray(manifest.files)) throw new Error('Файл не є коректною резервною копією CoDA.');
  if (manifest.files.length > 100000) throw new Error('Резервна копія містить забагато файлів.');
  const dataRoot = path.join(folder, 'data'), actual = await archiveFileList(dataRoot);
  const declaredPaths = new Set(), declaredByPath = new Map(); let totalSize = 0;
  for (const file of manifest.files) {
    const normalizedPath = String(file?.path || '').replaceAll('\\', '/');
    const normalizedKey = normalizedPath.toLowerCase();
    if (!archiveEntryAllowed(normalizedPath) || declaredPaths.has(normalizedKey)) throw new Error('Резервна копія містить некоректні шляхи.');
    if (!Number.isSafeInteger(file.size) || file.size < 0 || !/^[a-f0-9]{64}$/i.test(String(file.sha256 || ''))) throw new Error(`Некоректний опис файла у резервній копії: ${normalizedPath}`);
    declaredPaths.add(normalizedKey); declaredByPath.set(normalizedKey, file); totalSize += file.size;
    if (totalSize > 8 * 1024 * 1024 * 1024) throw new Error('Обсяг резервної копії перевищує 8 ГБ.');
  }
  if (actual.length !== declaredPaths.size) throw new Error('Склад резервної копії не відповідає її опису.');
  for (const file of actual) {
    const declared = declaredByPath.get(file.path.toLowerCase());
    if (!declared || file.size !== declared.size || file.sha256 !== String(declared.sha256).toLowerCase()) throw new Error(`Пошкоджений файл у резервній копії: ${file.path}`);
    if (file.path.toLowerCase().endsWith('.json')) { try { JSON.parse(await fs.readFile(resolveInside(dataRoot, file.path), 'utf8')); } catch { throw new Error(`Некоректний JSON у резервній копії: ${file.path}`); } }
  }
  return { manifest, dataRoot };
}
async function importUserDataArchive() {
  const selected = await dialog.showOpenDialog({ title: 'Імпортувати дані користувача CoDA', properties: ['openFile'], filters: [{ name: 'Резервна копія CoDA', extensions: ['codasaves'] }] });
  if (selected.canceled || !selected.filePaths[0]) return { canceled: true };
  const tempRoot = await fs.mkdtemp(path.join(writableDataDir, '.coda-saves-import-'));
  try {
    const archive = path.join(tempRoot, 'user-data.zip'), unpacked = path.join(tempRoot, 'unpacked'), rollback = path.join(tempRoot, 'rollback');
    await fs.copyFile(selected.filePaths[0], archive); await expandZip(archive, unpacked);
    const { manifest, dataRoot } = await validateUserDataArchive(unpacked);
    await fs.mkdir(rollback, { recursive: true });
    for (const entry of USER_DATA_ENTRIES) await copyUserDataEntry(path.join(writableDataDir, entry), path.join(rollback, entry));
    try {
      for (const entry of USER_DATA_ENTRIES) await fs.rm(path.join(writableDataDir, entry), { recursive: true, force: true });
      for (const entry of USER_DATA_ENTRIES) await copyUserDataEntry(path.join(dataRoot, entry), path.join(writableDataDir, entry));
      await ensureDataDirs();
    } catch (error) {
      for (const entry of USER_DATA_ENTRIES) await fs.rm(path.join(writableDataDir, entry), { recursive: true, force: true });
      for (const entry of USER_DATA_ENTRIES) await copyUserDataEntry(path.join(rollback, entry), path.join(writableDataDir, entry));
      await ensureDataDirs(); throw error;
    }
    return { canceled: false, importedAt: new Date().toISOString(), exportedAt: manifest.exportedAt || '', fileCount: manifest.files.length };
  } finally { await fs.rm(tempRoot, { recursive: true, force: true }); }
}
function safePdfDocumentId(value) {
  const id = String(value || '').trim();
  if (!id || id.length > 1000 || /[\u0000-\u001f]/.test(id)) throw new Error('Некоректний ідентифікатор PDF-документа');
  return id;
}
async function readPdfBookmarkStore() {
  const store = await readJson(pdfBookmarksFile, { version: 1, documents: {} });
  if (!store || typeof store !== 'object' || !store.documents || typeof store.documents !== 'object') return { version: 1, documents: {} };
  return store;
}
async function listPdfBookmarks(documentId) {
  const store = await readPdfBookmarkStore();
  const list = store.documents[safePdfDocumentId(documentId)];
  return Array.isArray(list) ? list : [];
}
async function addPdfBookmark(payload) {
  await ensureDataDirs();
  const documentId = safePdfDocumentId(payload?.documentId);
  const page = Number(payload?.page);
  const start = Number(payload?.start);
  const end = Number(payload?.end);
  const text = String(payload?.text || '').trim().slice(0, 4000);
  const label = String(payload?.label || '').trim().slice(0, 120) || text.slice(0, 60) || 'Закладка';
  const colors = new Set(['yellow', 'green', 'blue', 'pink', 'orange']);
  if (!Number.isInteger(page) || page < 1 || !Number.isInteger(start) || start < 0 || !Number.isInteger(end) || end <= start || !text) throw new Error('Некоректні дані текстової закладки');
  const geometry = Array.isArray(payload?.geometry) ? payload.geometry.slice(0, 1000).map(rect => {
    const x = Number(rect?.x), y = Number(rect?.y), width = Number(rect?.width), height = Number(rect?.height);
    if (![x, y, width, height].every(Number.isFinite) || x < 0 || y < 0 || width <= 0 || height <= 0 || x >= 1 || y >= 1) return null;
    return { x, y, width: Math.min(width, 1 - x), height: Math.min(height, 1 - y) };
  }).filter(Boolean) : [];
  const bookmark = { id: crypto.randomUUID(), page, start, end, text, label, color: colors.has(payload?.color) ? payload.color : 'yellow', geometry, geometryVersion: geometry.length ? 2 : 0, createdAt: new Date().toISOString() };
  const store = await readPdfBookmarkStore();
  const list = Array.isArray(store.documents[documentId]) ? store.documents[documentId] : [];
  store.documents[documentId] = [bookmark, ...list];
  await writeJson(pdfBookmarksFile, store);
  return bookmark;
}
async function deletePdfBookmark(payload) {
  await ensureDataDirs();
  const documentId = safePdfDocumentId(payload?.documentId);
  const id = String(payload?.id || '');
  const store = await readPdfBookmarkStore();
  const list = Array.isArray(store.documents[documentId]) ? store.documents[documentId] : [];
  store.documents[documentId] = list.filter(item => item?.id !== id);
  await writeJson(pdfBookmarksFile, store);
  return true;
}
function safeNoteId(value) {
  const id = String(value || '');
  if (!/^[a-zA-Z0-9_-]+$/.test(id)) throw new Error('Некоректний ідентифікатор нотатки');
  return id;
}
function resolveAttachment(relativePath) {
  const target = path.resolve(writableDataDir, String(relativePath || ''));
  const rootPath = `${path.resolve(attachmentsDir)}${path.sep}`;
  if (!target.startsWith(rootPath)) throw new Error('Некоректний шлях до вкладення');
  return target;
}
function resolveReadableAttachment(relativePath) {
  const value = String(relativePath || '').replaceAll('\\', '/');
  return resolveAttachment(value);
}
function attachmentMime(filePath) {
  return ({
    '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.gif': 'image/gif',
    '.webp': 'image/webp', '.bmp': 'image/bmp', '.svg': 'image/svg+xml', '.pdf': 'application/pdf',
  })[path.extname(filePath).toLowerCase()] || 'application/octet-stream';
}
function readCatalogFromExcel() {
  const workbook = XLSX.readFile(path.join(catalogDataDir, 'dbn_catalog.xlsx'), { cellDates: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  return rows.filter(row => row['Файл ДБН']).map(row => {
    const requested = String(row['Файл ДБН']).trim();
    const pdfFallback = requested.replace(/\.doc$/i, '.pdf');
    const filename = requested !== pdfFallback
      && !fsSync.existsSync(path.join(databaseRoot, 'building_codes', requested))
      && fsSync.existsSync(path.join(databaseRoot, 'building_codes', pdfFallback)) ? pdfFallback : requested;
    return {
      filename,
      number: String(row['Номер ДБН'] || '').trim(),
      title: String(row['Найменування ДБН'] || '').trim(),
      category: String(row['Категорія'] || '').trim(),
      edessb_url: String(row['Посилання ЄДЕССБ'] || '').trim(),
      path: `building_codes/${filename}`,
    };
  });
}
function excelDateToIso(cell) {
  if (!cell) return '';
  if (cell.v instanceof Date && !Number.isNaN(cell.v.valueOf())) return cell.v.toISOString().slice(0, 10);
  if (cell.t === 'n') {
    const parsed = XLSX.SSF.parse_date_code(cell.v);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
  }
  return String(cell.v || '').trim();
}
function readCatalogMetadata() {
  const workbook = XLSX.readFile(path.join(catalogDataDir, 'dbn_catalog.xlsx'), { cellDates: true });
  return { updated_at: excelDateToIso(workbook.Sheets['Налаштування']?.B2) };
}
function readUpdateHistory() {
  const workbook = XLSX.readFile(path.join(resourcesRoot, 'data', 'update_history.xlsx'), { cellDates: false });
  const sheet = workbook.Sheets['Історія оновлень'] || workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false }).map(row => ({
    version: String(row['Версія'] || '').trim(),
    description: String(row['Опис оновлення'] || '').trim(),
  })).filter(item => item.version || item.description);
}
async function listNotes() {
  await ensureDataDirs();
  const files = await fs.readdir(notesDir);
  const notes = await Promise.all(files.filter(file => file.endsWith('.json')).map(file => readJson(path.join(notesDir, file), null)));
  return notes.filter(Boolean).sort((a, b) => Number(Boolean(b.pinned)) - Number(Boolean(a.pinned)) || b.updated_at.localeCompare(a.updated_at));
}
async function saveNote(note) {
  await ensureDataDirs();
  const now = new Date().toISOString();
  const id = note.id || crypto.randomUUID().replaceAll('-', '');
  const previous = await readJson(path.join(notesDir, `${id}.json`), {});
  const saved = { id, title: note.title?.trim() || 'Без назви', body: note.body || '', messages: Array.isArray(note.messages) ? note.messages : [], attachments: note.attachments || [], pinned: Boolean(note.pinned), created_at: previous.created_at || now, updated_at: now };
  await writeJson(path.join(notesDir, `${id}.json`), saved);
  return saved;
}

async function resolveArchicadProjectConnection(projectId) {
  const project = await archicadProjects.getProject(writableDataDir, projectId);
  const connections = await archicad.getConnections();
  const savedPort = project.metadata.connection?.port;
  if (savedPort) {
    const saved = connections.find(connection => connection.port === Number(savedPort));
    if (saved) return saved;
    throw new Error(`Збережене підключення Archicad на порті ${savedPort} недоступне. Перепідключіть проєкт.`);
  }
  if (connections.length === 1) {
    await archicadProjects.setConnection(writableDataDir, projectId, connections[0]);
    return connections[0];
  }
  if (!connections.length) throw new Error('Archicad не знайдено. Відкрийте потрібний проєкт Archicad.');
  throw new Error('Відкрито кілька проєктів Archicad. Виберіть підключення для цього проєкту CoDA.');
}

function releaseNotesText(notes) {
  if (typeof notes === 'string') return notes;
  if (!Array.isArray(notes)) return '';
  return notes.map(note => typeof note === 'string' ? note : note?.note).filter(Boolean).join('\n\n');
}

async function checkProgramUpdate() {
  if (!app.isPackaged) return { supported: false, updateAvailable: false, currentVersion: app.getVersion(), message: 'Перевірка оновлень програми працює у встановленій версії CoDA.' };
  const result = await autoUpdater.checkForUpdates();
  return {
    supported: true,
    updateAvailable: Boolean(result?.isUpdateAvailable),
    currentVersion: app.getVersion(),
    latestVersion: result?.updateInfo?.version || app.getVersion(),
    releaseName: result?.updateInfo?.releaseName || '',
    releaseNotes: releaseNotesText(result?.updateInfo?.releaseNotes),
  };
}

async function installProgramUpdate(sender) {
  if (!app.isPackaged) throw new Error('Оновлення можна встановити лише у встановленій версії CoDA.');
  const progress = value => sender.send('program-update:progress', {
    percent: Math.max(0, Math.min(100, Number(value.percent) || 0)),
    message: `Завантаження оновлення: ${Math.round(Number(value.percent) || 0)}%`,
  });
  autoUpdater.on('download-progress', progress);
  try {
    await autoUpdater.downloadUpdate();
    sender.send('program-update:progress', { percent: 100, message: 'Оновлення завантажено. CoDA перезапускається…' });
    setTimeout(() => autoUpdater.quitAndInstall(false, true), 700);
    return true;
  } finally {
    autoUpdater.removeListener('download-progress', progress);
  }
}

function createWindow() {
  const win = new BrowserWindow({ width: 1440, height: 900, minWidth: 1040, minHeight: 680, frame: false, icon: path.join(root, 'build', 'icon.png'), backgroundColor: '#eaf3f7', webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, zoomFactor: 0.8 } });
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

app.whenReady().then(async () => {
  if (app.isPackaged) {
    writableDataDir = app.getPath('userData');
    notesDir = path.join(writableDataDir, 'notes');
    attachmentsDir = path.join(writableDataDir, 'attachments');
    stateFile = path.join(writableDataDir, 'ui_state.json');
    updateStateFile = path.join(writableDataDir, 'database_update_state.json');
    importedNotesDir = path.join(writableDataDir, 'imported_notes');
    pdfBookmarksFile = path.join(writableDataDir, 'pdf_bookmarks.json');
  }
  const installedDatabase = path.join(writableDataDir, 'dbn_database');
  if (await fs.stat(path.join(installedDatabase, 'dbn_catalog.xlsx')).catch(() => null)) { databaseRoot = installedDatabase; catalogDataDir = databaseRoot; }
  ipcMain.handle('catalog', async () => {
    if (!catalogDataDir) return [];
    try { return readCatalogFromExcel(); }
    catch { return readJson(path.join(catalogDataDir, 'dbn_catalog.json'), []); }
  });
  ipcMain.handle('catalog:metadata', async () => {
    try { return readCatalogMetadata(); } catch { return { updated_at: '' }; }
  });
  ipcMain.handle('updates:history', async () => {
    try { return readUpdateHistory(); } catch { return []; }
  });
  ipcMain.handle('database:update-status', async () => {
    const local = await readJson(updateStateFile, { version: BUILTIN_DATABASE_VERSION });
    return { version: local.version || BUILTIN_DATABASE_VERSION, installed: true, manifestUrl: MANIFEST_URL };
  });
  ipcMain.handle('database:check-update', async () => {
    const manifest = await fetchManifest();
    const local = await readJson(updateStateFile, { version: BUILTIN_DATABASE_VERSION });
    return { updateAvailable: (local.version || BUILTIN_DATABASE_VERSION) !== manifest.latestVersion, currentVersion: local.version || BUILTIN_DATABASE_VERSION, latestVersion: manifest.latestVersion, databaseDate: manifest.databaseDate, manifest };
  });
  ipcMain.handle('database:install-update', async event => installDatabaseUpdate(await fetchManifest(), event.sender));
  ipcMain.handle('program-update:check', checkProgramUpdate);
  ipcMain.handle('program-update:install', event => installProgramUpdate(event.sender));
  ipcMain.handle('app:version', () => app.getVersion());
  ipcMain.handle('state:get', () => readJson(stateFile, { theme: 'light', pinned_dbn: [], recent_dbn: [], recent_notes: [] }));
  ipcMain.handle('state:set', (_event, state) => writeJson(stateFile, state));
  ipcMain.handle('user-data:export', exportUserDataArchive);
  ipcMain.handle('user-data:import', importUserDataArchive);
  ipcMain.handle('pdf-bookmarks:list', (_event, documentId) => listPdfBookmarks(documentId));
  ipcMain.handle('pdf-bookmarks:add', (_event, payload) => addPdfBookmark(payload));
  ipcMain.handle('pdf-bookmarks:delete', (_event, payload) => deletePdfBookmark(payload));
  ipcMain.handle('archicad:status', archicad.getStatus);
  ipcMain.handle('archicad:connections', archicad.getConnections);
  ipcMain.handle('archicad:foreground-connection', async event => {
    try { return await archicad.getForegroundConnection(); }
    finally {
      const win = BrowserWindow.fromWebContents(event.sender);
      if (win && !win.isDestroyed()) {
        if (win.isMinimized()) win.restore();
        win.setAlwaysOnTop(true);
        win.show();
        win.focus();
      }
    }
  });
  ipcMain.handle('archicad-projects:list', () => archicadProjects.listProjects(writableDataDir));
  ipcMain.handle('archicad-projects:create', (_event, title) => archicadProjects.createProject(writableDataDir, title));
  ipcMain.handle('archicad-projects:set-connection', async (_event, payload) => {
    const connections = await archicad.getConnections();
    const connection = connections.find(entry => entry.port === Number(payload?.port));
    if (!connection) throw new Error('Вибране підключення Archicad більше не доступне.');
    return archicadProjects.setConnection(writableDataDir, payload?.projectId, connection);
  });
  ipcMain.handle('archicad-projects:export-elements', async (_event, projectId) => {
    const connection = await resolveArchicadProjectConnection(projectId);
    return archicadProjects.exportElements(writableDataDir, projectId, await archicad.getSnapshot(connection.port));
  });
  ipcMain.handle('archicad-projects:open', async (_event, payload) => {
    const project = await archicadProjects.getProjectFile(writableDataDir, payload?.projectId, payload?.fileId);
    const error = await shell.openPath(project.filePath);
    if (error) throw new Error(error);
    return true;
  });
  ipcMain.handle('archicad-projects:save-copy', async (_event, payload) => {
    const project = await archicadProjects.getProjectFile(writableDataDir, payload?.projectId, payload?.fileId);
    const result = await dialog.showSaveDialog({
      title: 'Зберегти Excel-файл проєкту',
      defaultPath: project.file.name,
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    await fs.copyFile(project.filePath, result.filePath);
    return { canceled: false, filePath: result.filePath };
  });
  ipcMain.handle('archicad-projects:import', async (_event, payload) => {
    const connection = await resolveArchicadProjectConnection(payload?.projectId);
    return archicad.importRows(await archicadProjects.readElementRows(writableDataDir, payload?.projectId, payload?.fileId), connection.port);
  });
  ipcMain.handle('archicad-projects:delete', (_event, projectId) => archicadProjects.deleteProject(writableDataDir, projectId));
  ipcMain.handle('archicad-projects:delete-file', (_event, payload) => archicadProjects.deleteProjectFile(writableDataDir, payload?.projectId, payload?.fileId));
  ipcMain.handle('notes:list', listNotes);
  ipcMain.handle('imported-notes:list', listImportedNotes);
  ipcMain.handle('imported-notes:delete', (_event, id) => deleteImportedNote(id));
  ipcMain.handle('notes:export', (_event, payload) => exportNotesLibrary(payload?.noteIds, payload?.libraryName, payload?.author));
  ipcMain.handle('notes:import', importNotesLibrary);
  ipcMain.handle('imported-notes:promote', (_event, note) => promoteImportedNote(note));
  ipcMain.handle('notes:save', (_event, note) => saveNote(note));
  ipcMain.handle('notes:delete', async (_event, id) => {
    const safeId = safeNoteId(id);
    await Promise.all([
      fs.rm(path.join(notesDir, `${safeId}.json`), { force: true }),
      fs.rm(path.join(attachmentsDir, safeId), { recursive: true, force: true }),
    ]);
  });
  ipcMain.handle('attachments:add', async (_event, noteId) => {
    await ensureDataDirs();
    noteId = safeNoteId(noteId);
    const result = await dialog.showOpenDialog({ properties: ['openFile', 'multiSelections'] });
    if (result.canceled) return [];
    const folder = path.join(attachmentsDir, noteId);
    await fs.mkdir(folder, { recursive: true });
    return Promise.all(result.filePaths.map(async source => {
      const filename = path.basename(source);
      const target = path.join(folder, `${crypto.randomUUID().slice(0, 8)}_${filename}`);
      await fs.copyFile(source, target);
      return { id: crypto.randomUUID(), name: filename, path: path.relative(writableDataDir, target), created_at: new Date().toISOString() };
    }));
  });
  ipcMain.handle('attachments:delete', async (_event, relativePath) => {
    const target = resolveAttachment(relativePath);
    await fs.rm(target, { force: true });
    try {
      const parent = path.dirname(target);
      if ((await fs.readdir(parent)).length === 0) await fs.rmdir(parent);
    } catch {}
  });
  ipcMain.handle('attachments:read', async (_event, relativePath) => {
    const target = resolveReadableAttachment(relativePath);
    return { data: await fs.readFile(target), mime: attachmentMime(target) };
  });
  ipcMain.handle('file:open', (_event, relativePath) => shell.openPath(resolveReadableAttachment(relativePath)));
  ipcMain.handle('dbn:open', (_event, relativePath) => {
    if (!databaseRoot) throw new Error('База ДБН ще не встановлена.');
    return shell.openPath(path.join(databaseRoot, relativePath));
  });
  ipcMain.handle('dbn:read-pdf', async (_event, relativePath) => {
    if (!databaseRoot) throw new Error('База ДБН ще не встановлена.');
    const pdfPath = path.resolve(databaseRoot, relativePath);
    const codesRoot = `${path.resolve(databaseRoot, 'building_codes')}${path.sep}`;
    if (!pdfPath.startsWith(codesRoot) || path.extname(pdfPath).toLowerCase() !== '.pdf') throw new Error('Недопустимий шлях до PDF');
    return fs.readFile(pdfPath);
  });
  ipcMain.handle('external:open', (_event, value) => {
    try {
      const url = new URL(value);
      if (!['http:', 'https:'].includes(url.protocol)) return false;
      shell.openExternal(url.toString());
      return true;
    } catch { return false; }
  });
  ipcMain.handle('window:minimize', event => BrowserWindow.fromWebContents(event.sender).minimize());
  ipcMain.handle('window:maximize', event => { const win = BrowserWindow.fromWebContents(event.sender); win.isMaximized() ? win.unmaximize() : win.maximize(); });
  ipcMain.handle('window:release-always-on-top', event => { const win = BrowserWindow.fromWebContents(event.sender); if (win && !win.isDestroyed()) win.setAlwaysOnTop(false); });
  ipcMain.handle('window:close', event => BrowserWindow.fromWebContents(event.sender).close());
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
