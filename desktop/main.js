const { app, BrowserWindow, dialog, ipcMain, shell, session } = require('electron');
const fs = require('fs/promises');
const fsSync = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFile } = require('child_process');
const { promisify } = require('util');
const XLSX = require('xlsx');

const execFileAsync = promisify(execFile);
const MANIFEST_URL = 'https://raw.githubusercontent.com/Wild-Architect/CoDA-Database/main/manifest.json';
const NOTES_MANIFEST_URL = 'https://raw.githubusercontent.com/Wild-Architect/CoDA-Database/main/notes-manifest.json';
const BUILTIN_DATABASE_VERSION = '2026.08.09.2';
const ADMIN_LOGIN = 'AlyaAdmin';
const ADMIN_PASSWORD_HASH = 'ccc29f81888b79bcbcea80c3d4a1adb8f716a3da291b838ebd27b20c994922aa';
const ADMIN_PASSWORD_SALT = 'coda-admin-v1';

const root = path.resolve(__dirname, '..');
const resourcesRoot = app.isPackaged ? process.resourcesPath : root;
let databaseRoot = resourcesRoot;
let catalogDataDir = path.join(databaseRoot, 'data');
let writableDataDir = path.join(root, 'data');
let notesDir = path.join(writableDataDir, 'notes');
let attachmentsDir = path.join(writableDataDir, 'attachments');
let stateFile = path.join(writableDataDir, 'ui_state.json');
let updateStateFile = path.join(writableDataDir, 'database_update_state.json');
let publicNotesDir = path.join(writableDataDir, 'public_notes_library');
let publicNotesStateFile = path.join(writableDataDir, 'public_notes_state.json');
const adminSessions = new Set();

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
async function fetchPublicNotesManifest() {
  let manifest;
  try {
    const response = await fetchDownload(NOTES_MANIFEST_URL);
    manifest = JSON.parse((await response.text()).replace(/^\uFEFF/, ''));
  } catch (error) {
    throw new Error(`Не вдалося прочитати notes-manifest.json: ${error?.message || error}`);
  }
  if (manifest.schemaVersion !== 1 || !manifest.latestVersion || !manifest.package) throw new Error('Формат notes-manifest.json не підтримується.');
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

async function listPublicNotes() {
  const library = await readJson(path.join(publicNotesDir, 'library.json'), { notes: [] });
  return (Array.isArray(library.notes) ? library.notes : []).map(note => ({
    ...note,
    public: true,
    readonly: true,
    pinned: false,
    attachments: (note.attachments || []).map(file => ({ ...file, path: `public_notes_library/${file.path}` })),
  })).sort((a, b) => String(b.updated_at || '').localeCompare(String(a.updated_at || '')));
}
async function installPublicNotesLibrary(manifest, sender) {
  const tempRoot = await fs.mkdtemp(path.join(writableDataDir, '.coda-notes-update-'));
  try {
    sender.send('public-notes:update-progress', { message: 'Завантаження бібліотеки нотаток…' });
    const zip = await downloadPackage(manifest.package, tempRoot);
    const unpacked = path.join(tempRoot, 'library');
    await expandZip(zip, unpacked);
    const library = await readJson(path.join(unpacked, 'library.json'), null);
    if (library?.schemaVersion !== 1 || !Array.isArray(library.notes)) throw new Error('Пакет не містить коректної бібліотеки нотаток.');
    for (const note of library.notes) {
      safeNoteId(note.id);
      for (const file of note.attachments || []) {
        const target = resolveInside(unpacked, file.path);
        if (!(await fs.stat(target).catch(() => null))?.isFile()) throw new Error(`У пакеті відсутнє вкладення ${file.name || file.path}.`);
      }
    }
    sender.send('public-notes:update-progress', { message: 'Установлення бібліотеки…' });
    const backup = `${publicNotesDir}.backup`;
    await fs.rm(backup, { recursive: true, force: true });
    if (await fs.stat(publicNotesDir).catch(() => null)) await fs.rename(publicNotesDir, backup);
    try { await fs.rename(unpacked, publicNotesDir); }
    catch (error) { if (await fs.stat(backup).catch(() => null)) await fs.rename(backup, publicNotesDir); throw error; }
    await fs.rm(backup, { recursive: true, force: true });
    await writeJson(publicNotesStateFile, { version: manifest.latestVersion, installedAt: new Date().toISOString() });
    return { version: manifest.latestVersion, count: library.notes.length };
  } finally { await fs.rm(tempRoot, { recursive: true, force: true }); }
}

function verifyAdminPassword(login, password) {
  const supplied = crypto.scryptSync(String(password || ''), ADMIN_PASSWORD_SALT, 32);
  const expected = Buffer.from(ADMIN_PASSWORD_HASH, 'hex');
  return String(login || '') === ADMIN_LOGIN && supplied.length === expected.length && crypto.timingSafeEqual(supplied, expected);
}
async function exportPublicNotesLibrary(noteIds, version, author) {
  const normalizedVersion = String(version || '').trim();
  if (!/^[0-9]{4}\.[0-9]{2}\.[0-9]{2}(?:\.[0-9]+)?$/.test(normalizedVersion)) throw new Error('Версія повинна мати формат 2026.09.01 або 2026.09.01.1.');
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
      const exported = { ...note, author: String(author || '').trim() || 'Alya', public: true, readonly: true, pinned: false, attachments: [] };
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
    await writeJson(path.join(stage, 'library.json'), { schemaVersion: 1, version: normalizedVersion, exportedAt: new Date().toISOString(), notes: exportedNotes });
    const packageName = `CoDA-notes-${normalizedVersion}.zip`;
    const packagePath = path.join(destination.filePaths[0], packageName);
    await fs.rm(packagePath, { force: true });
    const command = `Compress-Archive -Path '${path.join(stage, '*').replaceAll("'", "''")}' -DestinationPath '${packagePath.replaceAll("'", "''")}' -CompressionLevel Optimal -Force`;
    await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', command], { windowsHide: true, maxBuffer: 1024 * 1024 });
    const info = await fs.stat(packagePath);
    const manifest = {
      schemaVersion: 1,
      latestVersion: normalizedVersion,
      updatedAt: new Date().toISOString(),
      package: {
        file: packageName,
        url: `https://github.com/Wild-Architect/CoDA-Database/releases/download/notes-${normalizedVersion}/${packageName}`,
        size: info.size,
        sha256: await sha256File(packagePath),
      },
    };
    const manifestPath = path.join(destination.filePaths[0], 'notes-manifest.json');
    await writeJson(manifestPath, manifest);
    return { canceled: false, packagePath, manifestPath, count: exportedNotes.length, version: normalizedVersion };
  } finally { await fs.rm(tempRoot, { recursive: true, force: true }); }
}

async function ensureDataDirs() { await Promise.all([fs.mkdir(notesDir, { recursive: true }), fs.mkdir(attachmentsDir, { recursive: true }), fs.mkdir(publicNotesDir, { recursive: true })]); }
async function readJson(file, fallback) { try { return JSON.parse(await fs.readFile(file, 'utf8')); } catch { return fallback; } }
async function writeJson(file, value) { await fs.writeFile(file, JSON.stringify(value, null, 2), 'utf8'); }
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
  if (value.startsWith('public_notes_library/')) return resolveInside(publicNotesDir, value.slice('public_notes_library/'.length));
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

function createWindow() {
  const win = new BrowserWindow({ width: 1440, height: 900, minWidth: 1040, minHeight: 680, frame: false, backgroundColor: '#eaf3f7', webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, zoomFactor: 0.8 } });
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

app.whenReady().then(async () => {
  if (app.isPackaged) {
    writableDataDir = app.getPath('userData');
    notesDir = path.join(writableDataDir, 'notes');
    attachmentsDir = path.join(writableDataDir, 'attachments');
    stateFile = path.join(writableDataDir, 'ui_state.json');
    updateStateFile = path.join(writableDataDir, 'database_update_state.json');
    publicNotesDir = path.join(writableDataDir, 'public_notes_library');
    publicNotesStateFile = path.join(writableDataDir, 'public_notes_state.json');
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
  ipcMain.handle('state:get', () => readJson(stateFile, { theme: 'light', pinned_dbn: [], recent_dbn: [], recent_notes: [] }));
  ipcMain.handle('state:set', (_event, state) => writeJson(stateFile, state));
  ipcMain.handle('notes:list', listNotes);
  ipcMain.handle('public-notes:list', listPublicNotes);
  ipcMain.handle('public-notes:update-status', async () => {
    const local = await readJson(publicNotesStateFile, { version: '' });
    return { version: local.version || '', manifestUrl: NOTES_MANIFEST_URL };
  });
  ipcMain.handle('public-notes:check-update', async () => {
    const manifest = await fetchPublicNotesManifest();
    const local = await readJson(publicNotesStateFile, { version: '' });
    return { updateAvailable: local.version !== manifest.latestVersion, currentVersion: local.version || '', latestVersion: manifest.latestVersion, manifest };
  });
  ipcMain.handle('public-notes:install-update', async event => installPublicNotesLibrary(await fetchPublicNotesManifest(), event.sender));
  ipcMain.handle('admin:login', (event, credentials) => {
    const authorized = verifyAdminPassword(credentials?.login, credentials?.password);
    if (authorized) adminSessions.add(event.sender.id);
    return authorized;
  });
  ipcMain.handle('admin:logout', event => { adminSessions.delete(event.sender.id); return true; });
  ipcMain.handle('admin:export-library', async (event, payload) => {
    if (!adminSessions.has(event.sender.id)) throw new Error('Потрібно повторно увійти в режим адміністратора.');
    return exportPublicNotesLibrary(payload?.noteIds, payload?.version, payload?.author);
  });
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
  ipcMain.handle('window:close', event => BrowserWindow.fromWebContents(event.sender).close());
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
