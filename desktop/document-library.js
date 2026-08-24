const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');

const CATEGORIES = Object.freeze([
  { id: 'dstu', label: 'ДСТУ' },
  { id: 'other', label: 'Інші норми' },
]);
const CATEGORY_IDS = new Set(CATEGORIES.map(item => item.id));

function libraryRoot(writableDataDir) { return path.join(writableDataDir, 'document_library'); }
function filesRoot(writableDataDir) { return path.join(libraryRoot(writableDataDir), 'files'); }
function indexPath(writableDataDir) { return path.join(libraryRoot(writableDataDir), 'library.json'); }

async function readJson(filePath, fallback) {
  try { return JSON.parse(await fs.readFile(filePath, 'utf8')); }
  catch { return fallback; }
}

async function writeJson(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const temporary = `${filePath}.${crypto.randomUUID()}.tmp`;
  await fs.writeFile(temporary, JSON.stringify(value, null, 2), 'utf8');
  await fs.rename(temporary, filePath);
}

function safeId(value) {
  const id = String(value || '');
  if (!/^[a-f0-9]{32}$/i.test(id)) throw new Error('Некоректний ідентифікатор документа бібліотеки.');
  return id;
}

function safeStoredName(value) {
  const name = path.basename(String(value || ''));
  if (!name || name !== value || !/^[a-f0-9]{32}(\.[a-z0-9]{1,16})?$/i.test(name)) throw new Error('Некоректний шлях до файла бібліотеки.');
  return name;
}

async function readStore(writableDataDir) {
  const value = await readJson(indexPath(writableDataDir), { schemaVersion: 2, categories: [], items: [] });
  const categories = Array.isArray(value?.categories) ? value.categories.filter(category =>
    /^custom-[a-f0-9]{32}$/i.test(String(category?.id || '')) && String(category?.label || '').trim()
  ).map(category => ({ id: category.id, label: String(category.label).trim().replace(/\s+/g, ' ').slice(0, 120) })) : [];
  return { schemaVersion: 2, categories, items: Array.isArray(value?.items) ? value.items : [] };
}

async function getConfig(writableDataDir) {
  const store = await readStore(writableDataDir);
  return { categories: [...CATEGORIES, ...store.categories] };
}

async function addCategory(writableDataDir, value) {
  const label = String(value || '').trim().replace(/\s+/g, ' ').slice(0, 120);
  if (!label) throw new Error('Вкажіть назву категорії.');
  const store = await readStore(writableDataDir);
  const allCategories = [...CATEGORIES, ...store.categories];
  if (allCategories.some(category => category.label.localeCompare(label, 'uk', { sensitivity: 'accent' }) === 0)) {
    throw new Error('Категорія з такою назвою вже існує.');
  }
  const category = { id: `custom-${crypto.randomUUID().replaceAll('-', '')}`, label };
  store.categories.push(category);
  await writeJson(indexPath(writableDataDir), store);
  return category;
}

async function resolveItem(writableDataDir, itemId) {
  const id = safeId(itemId), store = await readStore(writableDataDir);
  const item = store.items.find(entry => entry?.id === id);
  if (!item) throw new Error('Документ бібліотеки не знайдено.');
  const filePath = path.join(filesRoot(writableDataDir), safeStoredName(item.storedName));
  if (!(await fs.stat(filePath).catch(() => null))?.isFile()) throw new Error('Файл документа відсутній у сховищі бібліотеки.');
  return { item, filePath, store };
}

async function listItems(writableDataDir) {
  const store = await readStore(writableDataDir);
  const root = filesRoot(writableDataDir);
  return Promise.all(store.items.map(async item => {
    let exists = false;
    try { exists = Boolean((await fs.stat(path.join(root, safeStoredName(item.storedName))).catch(() => null))?.isFile()); }
    catch {}
    return { ...item, exists };
  })).then(items => items.sort((left, right) => String(right.createdAt || '').localeCompare(String(left.createdAt || ''))));
}

async function addItem(writableDataDir, payload) {
  const category = String(payload?.category || '');
  const name = String(payload?.name || '').trim().replace(/\s+/g, ' ').slice(0, 240);
  const sourcePath = String(payload?.sourcePath || '');
  const store = await readStore(writableDataDir);
  if (!CATEGORY_IDS.has(category) && !store.categories.some(item => item.id === category)) throw new Error('Оберіть категорію документа.');
  if (!name) throw new Error('Вкажіть найменування документа.');
  const source = await fs.stat(sourcePath).catch(() => null);
  if (!source?.isFile()) throw new Error('Обраний файл більше не існує.');
  const id = crypto.randomUUID().replaceAll('-', '');
  const extension = path.extname(sourcePath).toLowerCase().replace(/[^.a-z0-9]/g, '').slice(0, 17);
  const storedName = `${id}${extension}`;
  const destination = path.join(filesRoot(writableDataDir), storedName);
  await fs.mkdir(path.dirname(destination), { recursive: true });
  await fs.copyFile(sourcePath, destination);
  try {
    const now = new Date().toISOString();
    const item = { id, category, name, originalName: path.basename(sourcePath), storedName, extension: extension.slice(1), isPdf: extension === '.pdf', size: source.size, createdAt: now, updatedAt: now };
    store.items.unshift(item);
    await writeJson(indexPath(writableDataDir), store);
    return { ...item, exists: true };
  } catch (error) {
    await fs.rm(destination, { force: true });
    throw error;
  }
}

async function deleteItem(writableDataDir, itemId) {
  const { item, filePath, store } = await resolveItem(writableDataDir, itemId);
  const temporary = `${filePath}.${crypto.randomUUID()}.delete`;
  await fs.rename(filePath, temporary);
  store.items = store.items.filter(entry => entry.id !== item.id);
  try {
    await writeJson(indexPath(writableDataDir), store);
    await fs.rm(temporary, { force: true });
  } catch (error) {
    await fs.rename(temporary, filePath).catch(() => {});
    throw error;
  }
  return true;
}

async function getItemFile(writableDataDir, itemId) {
  const { item, filePath } = await resolveItem(writableDataDir, itemId);
  return { item, filePath };
}

module.exports = { CATEGORIES, libraryRoot, filesRoot, getConfig, addCategory, listItems, addItem, deleteItem, getItemFile };
