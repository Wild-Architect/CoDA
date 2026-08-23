const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');

const DATABASE_FILE = 'materials_database.xlsx';
const ATTACHMENTS_DIR = 'materials_attachments';
const LEGACY_USER_STORE_FILE = 'user_materials.json';
const USER_CODE_MARKER = 'К';
const CODE_DIGITS = 4;
const SOURCE_BUILTIN = 'builtin';
const SOURCE_USER = 'user';

function cleanText(value, maxLength) { return String(value || '').trim().slice(0, maxLength); }
function normalizedPrefix(value) {
  const prefix = cleanText(value, 8).toUpperCase().replace(/[^А-ЯІЇЄҐA-Z]/g, '');
  if (!/^[А-ЯІЇЄҐA-Z]{2,8}$/.test(prefix)) throw new Error('Категорія має некоректний префікс коду.');
  return prefix;
}
function sourceFromRow(value, code) {
  const normalized = cleanText(value, 40).toLocaleLowerCase('uk-UA');
  if (normalized === 'користувацький' || normalized === SOURCE_USER) return SOURCE_USER;
  if (normalized === 'базовий' || normalized === SOURCE_BUILTIN) return SOURCE_BUILTIN;
  return code.replace(/\d+$/, '').endsWith(USER_CODE_MARKER) ? SOURCE_USER : SOURCE_BUILTIN;
}
function sourceLabel(source) { return source === SOURCE_USER ? 'Користувацький' : 'Базовий'; }
function builtinDatabasePath(resourcesRoot) { return path.join(resourcesRoot, 'data', DATABASE_FILE); }
function workingDatabasePath(dataDir) { return path.join(dataDir, DATABASE_FILE); }
function builtinAttachmentsRoot(resourcesRoot) { return path.join(resourcesRoot, 'data', ATTACHMENTS_DIR); }
function attachmentsRoot(dataDir) { return path.join(dataDir, ATTACHMENTS_DIR); }
function legacyStorePath(dataDir) { return path.join(dataDir, LEGACY_USER_STORE_FILE); }
function rowsFromSheet(workbook, sheetName, range) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) throw new Error(`У базі матеріалів відсутній аркуш «${sheetName}».`);
  return XLSX.utils.sheet_to_json(sheet, { range, defval: '', raw: false });
}
function safeStoredName(value) {
  const storedName = cleanText(value, 160);
  if (!storedName || path.basename(storedName) !== storedName || storedName === '.' || storedName === '..') throw new Error('Вкладення матеріалу має некоректне ім’я у сховищі.');
  return storedName;
}
function normalizeAttachment(value) {
  if (!value || typeof value !== 'object') return null;
  try {
    const storedName = safeStoredName(value.storedName);
    return { id: cleanText(value.id, 80) || crypto.randomUUID(), name: cleanText(value.name, 260) || storedName, storedName, createdAt: cleanText(value.createdAt || value.created_at, 60) || new Date().toISOString() };
  } catch { return null; }
}
function parseAttachments(value) {
  if (!value) return [];
  try {
    const parsed = typeof value === 'string' ? JSON.parse(value) : value;
    return Array.isArray(parsed) ? parsed.map(normalizeAttachment).filter(Boolean) : [];
  } catch { return []; }
}
function serializeAttachments(value) {
  const attachments = Array.isArray(value) ? value.map(normalizeAttachment).filter(Boolean) : [];
  return attachments.length ? JSON.stringify(attachments) : '';
}
function materialAttachmentPath(dataDir, storedName) {
  const root = path.resolve(attachmentsRoot(dataDir));
  const target = path.resolve(root, safeStoredName(storedName));
  if (!target.startsWith(`${root}${path.sep}`)) throw new Error('Некоректний шлях до вкладення матеріалу.');
  return target;
}
function readDatabase(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: false, cellStyles: true });
  const categories = rowsFromSheet(workbook, 'Категорії', 3).map(row => ({ code: normalizedPrefix(row['Префікс']), name: cleanText(row['Назва категорії'], 160), description: cleanText(row['Опис'], 1000) })).filter(category => category.code && category.name);
  const categoryMap = new Map(categories.map(category => [category.code, category]));
  if (categoryMap.size !== categories.length) throw new Error('У базі повторюються префікси категорій.');
  const seenCodes = new Set();
  const materials = rowsFromSheet(workbook, 'Матеріали', 3).map(row => {
    const code = cleanText(row['Ідентифікаційний код'], 32).toUpperCase(); if (!code) return null;
    const categoryCode = normalizedPrefix(row['Код категорії']), category = categoryMap.get(categoryCode);
    if (!category) throw new Error(`Матеріал ${code} посилається на невідому категорію ${categoryCode}.`);
    const source = sourceFromRow(row['Тип запису'], code), marker = source === SOURCE_USER ? USER_CODE_MARKER : '';
    if (!new RegExp(`^${categoryCode}${marker}(?:\\d{${CODE_DIGITS}}|\\d{6})$`).test(code)) throw new Error(`Матеріал ${code} має некоректний код для типу «${sourceLabel(source)}».`);
    if (seenCodes.has(code)) throw new Error(`У базі повторюється код ${code}.`); seenCodes.add(code);
    return { id: `material:${code}`, code, categoryCode, category: category.name, name: cleanText(row['Назва'], 240), description: cleanText(row['Опис'], 10000), installationRules: cleanText(row['Правила монтажу'], 20000), source, attachments: parseAttachments(row['Вкладення (JSON)']) };
  }).filter(Boolean);
  const aboutRows = rowsFromSheet(workbook, 'Про базу', 2), about = Object.fromEntries(aboutRows.map(row => [String(row['Параметр'] || ''), String(row['Значення'] || '')]));
  return { workbook, categories, materials, version: about['Версія вмісту'] || '', filePath };
}
function clearDataRows(sheet, startRow = 5) {
  for (const key of Object.keys(sheet)) {
    if (key.startsWith('!')) continue;
    const cell = XLSX.utils.decode_cell(key);
    if (cell.r >= startRow - 1) delete sheet[key];
  }
}
function captureRowStyles(sheet, rowNumber, columnCount) {
  return Array.from({ length: columnCount }, (_value, column) => sheet[XLSX.utils.encode_cell({ r: rowNumber - 1, c: column })]?.s || null);
}
function applyRowStyles(sheet, startRow, rowCount, styles) {
  for (let row = startRow; row < startRow + rowCount; row++) for (let column = 0; column < styles.length; column++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column })];
    if (cell && styles[column]) cell.s = styles[column];
  }
}
async function writeDatabase(templatePath, targetPath, categories, materials) {
  const workbook = XLSX.readFile(templatePath, { cellDates: false, cellStyles: true });
  const categoriesSheet = workbook.Sheets['Категорії'], materialsSheet = workbook.Sheets['Матеріали'];
  if (!categoriesSheet || !materialsSheet) throw new Error('Excel-база має некоректну структуру.');
  const categoryStyles = captureRowStyles(categoriesSheet, 5, 3), materialStyles = captureRowStyles(materialsSheet, 5, 8);
  clearDataRows(categoriesSheet);
  XLSX.utils.sheet_add_aoa(categoriesSheet, categories.map(category => [category.code, category.name, category.description]), { origin: 'A5' });
  applyRowStyles(categoriesSheet, 5, Math.max(1, categories.length), categoryStyles);
  categoriesSheet['!ref'] = `A1:C${Math.max(5, 4 + categories.length)}`;
  categoriesSheet['!autofilter'] = { ref: `A4:C${Math.max(5, 4 + categories.length)}` };
  clearDataRows(materialsSheet);
  const rows = materials.map(material => [material.code, material.categoryCode, categories.find(category => category.code === material.categoryCode)?.name || material.category || '', material.name, material.description, material.installationRules, sourceLabel(material.source), serializeAttachments(material.attachments)]);
  XLSX.utils.sheet_add_aoa(materialsSheet, rows.length ? rows : [['', '', '', '', '', '', '', '']], { origin: 'A5' });
  applyRowStyles(materialsSheet, 5, Math.max(1, rows.length), materialStyles);
  materialsSheet['!ref'] = `A1:H${Math.max(5, 4 + rows.length)}`;
  materialsSheet['!autofilter'] = { ref: `A4:H${Math.max(5, 4 + rows.length)}` };
  await fs.mkdir(path.dirname(targetPath), { recursive: true });
  const tempPath = path.join(path.dirname(targetPath), `.${path.basename(targetPath)}.${crypto.randomUUID()}.tmp.xlsx`);
  try { XLSX.writeFile(workbook, tempPath, { bookType: 'xlsx', cellStyles: true, compression: true }); await fs.copyFile(tempPath, targetPath); }
  catch (error) { if (['EBUSY', 'EPERM', 'EACCES'].includes(error?.code)) throw new Error('Не вдалося оновити Excel-базу. Закрийте файл у Excel і повторіть дію.'); throw error; }
  finally { await fs.rm(tempPath, { force: true }); }
}
function comparable(material) { return [material.code, material.categoryCode, material.name, material.description, material.installationRules, material.source, serializeAttachments(material.attachments)].join('\u0000'); }
function normalizeMaterialCodes(materials) {
  const used = new Map();
  return [...materials].sort((a, b) => a.code.localeCompare(b.code, 'uk')).map(material => {
    const marker = material.source === SOURCE_USER ? USER_CODE_MARKER : '', namespace = `${material.categoryCode}:${marker}`;
    if (!used.has(namespace)) used.set(namespace, new Set());
    const occupied = used.get(namespace), match = new RegExp(`^${material.categoryCode}${marker}(\\d+)$`).exec(material.code);
    let sequence = Number(match?.[1] || 0);
    if (sequence < 1 || sequence > 9999 || occupied.has(sequence)) { sequence = 1; while (occupied.has(sequence) && sequence <= 9999) sequence++; }
    if (sequence > 9999) throw new Error(`Для категорії ${material.categoryCode} вичерпано діапазон чотиризначних кодів.`);
    occupied.add(sequence);
    const code = `${material.categoryCode}${marker}${String(sequence).padStart(CODE_DIGITS, '0')}`;
    return { ...material, id: `material:${code}`, code };
  });
}
async function readLegacyMaterials(dataDir) {
  try { const store = JSON.parse(await fs.readFile(legacyStorePath(dataDir), 'utf8')); return Array.isArray(store?.materials) ? store.materials : []; } catch { return []; }
}
async function ensureBuiltinAttachmentFiles(resourcesRoot, dataDir, materials) {
  const sourceRoot = path.resolve(builtinAttachmentsRoot(resourcesRoot)), targetRoot = path.resolve(attachmentsRoot(dataDir));
  await fs.mkdir(targetRoot, { recursive: true });
  if (sourceRoot === targetRoot) return;
  for (const attachment of materials.flatMap(material => material.attachments || [])) {
    const target = materialAttachmentPath(dataDir, attachment.storedName);
    if ((await fs.stat(target).catch(() => null))?.isFile()) continue;
    const source = path.resolve(sourceRoot, safeStoredName(attachment.storedName));
    if (!source.startsWith(`${sourceRoot}${path.sep}`) || !(await fs.stat(source).catch(() => null))?.isFile()) continue;
    await fs.copyFile(source, target);
  }
}
async function removeUnreferencedAttachments(dataDir, materials) {
  const root = attachmentsRoot(dataDir), referenced = new Set(materials.flatMap(material => material.attachments || []).map(item => item.storedName));
  const entries = await fs.readdir(root, { withFileTypes: true }).catch(() => []);
  await Promise.all(entries.filter(entry => entry.isFile() && entry.name !== '.gitkeep' && !referenced.has(entry.name)).map(entry => fs.rm(path.join(root, entry.name), { force: true })));
}
async function ensureWorkingDatabase(resourcesRoot, dataDir) {
  const builtinPath = builtinDatabasePath(resourcesRoot), activePath = workingDatabasePath(dataDir);
  if (!(await fs.stat(builtinPath).catch(() => null))?.isFile()) throw new Error('Вбудовану Excel-базу матеріалів не знайдено.');
  if (!(await fs.stat(activePath).catch(() => null))?.isFile()) { await fs.mkdir(path.dirname(activePath), { recursive: true }); await fs.copyFile(builtinPath, activePath); }
  const builtin = readDatabase(builtinPath), active = readDatabase(activePath), legacy = await readLegacyMaterials(dataDir);
  builtin.materials = normalizeMaterialCodes(builtin.materials); active.materials = normalizeMaterialCodes(active.materials);
  const categories = [...builtin.categories], categoryCodes = new Set(categories.map(category => category.code));
  for (const material of [...active.materials.filter(item => item.source === SOURCE_USER), ...legacy]) {
    const code = cleanText(material.categoryCode, 8).toUpperCase();
    if (!categoryCodes.has(code)) { const previous = active.categories.find(category => category.code === code); if (previous) { categories.push(previous); categoryCodes.add(code); } }
  }
  const userMaterials = [...active.materials.filter(item => item.source === SOURCE_USER)];
  for (const item of legacy) {
    const categoryCode = cleanText(item.categoryCode, 8).toUpperCase(), code = cleanText(item.code, 32).toUpperCase();
    if (!categoryCodes.has(categoryCode) || !code) continue;
    userMaterials.push({ id: `material:${code}`, code, categoryCode, category: categories.find(category => category.code === categoryCode)?.name || '', name: cleanText(item.name, 240), description: cleanText(item.description, 10000), installationRules: cleanText(item.installationRules, 20000), source: SOURCE_USER, attachments: [] });
  }
  const normalizedUsers = normalizeMaterialCodes(userMaterials), userByCode = new Map(normalizedUsers.map(item => [item.code, item]));
  const sameDatabaseVersion = Boolean(active.version) && active.version === builtin.version;
  const builtinMaterials = sameDatabaseVersion ? active.materials.filter(item => item.source === SOURCE_BUILTIN) : builtin.materials.filter(item => item.source === SOURCE_BUILTIN);
  const merged = [...builtinMaterials, ...userByCode.values()];
  const current = [...readDatabase(activePath).materials].sort((a, b) => a.code.localeCompare(b.code)).map(comparable).join('\n'), next = [...merged].sort((a, b) => a.code.localeCompare(b.code)).map(comparable).join('\n');
  const categoriesChanged = categories.map(item => `${item.code}:${item.name}:${item.description}`).join('|') !== active.categories.map(item => `${item.code}:${item.name}:${item.description}`).join('|');
  if (current !== next || categoriesChanged || legacy.length) await writeDatabase(builtinPath, activePath, categories, merged);
  await ensureBuiltinAttachmentFiles(resourcesRoot, dataDir, builtinMaterials);
  await removeUnreferencedAttachments(dataDir, merged);
  if (legacy.length) await fs.rm(legacyStorePath(dataDir), { force: true });
  return readDatabase(activePath);
}
function validateMaterialPayload(payload, categoryMap) {
  const categoryCode = normalizedPrefix(payload?.categoryCode), category = categoryMap.get(categoryCode);
  if (!category) throw new Error('Оберіть наявну категорію матеріалу.');
  const name = cleanText(payload?.name, 240), description = cleanText(payload?.description, 10000), installationRules = cleanText(payload?.installationRules, 20000);
  if (!name) throw new Error('Вкажіть назву матеріалу.');
  return { categoryCode, name, description, installationRules };
}
function nextCode(categoryCode, source, allCodes) {
  const marker = source === SOURCE_USER ? USER_CODE_MARKER : '', expression = new RegExp(`^${categoryCode}${marker}(\\d{${CODE_DIGITS}})$`);
  const last = Math.max(0, ...allCodes.map(code => expression.exec(code)?.[1]).filter(Boolean).map(Number));
  if (last >= 9999) throw new Error(`Для категорії ${categoryCode} вичерпано діапазон чотиризначних кодів.`);
  return `${categoryCode}${marker}${String(last + 1).padStart(CODE_DIGITS, '0')}`;
}
async function attachmentsWithAvailability(dataDir, attachments) {
  return Promise.all((attachments || []).map(async attachment => ({ ...attachment, exists: Boolean((await fs.stat(materialAttachmentPath(dataDir, attachment.storedName)).catch(() => null))?.isFile()) })));
}
async function listLibrary(resourcesRoot, dataDir, authorMode = false) {
  const active = await ensureWorkingDatabase(resourcesRoot, dataDir), sorted = [...active.materials].sort((a, b) => a.category.localeCompare(b.category, 'uk') || a.name.localeCompare(b.name, 'uk') || a.code.localeCompare(b.code, 'uk'));
  const materials = await Promise.all(sorted.map(async material => ({ ...material, attachments: await attachmentsWithAvailability(dataDir, material.attachments) })));
  return { schemaVersion: 3, databaseVersion: active.version, authorMode: Boolean(authorMode), databasePath: active.filePath, attachmentsPath: attachmentsRoot(dataDir), categories: active.categories.map(category => ({ ...category, count: materials.filter(item => item.categoryCode === category.code).length })), materials };
}
async function prepareAttachments(dataDir, requested, existingAttachments) {
  const existingById = new Map((existingAttachments || []).map(item => [item.id, item])), kept = [], createdPaths = [], seenExisting = new Set();
  await fs.mkdir(attachmentsRoot(dataDir), { recursive: true });
  try {
    for (const item of Array.isArray(requested) ? requested : []) {
      if (item?.sourcePath) {
        const sourcePath = path.resolve(String(item.sourcePath)), stat = await fs.stat(sourcePath).catch(() => null);
        if (!stat?.isFile()) throw new Error(`Не вдалося знайти файл «${cleanText(item.name, 260) || path.basename(sourcePath)}».`);
        const extension = path.extname(sourcePath).slice(0, 20), storedName = `${crypto.randomUUID()}${extension}`, target = materialAttachmentPath(dataDir, storedName);
        await fs.copyFile(sourcePath, target); createdPaths.push(target);
        kept.push({ id: crypto.randomUUID(), name: cleanText(item.name, 260) || path.basename(sourcePath), storedName, createdAt: new Date().toISOString() });
      } else {
        const existing = existingById.get(cleanText(item?.id, 80));
        if (existing && !seenExisting.has(existing.id)) { kept.push(existing); seenExisting.add(existing.id); }
      }
    }
    return { attachments: kept, createdPaths };
  } catch (error) { await Promise.all(createdPaths.map(file => fs.rm(file, { force: true }))); throw error; }
}
async function saveMaterial(resourcesRoot, dataDir, payload, authorMode = false) {
  const active = await ensureWorkingDatabase(resourcesRoot, dataDir), categoryMap = new Map(active.categories.map(category => [category.code, category])), values = validateMaterialPayload(payload, categoryMap);
  const existing = payload?.id ? active.materials.find(item => item.id === String(payload.id)) : null;
  if (payload?.id && !existing) throw new Error('Матеріал не знайдено.');
  const source = payload?.source === SOURCE_BUILTIN ? SOURCE_BUILTIN : SOURCE_USER;
  const changedNamespace = existing && (existing.categoryCode !== values.categoryCode || existing.source !== source), code = existing && !changedNamespace ? existing.code : nextCode(values.categoryCode, source, active.materials.map(item => item.code));
  const prepared = await prepareAttachments(dataDir, payload?.attachments, existing?.attachments || []);
  const material = { id: `material:${code}`, code, ...values, category: categoryMap.get(values.categoryCode).name, source, attachments: prepared.attachments };
  const materials = [material, ...active.materials.filter(item => item.id !== existing?.id && item.code !== material.code)];
  try { await writeDatabase(active.filePath, active.filePath, active.categories, materials); }
  catch (error) { await Promise.all(prepared.createdPaths.map(file => fs.rm(file, { force: true }))); throw error; }
  await removeUnreferencedAttachments(dataDir, materials);
  return { ...material, attachments: await attachmentsWithAvailability(dataDir, material.attachments) };
}
async function deleteMaterial(resourcesRoot, dataDir, id, authorMode = false) {
  const active = await ensureWorkingDatabase(resourcesRoot, dataDir), existing = active.materials.find(item => item.id === cleanText(id, 100));
  if (!existing) throw new Error('Матеріал не знайдено.');
  const materials = active.materials.filter(item => item.id !== existing.id);
  await writeDatabase(active.filePath, active.filePath, active.categories, materials);
  await removeUnreferencedAttachments(dataDir, materials);
  return true;
}

module.exports = { ATTACHMENTS_DIR, CODE_DIGITS, DATABASE_FILE, LEGACY_USER_STORE_FILE, attachmentsRoot, deleteMaterial, ensureWorkingDatabase, listLibrary, materialAttachmentPath, readDatabase, saveMaterial, workingDatabasePath };
