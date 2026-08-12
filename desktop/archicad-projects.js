const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');

const FILE_NAME = 'Елементи Archicad.xlsx';
const ELEMENTS_FILE_ID = 'all-elements';
const HEADERS = ['Унікальний ID', 'ID елемента', 'Слой', 'Класифікація'];

function projectsRoot(dataDir) { return path.join(dataDir, 'archicad_projects'); }
function projectDir(dataDir, id) { return path.join(projectsRoot(dataDir), safeId(id)); }
function safeId(value) {
  const id = String(value || '');
  if (!/^[a-zA-Z0-9_-]+$/.test(id)) throw new Error('Некоректний ідентифікатор проєкту.');
  return id;
}
async function readJson(file, fallback = null) {
  try { return JSON.parse(await fs.readFile(file, 'utf8')); } catch { return fallback; }
}
async function writeJson(file, value) { await fs.writeFile(file, JSON.stringify(value, null, 2), 'utf8'); }

function normalizeMetadata(metadata) {
  if (!metadata) return null;
  if (Array.isArray(metadata.files)) return metadata;
  const files = metadata.fileName ? [{ id: ELEMENTS_FILE_ID, kind: ELEMENTS_FILE_ID, name: metadata.fileName, rowCount: metadata.elementCount || 0, updatedAt: metadata.createdAt }] : [];
  return { ...metadata, files };
}

async function listProjects(dataDir) {
  const root = projectsRoot(dataDir);
  await fs.mkdir(root, { recursive: true });
  const folders = await fs.readdir(root, { withFileTypes: true });
  const projects = await Promise.all(folders.filter(entry => entry.isDirectory()).map(async entry => {
    const folder = path.join(root, entry.name);
    const metadata = normalizeMetadata(await readJson(path.join(folder, 'project.json')));
    if (!metadata) return null;
    const files = await Promise.all(metadata.files.map(async file => {
      const stat = await fs.stat(path.join(folder, path.basename(file.name))).catch(() => null);
      return { ...file, exists: Boolean(stat), updatedAt: stat?.mtime?.toISOString() || file.updatedAt || '' };
    }));
    return { ...metadata, files };
  }));
  return projects.filter(Boolean).sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
}

function createWorkbook(snapshot) {
  const workbook = XLSX.utils.book_new();
  const rows = snapshot.rows.map(row => ({
    'Унікальний ID': row.uniqueId,
    'ID елемента': row.elementId,
    'Слой': row.layer,
    'Класифікація': row.classification,
  }));
  const elementsSheet = XLSX.utils.json_to_sheet(rows, { header: HEADERS });
  elementsSheet['!cols'] = [{ wch: 39 }, { wch: 20 }, { wch: 34 }, { wch: 42 }];
  elementsSheet['!autofilter'] = { ref: elementsSheet['!ref'] || 'A1:D1' };
  XLSX.utils.book_append_sheet(workbook, elementsSheet, 'Елементи');

  const classificationRows = snapshot.classification.items.map(item => ({
    'Код': item.id,
    'Назва': item.name,
    'Значення для таблиці': item.label,
  }));
  const classificationsSheet = XLSX.utils.json_to_sheet(classificationRows, { header: ['Код', 'Назва', 'Значення для таблиці'] });
  classificationsSheet['!cols'] = [{ wch: 18 }, { wch: 42 }, { wch: 58 }];
  classificationsSheet['!autofilter'] = { ref: classificationsSheet['!ref'] || 'A1:C1' };
  XLSX.utils.book_append_sheet(workbook, classificationsSheet, 'Класифікації');

  const info = [
    ['Формат', 'CoDA Archicad Project 1'],
    ['Версія Archicad', String(snapshot.product.version || '')],
    ['Build Archicad', String(snapshot.product.buildNumber || '')],
    ['Система класифікації', snapshot.classification.system?.name || ''],
    ['Дата експорту', new Date().toISOString()],
    ['Примітка', 'Унікальний ID не змінюйте. ID елемента та класифікацію можна імпортувати назад. Слой у поточному API доступний лише для читання.'],
  ];
  const infoSheet = XLSX.utils.aoa_to_sheet(info);
  infoSheet['!cols'] = [{ wch: 28 }, { wch: 95 }];
  XLSX.utils.book_append_sheet(workbook, infoSheet, 'Про файл');
  return workbook;
}

async function createProject(dataDir, requestedTitle) {
  const id = crypto.randomUUID().replaceAll('-', '');
  const createdAt = new Date().toISOString();
  const formatter = new Intl.DateTimeFormat('uk-UA', { dateStyle: 'medium', timeStyle: 'short' });
  const title = String(requestedTitle || '').trim() || `Проєкт Archicad — ${formatter.format(new Date())}`;
  const folder = projectDir(dataDir, id);
  await fs.mkdir(folder, { recursive: true });
  const metadata = {
    id,
    title,
    createdAt,
    files: [],
  };
  await writeJson(path.join(folder, 'project.json'), metadata);
  return metadata;
}

async function getProject(dataDir, id) {
  const folder = projectDir(dataDir, id);
  const metadata = normalizeMetadata(await readJson(path.join(folder, 'project.json')));
  if (!metadata) throw new Error('Проєкт CoDA не знайдено.');
  return { metadata, folder };
}

async function setConnection(dataDir, id, connection) {
  const project = await getProject(dataDir, id);
  const metadata = {
    ...project.metadata,
    connection: connection ? {
      port: Number(connection.port),
      version: connection.version,
      buildNumber: connection.buildNumber,
      languageCode: connection.languageCode || '',
      label: connection.label || `Archicad ${connection.version} — порт ${connection.port}`,
    } : null,
  };
  await writeJson(path.join(project.folder, 'project.json'), metadata);
  return metadata;
}

async function deleteProject(dataDir, id) {
  const folder = projectDir(dataDir, id);
  const root = `${path.resolve(projectsRoot(dataDir))}${path.sep}`;
  if (!path.resolve(folder).startsWith(root)) throw new Error('Небезпечний шлях до проєкту.');
  await fs.rm(folder, { recursive: true, force: true });
  return true;
}

async function deleteProjectFile(dataDir, id, fileId) {
  const project = await getProject(dataDir, id);
  const safeFileId = safeId(fileId);
  const file = project.metadata.files.find(entry => entry.id === safeFileId);
  if (!file) throw new Error('Excel-файл у проєкті не знайдено.');
  const filePath = path.resolve(project.folder, path.basename(file.name));
  if (!filePath.startsWith(`${path.resolve(project.folder)}${path.sep}`)) throw new Error('Небезпечний шлях до Excel-файла.');
  await fs.rm(filePath, { force: true });
  const metadata = { ...project.metadata, files: project.metadata.files.filter(entry => entry.id !== safeFileId) };
  await writeJson(path.join(project.folder, 'project.json'), metadata);
  return true;
}

async function exportElements(dataDir, id, snapshot) {
  const project = await getProject(dataDir, id);
  const target = path.join(project.folder, FILE_NAME);
  XLSX.writeFile(createWorkbook(snapshot), target, { compression: true });
  const updatedAt = new Date().toISOString();
  const file = { id: ELEMENTS_FILE_ID, kind: ELEMENTS_FILE_ID, name: FILE_NAME, rowCount: snapshot.rows.length, updatedAt };
  const metadata = {
    ...project.metadata,
    archicadVersion: snapshot.product.version,
    archicadBuild: snapshot.product.buildNumber,
    classificationSystem: snapshot.classification.system?.name || '',
    files: [file, ...project.metadata.files.filter(entry => entry.id !== ELEMENTS_FILE_ID)],
  };
  await writeJson(path.join(project.folder, 'project.json'), metadata);
  return file;
}

async function getProjectFile(dataDir, id, fileId) {
  const project = await getProject(dataDir, id);
  const file = project.metadata.files.find(entry => entry.id === safeId(fileId));
  if (!file) throw new Error('Excel-файл у проєкті не знайдено.');
  const filePath = path.join(project.folder, path.basename(file.name));
  if (!(await fs.stat(filePath).catch(() => null))?.isFile()) throw new Error('Excel-файл проєкту відсутній. Створіть його повторно.');
  return { ...project, file, filePath };
}

async function readElementRows(dataDir, id, fileId = ELEMENTS_FILE_ID) {
  const project = await getProjectFile(dataDir, id, fileId);
  const workbook = XLSX.readFile(project.filePath, { cellDates: false });
  const sheet = workbook.Sheets['Елементи'] || workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error('У файлі немає аркуша «Елементи».');
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  return rows.map(row => ({
    uniqueId: String(row['Унікальний ID'] || '').trim(),
    elementId: String(row['ID елемента'] ?? '').trim(),
    layer: String(row['Слой'] ?? '').trim(),
    classification: String(row['Класифікація'] ?? '').trim(),
  }));
}

module.exports = { createProject, deleteProject, deleteProjectFile, exportElements, getProject, getProjectFile, listProjects, readElementRows, setConnection };
