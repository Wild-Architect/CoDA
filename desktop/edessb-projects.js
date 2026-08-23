const fs = require('fs/promises');
const path = require('path');
const crypto = require('crypto');

const PORTAL_LINKS = Object.freeze([
  { id: 'project-documentation', label: 'Проєктна документація', url: 'https://admin.e-construction.gov.ua/project_documentation_employee_ARM' },
  { id: 'estimate-documentation', label: 'Кошторисна документація', url: 'https://admin.e-construction.gov.ua/estimate_construction_participant' },
  { id: 'design-task', label: 'Завдання на проєктування', url: 'https://admin.e-construction.gov.ua/edesb.design_tasks' },
  { id: 'document-access', label: 'Запит на надання доступу до документу', url: 'https://admin.e-construction.gov.ua/doc_access_participant' },
  { id: 'price-analysis', label: 'Звіт з аналізу цін', url: 'https://admin.e-construction.gov.ua/price_analysis_report' },
  { id: 'material-resources', label: 'Перелік матеріальних ресурсів та їх ціни', url: 'https://admin.e-construction.gov.ua/edesb_protocol_approval_price' },
]);

const DOCUMENT_TYPES = Object.freeze([
  { id: 'project-documentation', label: 'Проєктна документація', short: 'ПД' },
  { id: 'estimate-documentation', label: 'Кошторисна документація', short: 'КД' },
  { id: 'design-task', label: 'Завдання на проєктування', short: 'ЗП' },
  { id: 'price-analysis', label: 'Звіт з аналізу цін', short: 'ЗАЦ' },
  { id: 'material-resources', label: 'Перелік матеріальних ресурсів та їх ціни', short: 'ПМР' },
  { id: 'expert-report', label: 'Експертний звіт', short: 'ЕЗ' },
]);

const documentTypeMap = new Map(DOCUMENT_TYPES.map(item => [item.id, item]));
const projectsRoot = dataDir => path.join(dataDir, 'edessb_projects');
const documentsRoot = dataDir => path.join(dataDir, 'edessb_documents');
function safeId(value, label = 'проєкту') { const id = String(value || ''); if (!/^[a-zA-Z0-9_-]+$/.test(id)) throw new Error(`Некоректний ідентифікатор ${label}.`); return id; }
function projectDir(dataDir, id) { return path.join(projectsRoot(dataDir), safeId(id)); }
function cleanText(value, maxLength) { return String(value || '').trim().slice(0, maxLength); }
function safeIdentifier(value) { const identifier = cleanText(value, 40); if (!/^\d+$/.test(identifier)) throw new Error('Ідентифікаційний номер проєкту повинен складатися з цифр.'); return identifier.replace(/^0+(?=\d)/, ''); }
function safeDocumentType(value) { const type = documentTypeMap.get(String(value || '')); if (!type) throw new Error('Невідомий тип документа ЄДЕССБ.'); return type; }
function normalizeDocumentUrl(value) {
  const text = String(value || '').trim(); if (!text) return '';
  let url; try { url = new URL(text); } catch { throw new Error('Вкажіть правильне посилання на документ.'); }
  if (url.protocol !== 'https:' || (url.hostname !== 'e-construction.gov.ua' && !url.hostname.endsWith('.e-construction.gov.ua'))) throw new Error('Посилання повинно вести на захищену сторінку e-construction.gov.ua.');
  return url.toString();
}
async function readJson(file, fallback = null) { try { return JSON.parse(await fs.readFile(file, 'utf8')); } catch { return fallback; } }
async function writeJson(file, value) { await fs.mkdir(path.dirname(file), { recursive: true }); await fs.writeFile(file, JSON.stringify(value, null, 2), 'utf8'); }
function normalizeProject(metadata) {
  if (!metadata?.id) return null;
  const documents = [];
  for (const raw of Array.isArray(metadata.documents) ? metadata.documents : []) {
    const type = documentTypeMap.get(raw?.type); if (!type) continue;
    const rawRevisions = Array.isArray(raw.revisions) ? raw.revisions : [raw];
    const revisions = rawRevisions.map((revision, index) => ({ number: Number(revision.number) > 0 ? Number(revision.number) : index + 1, url: String(revision.url || ''), originalName: String(revision.originalName || ''), storedName: path.basename(String(revision.storedName || '')), createdAt: revision.createdAt || metadata.createdAt, updatedAt: revision.updatedAt || metadata.updatedAt })).sort((a, b) => a.number - b.number);
    documents.push({ type: type.id, label: type.label, short: type.short, revisions });
  }
  return { ...metadata, schemaVersion: 2, identifier: String(metadata.identifier || ''), documents };
}
async function getProject(dataDir, id) { const folder = projectDir(dataDir, id), metadata = normalizeProject(await readJson(path.join(folder, 'project.json'))); if (!metadata) throw new Error('Проєкт ЄДЕССБ не знайдено.'); return { folder, metadata }; }
function projectDocumentsDir(dataDir, identifier) { return path.join(documentsRoot(dataDir), safeIdentifier(identifier)); }
async function migrateLegacyFiles(dataDir, project) {
  if (!project.metadata.identifier) return;
  let changed = project.metadata.schemaVersion !== 2;
  for (const document of project.metadata.documents) for (const revision of document.revisions) {
    if (!revision.storedName) continue;
    const targetName = `${project.metadata.identifier}-${document.short}${revision.number}.pdf`, target = path.join(projectDocumentsDir(dataDir, project.metadata.identifier), targetName);
    if (revision.storedName === targetName && (await fs.stat(target).catch(() => null))?.isFile()) continue;
    const candidates = [path.join(project.folder, 'documents', path.basename(revision.storedName)), path.join(projectDocumentsDir(dataDir, project.metadata.identifier), path.basename(revision.storedName))];
    let source = null; for (const candidate of candidates) if ((await fs.stat(candidate).catch(() => null))?.isFile()) { source = candidate; break; }
    if (source) { await fs.mkdir(path.dirname(target), { recursive: true }); await fs.copyFile(source, target); revision.storedName = targetName; changed = true; }
  }
  if (changed) await writeJson(path.join(project.folder, 'project.json'), project.metadata);
}
async function listProjects(dataDir) {
  const root = projectsRoot(dataDir); await Promise.all([fs.mkdir(root, { recursive: true }), fs.mkdir(documentsRoot(dataDir), { recursive: true })]);
  const entries = await fs.readdir(root, { withFileTypes: true });
  const loaded = [];
  for (const entry of entries.filter(item => item.isDirectory())) { const folder = path.join(root, entry.name), metadata = normalizeProject(await readJson(path.join(folder, 'project.json'))); if (metadata) loaded.push({ folder, metadata }); }
  const used = new Set(loaded.map(project => project.metadata.identifier).filter(value => /^\d+$/.test(value))); let nextIdentifier = 1;
  for (const project of loaded.filter(item => !/^\d+$/.test(item.metadata.identifier))) { while (used.has(String(nextIdentifier))) nextIdentifier += 1; project.metadata.identifier = String(nextIdentifier); used.add(String(nextIdentifier)); await writeJson(path.join(project.folder, 'project.json'), project.metadata); }
  const projects = await Promise.all(loaded.map(async project => {
    await migrateLegacyFiles(dataDir, project);
    const documents = await Promise.all(project.metadata.documents.map(async document => ({ ...document, revisions: await Promise.all(document.revisions.map(async revision => ({ ...revision, fileExists: Boolean(revision.storedName && (await fs.stat(path.join(projectDocumentsDir(dataDir, project.metadata.identifier), revision.storedName)).catch(() => null))?.isFile()) }))) })));
    return { ...project.metadata, documents };
  }));
  return projects.filter(Boolean).sort((a, b) => Number(a.identifier || Number.MAX_SAFE_INTEGER) - Number(b.identifier || Number.MAX_SAFE_INTEGER) || String(a.identifier).localeCompare(String(b.identifier), 'uk', { numeric: true }) || a.title.localeCompare(b.title, 'uk'));
}
async function createProject(dataDir, payload) {
  const title = cleanText(payload?.title, 160), organization = cleanText(payload?.organization, 200), identifier = safeIdentifier(payload?.identifier);
  if (!title) throw new Error('Вкажіть назву проєкту.'); if (!organization) throw new Error('Вкажіть проєктну організацію.');
  if ((await listProjects(dataDir)).some(project => project.identifier === identifier)) throw new Error(`Проєкт з ідентифікаційним номером ${identifier} вже існує.`);
  const id = crypto.randomUUID().replaceAll('-', ''), now = new Date().toISOString(), metadata = { schemaVersion: 2, id, identifier, title, organization, createdAt: now, updatedAt: now, documents: [] };
  await Promise.all([writeJson(path.join(projectDir(dataDir, id), 'project.json'), metadata), fs.mkdir(projectDocumentsDir(dataDir, identifier), { recursive: true })]); return metadata;
}
async function validatePdf(sourcePath) {
  if (!sourcePath) return '';
  const source = path.resolve(String(sourcePath)); if (path.extname(source).toLowerCase() !== '.pdf' || !(await fs.stat(source).catch(() => null))?.isFile()) throw new Error('Оберіть наявний PDF-файл.');
  const handle = await fs.open(source, 'r'); try { const signature = Buffer.alloc(5); await handle.read(signature, 0, 5, 0); if (signature.toString('ascii') !== '%PDF-') throw new Error('Вибраний файл не є PDF-документом.'); } finally { await handle.close(); } return source;
}
async function addRevision(dataDir, projectId, payload) {
  const project = await getProject(dataDir, projectId), type = safeDocumentType(payload?.type), url = normalizeDocumentUrl(payload?.url), source = await validatePdf(payload?.sourcePath);
  if (!url && !source) throw new Error('Додайте посилання, PDF-файл або обидва.');
  const group = project.metadata.documents.find(item => item.type === type.id), number = Math.max(0, ...(group?.revisions || []).map(item => Number(item.number) || 0)) + 1;
  const storedName = source ? `${project.metadata.identifier}-${type.short}${number}.pdf` : '', target = storedName ? path.join(projectDocumentsDir(dataDir, project.metadata.identifier), storedName) : '';
  if (source) { await fs.mkdir(path.dirname(target), { recursive: true }); await fs.copyFile(source, target); }
  const now = new Date().toISOString(), revision = { number, url, originalName: source ? path.basename(source) : '', storedName, createdAt: now, updatedAt: now };
  const document = group ? { ...group, revisions: [...group.revisions, revision] } : { type: type.id, label: type.label, short: type.short, revisions: [revision] };
  const metadata = { ...project.metadata, updatedAt: now, documents: [...project.metadata.documents.filter(item => item.type !== type.id), document] };
  try { await writeJson(path.join(project.folder, 'project.json'), metadata); } catch (error) { if (target) await fs.rm(target, { force: true }); throw error; } return revision;
}
async function getRevision(dataDir, projectId, typeId, number) {
  const project = await getProject(dataDir, projectId), type = safeDocumentType(typeId), revision = project.metadata.documents.find(item => item.type === type.id)?.revisions.find(item => item.number === Number(number));
  if (!revision) throw new Error('Редакцію документа не знайдено.'); if (!revision.storedName) throw new Error('Для цієї редакції PDF-файл не додано.');
  const filePath = path.join(projectDocumentsDir(dataDir, project.metadata.identifier), path.basename(revision.storedName)); if (!(await fs.stat(filePath).catch(() => null))?.isFile()) throw new Error('PDF-файл редакції відсутній.'); return { revision, filePath };
}
async function deleteRevision(dataDir, projectId, typeId, number) {
  const project = await getProject(dataDir, projectId), type = safeDocumentType(typeId), revisionNumber = Number(number), group = project.metadata.documents.find(item => item.type === type.id), revision = group?.revisions.find(item => item.number === revisionNumber); if (!revision) return true;
  if (revision.storedName) await fs.rm(path.join(projectDocumentsDir(dataDir, project.metadata.identifier), path.basename(revision.storedName)), { force: true });
  const revisions = group.revisions.filter(item => item.number !== revisionNumber), documents = revisions.length ? [...project.metadata.documents.filter(item => item.type !== type.id), { ...group, revisions }] : project.metadata.documents.filter(item => item.type !== type.id);
  await writeJson(path.join(project.folder, 'project.json'), { ...project.metadata, updatedAt: new Date().toISOString(), documents }); return true;
}
async function deleteProject(dataDir, id) {
  const project = await getProject(dataDir, id), root = `${path.resolve(projectsRoot(dataDir))}${path.sep}`; if (!path.resolve(project.folder).startsWith(root)) throw new Error('Небезпечний шлях до проєкту ЄДЕССБ.');
  await Promise.all([fs.rm(project.folder, { recursive: true, force: true }), project.metadata.identifier ? fs.rm(projectDocumentsDir(dataDir, project.metadata.identifier), { recursive: true, force: true }) : Promise.resolve()]); return true;
}

module.exports = { DOCUMENT_TYPES, PORTAL_LINKS, addRevision, createProject, deleteProject, deleteRevision, documentsRoot, getRevision, listProjects };
