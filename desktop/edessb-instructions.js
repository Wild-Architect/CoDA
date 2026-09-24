const fs = require('fs/promises');
const path = require('path');
const XLSX = require('xlsx');

const PDF_FOLDER = 'edessb instructions';
const CATALOG_FILE = path.join('data', 'edessb_instructions_catalog.xlsx');

function safeFilename(value) {
  const filename = String(value || '').trim();
  if (!filename || filename !== path.basename(filename) || path.extname(filename).toLowerCase() !== '.pdf') {
    throw new Error('Некоректна назва PDF-інструкції ЄДЕССБ.');
  }
  return filename;
}

function pdfPath(resourcesRoot, filename) {
  return path.join(resourcesRoot, PDF_FOLDER, safeFilename(filename));
}

function list(resourcesRoot) {
  const workbook = XLSX.readFile(path.join(resourcesRoot, CATALOG_FILE), { cellDates: false });
  const sheet = workbook.Sheets['Каталог інструкцій'] || workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false })
    .filter(row => row['Файл'])
    .map((row, index) => {
      const filename = safeFilename(row['Файл']);
      return {
        id: filename,
        filename,
        title: String(row['Найменування'] || '').trim() || `Інструкція ${index + 1}`,
        category: String(row['Категорія'] || '').trim() || 'Без категорії',
        version: String(row['Версія інструкції'] || '').trim(),
        pages: Math.max(0, Number(row['Сторінок']) || 0),
        order: Math.max(0, Number(row['Порядок']) || index + 1),
      };
    })
    .sort((a, b) => a.order - b.order || a.title.localeCompare(b.title, 'uk'));
}

async function readPdf(resourcesRoot, filename) {
  const target = pdfPath(resourcesRoot, filename);
  const stat = await fs.stat(target).catch(() => null);
  if (!stat?.isFile()) throw new Error('PDF-інструкцію ЄДЕССБ не знайдено.');
  return fs.readFile(target);
}

module.exports = { CATALOG_FILE, PDF_FOLDER, list, pdfPath, readPdf, safeFilename };
