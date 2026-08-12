const { execFile } = require('child_process');
const { promisify } = require('util');

const execFileAsync = promisify(execFile);
const ARCHICAD_PORTS = Array.from({ length: 21 }, (_, index) => 19723 + index);

async function post(port, command, parameters) {
  const response = await fetch(`http://127.0.0.1:${port}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(parameters ? { command, parameters } : { command }),
    signal: AbortSignal.timeout(8000),
  });
  if (!response.ok) throw new Error(`Archicad API повернув HTTP ${response.status}.`);
  const payload = await response.json();
  if (!payload.succeeded) throw new Error(payload.error?.message || `Команда ${command} не виконана.`);
  return payload.result || {};
}

async function getConnections() {
  const connections = await Promise.all(ARCHICAD_PORTS.map(async port => {
    try {
      const alive = await post(port, 'API.IsAlive');
      if (!alive.isAlive) return null;
      const product = await post(port, 'API.GetProductInfo');
      return { port, ...product, label: `Archicad ${product.version}, build ${product.buildNumber} — порт ${port}` };
    } catch { return null; }
  }));
  return connections.filter(Boolean).sort((a, b) => a.port - b.port);
}

function listeningArchicadPorts(output) {
  const byProcessId = new Map();
  for (const line of String(output || '').split(/\r?\n/)) {
    const match = line.match(/^\s*TCP\s+\S+:(\d+)\s+\S+\s+LISTENING\s+(\d+)\s*$/i);
    if (!match) continue;
    const port = Number(match[1]);
    if (ARCHICAD_PORTS.includes(port)) byProcessId.set(Number(match[2]), port);
  }
  return byProcessId;
}

async function getForegroundProcessId(delayMs) {
  const safeDelay = Math.max(1000, Math.min(Number(delayMs) || 5000, 10000));
  const script = `
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CodaForegroundWindow {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@
Start-Sleep -Milliseconds ${safeDelay}
[uint32]$codaForegroundPid = 0
[void][CodaForegroundWindow]::GetWindowThreadProcessId([CodaForegroundWindow]::GetForegroundWindow(), [ref]$codaForegroundPid)
Write-Output $codaForegroundPid
`;
  const { stdout } = await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script], {
    windowsHide: true,
    timeout: safeDelay + 6000,
  });
  const processId = Number(String(stdout).trim().split(/\s+/).pop());
  if (!processId) throw new Error('Не вдалося визначити активне вікно Windows.');
  return processId;
}

async function getForegroundConnection(delayMs = 5000) {
  if (process.platform !== 'win32') throw new Error('Вибір через активне вікно підтримується лише у Windows.');
  const processId = await getForegroundProcessId(delayMs);
  const { stdout } = await execFileAsync('netstat.exe', ['-ano', '-p', 'tcp'], { windowsHide: true, timeout: 5000 });
  const port = listeningArchicadPorts(stdout).get(processId);
  if (!port) throw new Error('Активним було не вікно Archicad. Повторіть і протягом відліку перейдіть саме у потрібний проєкт Archicad.');
  const connection = await connect(port);
  return { port, ...connection.product, label: `Archicad ${connection.product.version}, build ${connection.product.buildNumber} — порт ${port}` };
}

async function connect(preferredPort) {
  if (preferredPort !== undefined && preferredPort !== null && preferredPort !== '') {
    const port = Number(preferredPort);
    if (!ARCHICAD_PORTS.includes(port)) throw new Error('Некоректний порт підключення Archicad.');
    try {
      const alive = await post(port, 'API.IsAlive');
      if (alive.isAlive) return { port, product: await post(port, 'API.GetProductInfo') };
    } catch {}
    throw new Error(`Archicad на порті ${port} недоступний. Виберіть інше підключення.`);
  }
  const [first] = await getConnections();
  if (first) return { port: first.port, product: first };
  throw new Error('Archicad не знайдено. Відкрийте проєкт в Archicad і повторіть спробу.');
}

function propertyValue(entry) {
  const value = entry?.propertyValue;
  return value?.status === 'normal' ? String(value.value ?? '') : '';
}

function flattenClassifications(items, output = []) {
  for (const wrapper of items || []) {
    const item = wrapper.classificationItem;
    if (!item) continue;
    const guid = item.classificationItemId?.guid;
    const id = String(item.id || '').trim();
    const name = String(item.name || '').trim();
    if (guid) output.push({ guid, id, name, label: [id, name].filter(Boolean).join(' — ') });
    flattenClassifications(item.children, output);
  }
  return output;
}

async function getPropertyIds(port) {
  const names = ['General_ElementID', 'General_UniqueID', 'ModelView_LayerName'];
  const result = await post(port, 'API.GetPropertyIds', {
    properties: names.map(nonLocalizedName => ({ type: 'BuiltIn', nonLocalizedName })),
  });
  const ids = result.properties || [];
  if (ids.some(entry => !entry.propertyId?.guid)) throw new Error('Archicad не надав системні поля елементів.');
  return Object.fromEntries(names.map((name, index) => [name, ids[index].propertyId]));
}

async function getClassificationContext(port) {
  const idsResult = await post(port, 'API.GetClassificationSystemIds');
  const systemIds = idsResult.classificationSystemIds || [];
  if (!systemIds.length) return { systemIds: [], system: null, items: [] };
  const systemsResult = await post(port, 'API.GetClassificationSystems', { classificationSystemIds: systemIds });
  const systems = systemsResult.classificationSystems || [];
  const first = systems.find(entry => entry.classificationSystem)?.classificationSystem || null;
  if (!first) return { systemIds, system: null, items: [] };
  const tree = await post(port, 'API.GetAllClassificationsInSystem', { classificationSystemId: first.classificationSystemId });
  return { systemIds, system: first, items: flattenClassifications(tree.classificationItems) };
}

async function getSnapshot(preferredPort) {
  const { port, product } = await connect(preferredPort);
  const elementsResult = await post(port, 'API.GetAllElements');
  const elements = elementsResult.elements || [];
  const propertyIds = await getPropertyIds(port);
  const classification = await getClassificationContext(port);
  if (!elements.length) return { port, product, propertyIds, classification, rows: [] };

  const [typesResult, valuesResult, classesResult] = await Promise.all([
    post(port, 'API.GetTypesOfElements', { elements }),
    post(port, 'API.GetPropertyValuesOfElements', {
      elements,
      properties: Object.values(propertyIds).map(propertyId => ({ propertyId })),
    }),
    classification.systemIds.length
      ? post(port, 'API.GetClassificationsOfElements', { elements, classificationSystemIds: classification.systemIds })
      : Promise.resolve({ elementClassifications: [] }),
  ]);

  const itemByGuid = new Map(classification.items.map(item => [item.guid, item]));
  const types = typesResult.typesOfElements || [];
  const values = valuesResult.propertyValuesForElements || [];
  const classes = classesResult.elementClassifications || [];
  const rows = elements.map((element, index) => {
    const properties = values[index]?.propertyValues || [];
    const classIds = classes[index]?.classificationIds || [];
    const selected = classIds.find(entry => entry.classificationId?.classificationSystemId?.guid === classification.system?.classificationSystemId?.guid);
    const classGuid = selected?.classificationId?.classificationItemId?.guid || '';
    return {
      guid: element.elementId.guid,
      uniqueId: propertyValue(properties[1]) || element.elementId.guid,
      elementId: propertyValue(properties[0]),
      layer: propertyValue(properties[2]),
      classification: itemByGuid.get(classGuid)?.label || '',
      classificationGuid: classGuid,
      elementType: types[index]?.typeOfElement?.elementType || '',
    };
  });
  return { port, product, propertyIds, classification, rows };
}

function resolveClassification(value, items) {
  const normalized = String(value || '').trim().toLocaleLowerCase('uk');
  if (!normalized) return null;
  const matches = items.filter(item => [item.label, item.id, item.name].some(candidate => candidate.toLocaleLowerCase('uk') === normalized));
  if (matches.length === 1) return matches[0];
  if (!matches.length) throw new Error(`Класифікацію «${value}» не знайдено в поточному проєкті Archicad.`);
  throw new Error(`Назва класифікації «${value}» неоднозначна. Вкажіть її код, наприклад «${matches[0].id}».`);
}

function failedResults(results) {
  return (results || []).filter(entry => entry.error || entry.success === false);
}

async function importRows(sheetRows, preferredPort) {
  const snapshot = await getSnapshot(preferredPort);
  const liveByUniqueId = new Map(snapshot.rows.map(row => [row.uniqueId.toLocaleLowerCase(), row]));
  const seen = new Set();
  const propertyChanges = [];
  const classificationChanges = [];
  let unchanged = 0;
  let layerChangesSkipped = 0;

  for (const row of sheetRows) {
    const uniqueId = String(row.uniqueId || '').trim();
    if (!uniqueId) throw new Error('У таблиці є рядок без «Унікального ID».');
    const key = uniqueId.toLocaleLowerCase();
    if (seen.has(key)) throw new Error(`Унікальний ID ${uniqueId} повторюється в таблиці.`);
    seen.add(key);
    const live = liveByUniqueId.get(key);
    if (!live) throw new Error(`Елемент ${uniqueId} не знайдено у відкритому проєкті Archicad.`);

    let changed = false;
    const nextElementId = String(row.elementId ?? '').trim();
    if (nextElementId !== live.elementId) {
      propertyChanges.push({
        elementId: { guid: live.guid },
        propertyId: snapshot.propertyIds.General_ElementID,
        propertyValue: { type: 'string', status: 'normal', value: nextElementId },
      });
      changed = true;
    }

    const nextClass = resolveClassification(row.classification, snapshot.classification.items);
    const nextClassGuid = nextClass?.guid || '';
    if (snapshot.classification.system && nextClassGuid !== live.classificationGuid) {
      const classificationId = { classificationSystemId: snapshot.classification.system.classificationSystemId };
      if (nextClass) classificationId.classificationItemId = { guid: nextClass.guid };
      classificationChanges.push({ elementId: { guid: live.guid }, classificationId });
      changed = true;
    }

    if (String(row.layer ?? '').trim() !== live.layer) layerChangesSkipped += 1;
    if (!changed) unchanged += 1;
  }

  const propertyResult = propertyChanges.length
    ? await post(snapshot.port, 'API.SetPropertyValuesOfElements', { elementPropertyValues: propertyChanges })
    : { executionResults: [] };
  const propertyFailures = failedResults(propertyResult.executionResults);
  if (propertyFailures.length) throw new Error(`Не вдалося змінити ID у ${propertyFailures.length} елементів.`);

  const classificationResult = classificationChanges.length
    ? await post(snapshot.port, 'API.SetClassificationsOfElements', { elementClassifications: classificationChanges })
    : { executionResults: [] };
  const classificationFailures = failedResults(classificationResult.executionResults);
  if (classificationFailures.length) throw new Error(`Не вдалося змінити класифікацію у ${classificationFailures.length} елементів.`);

  return {
    elementIdsUpdated: propertyChanges.length,
    classificationsUpdated: classificationChanges.length,
    layerChangesSkipped,
    unchanged,
    total: sheetRows.length,
  };
}

async function getStatus(preferredPort) {
  try {
    const connection = await connect(preferredPort);
    return { connected: true, port: connection.port, ...connection.product };
  } catch (error) {
    return { connected: false, message: error.message };
  }
}

module.exports = { getConnections, getForegroundConnection, getSnapshot, getStatus, importRows };
