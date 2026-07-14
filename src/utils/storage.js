import { buildActionKey, normalizeAction, mergeActionRecords, normalizePlatform, normalizeSku } from './actions.js';
import { normalizeDateKey } from './date.js';

const DB_NAME = 'wb-daily-ops-review';
const DB_VERSION = 5;
const RECORD_STORE = 'dailyRecords';
const ACTION_STORE = 'dailyActions';
const RECOMMENDATION_STORE = 'recommendationHistory';
const SKU_CATALOG_STORE = 'skuCatalog';

const ensureStore = (db, name, transaction) => {
  const store = db.objectStoreNames.contains(name)
    ? transaction.objectStore(name)
    : db.createObjectStore(name, { keyPath: 'uniqueKey' });
  if (!store.indexNames.contains('date')) store.createIndex('date', 'date', { unique: false });
  if (!store.indexNames.contains('sku')) store.createIndex('sku', 'sku', { unique: false });
  if (!store.indexNames.contains('platform')) store.createIndex('platform', 'platform', { unique: false });
};

const openDb = () => new Promise((resolve, reject) => {
  const request = indexedDB.open(DB_NAME, DB_VERSION);
  request.onupgradeneeded = () => {
    const db = request.result;
    const transaction = request.transaction;
    ensureStore(db, RECORD_STORE, transaction);
    ensureStore(db, ACTION_STORE, transaction);
    ensureStore(db, RECOMMENDATION_STORE, transaction);
    ensureStore(db, SKU_CATALOG_STORE, transaction);
  };
  request.onsuccess = () => resolve(request.result);
  request.onerror = () => reject(request.error);
});

const txDone = (transaction) => new Promise((resolve, reject) => {
  transaction.oncomplete = resolve;
  transaction.onerror = () => reject(transaction.error);
  transaction.onabort = () => reject(transaction.error);
});

const getAllFromStore = async (storeName) => {
  const db = await openDb();
  const transaction = db.transaction(storeName, 'readonly');
  const store = transaction.objectStore(storeName);
  const rows = await new Promise((resolve, reject) => {
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return rows.sort((a, b) => `${b.platform || 'WB'}${b.date}${b.sku}`.localeCompare(`${a.platform || 'WB'}${a.date}${a.sku}`));
};

const normalizeRecord = (record) => {
  const date = normalizeDateKey(record.date);
  const sku = normalizeSku(record.sku);
  const platform = normalizePlatform(record.platform);
  return { ...record, platform, date, sku, uniqueKey: buildActionKey(date, sku, platform) };
};

const getFromStore = (store, key) => new Promise((resolve, reject) => {
  const request = store.get(key);
  request.onsuccess = () => resolve(request.result);
  request.onerror = () => reject(request.error);
});

export const saveRecords = async (records) => {
  const db = await openDb();
  const transaction = db.transaction(RECORD_STORE, 'readwrite');
  const store = transaction.objectStore(RECORD_STORE);
  let added = 0;
  let overwritten = 0;
  await Promise.all(records.map(async (record) => {
    const normalized = normalizeRecord(record);
    const existing = await getFromStore(store, normalized.uniqueKey);
    if (existing) overwritten += 1;
    else added += 1;
    store.put({ ...normalized, importedAt: new Date().toISOString() });
  }));
  await txDone(transaction);
  db.close();
  return { added, overwritten, total: records.length };
};


const normalizeSkuCatalogItem = (item) => {
  const platform = normalizePlatform(item.platform);
  const sku = normalizeSku(item.sku);
  const parsedRowCount = Number(item.parsedRowCount) || 0;
  return {
    ...item,
    platform,
    sku,
    sheetName: String(item.sheetName || sku).trim(),
    importedAt: item.importedAt || new Date().toISOString(),
    hasParsedRows: Boolean(item.hasParsedRows || parsedRowCount > 0),
    parsedRowCount,
    uniqueKey: `${platform}_${sku}`,
  };
};

export const saveSkuCatalog = async (items = []) => {
  const db = await openDb();
  const transaction = db.transaction(SKU_CATALOG_STORE, 'readwrite');
  const store = transaction.objectStore(SKU_CATALOG_STORE);
  let added = 0;
  let overwritten = 0;
  await Promise.all((Array.isArray(items) ? items : []).map(async (item) => {
    const normalized = normalizeSkuCatalogItem(item);
    if (!normalized.sku) return;
    const existing = await getFromStore(store, normalized.uniqueKey);
    if (existing) overwritten += 1;
    else added += 1;
    store.put(normalized);
  }));
  await txDone(transaction);
  db.close();
  return { added, overwritten, total: items.length };
};

export const getAllSkuCatalog = () => getAllFromStore(SKU_CATALOG_STORE);

export const getAllRecords = () => getAllFromStore(RECORD_STORE);
export const getAllActions = async () => {
  const rows = (await getAllFromStore(ACTION_STORE)).map(normalizeAction);
  const deduped = new Map();
  rows.forEach((action) => deduped.set(action.uniqueKey, action));
  return [...deduped.values()].sort((a, b) => `${b.platform}${b.date}${b.sku}`.localeCompare(`${a.platform}${a.date}${a.sku}`));
};
export const getAllRecommendationHistory = () => getAllFromStore(RECOMMENDATION_STORE);

export const saveAction = async (action) => {
  const db = await openDb();
  const transaction = db.transaction(ACTION_STORE, 'readwrite');
  transaction.objectStore(ACTION_STORE).put(normalizeAction(action));
  await txDone(transaction);
  db.close();
};

export const deleteAction = async (uniqueKey) => {
  const db = await openDb();
  const transaction = db.transaction(ACTION_STORE, 'readwrite');
  const store = transaction.objectStore(ACTION_STORE);
  const separator = String(uniqueKey).includes('__') ? '__' : '_';
  const parts = String(uniqueKey || '').split(separator);
  const [platformMaybe = '', dateMaybe = '', ...skuParts] = parts;
  const hasPlatform = ['WB', 'Ozon'].includes(platformMaybe);
  const platform = hasPlatform ? platformMaybe : 'WB';
  const date = hasPlatform ? dateMaybe : platformMaybe;
  const sku = hasPlatform ? skuParts.join(separator) : [dateMaybe, ...skuParts].join(separator);
  const normalizedKey = buildActionKey(date, sku, platform);
  store.delete(uniqueKey);
  store.delete(normalizedKey);
  store.delete(`${normalizeDateKey(date)}__${normalizeSku(sku)}`);
  store.delete(`${normalizeDateKey(date)}_${normalizeSku(sku)}`);
  await txDone(transaction);
  db.close();
};

export const replaceAllRecords = async (records) => {
  const db = await openDb();
  const transaction = db.transaction(RECORD_STORE, 'readwrite');
  const store = transaction.objectStore(RECORD_STORE);
  store.clear();
  records.forEach((record) => store.put(record));
  await txDone(transaction);
  db.close();
};

export const replaceAllActions = async (actions) => {
  const db = await openDb();
  const transaction = db.transaction(ACTION_STORE, 'readwrite');
  const store = transaction.objectStore(ACTION_STORE);
  store.clear();
  actions.forEach((action) => store.put(action));
  await txDone(transaction);
  db.close();
};

export const saveExcelActions = async (actions = []) => {
  const db = await openDb();
  const transaction = db.transaction(ACTION_STORE, 'readwrite');
  const store = transaction.objectStore(ACTION_STORE);
  let added = 0;
  let overwritten = 0;
  let keptManual = 0;
  await Promise.all(actions.map(async (action) => {
    const incoming = normalizeAction({ ...action, source: action.source || 'excel_auto' });
    const existing = await getFromStore(store, incoming.uniqueKey);
    if (!existing) {
      store.put(incoming);
      added += 1;
    } else {
      const merged = mergeActionRecords([normalizeAction(existing)], [incoming]);
      const next = merged.actions.find((item) => item.uniqueKey === incoming.uniqueKey) || existing;
      if (merged.stats.overwritten > 0) {
        store.put(next);
        overwritten += 1;
      } else {
        keptManual += ['manual', 'manual_modified'].includes(existing.source) ? 1 : 0;
      }
    }
  }));
  await txDone(transaction);
  db.close();
  return { autoActionAdded: added, autoActionOverwritten: overwritten, keptManualActions: keptManual };
};

const normalizeRecommendation = (item) => {
  const date = normalizeDateKey(item.date);
  const sku = normalizeSku(item.sku);
  const type = item.recommendationType || item.type || item.id || 'unknown';
  return { ...item, date, sku, uniqueKey: item.uniqueKey || `${date}__${sku}__${type}` };
};

export const exportBackup = async () => ({
  exportedAt: new Date().toISOString(),
  version: 5,
  records: await getAllRecords(),
  skuCatalog: await getAllSkuCatalog(),
  actionRecords: await getAllActions(),
  actions: await getAllActions(),
  recommendationHistory: await getAllRecommendationHistory(),
  settings: {},
});

export const importBackup = async (backup) => {
  if (!backup || !Array.isArray(backup.records)) throw new Error('备份 JSON 格式不正确，缺少 records 数组。');
  const incomingRecords = backup.records.map(normalizeRecord);
  const incomingActions = (Array.isArray(backup.actionRecords) ? backup.actionRecords : backup.actions || [])
    .map((action) => normalizeAction({ ...action, source: action.source || 'json_import' }));
  const incomingRecommendations = (backup.recommendationHistory || []).map(normalizeRecommendation);
  const incomingSkuCatalog = (backup.skuCatalog || []).map(normalizeSkuCatalogItem);
  const db = await openDb();
  const transaction = db.transaction([RECORD_STORE, ACTION_STORE, RECOMMENDATION_STORE, SKU_CATALOG_STORE], 'readwrite');
  const recordStore = transaction.objectStore(RECORD_STORE);
  const actionStore = transaction.objectStore(ACTION_STORE);
  const recommendationStore = transaction.objectStore(RECOMMENDATION_STORE);
  const skuCatalogStore = transaction.objectStore(SKU_CATALOG_STORE);
  const stats = { recordsAdded: 0, recordsOverwritten: 0, actionsAdded: 0, actionsOverwritten: 0, actionsKeptLocal: 0, currentActionTotal: 0, recommendationsAdded: 0, recommendationsOverwritten: 0, skuCatalogAdded: 0, skuCatalogOverwritten: 0 };
  await Promise.all(incomingRecords.map(async (record) => {
    const existing = await getFromStore(recordStore, record.uniqueKey);
    if (existing) stats.recordsOverwritten += 1;
    else stats.recordsAdded += 1;
    recordStore.put(record);
  }));
  const localActions = await new Promise((resolve, reject) => {
    const request = actionStore.getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
  const { actions, stats: actionStats } = mergeActionRecords(localActions.map(normalizeAction), incomingActions);
  stats.actionsAdded = actionStats.added;
  stats.actionsOverwritten = actionStats.overwritten;
  stats.actionsKeptLocal = actionStats.keptLocal;
  stats.currentActionTotal = actions.length;
  actions.forEach((action) => actionStore.put(action));
  await Promise.all(incomingSkuCatalog.map(async (item) => {
    const existing = await getFromStore(skuCatalogStore, item.uniqueKey);
    if (existing) stats.skuCatalogOverwritten += 1;
    else stats.skuCatalogAdded += 1;
    skuCatalogStore.put(item);
  }));
  await Promise.all(incomingRecommendations.map(async (item) => {
    const existing = await getFromStore(recommendationStore, item.uniqueKey);
    if (existing) stats.recommendationsOverwritten += 1;
    else stats.recommendationsAdded += 1;
    recommendationStore.put(item);
  }));
  await txDone(transaction);
  db.close();
  return stats;
};
