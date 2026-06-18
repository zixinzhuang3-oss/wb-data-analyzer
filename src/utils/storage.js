const DB_NAME = 'wb-daily-ops-review';
const DB_VERSION = 2;
const RECORD_STORE = 'dailyRecords';
const ACTION_STORE = 'dailyActions';

const ensureStore = (db, name) => {
  if (!db.objectStoreNames.contains(name)) {
    const store = db.createObjectStore(name, { keyPath: 'uniqueKey' });
    store.createIndex('date', 'date', { unique: false });
    store.createIndex('sku', 'sku', { unique: false });
  }
};

const openDb = () => new Promise((resolve, reject) => {
  const request = indexedDB.open(DB_NAME, DB_VERSION);
  request.onupgradeneeded = () => {
    const db = request.result;
    ensureStore(db, RECORD_STORE);
    ensureStore(db, ACTION_STORE);
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
  return rows.sort((a, b) => `${b.date}${b.sku}`.localeCompare(`${a.date}${a.sku}`));
};

export const saveRecords = async (records) => {
  const db = await openDb();
  const transaction = db.transaction(RECORD_STORE, 'readwrite');
  const store = transaction.objectStore(RECORD_STORE);
  records.forEach((record) => store.put({ ...record, importedAt: new Date().toISOString() }));
  await txDone(transaction);
  db.close();
};

export const getAllRecords = () => getAllFromStore(RECORD_STORE);
export const getAllActions = () => getAllFromStore(ACTION_STORE);

export const saveAction = async (action) => {
  const db = await openDb();
  const transaction = db.transaction(ACTION_STORE, 'readwrite');
  transaction.objectStore(ACTION_STORE).put(action);
  await txDone(transaction);
  db.close();
};

export const deleteAction = async (uniqueKey) => {
  const db = await openDb();
  const transaction = db.transaction(ACTION_STORE, 'readwrite');
  transaction.objectStore(ACTION_STORE).delete(uniqueKey);
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

export const exportBackup = async () => ({
  exportedAt: new Date().toISOString(),
  version: 2,
  records: await getAllRecords(),
  actions: await getAllActions(),
});

export const importBackup = async (backup) => {
  if (!backup || !Array.isArray(backup.records)) throw new Error('备份 JSON 格式不正确，缺少 records 数组。');
  const records = backup.records.map((record) => ({ ...record, uniqueKey: record.uniqueKey || `${record.date}__${record.sku}` }));
  const actions = Array.isArray(backup.actions) ? backup.actions.map((action) => ({ ...action, uniqueKey: action.uniqueKey || `${action.date}__${action.sku}` })) : [];
  await replaceAllRecords(records);
  await replaceAllActions(actions);
  return { records: records.length, actions: actions.length };
};
