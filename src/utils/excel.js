import { DAILY_FIELDS, NUMERIC_FIELD_KEYS, fieldLabels, isSkuSheet } from './fields.js';

const SHEETJS_CDN = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';

const loadSheetJs = () => new Promise((resolve, reject) => {
  if (window.XLSX) {
    resolve(window.XLSX);
    return;
  }
  const script = document.createElement('script');
  script.src = SHEETJS_CDN;
  script.async = true;
  script.onload = () => resolve(window.XLSX);
  script.onerror = () => reject(new Error('Excel 解析库加载失败，请检查网络后重试，或后续改为本地依赖版本。'));
  document.head.appendChild(script);
});

const normalizeHeader = (value) => String(value ?? '').trim().replace(/\s+/g, '').toLowerCase();
const excelDateEpoch = Date.UTC(1899, 11, 30);

const toIsoDate = (value) => {
  if (!value) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') return new Date(excelDateEpoch + value * 86400000).toISOString().slice(0, 10);
  const text = String(value).trim();
  const match = text.match(/^(\d{4})[\-/\.年](\d{1,2})[\-/\.月](\d{1,2})/);
  if (match) return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? text : parsed.toISOString().slice(0, 10);
};

const toNumber = (value) => {
  if (value === undefined || value === null || value === '') return 0;
  const raw = String(value).trim();
  const cleaned = raw.replace(/[%,$￥₽,\s]/g, '');
  const parsed = Number(cleaned);
  if (Number.isNaN(parsed)) return 0;
  return raw.includes('%') ? parsed / 100 : parsed;
};

const buildHeaderMap = (headers) => {
  const normalizedHeaders = headers.map(normalizeHeader);
  return Object.fromEntries(DAILY_FIELDS.map((field) => {
    const aliases = field.aliases.map(normalizeHeader);
    const index = normalizedHeaders.findIndex((header) => aliases.includes(header));
    return [field.key, index];
  }));
};

const rowIsEmpty = (row) => row.every((cell) => cell === undefined || cell === null || String(cell).trim() === '');

const findHeaderRowIndex = (rows) => rows.findIndex((row) => {
  const headers = row.map(normalizeHeader);
  const hasDate = DAILY_FIELDS.find((field) => field.key === 'date').aliases.map(normalizeHeader).some((alias) => headers.includes(alias));
  const hasOrdersOrProfit = ['totalOrders', 'profit', 'adSpend'].some((key) => {
    const field = DAILY_FIELDS.find((item) => item.key === key);
    return field.aliases.map(normalizeHeader).some((alias) => headers.includes(alias));
  });
  return hasDate && hasOrdersOrProfit;
});

const normalizeSheetRow = (sheetName, headers, row) => {
  const headerMap = buildHeaderMap(headers);
  const record = { sku: sheetName.trim(), sourceSheet: sheetName.trim() };
  DAILY_FIELDS.forEach((field) => {
    if (field.key === 'sku') return;
    const index = headerMap[field.key];
    const value = index >= 0 ? row[index] : '';
    if (field.key === 'date') record.date = toIsoDate(value);
    else if (NUMERIC_FIELD_KEYS.has(field.key)) record[field.key] = toNumber(value);
    else record[field.key] = String(value ?? '').trim();
  });
  record.uniqueKey = `${record.date}__${record.sku}`;
  return record;
};

export const parseExcelWorkbook = async (file) => {
  const XLSX = await loadSheetJs();
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const skippedSheets = workbook.SheetNames.filter((name) => !isSkuSheet(name));
  const skuSheets = workbook.SheetNames.filter(isSkuSheet);
  const records = [];

  skuSheets.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false, blankrows: false });
    const headerIndex = findHeaderRowIndex(rows);
    if (headerIndex < 0) return;
    const headers = rows[headerIndex];
    rows.slice(headerIndex + 1).filter((row) => !rowIsEmpty(row)).forEach((row) => {
      const record = normalizeSheetRow(sheetName, headers, row);
      if (record.date) records.push(record);
    });
  });

  return { records, skuSheets, skippedSheets, fieldLabels };
};
