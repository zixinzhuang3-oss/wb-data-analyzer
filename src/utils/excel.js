import { DAILY_FIELDS, NUMERIC_FIELD_KEYS, fieldLabels, isSkuSheet } from './fields.js';
import { parseOperationActionText } from './actions.js';

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
const pad2 = (value) => String(value).padStart(2, '0');
const formatYmd = (year, month, day) => `${year}-${pad2(month)}-${pad2(day)}`;

export const toIsoDate = (value, XLSX = window.XLSX) => {
  if (value === undefined || value === null || value === '') return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return formatYmd(value.getFullYear(), value.getMonth() + 1, value.getDate());
  if (typeof value === 'number') {
    const parsed = XLSX?.SSF?.parse_date_code?.(value);
    return parsed ? formatYmd(parsed.y, parsed.m, parsed.d) : '';
  }
  const text = String(value).trim();
  let match = text.match(/^(\d{4})[\-/\.年](\d{1,2})[\-/\.月](\d{1,2})/);
  if (match) return formatYmd(match[1], match[2], match[3]);
  match = text.match(/^(\d{1,2})\s*月\s*(\d{1,2})\s*日?$/);
  if (match) return formatYmd(new Date().getFullYear(), match[1], match[2]);
  match = text.match(/^(\d{1,2})[\-/\.](\d{1,2})$/);
  if (match) return formatYmd(new Date().getFullYear(), match[1], match[2]);
  return text;
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

const AD_FIELD_KEYS = ['adSpend', 'adOrders', 'adShare', 'adImpressions', 'adClicks'];
const isBlank = (value) => value === undefined || value === null || String(value).trim() === '';

const detectAdStatus = (record, headers, row, headerMap) => {
  const present = AD_FIELD_KEYS.filter((key) => headerMap[key] >= 0);
  const allBlank = present.length === 0 || present.every((key) => isBlank(row[headerMap[key]]));
  if (allBlank && (Number(record.totalOrders) > 0 || Number(record.revenue) > 0 || Number(record.profit) !== 0)) return '无广告数据';
  if (AD_FIELD_KEYS.some((key) => Number(record[key]) > 0)) return '开启';
  return allBlank ? '无广告数据' : '关闭';
};

const normalizeSheetRow = (sheetName, headers, row, XLSX) => {
  const headerMap = buildHeaderMap(headers);
  const record = { sku: sheetName.trim(), sourceSheet: sheetName.trim() };
  DAILY_FIELDS.forEach((field) => {
    if (field.key === 'sku') return;
    const index = headerMap[field.key];
    const value = index >= 0 ? row[index] : '';
    if (field.key === 'date') record.date = toIsoDate(value, XLSX);
    else if (NUMERIC_FIELD_KEYS.has(field.key)) record[field.key] = toNumber(value);
    else record[field.key] = String(value ?? '').trim();
  });
  record.adStatus = detectAdStatus(record, headers, row, headerMap);
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
  const actions = [];

  skuSheets.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: true, blankrows: false });
    const headerIndex = findHeaderRowIndex(rows);
    if (headerIndex < 0) return;
    const headers = rows[headerIndex];
    rows.slice(headerIndex + 1).filter((row) => !rowIsEmpty(row)).forEach((row) => {
      const record = normalizeSheetRow(sheetName, headers, row, XLSX);
      if (record.date) {
        records.push(record);
        const parsedAction = parseOperationActionText(record.operationAction, record.date, record.sku);
        if (parsedAction) actions.push(parsedAction);
      }
    });
  });

  return { records, actions, skuSheets, skippedSheets, fieldLabels };
};
