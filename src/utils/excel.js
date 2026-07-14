import { DAILY_FIELDS, NUMERIC_FIELD_KEYS, fieldLabels, isSkuSheet, normalizeSkuSheetName } from './fields.js';
import { buildActionKey, normalizePlatform, parseOperationActionText } from './actions.js';
import { CNY_TO_RUB } from './currency.js';
import { normalizeDateKey } from './date.js';

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
export const toIsoDate = (value, XLSX = window.XLSX) => normalizeDateKey(value, XLSX);

export const toNumber = (value) => {
  if (value === undefined || value === null || String(value).trim() === '') return 0;
  const raw = String(value).trim();
  const cleaned = raw.replace(/[₽￥¥$,%\s]/g, '');
  const parsed = Number(cleaned);
  if (Number.isNaN(parsed)) return 0;
  return raw.includes('%') ? parsed / 100 : parsed;
};
const safeDivide = (a, b) => (b ? a / b : 0);

const FIXED_COLUMNS = {
  date: 0,
  dealPriceRub: 4,
  impressions: 13,
  clicks: 14,
  ctr: 15,
  addToCart: 16,
  conversionRate: 17,
  revenue: 26,
  adSpend: 27,
  adOrders: 28,
  adShare: 29,
  adCtr: 30,
  adImpressions: 31,
  adClicks: 32,
  adClickAddToCartRate: 33,
  adAddToCart: 34,
  adCostPerOrder: 35,
  adAvgClickCost: 36,
};

const FIXED_COLUMN_LABELS = {
  date: 'A列',
  dealPriceRub: 'E列',
  impressions: 'N列',
  clicks: 'O列',
  ctr: 'P列',
  addToCart: 'Q列',
  conversionRate: 'R列',
  revenue: 'AA列',
  adSpend: 'AB列',
  adOrders: 'AC列',
  adShare: 'AD列',
  adCtr: 'AE列',
  adImpressions: 'AF列',
  adClicks: 'AG列',
  adClickAddToCartRate: 'AH列',
  adAddToCart: 'AI列',
  adCostPerOrder: 'AJ列',
  adAvgClickCost: 'AK列（卢布 ₽）',
  profit: '利润列，单位 ¥',
};


const DIAGNOSTIC_FIELD_LABELS = {
  dealPriceRub: '成交价',
  revenue: '销售额',
  adSpend: '广告费',
  profit: '利润',
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

const AD_FIELD_KEYS = ['adSpend', 'adOrders', 'adShare', 'adCtr', 'adImpressions', 'adClicks', 'adClickAddToCartRate', 'adAddToCart', 'adCostPerOrder', 'adAvgClickCost'];
const isBlank = (value) => value === undefined || value === null || String(value).trim() === '';

const detectAdStatus = (record, headers, row, headerMap) => {
  const adColumns = AD_FIELD_KEYS.map((key) => FIXED_COLUMNS[key]).filter((index) => index !== undefined);
  const allBlank = adColumns.every((index) => isBlank(row[index]));
  if (allBlank && (Number(record.totalOrders) > 0 || Number(record.revenue) > 0 || Number(record.profit) !== 0)) return '无广告数据';
  if (AD_FIELD_KEYS.some((key) => Number(record[key]) > 0)) return '开启';
  return allBlank ? '无广告数据' : '关闭';
};

const getFieldValue = (field, headerMap, row) => {
  const fixedIndex = FIXED_COLUMNS[field.key];
  if (fixedIndex !== undefined) return row[fixedIndex];
  const index = headerMap[field.key];
  return index >= 0 ? row[index] : '';
};

const BUSINESS_DATA_KEYS = ['dealPriceRub', 'totalOrders', 'revenue', 'impressions', 'clicks', 'addToCart', 'adSpend', 'adOrders', 'adImpressions', 'adClicks', 'adAddToCart', 'profit', 'profitCny', 'profitRub', 'stock', 'price', 'reviews', 'rating'];
const POSITIVE_BUSINESS_KEYS = new Set(['totalOrders', 'revenue', 'impressions', 'clicks', 'addToCart', 'adSpend', 'adOrders', 'adImpressions', 'adClicks', 'adAddToCart']);
const PRESENT_BUSINESS_KEYS = new Set(['dealPriceRub', 'stock', 'price', 'reviews', 'rating', 'profit', 'profitCny', 'profitRub']);

const hasRawBusinessCell = (fieldKey, rawValue) => {
  if (!BUSINESS_DATA_KEYS.includes(fieldKey) || isBlank(rawValue)) return false;
  const value = toNumber(rawValue);
  if (POSITIVE_BUSINESS_KEYS.has(fieldKey)) return value > 0;
  if (PRESENT_BUSINESS_KEYS.has(fieldKey)) return Number.isFinite(value);
  return false;
};

export const hasValidBusinessData = (record = {}) => {
  if (typeof record.hasValidBusinessData === 'boolean') return record.hasValidBusinessData;
  if (typeof record.__hasValidBusinessData === 'boolean') return record.__hasValidBusinessData;
  return BUSINESS_DATA_KEYS.some((key) => {
    const raw = record[key];
    if (raw === undefined || raw === null || String(raw).trim?.() === '') return false;
    const value = Number(raw);
    if (!Number.isFinite(value)) return false;
    if (POSITIVE_BUSINESS_KEYS.has(key)) return value > 0;
    if (['dealPriceRub', 'stock', 'price', 'reviews', 'rating'].includes(key)) return true;
    return value !== 0;
  });
};

export const hasEffectiveDailyData = hasValidBusinessData;

export const normalizeSheetRow = (sheetName, headers, row, XLSX, platform = 'WB') => {
  const normalizedPlatform = normalizePlatform(platform);
  const headerMap = buildHeaderMap(headers);
  const normalizedSheetName = normalizeSkuSheetName(sheetName);
  const record = { platform: normalizedPlatform, sku: normalizedSheetName, sourceSheet: String(sheetName || '').trim() };
  record.dealPriceRubRaw = row[FIXED_COLUMNS.dealPriceRub] ?? '';
  const rawBusinessFields = {};
  DAILY_FIELDS.forEach((field) => {
    if (field.key === 'sku' || field.key === 'platform') return;
    const value = getFieldValue(field, headerMap, row);
    if (hasRawBusinessCell(field.key, value)) rawBusinessFields[field.key] = true;
    if (field.key === 'date') record.date = toIsoDate(value, XLSX);
    else if (NUMERIC_FIELD_KEYS.has(field.key)) record[field.key] = toNumber(value);
    else record[field.key] = String(value ?? '').trim();
  });
  record.ctr = safeDivide(record.clicks, record.impressions) || record.ctr || 0;
  record.conversionRate = safeDivide(record.addToCart, record.clicks) || record.conversionRate || 0;
  record.orderConversionRate = safeDivide(record.totalOrders, record.clicks);
  record.adCtr = safeDivide(record.adClicks, record.adImpressions) || record.adCtr || 0;
  record.adClickAddToCartRate = safeDivide(record.adAddToCart, record.adClicks) || record.adClickAddToCartRate || 0;
  record.adShare = safeDivide(record.adSpend, record.revenue) || record.adShare || 0;
  record.adCostPerOrder = safeDivide(record.adSpend, record.adOrders) || record.adCostPerOrder || 0;
  record.adAvgClickCost = safeDivide(record.adSpend, record.adClicks) || record.adAvgClickCost || 0;
  record.profitCny = Number(record.profit) || 0;
  record.profitRub = Math.round(record.profitCny * CNY_TO_RUB * 100) / 100;
  record.profit = record.profitCny;
  record.hasValidBusinessData = Object.keys(rawBusinessFields).length > 0;
  record.__hasValidBusinessData = record.hasValidBusinessData;
  record.adStatus = detectAdStatus(record, headers, row, headerMap);
  record.uniqueKey = buildActionKey(record.date, record.sku, record.platform);
  return record;
};

export const buildSheetDiagnostics = (sheetName, rows, headerIndex, XLSX) => ({
  sheetName: normalizeSkuSheetName(sheetName),
  fields: Object.fromEntries(Object.entries(FIXED_COLUMN_LABELS).map(([key, column]) => [DIAGNOSTIC_FIELD_LABELS[key] || fieldLabels[key] || key, column])),
  dealPriceSamples: rows.slice(headerIndex + 1).filter((row) => !rowIsEmpty(row)).slice(0, 5).map((row, offset) => ({ sku: normalizeSkuSheetName(sheetName), date: toIsoDate(row[FIXED_COLUMNS.date], XLSX), rowNumber: headerIndex + 2 + offset, raw: row[FIXED_COLUMNS.dealPriceRub] ?? '', parsed: toNumber(row[FIXED_COLUMNS.dealPriceRub]), field: 'dealPriceRub', column: 'E列' })),
  blankAdDates: rows.slice(headerIndex + 1).filter((row) => !rowIsEmpty(row) && AD_FIELD_KEYS.every((key) => isBlank(row[FIXED_COLUMNS[key]]))).map((row) => toIsoDate(row[FIXED_COLUMNS.date], XLSX)).filter(Boolean),
});


export const PRICE_CHANGE_THRESHOLD = 0.5;

export const detectPriceAction = (currentPrice, previousPrice, threshold = PRICE_CHANGE_THRESHOLD) => {
  const hasCurrent = currentPrice !== null && currentPrice !== undefined && String(currentPrice).trim?.() !== '';
  const hasPrevious = previousPrice !== null && previousPrice !== undefined && String(previousPrice).trim?.() !== '';
  if (!hasCurrent || !Number.isFinite(Number(currentPrice))) return { priceAction: '', message: '当天无成交价，无法判断', previousPrice: null, source: 'price_auto' };
  if (!hasPrevious || !Number.isFinite(Number(previousPrice))) return { priceAction: '', message: '暂无上期价格，无法判断', previousPrice: null, source: 'price_auto' };
  const diff = Number(currentPrice) - Number(previousPrice);
  const priceAction = Math.abs(diff) < threshold ? '保持价格' : diff > 0 ? '涨价' : '降价';
  return { priceAction, priceActionSource: 'price_auto', source: 'price_auto', previousPrice: Number(previousPrice), priceChangeRub: diff, priceChangePercent: previousPrice ? diff / Math.abs(previousPrice) : 0, message: `自动识别为${priceAction}` };
};

export const buildPriceAutoActions = (records = []) => {
  const bySku = new Map();
  records.forEach((record) => {
    if (!record?.sku || !record?.date) return;
    const key = `${normalizePlatform(record.platform)}__${record.sku}`;
    if (!bySku.has(key)) bySku.set(key, []);
    bySku.get(key).push(record);
  });
  const actions = [];
  bySku.forEach((rows) => {
    let previousPrice = null;
    rows.sort((a, b) => a.date.localeCompare(b.date)).forEach((record) => {
      const hasRawPrice = record.dealPriceRubRaw !== null && record.dealPriceRubRaw !== undefined && String(record.dealPriceRubRaw).trim() !== '';
      const currentPrice = Number(record.dealPriceRub);
      if (!hasRawPrice || !Number.isFinite(currentPrice)) return;
      const detected = detectPriceAction(currentPrice, previousPrice);
      record.previousDealPriceRub = Number.isFinite(Number(previousPrice)) ? Number(previousPrice) : null;
      record.priceChangeRub = detected.priceChangeRub ?? 0;
      record.priceChangePercent = detected.priceChangePercent ?? 0;
      record.priceActionAuto = detected.priceAction || '';
      record.priceActionSource = detected.priceAction ? 'price_auto' : '';
      record.priceActionMessage = detected.message;
      if (detected.priceAction) actions.push(parseOperationActionText(`价格动作=${detected.priceAction}`, record.date, record.sku, record.platform) || null);
      const last = actions.at(-1);
      if (last && last.date === record.date && last.sku === record.sku) {
        last.source = 'price_auto';
        last.priceActionSource = 'price_auto';
        last.previousDealPriceRub = record.previousDealPriceRub;
        last.dealPriceRub = currentPrice;
        last.priceChangeRub = record.priceChangeRub;
        last.priceChangePercent = record.priceChangePercent;
      }
      previousPrice = currentPrice;
    });
  });
  return actions.filter(Boolean);
};


export const buildSkuCatalogKey = (platform = 'WB', sku = '') => `${normalizePlatform(platform)}_${normalizeSkuSheetName(sku)}`;

export const buildSkuCatalogFromSheetNames = (sheetNames = [], platform = 'WB', parsedRowCounts = {}, importedAt = new Date().toISOString()) => {
  const normalizedPlatform = normalizePlatform(platform);
  const catalog = new Map();
  (Array.isArray(sheetNames) ? sheetNames : []).forEach((rawSheetName) => {
    if (!isSkuSheet(rawSheetName)) return;
    const sku = normalizeSkuSheetName(rawSheetName);
    const parsedRowCount = Number(parsedRowCounts[sku] ?? parsedRowCounts[String(rawSheetName || '').trim()] ?? 0) || 0;
    catalog.set(buildSkuCatalogKey(normalizedPlatform, sku), {
      uniqueKey: buildSkuCatalogKey(normalizedPlatform, sku),
      platform: normalizedPlatform,
      sku,
      sheetName: String(rawSheetName || '').trim(),
      importedAt,
      hasParsedRows: parsedRowCount > 0,
      parsedRowCount,
    });
  });
  return [...catalog.values()].sort((a, b) => a.sku.localeCompare(b.sku));
};

export const buildWorkbookSheetDiagnostics = (sheetNames = [], skuCatalog = [], parseReasons = {}) => {
  const rawSheets = (Array.isArray(sheetNames) ? sheetNames : []).map((name) => String(name || '').trim());
  const skuSheets = rawSheets.filter(isSkuSheet).map(normalizeSkuSheetName);
  const skippedSheets = rawSheets.filter((name) => !isSkuSheet(name));
  const counts = Object.fromEntries((Array.isArray(skuCatalog) ? skuCatalog : []).map((item) => [normalizeSkuSheetName(item.sku), Number(item.parsedRowCount) || 0]));
  const skuParseDiagnostics = skuSheets.map((sku) => ({
    sku,
    parsedRowCount: counts[sku] || 0,
    message: counts[sku] > 0 ? `${sku}：解析到 ${counts[sku]} 行有效数据` : (parseReasons[sku] ? `${sku}：字段读取失败，原因：${parseReasons[sku]}` : `${sku}：已识别为 SKU，但当前未解析到有效数据`),
  }));
  return { rawSheets, skuSheets, skippedSheets, skuParseDiagnostics };
};

const parseSheetByFixedColumns = (sheetName, headers, row, XLSX, platform) => normalizeSheetRow(sheetName, headers, row, XLSX, platform);
export const parseWbSheet = parseSheetByFixedColumns;
export const parseOzonSheet = parseSheetByFixedColumns;

export const parseExcelWorkbook = async (file, platform = 'WB') => {
  const normalizedPlatform = normalizePlatform(platform);
  const XLSX = await loadSheetJs();
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const rawSheets = workbook.SheetNames.map((name) => String(name || '').trim());
  const skippedSheets = rawSheets.filter((name) => !isSkuSheet(name));
  const skuSheets = rawSheets.filter(isSkuSheet).map(normalizeSkuSheetName);
  const records = [];
  const actions = [];
  const diagnostics = [];
  const parsedRowCounts = {};
  const parseReasons = {};

  workbook.SheetNames.filter(isSkuSheet).forEach((rawSheetName) => {
    const sheetName = normalizeSkuSheetName(rawSheetName);
    const worksheet = workbook.Sheets[rawSheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: true, blankrows: false });
    const headerIndex = findHeaderRowIndex(rows);
    if (headerIndex < 0) {
      parseReasons[sheetName] = '未找到包含日期和订单/利润/广告费的表头行';
      return;
    }
    diagnostics.push(buildSheetDiagnostics(sheetName, rows, headerIndex, XLSX));
    const headers = rows[headerIndex];
    rows.slice(headerIndex + 1).filter((row) => !rowIsEmpty(row)).forEach((row) => {
      const parser = normalizedPlatform === 'Ozon' ? parseOzonSheet : parseWbSheet;
      const record = parser(sheetName, headers, row, XLSX, normalizedPlatform);
      if (record.date && hasEffectiveDailyData(record)) {
        records.push(record);
        parsedRowCounts[sheetName] = (parsedRowCounts[sheetName] || 0) + 1;
        const parsedAction = parseOperationActionText(record.operationAction, record.date, record.sku, record.platform);
        if (parsedAction) actions.push(parsedAction);
      }
    });
  });

  actions.push(...buildPriceAutoActions(records));
  const importedAt = new Date().toISOString();
  const skuCatalog = buildSkuCatalogFromSheetNames(workbook.SheetNames, normalizedPlatform, parsedRowCounts, importedAt);
  const sheetDiagnostics = buildWorkbookSheetDiagnostics(workbook.SheetNames, skuCatalog, parseReasons);

  return { platform: normalizedPlatform, records, actions, skuCatalog, rawSheets, skuSheets, skippedSheets, sheetDiagnostics, fieldLabels, diagnostics };
};
