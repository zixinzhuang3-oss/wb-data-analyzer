import { addDays } from './date.js';
import { toProfitCny, toProfitRub } from './currency.js';
import { hasValidBusinessData } from './excel.js';
import { buildPlatformSkuKey, normalizePlatform, normalizeSku } from './actions.js';

const DAY_MS = 86400000;
const sum = (rows, key) => rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);
const sumProfitCny = (rows) => rows.reduce((total, row) => total + toProfitCny(row), 0);
const sumProfitRub = (rows) => rows.reduce((total, row) => total + toProfitRub(row), 0);
const safeDivide = (a, b) => (b ? a / b : 0);
const avgPresent = (rows, key) => { const values = rows.map((row) => Number(row[key])).filter(Number.isFinite); return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0; };
const toDate = (date) => new Date(`${date}T00:00:00Z`);
export { addDays };

export const getPlatformOptions = () => ['all', 'WB', 'Ozon'];
export const getSkuOptions = (records, platform = 'all') => {
  const targetPlatform = platform === 'all' || !platform ? 'all' : normalizePlatform(platform);
  const rows = records.filter((record) => targetPlatform === 'all' || normalizePlatform(record.platform) === targetPlatform);
  return [...new Set(rows.map((record) => targetPlatform === 'all' ? buildPlatformSkuKey(record.platform, record.sku) : normalizeSku(record.sku)).filter(Boolean))].sort();
};
export const getDateOptions = (records) => [...new Set(records.map((record) => record.date).filter(Boolean))].sort().reverse();
export const getLatestDate = (records) => getDateOptions(records)[0] || '';

export const resolveDateRange = (records, filters = {}) => {
  const dates = getDateOptions(records).slice().reverse();
  const latest = dates.at(-1) || '';
  if (filters.allDates) return { allDates: true, startDate: dates[0] || '', endDate: latest };
  const startDate = filters.startDate || filters.date || '';
  const endDate = filters.endDate || filters.date || '';
  if (startDate && endDate && startDate > endDate) return { allDates: false, startDate: endDate, endDate: startDate };
  return { allDates: false, startDate, endDate };
};

export const getPreviousRange = ({ startDate, endDate, allDates }) => {
  if (allDates || !startDate || !endDate) return null;
  const length = Math.floor((toDate(endDate).getTime() - toDate(startDate).getTime()) / DAY_MS) + 1;
  return { startDate: addDays(startDate, -length), endDate: addDays(startDate, -1), length };
};

export const filterRecords = (records, filters = {}) => {
  const range = resolveDateRange(records, filters);
  return records.filter((record) => {
    if (!hasValidBusinessData(record)) return false;
    const dateMatched = range.allDates || (!range.startDate && !range.endDate) || (record.date >= range.startDate && record.date <= range.endDate);
    const platform = normalizePlatform(record.platform);
    const platformMatched = !filters.platform || filters.platform === 'all' || platform === normalizePlatform(filters.platform);
    const skuMatched = !filters.sku || record.sku === filters.sku || buildPlatformSkuKey(platform, record.sku) === filters.sku;
    return dateMatched && platformMatched && skuMatched;
  });
};

export const summarizeRecords = (records) => {
  const totalOrders = sum(records, 'totalOrders');
  const totalAdSpend = sum(records, 'adSpend');
  const totalProfitCny = sumProfitCny(records);
  const totalProfitRub = sumProfitRub(records);
  const totalProfit = totalProfitRub;
  const totalRevenue = sum(records, 'revenue');
  const impressions = sum(records, 'impressions');
  const clicks = sum(records, 'clicks');
  const addToCart = sum(records, 'addToCart');
  const adImpressions = sum(records, 'adImpressions');
  const adClicks = sum(records, 'adClicks');
  const adAddToCart = sum(records, 'adAddToCart');
  const adOrders = sum(records, 'adOrders');
  const avgDealPriceRub = avgPresent(records, 'dealPriceRub');
  return {
    dateCount: new Set(records.map((record) => record.date).filter(Boolean)).size,
    skuCount: new Set(records.map((record) => buildPlatformSkuKey(record.platform, record.sku)).filter(Boolean)).size,
    totalOrders,
    totalAdSpend,
    totalProfit,
    totalProfitCny,
    totalProfitRub,
    totalRevenue,
    avgDealPriceRub,
    impressions,
    clicks,
    addToCart,
    adImpressions,
    adClicks,
    adAddToCart,
    adOrders,
    adCostPerOrder: safeDivide(totalAdSpend, adOrders),
    adAvgClickCost: safeDivide(totalAdSpend, adClicks),
    margin: safeDivide(totalProfitRub, totalRevenue),
    adShare: safeDivide(totalAdSpend, totalRevenue),
    roi: safeDivide(totalRevenue, totalAdSpend),
    acos: safeDivide(totalAdSpend, totalRevenue),
    ctr: safeDivide(clicks, impressions),
    cvr: safeDivide(addToCart, clicks),
    orderConversionRate: safeDivide(totalOrders, clicks),
    adCtr: safeDivide(adClicks, adImpressions),
    adClickAddToCartRate: safeDivide(adAddToCart, adClicks),
    adStatus: records.some((record) => record.adStatus === '开启') ? '开启' : records.some((record) => record.adStatus === '关闭') ? '关闭' : records.length ? '广告关闭 / 无广告数据' : '无数据',
  };
};

export const buildComparison = (records, filters = {}) => {
  const currentRange = resolveDateRange(records, filters);
  const currentRecords = filterRecords(records, { ...filters, ...currentRange });
  const previousRange = getPreviousRange(currentRange);
  const previousRecords = previousRange ? filterRecords(records, { ...filters, ...previousRange, allDates: false }) : [];
  return { currentRange, previousRange, currentRecords, previousRecords, current: summarizeRecords(currentRecords), previous: summarizeRecords(previousRecords), hasPreviousData: previousRecords.length > 0 };
};

export const summarizeByDate = (records) => {
  const grouped = new Map();
  records.filter(hasValidBusinessData).forEach((record) => {
    if (!grouped.has(record.date)) grouped.set(record.date, []);
    grouped.get(record.date).push(record);
  });
  return [...grouped.entries()].sort((a, b) => b[0].localeCompare(a[0])).map(([date, rows]) => ({
    date,
    skuCount: new Set(rows.map((row) => buildPlatformSkuKey(row.platform, row.sku))).size,
    totalOrders: sum(rows, 'totalOrders'),
    totalRevenue: sum(rows, 'revenue'),
    avgDealPriceRub: avgPresent(rows, 'dealPriceRub'),
    totalAdSpend: sum(rows, 'adSpend'),
    totalProfit: sumProfitRub(rows),
    totalProfitRub: sumProfitRub(rows),
    totalProfitCny: sumProfitCny(rows),
  }));
};
