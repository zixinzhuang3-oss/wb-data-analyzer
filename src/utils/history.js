const DAY_MS = 86400000;
const sum = (rows, key) => rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);
const safeDivide = (a, b) => (b ? a / b : 0);
const toDate = (date) => new Date(`${date}T00:00:00Z`);
export const addDays = (date, days) => new Date(toDate(date).getTime() + days * DAY_MS).toISOString().slice(0, 10);

export const getSkuOptions = (records) => [...new Set(records.map((record) => record.sku).filter(Boolean))].sort();
export const getDateOptions = (records) => [...new Set(records.map((record) => record.date).filter(Boolean))].sort().reverse();
export const getLatestDate = (records) => getDateOptions(records)[0] || '';

export const resolveDateRange = (records, filters = {}) => {
  const dates = getDateOptions(records).slice().reverse();
  const latest = dates.at(-1) || '';
  if (filters.allDates) return { allDates: true, startDate: dates[0] || '', endDate: latest };
  const startDate = filters.startDate || filters.date || latest;
  const endDate = filters.endDate || filters.date || latest;
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
    const dateMatched = range.allDates || (!range.startDate && !range.endDate) || (record.date >= range.startDate && record.date <= range.endDate);
    const skuMatched = !filters.sku || record.sku === filters.sku;
    return dateMatched && skuMatched;
  });
};

export const summarizeRecords = (records) => {
  const totalOrders = sum(records, 'totalOrders');
  const totalAdSpend = sum(records, 'adSpend');
  const totalProfit = sum(records, 'profit');
  const totalRevenue = sum(records, 'revenue');
  const impressions = sum(records, 'impressions') || sum(records, 'adImpressions');
  const clicks = sum(records, 'clicks') || sum(records, 'adClicks');
  return {
    dateCount: new Set(records.map((record) => record.date).filter(Boolean)).size,
    skuCount: new Set(records.map((record) => record.sku).filter(Boolean)).size,
    totalOrders,
    totalAdSpend,
    totalProfit,
    totalRevenue,
    impressions,
    clicks,
    margin: safeDivide(totalProfit, totalRevenue),
    adShare: safeDivide(totalAdSpend, totalRevenue),
    roi: safeDivide(totalRevenue, totalAdSpend),
    acos: safeDivide(totalAdSpend, totalRevenue),
    ctr: safeDivide(clicks, impressions),
    cvr: safeDivide(totalOrders, clicks),
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
  records.forEach((record) => {
    if (!grouped.has(record.date)) grouped.set(record.date, []);
    grouped.get(record.date).push(record);
  });
  return [...grouped.entries()].sort((a, b) => b[0].localeCompare(a[0])).map(([date, rows]) => ({
    date,
    skuCount: new Set(rows.map((row) => row.sku)).size,
    totalOrders: sum(rows, 'totalOrders'),
    totalAdSpend: sum(rows, 'adSpend'),
    totalProfit: sum(rows, 'profit'),
  }));
};
