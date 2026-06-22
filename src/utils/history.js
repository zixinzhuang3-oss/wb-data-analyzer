const DAY_MS = 86400000;
const sum = (rows, key) => rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);
const number = (value) => Number(value) || 0;
const safeDivide = (a, b) => (b ? a / b : 0);

export const getSkuOptions = (records) => [...new Set(records.map((record) => record.sku).filter(Boolean))].sort();
export const getDateOptions = (records) => [...new Set(records.map((record) => record.date).filter(Boolean))].sort().reverse();

export const parseIsoDateUtc = (date) => {
  const match = String(date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
};

export const formatIsoDateUtc = (date) => date instanceof Date && !Number.isNaN(date.getTime()) ? date.toISOString().slice(0, 10) : '';
export const addDays = (date, days) => {
  const parsed = parseIsoDateUtc(date);
  if (!parsed) return '';
  return formatIsoDateUtc(new Date(parsed.getTime() + days * DAY_MS));
};

export const countDaysInclusive = (startDate, endDate) => {
  const start = parseIsoDateUtc(startDate);
  const end = parseIsoDateUtc(endDate);
  if (!start || !end || start > end) return 0;
  return Math.round((end.getTime() - start.getTime()) / DAY_MS) + 1;
};

export const normalizeDateRange = (records = [], filters = {}) => {
  const dates = getDateOptions(records).sort();
  const fallbackEnd = dates.at(-1) || '';
  const fallbackStart = dates[0] || fallbackEnd;
  const isAllDates = filters.quickRange === '全部日期' || filters.allDates;
  let startDate = isAllDates ? fallbackStart : (filters.startDate || fallbackEnd);
  let endDate = isAllDates ? fallbackEnd : (filters.endDate || startDate || fallbackEnd);
  if (startDate && endDate && startDate > endDate) [startDate, endDate] = [endDate, startDate];
  return { startDate, endDate, isAllDates };
};

export const getPreviousDateRange = (startDate, endDate) => {
  const dayCount = countDaysInclusive(startDate, endDate);
  if (!dayCount) return { startDate: '', endDate: '', dayCount: 0 };
  const previousEndDate = addDays(startDate, -1);
  const previousStartDate = addDays(previousEndDate, -(dayCount - 1));
  return { startDate: previousStartDate, endDate: previousEndDate, dayCount };
};

export const filterRecords = (records, filters = {}) => {
  const { startDate, endDate } = normalizeDateRange(records, filters);
  return records.filter((record) => {
    const dateMatched = (!startDate || record.date >= startDate) && (!endDate || record.date <= endDate);
    const skuMatched = !filters.sku || record.sku === filters.sku;
    return dateMatched && skuMatched;
  });
};

export const summarizeRecords = (records) => {
  const totalOrders = sum(records, 'totalOrders');
  const totalRevenue = sum(records, 'revenue');
  const totalAdSpend = sum(records, 'adSpend');
  const totalProfit = sum(records, 'profit');
  const totalImpressions = records.reduce((total, row) => total + (number(row.impressions) || number(row.adImpressions)), 0);
  const totalClicks = records.reduce((total, row) => total + (number(row.clicks) || number(row.adClicks)), 0);
  return {
    dateCount: new Set(records.map((record) => record.date).filter(Boolean)).size,
    skuCount: new Set(records.map((record) => record.sku).filter(Boolean)).size,
    totalOrders,
    totalRevenue,
    totalAdSpend,
    totalProfit,
    averageMargin: safeDivide(totalProfit, totalRevenue),
    adShare: safeDivide(totalAdSpend, totalRevenue),
    roi: safeDivide(totalRevenue, totalAdSpend),
    acos: safeDivide(totalAdSpend, totalRevenue),
    totalImpressions,
    totalClicks,
    ctr: safeDivide(totalClicks, totalImpressions),
    cvr: safeDivide(totalOrders, totalClicks),
    totalStock: sum(records, 'stock'),
  };
};

export const METRIC_DEFINITIONS = [
  { key: 'totalOrders', label: '总订单', type: 'number', direction: 'positive' },
  { key: 'totalRevenue', label: '总销售额', type: 'money', direction: 'positive' },
  { key: 'totalAdSpend', label: '总广告费', type: 'money', direction: 'cost' },
  { key: 'totalProfit', label: '总利润', type: 'money', direction: 'positive' },
  { key: 'averageMargin', label: '平均利润率', type: 'percent', direction: 'positive' },
  { key: 'adShare', label: '广告占比', type: 'percent', direction: 'negative' },
  { key: 'roi', label: 'ROI', type: 'number', direction: 'positive' },
  { key: 'acos', label: 'ACOS', type: 'percent', direction: 'negative' },
  { key: 'totalImpressions', label: '总曝光', type: 'number', direction: 'neutral' },
  { key: 'totalClicks', label: '总点击', type: 'number', direction: 'neutral' },
  { key: 'ctr', label: 'CTR', type: 'percent', direction: 'positive' },
  { key: 'cvr', label: 'CVR', type: 'percent', direction: 'positive' },
];

export const compareMetric = (current, previous, definition, currentSummary = {}, previousSummary = {}) => {
  const currentValue = number(current[definition.key]);
  const previousValue = number(previous[definition.key]);
  const delta = currentValue - previousValue;
  const percent = previousValue ? delta / Math.abs(previousValue) : (currentValue ? 1 : 0);
  const trend = Math.abs(delta) < 0.000001 ? '持平' : delta > 0 ? '上升' : '下降';
  let quality = 'neutral';
  if (definition.direction === 'positive') quality = trend === '上升' ? 'good' : trend === '下降' ? 'bad' : 'neutral';
  if (definition.direction === 'negative') quality = trend === '下降' ? 'good' : trend === '上升' ? 'bad' : 'neutral';
  if (definition.direction === 'cost') {
    const ordersStable = number(currentSummary.totalOrders) >= number(previousSummary.totalOrders) * 0.95;
    const profitStable = number(currentSummary.totalProfit) >= number(previousSummary.totalProfit) * 0.95;
    quality = trend === '下降' && ordersStable && profitStable ? 'good' : trend === '上升' ? 'bad' : 'neutral';
  }
  if (definition.direction === 'stock') quality = currentValue < 30 ? 'risk' : 'neutral';
  return { ...definition, current: currentValue, previous: previousValue, delta, percent, trend, quality };
};

export const buildPeriodComparison = (records = [], filters = {}) => {
  const currentRange = normalizeDateRange(records, filters);
  const previousRange = currentRange.isAllDates ? { startDate: '', endDate: '', dayCount: 0 } : getPreviousDateRange(currentRange.startDate, currentRange.endDate);
  const withSku = (range) => records.filter((record) => {
    const dateMatched = (!range.startDate || record.date >= range.startDate) && (!range.endDate || record.date <= range.endDate);
    const skuMatched = !filters.sku || record.sku === filters.sku;
    return dateMatched && skuMatched;
  });
  const currentRecords = withSku(currentRange);
  const previousRecords = withSku(previousRange);
  const currentSummary = summarizeRecords(currentRecords);
  const previousSummary = summarizeRecords(previousRecords);
  const metrics = METRIC_DEFINITIONS.map((definition) => compareMetric(currentSummary, previousSummary, definition, currentSummary, previousSummary));
  return {
    filters,
    currentRange: { ...currentRange, dayCount: countDaysInclusive(currentRange.startDate, currentRange.endDate) },
    previousRange,
    currentRecords,
    previousRecords,
    currentSummary,
    previousSummary,
    hasComparison: !currentRange.isAllDates && previousRecords.length > 0,
    metrics,
    skuRows: buildSkuComparison(records, currentRange, previousRange),
  };
};

const judgeSku = (current, previous) => {
  if (current.totalStock > 0 && current.totalStock < 30) return '库存风险';
  if (current.totalProfit < 0 && current.totalAdSpend > 0) return '建议暂停广告';
  if (current.totalAdSpend > previous.totalAdSpend * 1.2 && current.totalOrders <= previous.totalOrders) return '广告费上升但订单未增长';
  if (current.totalProfit < previous.totalProfit) return '利润下降';
  if (current.totalAdSpend < previous.totalAdSpend && current.totalOrders >= previous.totalOrders * 0.95 && current.totalProfit >= previous.totalProfit) return '控费有效';
  if (current.totalOrders > previous.totalOrders && current.totalProfit > previous.totalProfit && current.roi > 2 && current.totalStock >= 30) return '建议加预算';
  if (current.totalOrders > previous.totalOrders || current.totalProfit > previous.totalProfit) return '表现改善';
  return '需要观察';
};

export function buildSkuComparison(records = [], currentRange = {}, previousRange = {}) {
  const skus = getSkuOptions(records);
  return skus.map((sku) => {
    const inRange = (range) => records.filter((record) => record.sku === sku && (!range.startDate || record.date >= range.startDate) && (!range.endDate || record.date <= range.endDate));
    const currentRows = inRange(currentRange);
    const previousRows = inRange(previousRange);
    const current = summarizeRecords(currentRows);
    const previous = summarizeRecords(previousRows);
    const currentMinStock = currentRows.reduce((min, row) => { const stock = number(row.stock); return stock > 0 ? Math.min(min, stock) : min; }, Infinity);
    return {
      sku,
      currentOrders: current.totalOrders,
      previousOrders: previous.totalOrders,
      orderDelta: current.totalOrders - previous.totalOrders,
      currentAdSpend: current.totalAdSpend,
      previousAdSpend: previous.totalAdSpend,
      adSpendDelta: current.totalAdSpend - previous.totalAdSpend,
      currentProfit: current.totalProfit,
      previousProfit: previous.totalProfit,
      profitDelta: current.totalProfit - previous.totalProfit,
      currentRoi: current.roi,
      previousRoi: previous.roi,
      currentAcos: current.acos,
      previousAcos: previous.acos,
      currentStock: Number.isFinite(currentMinStock) ? currentMinStock : current.totalStock,
      judgment: judgeSku({ ...current, totalStock: Number.isFinite(currentMinStock) ? currentMinStock : current.totalStock }, previous),
    };
  }).filter((row) => row.currentOrders || row.previousOrders || row.currentAdSpend || row.previousAdSpend || row.currentProfit || row.previousProfit || row.currentStock);
}

export const buildPeriodSummaryText = (comparison) => {
  const { currentRange, previousRange, currentSummary: current, previousSummary: previous, hasComparison } = comparison;
  const prefix = `当前时间段 ${currentRange.startDate || '-'} 至 ${currentRange.endDate || '-'}`;
  if (currentRange.isAllDates) return `${prefix}，全部日期暂无上期对比。建议结合 SKU 明细定位长期趋势。`;
  if (!hasComparison) return `${prefix}，上一时间段 ${previousRange.startDate || '-'} 至 ${previousRange.endDate || '-'} 无对比数据。建议先积累完整历史后再判断趋势。`;
  const orderPct = previous.totalOrders ? (current.totalOrders - previous.totalOrders) / Math.abs(previous.totalOrders) : 0;
  const spendPct = previous.totalAdSpend ? (current.totalAdSpend - previous.totalAdSpend) / Math.abs(previous.totalAdSpend) : 0;
  const profitDelta = current.totalProfit - previous.totalProfit;
  if (spendPct < -0.05 && current.totalOrders >= previous.totalOrders * 0.95 && profitDelta > 0) return `${prefix}，较上一时间段订单${orderPct >= 0 ? '增加' : '下降'} ${Math.abs(orderPct * 100).toFixed(0)}%，广告费下降 ${Math.abs(spendPct * 100).toFixed(0)}%，利润提升，说明控费效果较好。建议继续保持当前预算策略。`;
  if (spendPct > 0.25 && orderPct < 0.1 && profitDelta <= 0) return `${prefix}，较上一时间段广告费增加 ${Math.abs(spendPct * 100).toFixed(0)}%，但订单仅${orderPct >= 0 ? '增加' : '下降'} ${Math.abs(orderPct * 100).toFixed(0)}%，利润下降，说明广告放量效率较差。建议重点检查高花费 SKU，并降低 ROI 差的广告预算。`;
  if (current.totalOrders > previous.totalOrders && current.totalProfit > previous.totalProfit && current.roi > 2 && current.totalStock >= 30) return `${prefix}，订单和利润均较上一时间段增长，ROI ${current.roi.toFixed(2)}，库存充足，可考虑对表现好的 SKU 适当加预算。`;
  if (current.totalStock > 0 && current.totalStock < 30) return `${prefix}，库存低于 30，存在库存风险，不建议继续放大广告，优先补货。`;
  return `${prefix}，较上一时间段订单变化 ${orderPct >= 0 ? '+' : ''}${(orderPct * 100).toFixed(0)}%，广告费变化 ${spendPct >= 0 ? '+' : ''}${(spendPct * 100).toFixed(0)}%，利润变化 ${profitDelta >= 0 ? '+' : ''}${profitDelta.toFixed(0)}。建议继续观察高花费和利润波动 SKU。`;
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
