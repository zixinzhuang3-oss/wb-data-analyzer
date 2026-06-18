const safeDivide = (a, b) => (b ? a / b : 0);
const keyOf = (row) => row.link || row.sku || '未命名链接';

const sumBy = (rows, field) => rows.reduce((sum, row) => sum + (Number(row[field]) || 0), 0);

const mergeRows = (salesRows, adsRows, linkRows) => {
  const map = new Map();
  const ensure = (row) => {
    const key = keyOf(row);
    if (!map.has(key)) {
      map.set(key, {
        key,
        sku: row.sku || '',
        link: row.link || '',
        orders: 0,
        units: 0,
        revenue: 0,
        productCost: 0,
        shippingCost: 0,
        fee: 0,
        spend: 0,
        impressions: 0,
        clicks: 0,
        adOrders: 0,
        adRevenue: 0,
        visitors: 0,
        sessions: 0,
      });
    }
    const item = map.get(key);
    item.sku ||= row.sku;
    item.link ||= row.link;
    return item;
  };

  salesRows.forEach((row) => {
    const item = ensure(row);
    ['orders', 'units', 'revenue', 'productCost', 'shippingCost', 'fee'].forEach((field) => {
      item[field] += row[field] || 0;
    });
  });

  adsRows.forEach((row) => {
    const item = ensure(row);
    ['spend', 'impressions', 'clicks', 'adOrders', 'adRevenue'].forEach((field) => {
      item[field] += row[field] || 0;
    });
  });

  linkRows.forEach((row) => {
    const item = ensure(row);
    ['visitors', 'sessions'].forEach((field) => {
      item[field] += row[field] || 0;
    });
  });

  return [...map.values()].map((item) => {
    const totalCost = item.productCost + item.shippingCost + item.fee + item.spend;
    const profit = item.revenue - totalCost;
    return {
      ...item,
      totalCost,
      profit,
      margin: safeDivide(profit, item.revenue),
      roi: safeDivide(item.revenue, item.spend),
      acos: safeDivide(item.spend, item.adRevenue || item.revenue),
      ctr: safeDivide(item.clicks, item.impressions),
      cpc: safeDivide(item.spend, item.clicks),
      cvr: safeDivide(item.orders || item.adOrders, item.clicks || item.visitors),
      visitorValue: safeDivide(item.revenue, item.visitors),
    };
  });
};

const buildRecommendations = (items, totals) => {
  const recommendations = [];
  const lossItems = items.filter((item) => item.profit < 0).sort((a, b) => a.profit - b.profit).slice(0, 3);
  lossItems.forEach((item) => {
    recommendations.push({
      type: '止损优化',
      target: item.key,
      priority: '高',
      action: `该链接亏损 ${formatMoney(Math.abs(item.profit))}，建议先检查售价、物流/平台费和广告预算；若短期无法提升转化，降低出价或暂停低效广告。`,
    });
  });

  const highAcos = items.filter((item) => item.spend > 0 && item.acos > 0.35).sort((a, b) => b.acos - a.acos).slice(0, 3);
  highAcos.forEach((item) => {
    recommendations.push({
      type: '广告降本',
      target: item.key,
      priority: item.acos > 0.5 ? '高' : '中',
      action: `ACOS 为 ${formatPercent(item.acos)}，高于 35% 警戒线；建议拆分关键词，否词无转化词，保留高转化词并下调泛匹配预算。`,
    });
  });

  const scaleItems = items.filter((item) => item.profit > 0 && item.margin > 0.2 && item.roi > 3).sort((a, b) => b.profit - a.profit).slice(0, 3);
  scaleItems.forEach((item) => {
    recommendations.push({
      type: '放量增长',
      target: item.key,
      priority: '中',
      action: `利润率 ${formatPercent(item.margin)} 且 ROI ${item.roi.toFixed(2)}，可逐步增加 10%-20% 广告预算，同时保障库存和客服响应。`,
    });
  });

  if (totals.margin < 0.15) {
    recommendations.push({
      type: '利润结构',
      target: '全店',
      priority: '高',
      action: `整体利润率仅 ${formatPercent(totals.margin)}，建议优先压缩物流/采购成本、复核平台费，并提高低毛利 SKU 售价或减少促销。`,
    });
  }

  return recommendations.slice(0, 8);
};

export const analyzeBusiness = ({ salesRows = [], adsRows = [], linkRows = [] }) => {
  const items = mergeRows(salesRows, adsRows, linkRows).sort((a, b) => b.revenue - a.revenue);
  const totals = {
    revenue: sumBy(items, 'revenue'),
    productCost: sumBy(items, 'productCost'),
    shippingCost: sumBy(items, 'shippingCost'),
    fee: sumBy(items, 'fee'),
    spend: sumBy(items, 'spend'),
    orders: sumBy(items, 'orders'),
    units: sumBy(items, 'units'),
    clicks: sumBy(items, 'clicks'),
    impressions: sumBy(items, 'impressions'),
    visitors: sumBy(items, 'visitors'),
  };
  totals.totalCost = totals.productCost + totals.shippingCost + totals.fee + totals.spend;
  totals.profit = totals.revenue - totals.totalCost;
  totals.margin = safeDivide(totals.profit, totals.revenue);
  totals.roi = safeDivide(totals.revenue, totals.spend);
  totals.acos = safeDivide(totals.spend, totals.revenue);
  totals.ctr = safeDivide(totals.clicks, totals.impressions);
  totals.cvr = safeDivide(totals.orders, totals.clicks || totals.visitors);

  return { totals, items, recommendations: buildRecommendations(items, totals) };
};

export const formatMoney = (value) =>
  new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY', maximumFractionDigits: 2 }).format(value || 0);

export const formatPercent = (value) => `${((value || 0) * 100).toFixed(1)}%`;
