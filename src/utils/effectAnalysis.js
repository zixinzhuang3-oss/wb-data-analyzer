const DAY_MS = 86400000;

const METRICS = {
  totalOrders: '订单量',
  impressions: '曝光',
  clicks: '点击',
  revenue: '销售额',
  adSpend: '广告费',
  adShare: '广告占比',
  profit: '利润',
  margin: '利润率',
  ctr: 'CTR',
  cvr: 'CVR',
  acos: 'ACOS',
  roi: 'ROI',
  stock: '库存',
  price: '价格',
  reviews: '评论数',
};

const ADVICE_TYPES = [
  '保持当前策略', '降低搜索出价', '降低推荐位预算', '暂停广告', '加大预算', '提高出价', '优化主图',
  '优化标题关键词', '降价或参加活动', '控制广告花费', '补货', '观察1天',
];

const toDate = (date) => new Date(`${date}T00:00:00Z`);
const toIsoDate = (date) => date.toISOString().slice(0, 10);
const addDays = (date, days) => toIsoDate(new Date(toDate(date).getTime() + days * DAY_MS));
const safeDivide = (a, b) => (b ? a / b : 0);
const number = (value) => Number(value) || 0;
const pct = (value) => `${((value || 0) * 100).toFixed(1)}%`;
const signedPct = (value) => `${value >= 0 ? '+' : ''}${((value || 0) * 100).toFixed(1)}%`;
const money = (value) => new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'CNY', maximumFractionDigits: 0 }).format(value || 0);

const enrichRecord = (record) => {
  const revenue = number(record.revenue);
  const adSpend = number(record.adSpend);
  const totalOrders = number(record.totalOrders);
  const clicks = number(record.clicks || record.adClicks);
  const impressions = number(record.impressions || record.adImpressions);
  return {
    ...record,
    totalOrders,
    revenue,
    adSpend,
    impressions,
    clicks,
    adShare: number(record.adShare) || safeDivide(number(record.adOrders), totalOrders) || safeDivide(adSpend, revenue),
    profit: number(record.profit),
    margin: safeDivide(number(record.profit), revenue),
    ctr: number(record.ctr) || safeDivide(clicks, impressions),
    cvr: number(record.conversionRate) || safeDivide(totalOrders, clicks),
    acos: safeDivide(adSpend, revenue),
    roi: safeDivide(revenue, adSpend),
    stock: number(record.stock),
    price: number(record.price),
    reviews: number(record.reviews),
  };
};

const averageRecords = (records) => {
  const enriched = records.map(enrichRecord);
  const count = enriched.length || 1;
  return Object.fromEntries(Object.keys(METRICS).map((key) => [key, enriched.reduce((sum, row) => sum + number(row[key]), 0) / count]));
};

const delta = (current, previous) => ({ value: current - previous, rate: previous ? (current - previous) / Math.abs(previous) : 0 });

const buildMetricSnapshot = (todayRecord, yesterdayRecord, last3, last7, before3, after3) => {
  const today = enrichRecord(todayRecord || {});
  const yesterday = enrichRecord(yesterdayRecord || {});
  return Object.fromEntries(Object.keys(METRICS).map((key) => [key, {
    label: METRICS[key],
    today: number(today[key]),
    yesterday: number(yesterday[key]),
    todayVsYesterday: delta(number(today[key]), number(yesterday[key])),
    last3Avg: number(last3[key]),
    last7Avg: number(last7[key]),
    before3Avg: number(before3[key]),
    after3Avg: number(after3[key]),
    actionWindowDelta: delta(number(after3[key]), number(before3[key])),
  }]));
};

const actionForDate = (actionsByKey, date, sku) => actionsByKey.get(`${date}__${sku}`);

const latestBidText = (action) => {
  if (!action) return '';
  const parts = [];
  if (action.searchBid !== '' && action.searchBid !== undefined) parts.push(`搜索出价 ${action.searchBid}`);
  if (action.recommendBid !== '' && action.recommendBid !== undefined) parts.push(`推荐位出价 ${action.recommendBid}`);
  if (action.dailyBudget !== '' && action.dailyBudget !== undefined) parts.push(`预算 ${action.dailyBudget}`);
  return parts.join('，');
};

const makeRecommendation = (type, reason, priority = '中') => ({
  type: ADVICE_TYPES.includes(type) ? type : '观察1天',
  reason,
  priority,
});

const analyzeRules = ({ sku, today, yesterday, metrics, latestAction, previousAction, records }) => {
  const recommendations = [];
  const risks = [];
  const effects = [];
  const spendDelta = metrics.adSpend.todayVsYesterday.rate;
  const orderDelta = metrics.totalOrders.todayVsYesterday.rate;
  const profitDelta = metrics.profit.todayVsYesterday.value;
  const ctrDelta = metrics.ctr.todayVsYesterday.rate;
  const cvr = metrics.cvr.today;
  const roi = metrics.roi.today;
  const profit = metrics.profit.today;
  const stock = metrics.stock.today;
  const exposure = today.impressions || today.adImpressions || 0;
  const clicks = today.clicks || today.adClicks || 0;
  const stableOrders = Math.abs(orderDelta) <= 0.1;
  const bidIncreased = latestAction && previousAction && (number(latestAction.searchBid) > number(previousAction.searchBid) || number(latestAction.recommendBid) > number(previousAction.recommendBid) || number(latestAction.dailyBudget) > number(previousAction.dailyBudget));
  const budgetReduced = latestAction?.budgetAction === '降低预算' || (latestAction && previousAction && number(latestAction.dailyBudget) < number(previousAction.dailyBudget));
  const recommendRaised = latestAction?.adType === 'CPM推荐' && bidIncreased;

  if (bidIncreased && spendDelta > 0.1 && orderDelta <= 0.05 && profitDelta < 0) {
    effects.push({ level: '差', text: `${sku} 提高出价后广告费${signedPct(spendDelta)}，订单未增长，利润下降 ${money(Math.abs(profitDelta))}，动作效果差。` });
    recommendations.push(makeRecommendation('控制广告花费', `${sku} 提高出价后没有换来订单增长，建议明天控制广告花费，回调低效出价并观察 1 天。`, '高'));
  }

  if (budgetReduced && stableOrders && profitDelta > 0) {
    effects.push({ level: '好', text: `${sku} 降低预算后订单保持稳定，利润提升 ${money(profitDelta)}，说明控费有效。` });
    recommendations.push(makeRecommendation('保持当前策略', `${sku} ${latestAction?.adType || '广告'} 预算调整后订单稳定、利润提升，建议明天保持当前预算和出价，继续观察 1 天。`, '中'));
  }

  if (recommendRaised && (metrics.impressions.todayVsYesterday.rate > 0.1 || metrics.adSpend.todayVsYesterday.rate > 0.1) && ctrDelta < -0.1) {
    effects.push({ level: '差', text: `${sku} 提高推荐位预算后曝光上升但 CTR ${signedPct(ctrDelta)}，推荐流量质量偏差。` });
    recommendations.push(makeRecommendation('降低推荐位预算', `${sku} 推荐流量放量后点击率下降，建议明天降低推荐位预算，把预算转向搜索或高转化入口。`, '高'));
  }

  if ((latestAction?.adType === 'CPM搜索' || latestAction?.adType === 'CPC') && metrics.adSpend.today > metrics.adSpend.last3Avg && cvr < 0.02) {
    recommendations.push(makeRecommendation('降低搜索出价', `${sku} 搜索广告花费高于近 3 天均值，但 CVR 仅 ${pct(cvr)}，建议降低搜索出价或优化关键词。`, '高'));
  }

  if (exposure > 0 && clicks === 0) {
    risks.push(`${sku} 有曝光但没有点击，主图、标题或价格吸引力不足。`);
    recommendations.push(makeRecommendation('优化主图', `${sku} 今天有曝光但无点击，建议明天优先优化主图，同时检查标题和价格。`, '高'));
  }

  if (clicks > 0 && metrics.totalOrders.today === 0) {
    risks.push(`${sku} 有点击但没有转化，价格、评价或详情页可能阻碍下单。`);
    recommendations.push(makeRecommendation('降价或参加活动', `${sku} 今天有点击但无订单，建议优化价格、评价和详情页，可测试降价或参加活动。`, '高'));
  }

  const last3Positive = records.slice(-3).every((record) => enrichRecord(record).roi > 1 && enrichRecord(record).profit > 0);
  if (last3Positive && stock > 20) {
    recommendations.push(makeRecommendation('加大预算', `${sku} ROI 和利润近 3 天持续为正，且库存 ${stock} 充足，建议明天适当加大预算。`, '中'));
  }

  if (stock > 0 && stock < Math.max(5, metrics.totalOrders.last3Avg * 2)) {
    risks.push(`${sku} 库存 ${stock} 偏低，不适合继续放大广告。`);
    recommendations.push(makeRecommendation('补货', `${sku} 库存不足，建议不要继续加大广告，优先补货。`, '高'));
  }

  const last3NegativeProfit = records.slice(-3).length >= 3 && records.slice(-3).every((record) => enrichRecord(record).profit < 0);
  if (last3NegativeProfit) {
    risks.push(`${sku} 利润连续 3 天为负。`);
    recommendations.push(makeRecommendation('暂停广告', `${sku} 利润连续为负，建议明天降低广告预算或暂停广告，先排查成本和转化问题。`, '高'));
  }

  if (!recommendations.length) {
    recommendations.push(makeRecommendation('观察1天', `${sku} 当前数据未触发强干预规则，建议明天保持关键出价不变，继续观察 1 天。`, '低'));
  }

  const priorityWeight = { 高: 3, 中: 2, 低: 1 };
  const primary = recommendations.sort((a, b) => priorityWeight[b.priority] - priorityWeight[a.priority])[0];
  const actionText = latestBidText(latestAction);
  primary.reason = actionText ? `${sku} 最近动作：${actionText}。${primary.reason}` : primary.reason;

  return { recommendations, risks, effects };
};

export const buildEffectAnalysis = (records = [], actions = [], filters = {}) => {
  const actionsByKey = new Map(actions.map((action) => [action.uniqueKey, action]));
  const bySku = new Map();
  records.forEach((record) => {
    if (!record.date || !record.sku) return;
    if (filters.sku && record.sku !== filters.sku) return;
    if (!bySku.has(record.sku)) bySku.set(record.sku, []);
    bySku.get(record.sku).push(record);
  });

  const analyses = [...bySku.entries()].map(([sku, skuRecords]) => {
    const sorted = skuRecords.sort((a, b) => a.date.localeCompare(b.date));
    const targetDate = filters.date || sorted.at(-1)?.date;
    const today = sorted.find((record) => record.date === targetDate) || (!filters.date ? sorted.at(-1) : null);
    if (!today) return null;
    const todayIndex = sorted.findIndex((record) => record.uniqueKey === today.uniqueKey);
    const yesterdayDate = addDays(today.date, -1);
    const yesterday = sorted.find((record) => record.date === yesterdayDate) || sorted[todayIndex - 1] || {};
    const last3Rows = sorted.slice(Math.max(0, todayIndex - 2), todayIndex + 1);
    const last7Rows = sorted.slice(Math.max(0, todayIndex - 6), todayIndex + 1);
    const skuActions = actions.filter((action) => action.sku === sku && action.date <= today.date).sort((a, b) => a.date.localeCompare(b.date));
    const latestAction = actionForDate(actionsByKey, today.date, sku) || skuActions.at(-1);
    const actionDate = latestAction?.date || today.date;
    const before3Rows = sorted.filter((record) => record.date < actionDate).slice(-3);
    const after3Rows = sorted.filter((record) => record.date >= actionDate).slice(0, 3);
    const previousAction = actions.filter((action) => action.sku === sku && action.date < actionDate).sort((a, b) => a.date.localeCompare(b.date)).at(-1);
    const metrics = buildMetricSnapshot(today, yesterday, averageRecords(last3Rows), averageRecords(last7Rows), averageRecords(before3Rows), averageRecords(after3Rows));
    const ruleResult = analyzeRules({ sku, today: enrichRecord(today), yesterday: enrichRecord(yesterday), metrics, latestAction, previousAction, records: sorted.slice(0, todayIndex + 1) });

    return {
      sku,
      date: today.date,
      uniqueKey: today.uniqueKey,
      latestAction,
      previousAction,
      metrics,
      primaryRecommendation: ruleResult.recommendations[0],
      recommendations: ruleResult.recommendations,
      risks: ruleResult.risks,
      effects: ruleResult.effects,
    };
  });

  const priorityWeight = { 高: 3, 中: 2, 低: 1 };
  return analyses.filter(Boolean).sort((a, b) => priorityWeight[b.primaryRecommendation.priority] - priorityWeight[a.primaryRecommendation.priority] || a.sku.localeCompare(b.sku));
};

export const metricLabels = METRICS;
