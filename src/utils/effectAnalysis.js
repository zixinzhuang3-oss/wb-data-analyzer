import { toProfitRub } from './currency.js';
import { hasValidBusinessData } from './excel.js';
import { addDays as addDateDays, normalizeDateKey } from './date.js';
import { ACTION_HISTORY_DAYS, ACTION_LOOKBACK_DAYS, actionToSummary, buildActionKey, CPM_RECOMMEND_MIN_BID, CPM_SEARCH_MIN_BID, findRecentAction, getEffectiveAction, getSkuActionTimeline, normalizeAction, normalizeSku } from './actions.js';

const METRICS = {
  totalOrders: '订单',
  revenue: '销售额',
  adSpend: '广告费',
  profit: '利润',
  margin: '利润率',
  roi: 'ROI',
  acos: 'ACOS',
  impressions: '曝光',
  clicks: '点击',
  ctr: 'CTR',
  cvr: 'CVR',
  adImpressions: '广告曝光',
  adClicks: '广告点击',
  adCtr: '广告 CTR',
  adOrders: '广告订单',
  adShare: '广告费占比',
  stock: '库存',
  price: '价格',
  reviews: '评论数',
};

const ADVICE_TYPES = [
  '保持当前策略', '降低搜索出价', '降低推荐位预算', '暂停广告', '加大预算', '提高出价', '优化主图',
  '暂停搜索位',
  '优化标题关键词', '降价或参加活动', '控制广告花费', '补货', '观察1天', '恢复小预算搜索广告', '暂停推荐位', '保留CPC暂停CPM推荐', '保留CPM推荐', '关闭CPM推荐',
];

const safeDivide = (a, b) => (b ? a / b : 0);
const number = (value) => Number(value) || 0;
const pct = (value) => `${((value || 0) * 100).toFixed(1)}%`;
const signedPct = (value) => `${value >= 0 ? '+' : ''}${((value || 0) * 100).toFixed(1)}%`;
const money = (value) => new Intl.NumberFormat('zh-CN', { style: 'currency', currency: 'RUB', maximumFractionDigits: 0 }).format(value || 0);

const reducedBid = (current, min, pct = 0.15) => Math.max(min, Math.round(current * (1 - pct)));
const cpmMinBidHint = `CPM 搜索出价最低：${CPM_SEARCH_MIN_BID}；CPM 推荐出价最低：${CPM_RECOMMEND_MIN_BID}`;

const cpmBidContext = (action = {}) => {
  const position = action?.cpmPosition || action?.adPosition || '';
  const bidType = action?.cpmBidType || '手动出价';
  const hasSearch = ['仅搜索', '搜索', '搜索+推荐'].includes(position);
  const hasRecommend = ['仅推荐', '推荐', '搜索+推荐'].includes(position);
  const unified = bidType === '统一出价';
  const unifiedBid = number(action?.cpmUnifiedBid);
  return {
    position,
    bidType,
    unified,
    hasSearch,
    hasRecommend,
    searchBid: unified ? unifiedBid : number(action?.cpmSearchBid || action?.searchBid),
    recommendBid: unified ? unifiedBid : number(action?.cpmRecommendBid || action?.recommendBid),
    unifiedBid,
    unifiedMinBid: hasSearch ? CPM_SEARCH_MIN_BID : CPM_RECOMMEND_MIN_BID,
  };
};

const cpmSearchCostAdvice = (sku, ctx) => {
  if (ctx.unified) {
    if (ctx.unifiedBid > ctx.unifiedMinBid) return makeRecommendation('降低搜索出价', `${sku} CPM 使用统一出价且覆盖搜索位，当前统一出价 ${ctx.unifiedBid}，表现差时可下调 10%–20%，建议先降至 ${reducedBid(ctx.unifiedBid, ctx.unifiedMinBid)}，最低不能低于 ${ctx.unifiedMinBid}。${cpmMinBidHint}。`, '高');
    return makeRecommendation('控制广告花费', `${sku} 当前 CPM 统一出价已是搜索最低 ${CPM_SEARCH_MIN_BID}，无法继续降低统一出价。建议降低每日预算或观察 1 天，如仍亏损则暂停搜索位，并优化主图、价格、关键词、评价。${cpmMinBidHint}。`, '高');
  }
  if (ctx.searchBid > CPM_SEARCH_MIN_BID) return makeRecommendation('降低搜索出价', `${sku} CPM 搜索出价 ${ctx.searchBid} 高于最低 ${CPM_SEARCH_MIN_BID}，表现差可降低 10%–20%，建议先降至 ${reducedBid(ctx.searchBid, CPM_SEARCH_MIN_BID)}，但不能低于 ${CPM_SEARCH_MIN_BID}。${cpmMinBidHint}。`, '高');
  return makeRecommendation('控制广告花费', `${sku} 当前 CPM 搜索出价已是最低 ${CPM_SEARCH_MIN_BID}，无法继续降低出价。建议保持最低搜索出价观察 1 天，优先降低每日预算；如仍亏损则暂停搜索位，并优化主图、价格、关键词、评价；若搜索有订单但利润差，优先控制预算而不是降低出价。${cpmMinBidHint}。`, '高');
};

const cpmRecommendCostAdvice = (sku, ctx) => {
  const minBid = ctx.unified ? ctx.unifiedMinBid : CPM_RECOMMEND_MIN_BID;
  const bid = ctx.unified ? ctx.unifiedBid : ctx.recommendBid;
  if (bid > minBid) return makeRecommendation('降低推荐位预算', `${sku} CPM 推荐出价 ${bid} 高于最低 ${minBid}，表现差可降低 10%–20%，建议先降至 ${reducedBid(bid, minBid)}，但不能低于 ${minBid}。${cpmMinBidHint}。`, '高');
  return makeRecommendation('暂停推荐位', `${sku} 当前 CPM 推荐出价已是最低 ${CPM_RECOMMEND_MIN_BID}，无法继续降低出价。建议保持最低推荐出价观察 1 天，降低推荐预算或暂停推荐位，只保留搜索位观察；若推荐曝光高但订单差，优先暂停推荐而不是继续降出价。${cpmMinBidHint}。`, '高');
};

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
    adImpressions: number(record.adImpressions) || impressions,
    adClicks: number(record.adClicks) || clicks,
    adCtr: number(record.adCtr) || safeDivide(number(record.adClicks) || clicks, number(record.adImpressions) || impressions),
    adOrders: number(record.adOrders),
    adShare: number(record.adShare) || safeDivide(adSpend, revenue),
    profit: toProfitRub(record),
    profitRub: toProfitRub(record),
    margin: safeDivide(toProfitRub(record), revenue),
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

const summarizeChange = (after = {}, before = {}) => {
  const parts = [];
  ['totalOrders', 'profit', 'margin', 'adSpend', 'roi'].forEach((key) => {
    const d = delta(number(after[key]), number(before[key]));
    if (Math.abs(d.rate) >= 0.01 || Math.abs(d.value) > 0) parts.push(`${METRICS[key]}${signedPct(d.rate)}`);
  });
  return parts.length ? parts.join('，') : '暂无明显变化';
};

const buildActionWindows = (sorted, actionDate) => {
  const before = (days) => sorted.filter((record) => record.date < actionDate).slice(-days);
  const afterThrough = (days) => sorted.filter((record) => record.date > actionDate && record.date <= addDateDays(actionDate, days));
  const avg = (rows) => rows.length ? averageRecords(rows) : null;
  const before3 = avg(before(3));
  const before7 = avg(before(7));
  const make = (days) => {
    const rows = afterThrough(days);
    const after = avg(rows);
    return { days, hasData: rows.length > 0, recordCount: rows.length, summary: after && before3 ? summarizeChange(after, before3) : '暂无后续数据', metrics: after };
  };
  return { before3Avg: before3, before7Avg: before7, after1: make(1), after3: make(3), after7: make(7) };
};

const containsText = (action, pattern) => pattern.test([action?.budgetAction, action?.priceAction, action?.note, action?.rawOperationAction, action?.cpmNote, action?.cpcNote, action?.cpmPosition, action?.adStatus].join(' '));

const buildActionJudgement = (action, windows, nearby = []) => {
  const before = windows.before3Avg || windows.before7Avg || {};
  const after = windows.after3.metrics || windows.after1.metrics || windows.after7.metrics || {};
  const order = delta(number(after.totalOrders), number(before.totalOrders));
  const profit = delta(number(after.profit), number(before.profit));
  const margin = delta(number(after.margin), number(before.margin));
  const spend = delta(number(after.adSpend), number(before.adSpend));
  const roi = delta(number(after.roi), number(before.roi));
  const stableOrders = Math.abs(order.rate) <= 0.1;
  const messages = [];
  if (nearby.length >= 2 && nearby.some((a, i) => i && Math.abs((new Date(a.date) - new Date(nearby[i - 1].date)) / 86400000) <= 2)) messages.push('多个动作叠加影响，不能单独归因。');
  if (!windows.after1.hasData && !windows.after3.hasData) messages.push(`${action.date === windows.analysisDate ? '今日刚' : '动作后'}${action.priceAction || '调整'}，暂无法判断效果。建议明天观察点击转订单率、销售额和利润率变化。`);
  if (action.priceAction === '涨价') {
    if (order.rate < -0.05 && margin.value > 0 && profit.value >= 0) messages.push('涨价后订单下降，但利润率提升，说明涨价提高了单件收益，但可能影响转化。建议继续观察或小幅回调价格。');
    else if (order.rate < -0.05 && profit.value < 0) messages.push('涨价后订单和利润同时下降，说明价格可能影响成交，建议考虑恢复原价或参加活动。');
    else messages.push('涨价可能提高利润率但影响转化，建议结合 3 天与 7 天窗口继续观察。');
  }
  if (action.priceAction === '降价') {
    if (order.rate > 0.05 && margin.value < 0) messages.push('降价带动订单增长，但压缩利润率。若总利润提升，可以继续；若总利润下降，不建议继续降价。');
    else messages.push('降价可能拉动订单但压缩利润，需重点观察总利润是否同步提升。');
  }
  if (containsText(action, /暂停推荐|关闭CPM推荐|暂停.*推荐/)) {
    if (spend.value < 0 && stableOrders && profit.value > 0) messages.push('暂停推荐可能有效：广告费下降、订单稳定、利润提升。');
  }
  if (action.budgetAction === '降低预算' && spend.value < 0 && stableOrders && roi.value > 0) messages.push('降低预算后广告费下降、订单稳定、ROI 提升，判断控费有效。');
  if (action.budgetAction === '恢复广告') {
    if (order.value > 0 && profit.value > 0) messages.push('恢复广告后订单和利润增长，判断恢复广告有效。');
    else if (order.value > 0 && profit.value < 0) messages.push('恢复广告带来订单，但成本压力变大，建议降低预算或只保留高效位置。');
  }
  if (action.budgetAction === '暂停广告' || action.adStatus === '无广告') {
    if (stableOrders && profit.value > 0) messages.push('广告关闭后自然订单稳定、利润提升，建议继续观察，不急于恢复广告。');
    else if (order.rate < -0.15) messages.push('广告关闭后订单明显下降，建议恢复低预算 CPC 搜索或 CPM 搜索测试，不建议直接大预算投推荐。');
  }
  return messages.length ? messages.join(' ') : '暂无明确动作效果，建议继续观察。';
};

export const buildSkuActionHistory = (records = [], actions = [], sku = '', endDate = '', days = ACTION_HISTORY_DAYS) => {
  const recordList = Array.isArray(records) ? records : [];
  const actionList = Array.isArray(actions) ? actions : [];
  const targetDate = normalizeDateKey(endDate || recordList.map((r) => normalizeDateKey(r.date)).filter(Boolean).sort().at(-1) || '');
  const fromDate = targetDate ? addDateDays(targetDate, -(days - 1)) : '';
  const normalizedRecords = recordList.map((r) => ({ ...r, date: normalizeDateKey(r.date), sku: normalizeSku(r.sku) })).filter((r) => r.date && (!sku || r.sku === normalizeSku(sku))).sort((a, b) => a.date.localeCompare(b.date));
  const explicit = getSkuActionTimeline(actionList, sku, { fromDate, toDate: targetDate }).map((action) => ({ type: 'explicit', action }));
  const explicitDates = new Set(explicit.map((row) => row.action.date));
  const inherited = normalizedRecords
    .filter((record) => record.date >= fromDate && record.date <= targetDate && !explicitDates.has(record.date))
    .map((record) => ({ type: 'inherited', effective: getEffectiveAction(actionList, record.date, sku), record }))
    .filter((row) => row.effective.found && row.effective.isInherited)
    .map((row) => ({ type: 'inherited', action: row.effective.action, sourceActionDate: row.effective.sourceActionDate }));
  return [...explicit, ...inherited].map((entry) => {
    const action = entry.action;
    const windows = buildActionWindows(normalizedRecords, action.date);
    const nearby = getSkuActionTimeline(actionList, action.sku, { fromDate: addDateDays(action.date, -2), toDate: addDateDays(action.date, 2) });
    return { action, date: action.date, sku: action.sku, type: entry.type, sourceActionDate: entry.sourceActionDate || action.sourceActionDate || action.date, summary: entry.type === 'inherited' ? `继承 ${entry.sourceActionDate || action.sourceActionDate}` : actionToSummary(action), windows, judgement: entry.type === 'inherited' ? `自动继承动作：沿用 ${entry.sourceActionDate || action.sourceActionDate}。` : buildActionJudgement(action, windows, nearby) };
  }).sort((a, b) => b.date.localeCompare(a.date));
};

const noDataRecommendation = (sku, hasRecentAction) => makeRecommendation('观察1天', hasRecentAction
  ? `${sku} 已找到最近动作记录，但当前日期暂无有效数据，暂时无法判断动作效果。`
  : `${sku} 当前时间段暂无有效数据，无法生成策略建议。`, '低');

const latestBidText = (action) => {
  if (!action) return '';
  const parts = [];
  if (action.adStatus) parts.push(`整体广告状态 ${action.adStatus}`);
  if (action.cpcEnabled) parts.push(`CPC ${action.cpcEnabled}`);
  if (action.cpmEnabled) parts.push(`CPM ${action.cpmEnabled}`);
  if (action.cpmPosition) parts.push(`CPM位置 ${action.cpmPosition}`);
  if (action.cpmBidType) parts.push(`出价方式 ${action.cpmBidType}`);
  if (action.cpcSearchBid !== '' && action.cpcSearchBid !== undefined) parts.push(`CPC搜索出价 ${action.cpcSearchBid}`);
  if (action.cpmSearchBid !== '' && action.cpmSearchBid !== undefined) parts.push(`CPM搜索出价 ${action.cpmSearchBid}`);
  if (action.cpmRecommendBid !== '' && action.cpmRecommendBid !== undefined) parts.push(`CPM推荐出价 ${action.cpmRecommendBid}`);
  if (action.cpmUnifiedBid !== '' && action.cpmUnifiedBid !== undefined) parts.push(`统一出价 ${action.cpmUnifiedBid}`);
  if (action.searchBid !== '' && action.searchBid !== undefined) parts.push(`搜索出价 ${action.searchBid}`);
  if (action.recommendBid !== '' && action.recommendBid !== undefined) parts.push(`推荐位出价 ${action.recommendBid}`);
  if (action.cpcDailyBudget !== '' && action.cpcDailyBudget !== undefined) parts.push(`CPC预算 ${action.cpcDailyBudget}`);
  if (action.cpmDailyBudget !== '' && action.cpmDailyBudget !== undefined) parts.push(`CPM预算 ${action.cpmDailyBudget}`);
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
  const cpcOn = latestAction?.cpcEnabled === '开启' || latestAction?.adMode === 'CPC' || latestAction?.adStatus === '仅 CPC' || latestAction?.adStatus === 'CPC+CPM';
  const cpmOn = latestAction?.cpmEnabled === '开启' || latestAction?.adMode === 'CPM' || latestAction?.adStatus === '仅 CPM' || latestAction?.adStatus === 'CPC+CPM';
  const cpmPosition = latestAction?.cpmPosition || latestAction?.adPosition || '整体广告';
  const hasCpmSearch = cpmOn && ['仅搜索', '搜索', '搜索+推荐'].includes(cpmPosition);
  const hasCpmRecommend = cpmOn && ['仅推荐', '推荐', '搜索+推荐'].includes(cpmPosition);
  const cpcSearchBid = (action) => number(action?.cpcSearchBid || action?.searchBid);
  const cpmSearchBid = (action) => number(action?.cpmSearchBid || action?.cpmUnifiedBid || action?.searchBid);
  const recommendBid = (action) => number(action?.cpmRecommendBid || action?.recommendBid || action?.cpmUnifiedBid);
  const totalBudget = (action) => number(action?.cpcDailyBudget) + number(action?.cpmDailyBudget || action?.dailyBudget);
  const bidIncreased = latestAction && previousAction && (cpcSearchBid(latestAction) > cpcSearchBid(previousAction) || cpmSearchBid(latestAction) > cpmSearchBid(previousAction) || recommendBid(latestAction) > recommendBid(previousAction) || totalBudget(latestAction) > totalBudget(previousAction));
  const budgetReduced = latestAction?.budgetAction === '降低预算' || (latestAction && previousAction && totalBudget(latestAction) < totalBudget(previousAction));
  const adStatus = latestAction?.adStatus || today.adStatus || (today.adSpend > 0 || today.adClicks > 0 || today.adImpressions > 0 ? '仅 CPM' : '无广告');
  const adMode = latestAction?.adMode || adStatus;
  const recommendRaised = hasCpmRecommend && bidIncreased;
  const cpmCtx = cpmBidContext(latestAction);
  const bothCpmAtMin = hasCpmSearch && hasCpmRecommend && ((cpmCtx.unified && cpmCtx.unifiedBid <= CPM_SEARCH_MIN_BID) || (!cpmCtx.unified && cpmCtx.searchBid <= CPM_SEARCH_MIN_BID && cpmCtx.recommendBid <= CPM_RECOMMEND_MIN_BID));
  const cpmPoor = profit < 0 || (spendDelta > 0.1 && orderDelta <= 0.05) || (metrics.totalOrders.today <= 1 && metrics.adSpend.today > metrics.adSpend.last3Avg);

  if (bothCpmAtMin && cpmPoor) {
    recommendations.push(makeRecommendation('控制广告花费', `${sku} CPM 搜索和推荐均已在最低出价（搜索 ${CPM_SEARCH_MIN_BID}、推荐 ${CPM_RECOMMEND_MIN_BID}），广告表现差时不要继续降出价。建议降低整体预算，或暂停表现更差的位置。${cpmMinBidHint}。`, '高'));
  }

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

  if (cpcOn && cpmOn) {
    effects.push({ level: '观察', text: `${sku} 同时开启 CPC 和 CPM，应拆分观察 CPC搜索、CPM搜索、CPM推荐，以及整体广告投入与利润。` });
    if (profit < 0 && spendDelta > 0.1 && orderDelta <= 0.05) recommendations.push(makeRecommendation('控制广告花费', `${sku} CPC 和 CPM 都花费高、订单没有增长且利润下降，建议整体降低预算，优先保留转化更好的广告位置。`, '高'));
    if (roi > 2 && profit > 0 && hasCpmRecommend && ctrDelta < -0.1) recommendations.push(makeRecommendation('保留CPC暂停CPM推荐', `${sku} CPC 搜索表现较稳但 CPM 推荐点击率走弱，建议保留 CPC 搜索，降低或暂停 CPM 推荐。`, '高'));
    if (clicks >= 10 && cvr < 0.02 && hasCpmRecommend && metrics.totalOrders.today > 0 && metrics.adSpend.today <= metrics.adSpend.last3Avg) recommendations.push(makeRecommendation('保留CPM推荐', `${sku} CPC 搜索点击多但转化偏差，CPM 推荐花费较低且有订单，建议降低 CPC 搜索出价，保留 CPM 推荐继续观察。`, '高'));
  }

  if (cpcOn && clicks >= 10 && metrics.totalOrders.today <= 1) recommendations.push(makeRecommendation('降低搜索出价', `${sku} CPC 搜索点击多但订单少，建议降低 CPC 搜索出价，或优化关键词/主图/价格。`, '高'));
  if (cpcOn && metrics.adSpend.today > metrics.adSpend.last3Avg && profit < 0) recommendations.push(makeRecommendation('暂停广告', `${sku} CPC 花费高且利润为负，建议降低 CPC 出价或暂停 CPC 搜索广告。`, '高'));
  if (cpcOn && roi > 2 && profit > 0 && stock > 20) recommendations.push(makeRecommendation('加大预算', `${sku} CPC 搜索 ROI ${roi.toFixed(2)} 且利润为正，库存充足，可小幅提高搜索出价或预算。`, '中'));

  if (hasCpmSearch) {
    if (exposure > 1000 && metrics.ctr.today < 0.01) recommendations.push(makeRecommendation('优化主图', `${sku} CPM搜索曝光高但 CTR 仅 ${pct(metrics.ctr.today)}，建议优化主图、标题和价格。`, '高'));
    if (metrics.ctr.today >= 0.01 && cvr < 0.02) recommendations.push(makeRecommendation('降价或参加活动', `${sku} CPM搜索点击率正常但转化率低，建议优化价格、评价和详情页。`, '中'));
    if (spendDelta > 0.1 && orderDelta <= 0.05) recommendations.push(cpmSearchCostAdvice(sku, cpmCtx));
  }

  if (hasCpmRecommend) {
    if (metrics.impressions.todayVsYesterday.rate > 0.1 && ctrDelta < -0.1) {
      risks.push(`${sku} 推荐流量质量差。`);
      recommendations.push(cpmRecommendCostAdvice(sku, cpmCtx));
    }
    if (metrics.adSpend.today > metrics.adSpend.last3Avg && metrics.totalOrders.today <= 1) recommendations.push(makeRecommendation('暂停推荐位', `${sku} CPM推荐花费高、订单少，建议暂停推荐位，只保留搜索位观察。`, '高'));
    if (roi > 2 && profit > 0 && metrics.totalOrders.today > 0) recommendations.push(makeRecommendation('保持当前策略', `${sku} CPM推荐带来订单且 ROI 为正，可保留推荐位预算。`, '中'));
  }

  if (!cpcOn && hasCpmRecommend && !hasCpmSearch) {
    if (exposure > 1000 && metrics.ctr.today < 0.01) recommendations.push(makeRecommendation('暂停推荐位', `${sku} 只开 CPM 推荐但曝光高、点击率低。${cpmCtx.recommendBid <= CPM_RECOMMEND_MIN_BID ? `当前 CPM 推荐出价已是最低 ${CPM_RECOMMEND_MIN_BID}，无法继续降低出价，建议暂停推荐或降低推荐预算，并考虑恢复低预算搜索广告。${cpmMinBidHint}。` : `建议降低推荐出价至不低于 ${CPM_RECOMMEND_MIN_BID} 或暂停推荐，并考虑恢复低预算搜索广告。${cpmMinBidHint}。`}`, '高'));
    if (metrics.totalOrders.today > 0 && roi > 1 && profit > 0) recommendations.push(makeRecommendation('保持当前策略', `${sku} 只开 CPM 推荐且订单稳定、ROI 为正，建议继续保留推荐位，暂不加大预算，观察 1 天。`, '中'));
  }

  if ((!cpcOn && !cpmOn) || adStatus === '无广告' || adStatus === '关闭' || adStatus === '无广告数据') {
    if (stock > 0 && stock < Math.max(5, metrics.totalOrders.last3Avg * 2)) recommendations.push(makeRecommendation('补货', `${sku} 广告关闭且库存不足，不建议恢复广告，优先补货。`, '高'));
    else if (stableOrders && profitDelta > 0) recommendations.push(makeRecommendation('保持当前策略', `${sku} 广告关闭/无数据后订单稳定、利润提升，说明控费有效，建议继续观察，暂不恢复广告。`, '中'));
    else if (orderDelta < -0.15) { risks.push(`${sku} 广告关闭后订单下降。`); recommendations.push(makeRecommendation('恢复小预算搜索广告', `${sku} 广告关闭/无数据后订单下降，建议恢复低预算 CPC 搜索或 CPM 搜索测试，不建议直接大预算投推荐。`, '高')); }
    else if (metrics.totalOrders.today === 0) recommendations.push(makeRecommendation('优化标题关键词', `${sku} 广告关闭且无订单，建议检查价格、主图、评价、关键词，再决定是否恢复广告。`, '高'));
    else recommendations.push(makeRecommendation('观察1天', `${sku} 广告关闭/无数据但自然订单尚可，可继续观察；如订单下降，优先恢复低预算 CPC 搜索或 CPM 搜索测试，不建议直接大预算投推荐。`, '低'));
  }

  if (metrics.adSpend.today > metrics.revenue.today * 0.35 && metrics.adSpend.today > 0) risks.push(`${sku} 广告费过高。`);
  if (exposure > 0 && clicks === 0) {
    risks.push(`${sku} 有曝光无点击，主图、标题或价格吸引力不足。`);
    recommendations.push(makeRecommendation('优化主图', `${sku} 今天有曝光但无点击，建议明天优先优化主图，同时检查标题和价格。`, '高'));
  }

  if (clicks > 0 && metrics.totalOrders.today === 0) {
    risks.push(`${sku} 有点击无转化，价格、评价或详情页可能阻碍下单。`);
    recommendations.push(makeRecommendation('降价或参加活动', `${sku} 今天有点击但无订单，建议优化价格、评价和详情页，可测试降价或参加活动。`, '高'));
  }

  const last3Positive = records.slice(-3).every((record) => enrichRecord(record).roi > 1 && enrichRecord(record).profit > 0);
  if (last3Positive && stock > 20) {
    recommendations.push(makeRecommendation('加大预算', `${sku} ROI 和利润近 3 天持续为正，且库存 ${stock} 充足，建议明天适当加大预算。`, '中'));
  }

  if (stock > 0 && stock < Math.max(5, metrics.totalOrders.last3Avg * 2)) {
    risks.push(`${sku} 库存不足：当前库存 ${stock} 偏低，不适合继续放大广告。`);
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
  const recordList = Array.isArray(records) ? records : [];
  const actionList = Array.isArray(actions) ? actions : [];
  const normalizedActions = actionList.map((action) => normalizeAction(action)).filter((action) => action.date && action.sku);
  const bySku = new Map();
  const candidateSkus = new Set();
  recordList.forEach((record) => {
    const recordDate = normalizeDateKey(record.date);
    if (!recordDate || !record.sku) return;
    const recordSku = normalizeSku(record.sku);
    const filterSku = normalizeSku(filters.sku);
    if (filterSku && recordSku !== filterSku) return;
    candidateSkus.add(recordSku);
    if (!hasValidBusinessData(record)) return;
    if (!bySku.has(recordSku)) bySku.set(recordSku, []);
    bySku.get(recordSku).push({ ...record, date: recordDate, sku: recordSku });
  });
  actionList.forEach((action) => {
    const actionSku = normalizeSku(action.sku);
    const filterSku = normalizeSku(filters.sku);
    if (actionSku && (!filterSku || actionSku === filterSku)) candidateSkus.add(actionSku);
  });

  const analyses = [...candidateSkus].map((sku) => {
    const sorted = (bySku.get(sku) || []).sort((a, b) => a.date.localeCompare(b.date));
    const inRange = sorted.filter((record) => (filters.allDates || !filters.startDate || record.date >= filters.startDate) && (filters.allDates || !filters.endDate || record.date <= filters.endDate));
    const targetDate = filters.endDate || filters.date || inRange.at(-1)?.date || sorted.at(-1)?.date;
    const exactToday = targetDate ? inRange.find((record) => record.date === targetDate) : null;
    const today = exactToday || (!targetDate ? inRange.at(-1) || sorted.at(-1) : null);
    const requiredActionDate = targetDate ? addDateDays(targetDate, -1) : '';
    const effectiveLookup = targetDate ? getEffectiveAction(normalizedActions, targetDate, sku, ACTION_LOOKBACK_DAYS) : { action: null, found: false, isInherited: false };
    const recentLookup = targetDate ? findRecentAction(normalizedActions, targetDate, sku, ACTION_LOOKBACK_DAYS) : { action: null, found: false };
    const currentDateAction = effectiveLookup.found && !effectiveLookup.isInherited ? effectiveLookup.action : null;
    const latestActionForMissing = effectiveLookup.action || null;
    const missingCurrentData = !today || (targetDate && !exactToday);
    if (missingCurrentData) {
      const recommendation = noDataRecommendation(sku, Boolean(latestActionForMissing));
      return {
        sku,
        date: targetDate || '',
        uniqueKey: targetDate ? buildActionKey(targetDate, sku) : '',
        latestAction: latestActionForMissing,
        previousAction: null,
        metrics: buildMetricSnapshot({}, {}, {}, {}, {}, {}),
        actionWindows: latestActionForMissing ? buildActionWindows(sorted, latestActionForMissing.date) : null,
        noValidData: true,
        actionMeta: { analysisDate: targetDate || '', comparisonDate: requiredActionDate, requiredActionDate, usedActionDate: latestActionForMissing?.date || '', daysSinceAction: effectiveLookup.daysSinceAction, lookbackDays: ACTION_LOOKBACK_DAYS, previousDayHadAction: recentLookup.previousDayHadAction, missingMessage: latestActionForMissing ? (effectiveLookup.isInherited ? `未找到当前日期新动作，已沿用最近动作：${effectiveLookup.sourceActionDate}。` : `当前日期已有保存动作：${latestActionForMissing.date} / ${sku}。`) : `最近 ${ACTION_LOOKBACK_DAYS} 天未找到动作记录。`, effectiveActionDate: effectiveLookup.sourceActionDate || latestActionForMissing?.date || '', isTodayNewAction: Boolean(currentDateAction), isInherited: Boolean(effectiveLookup.isInherited), sourceActionDate: effectiveLookup.sourceActionDate || '', sku, source: latestActionForMissing?.source || 'IndexedDB 动作记录', found: Boolean(latestActionForMissing), lookupKey: buildActionKey(latestActionForMissing?.date || requiredActionDate, sku), noValidData: true },
        primaryRecommendation: recommendation,
        recommendations: [recommendation],
        risks: [],
        effects: [],
      };
    }
    const todayIndex = sorted.findIndex((record) => record.uniqueKey === today.uniqueKey);
    const yesterdayDate = addDateDays(today.date, -1);
    const yesterday = sorted.find((record) => record.date === yesterdayDate) || sorted[todayIndex - 1] || {};
    const last3Rows = sorted.slice(Math.max(0, todayIndex - 2), todayIndex + 1);
    const last7Rows = sorted.slice(Math.max(0, todayIndex - 6), todayIndex + 1);
    const latestAction = effectiveLookup.action || null;
    const actionDate = effectiveLookup.sourceActionDate || latestAction?.date || requiredActionDate;
    const before3Rows = sorted.filter((record) => record.date < actionDate).slice(-3);
    const after3Rows = sorted.filter((record) => record.date > actionDate && record.date <= addDateDays(actionDate, 3));
    const previousAction = normalizedActions.filter((action) => normalizeSku(action.sku) === sku && normalizeDateKey(action.date) < actionDate).sort((a, b) => a.date.localeCompare(b.date)).at(-1);
    const metrics = buildMetricSnapshot(today, yesterday, averageRecords(last3Rows), averageRecords(last7Rows), averageRecords(before3Rows), averageRecords(after3Rows));
    const actionWindows = latestAction ? buildActionWindows(sorted, actionDate) : null;
    if (actionWindows) actionWindows.analysisDate = today.date;
    const nearbyActions = latestAction ? getSkuActionTimeline(normalizedActions, sku, { fromDate: addDateDays(actionDate, -2), toDate: addDateDays(today.date, 0) }) : [];
    const specialJudgement = latestAction ? buildActionJudgement(latestAction, actionWindows, nearbyActions) : '';
    const ruleResult = latestAction
      ? analyzeRules({ sku, today: enrichRecord(today), yesterday: enrichRecord(yesterday), metrics, latestAction, previousAction, records: sorted.slice(0, todayIndex + 1) })
      : { recommendations: [makeRecommendation('观察1天', `最近 ${ACTION_LOOKBACK_DAYS} 天未找到动作记录。`, '低')], risks: [], effects: [] };
    if (specialJudgement && !ruleResult.effects.some((e) => e.text === specialJudgement)) ruleResult.effects.unshift({ level: specialJudgement.includes('有效') || specialJudgement.includes('提升') ? '好' : '观察', text: specialJudgement });
    return {
      sku,
      date: today.date,
      uniqueKey: today.uniqueKey,
      latestAction,
      previousAction,
      metrics,
      actionWindows,
      actionMeta: {
        analysisDate: today.date,
        comparisonDate: yesterdayDate,
        requiredActionDate,
        usedActionDate: actionDate || '',
        effectiveActionDate: actionDate || '',
        isTodayNewAction: Boolean(currentDateAction),
        isInherited: Boolean(effectiveLookup.isInherited),
        sourceActionDate: effectiveLookup.sourceActionDate || '',
        daysSinceAction: effectiveLookup.daysSinceAction,
        lookbackDays: ACTION_LOOKBACK_DAYS,
        previousDayHadAction: recentLookup.previousDayHadAction,
        missingMessage: latestAction ? (effectiveLookup.isInherited ? `未找到当前日期新动作，已沿用最近动作：${effectiveLookup.sourceActionDate}。` : `当前日期已有保存动作：${actionDate} / ${sku}。`) : `最近 ${ACTION_LOOKBACK_DAYS} 天未找到动作记录。`,
        sku,
        source: latestAction?.source || 'IndexedDB 动作记录',
        found: Boolean(latestAction),
        lookupKey: buildActionKey(actionDate || requiredActionDate, sku),
      },
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
