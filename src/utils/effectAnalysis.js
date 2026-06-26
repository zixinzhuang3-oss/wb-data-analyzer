import { toProfitRub } from './currency.js';
import { hasValidBusinessData } from './excel.js';
import { addDays as addDateDays, normalizeDateKey } from './date.js';
import { buildActionKey, CPM_RECOMMEND_MIN_BID, CPM_SEARCH_MIN_BID, getActionRecord, normalizeAction, normalizeSku } from './actions.js';

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
    adShare: number(record.adShare) || safeDivide(number(record.adOrders), totalOrders) || safeDivide(adSpend, revenue),
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

const actionForDate = (actions, date, sku) => getActionRecord(actions, date, sku);
const noDataRecommendation = (sku, hasPreviousAction) => makeRecommendation('观察1天', hasPreviousAction
  ? `${sku} 已找到上一日动作记录，但当前日期暂无有效数据，暂时无法判断动作效果。`
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
  const normalizedActions = actions.map((action) => normalizeAction(action));
  const bySku = new Map();
  const candidateSkus = new Set();
  records.forEach((record) => {
    if (!record.date || !record.sku) return;
    const recordSku = normalizeSku(record.sku);
    const filterSku = normalizeSku(filters.sku);
    if (filterSku && recordSku !== filterSku) return;
    candidateSkus.add(recordSku);
    if (!hasValidBusinessData(record)) return;
    if (!bySku.has(recordSku)) bySku.set(recordSku, []);
    bySku.get(recordSku).push({ ...record, date: normalizeDateKey(record.date), sku: recordSku });
  });
  actions.forEach((action) => {
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
    const missingCurrentData = !today || (targetDate && !exactToday);
    const latestActionForMissing = requiredActionDate ? actionForDate(normalizedActions, requiredActionDate, sku) || null : null;
    if (missingCurrentData) {
      const recommendation = noDataRecommendation(sku, Boolean(latestActionForMissing));
      return {
        sku,
        date: targetDate || '',
        uniqueKey: targetDate ? buildActionKey(targetDate, sku) : '',
        latestAction: latestActionForMissing,
        previousAction: null,
        metrics: buildMetricSnapshot({}, {}, {}, {}, {}, {}),
        noValidData: true,
        actionMeta: { analysisDate: targetDate || '', comparisonDate: requiredActionDate, requiredActionDate, usedActionDate: latestActionForMissing?.date || requiredActionDate, sku, source: latestActionForMissing?.source || 'IndexedDB 动作记录', found: Boolean(latestActionForMissing), lookupKey: buildActionKey(requiredActionDate, sku), noValidData: true },
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
    const latestAction = actionForDate(normalizedActions, requiredActionDate, sku) || null;
    const actionDate = requiredActionDate;
    const before3Rows = sorted.filter((record) => record.date < actionDate).slice(-3);
    const after3Rows = sorted.filter((record) => record.date >= actionDate).slice(0, 3);
    const previousAction = normalizedActions.filter((action) => normalizeSku(action.sku) === sku && normalizeDateKey(action.date) < actionDate).sort((a, b) => a.date.localeCompare(b.date)).at(-1);
    const metrics = buildMetricSnapshot(today, yesterday, averageRecords(last3Rows), averageRecords(last7Rows), averageRecords(before3Rows), averageRecords(after3Rows));
    const ruleResult = latestAction
      ? analyzeRules({ sku, today: enrichRecord(today), yesterday: enrichRecord(yesterday), metrics, latestAction, previousAction, records: sorted.slice(0, todayIndex + 1) })
      : { recommendations: [makeRecommendation('观察1天', `未找到 ${requiredActionDate} / ${sku} 的动作记录，无法判断动作效果。`, '低')], risks: [], effects: [] };

    return {
      sku,
      date: today.date,
      uniqueKey: today.uniqueKey,
      latestAction,
      previousAction,
      metrics,
      actionMeta: {
        analysisDate: today.date,
        comparisonDate: yesterdayDate,
        requiredActionDate,
        usedActionDate: latestAction?.date || requiredActionDate,
        sku,
        source: latestAction?.source || 'IndexedDB 动作记录',
        found: Boolean(latestAction),
        lookupKey: buildActionKey(requiredActionDate, sku),
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
