import assert from 'node:assert/strict';
import { hasEffectiveDailyData, toIsoDate } from '../src/utils/excel.js';
import { parseOperationActionText, normalizeAction, actionToSummary } from '../src/utils/actions.js';
import { buildEffectAnalysis } from '../src/utils/effectAnalysis.js';
import { buildComparison, filterRecords, resolveDateRange } from '../src/utils/history.js';
import { addDays, buildQuickRange, formatDateInTimeZone, getTodayDate } from '../src/utils/date.js';

const XLSX = { SSF: { parse_date_code(value) { if (value === 46190) return { y: 2026, m: 6, d: 17 }; if (value === 46191) return { y: 2026, m: 6, d: 18 }; return null; } } };
assert.equal(toIsoDate(46190, XLSX), '2026-06-17');
assert.equal(toIsoDate(46191, XLSX), '2026-06-18');
assert.equal(toIsoDate(new Date(2026, 5, 17), XLSX), '2026-06-17');
assert.equal(toIsoDate('2026/06/17', XLSX), '2026-06-17');
assert.match(toIsoDate('6月17日', XLSX), /^\d{4}-06-17$/);

const realToday = '2026-06-22';
assert.deepEqual(buildQuickRange('today', realToday), { allDates: false, startDate: '2026-06-22', endDate: '2026-06-22' });
assert.deepEqual(buildQuickRange('yesterday', realToday), { allDates: false, startDate: '2026-06-21', endDate: '2026-06-21' });
assert.deepEqual(buildQuickRange('3', realToday), { allDates: false, startDate: '2026-06-20', endDate: '2026-06-22' });
assert.deepEqual(buildQuickRange('7', realToday), { allDates: false, startDate: '2026-06-16', endDate: '2026-06-22' });
assert.deepEqual(buildQuickRange('14', realToday), { allDates: false, startDate: '2026-06-09', endDate: '2026-06-22' });
assert.deepEqual(buildQuickRange('30', realToday), { allDates: false, startDate: '2026-05-24', endDate: '2026-06-22' });
assert.equal(addDays('2026-06-22', -6), '2026-06-16');
assert.equal(addDays('2026-03-01', -1), '2026-02-28');
assert.equal(formatDateInTimeZone(new Date('2026-06-21T16:30:00.000Z'), 'Asia/Shanghai'), '2026-06-22');
const networkToday = await getTodayDate({ timeZone: 'Asia/Shanghai', fetchImpl: async () => ({ ok: true, json: async () => ({ datetime: '2026-06-21T16:30:00.000Z' }) }) });
assert.deepEqual(networkToday, { date: '2026-06-22', source: '网络时间', timeZone: 'Asia/Shanghai' });
const fallbackToday = await getTodayDate({ timeZone: 'Asia/Shanghai', fetchImpl: async () => { throw new Error('offline'); } });
assert.match(fallbackToday.date, /^\d{4}-\d{2}-\d{2}$/);
assert.equal(fallbackToday.source, '浏览器本地时间');

const scenarios = [
  ['只开 CPC 搜索', 'CPC=开启；CPC搜索出价=25；CPC预算=300；CPM=关闭；预算动作=保持预算；备注=CPC搜索保量', { adStatus: '仅 CPC', cpcEnabled: '开启', cpmEnabled: '关闭', cpcSearchBid: 25, cpcDailyBudget: 300 }],
  ['只开 CPM 搜索', 'CPC=关闭；CPM=开启；CPM位置=仅搜索；CPM出价方式=手动出价；CPM搜索出价=620；CPM预算=400；预算动作=保持预算；备注=搜索控费', { adStatus: '仅 CPM', cpmPosition: '仅搜索', cpmSearchBid: 620, cpmRecommendBid: '' }],
  ['只开 CPM 推荐', 'CPC=关闭；CPM=开启；CPM位置=仅推荐；CPM出价方式=手动出价；CPM推荐出价=200；CPM预算=300；预算动作=保持预算；备注=只测试推荐流量', { adStatus: '仅 CPM', cpmPosition: '仅推荐', cpmSearchBid: '', cpmRecommendBid: 200 }],
  ['CPM 搜索+推荐手动', 'CPC=关闭；CPM=开启；CPM位置=搜索+推荐；CPM出价方式=手动出价；CPM搜索出价=730；CPM推荐出价=200；CPM预算=500；预算动作=降低预算；备注=搜索推荐手动', { adStatus: '仅 CPM', cpmPosition: '搜索+推荐', cpmSearchBid: 730, cpmRecommendBid: 200 }],
  ['CPM 搜索+推荐统一', 'CPC=关闭；CPM=开启；CPM位置=搜索+推荐；CPM出价方式=统一出价；CPM统一出价=650；CPM预算=500；预算动作=保持预算；备注=统一出价测试搜索和推荐', { adStatus: '仅 CPM', cpmBidType: '统一出价', cpmUnifiedBid: 650, cpmSearchBid: '', cpmRecommendBid: '' }],
  ['CPC + CPM 推荐', 'CPC=开启；CPC搜索出价=25；CPC预算=300；CPM=开启；CPM位置=仅推荐；CPM出价方式=手动出价；CPM推荐出价=200；CPM预算=300；预算动作=保持预算；备注=CPC保搜索流量，CPM推荐低预算测试', { adStatus: 'CPC+CPM', cpcSearchBid: 25, cpmPosition: '仅推荐', cpmRecommendBid: 200 }],
  ['CPC + CPM 搜索推荐', 'CPC=开启；CPC搜索出价=30；CPC预算=500；CPM=开启；CPM位置=搜索+推荐；CPM出价方式=手动出价；CPM搜索出价=730；CPM推荐出价=200；CPM预算=500；预算动作=降低预算；备注=搜索位保量，推荐位控费', { adStatus: 'CPC+CPM', cpcSearchBid: 30, cpmSearchBid: 730, cpmRecommendBid: 200 }],
  ['广告关闭', 'CPC=关闭；CPM=关闭；预算动作=暂停广告；备注=广告关闭，观察自然流量', { adStatus: '无广告', cpcEnabled: '关闭', cpmEnabled: '关闭' }],
];

scenarios.forEach(([name, text, expected], index) => {
  const parsed = parseOperationActionText(text, '2026-06-17', `ES${String(index).padStart(3, '0')}BK`);
  assert.ok(parsed, name);
  assert.equal(parsed.source, 'Excel 自动识别');
  Object.entries(expected).forEach(([key, value]) => assert.equal(parsed[key], value, `${name} ${key}`));
});

assert.equal(parseOperationActionText('今天先观察自然流量，不要强行解析', '2026-06-17', 'ES033BK'), null);
assert.equal(parseOperationActionText('CPC=开启；随便写=不要解析', '2026-06-17', 'ES033BK'), null);

const cpcOff = normalizeAction({ date: '2026-06-17', sku: 'ES040BK', cpcEnabled: '关闭', cpcSearchBid: '30', cpcDailyBudget: '300' });
assert.equal(cpcOff.cpcSearchBid, '');
assert.equal(cpcOff.cpcDailyBudget, '');

const summary = actionToSummary(parseOperationActionText(scenarios[5][1], '2026-06-17', 'ES035BK'));
assert.match(summary, /整体广告状态：CPC\+CPM/);
assert.match(summary, /CPC搜索出价：25/);
assert.match(summary, /CPM推荐出价：200/);

const records = [
  { date: '2026-06-17', sku: 'ES035BK', uniqueKey: '2026-06-17__ES035BK', totalOrders: 10, adSpend: 200, adImpressions: 10000, adClicks: 300, revenue: 1000, profit: 300, stock: 50 },
  { date: '2026-06-18', sku: 'ES035BK', uniqueKey: '2026-06-18__ES035BK', totalOrders: 10, adSpend: 260, adImpressions: 18000, adClicks: 250, revenue: 1100, profit: 360, stock: 50 },
];
const actions = [parseOperationActionText(scenarios[5][1], '2026-06-18', 'ES035BK')];
const analysis = buildEffectAnalysis(records, actions, { date: '2026-06-18', sku: 'ES035BK' })[0];
assert.ok(analysis.effects.some((item) => item.text.includes('CPC搜索、CPM搜索、CPM推荐')));
assert.ok(analysis.recommendations.some((item) => item.reason.includes('CPC') || item.reason.includes('CPM')));

const rangeRecords = [
  { date: '2026-06-10', sku: 'ES035BK', uniqueKey: '2026-06-10__ES035BK', totalOrders: 1, adSpend: 10, adImpressions: 100, adClicks: 10, revenue: 100, profit: 10 },
  { date: '2026-06-16', sku: 'ES035BK', uniqueKey: '2026-06-16__ES035BK', totalOrders: 2, adSpend: 20, adImpressions: 200, adClicks: 20, revenue: 200, profit: 20 },
  { date: '2026-06-17', sku: 'ES035BK', uniqueKey: '2026-06-17__ES035BK', totalOrders: 3, adSpend: 30, adImpressions: 300, adClicks: 30, revenue: 300, profit: 30 },
  { date: '2026-06-17', sku: 'ES040BK', uniqueKey: '2026-06-17__ES040BK', totalOrders: 4, adSpend: 40, adImpressions: 400, adClicks: 40, revenue: 400, profit: 40 },
  { date: '2026-06-18', sku: 'ES035BK', uniqueKey: '2026-06-18__ES035BK', totalOrders: 5, adSpend: 50, adImpressions: 500, adClicks: 50, revenue: 500, profit: 50 },
  { date: '2026-06-19', sku: 'ES035BK', uniqueKey: '2026-06-19__ES035BK', totalOrders: 6, adSpend: 60, adImpressions: 600, adClicks: 60, revenue: 600, profit: 60 },
  { date: '2026-06-20', sku: 'ES035BK', uniqueKey: '2026-06-20__ES035BK', totalOrders: 7, adSpend: 70, adImpressions: 700, adClicks: 70, revenue: 700, profit: 70 },
  { date: '2026-06-21', sku: 'ES035BK', uniqueKey: '2026-06-21__ES035BK', totalOrders: 8, adSpend: 80, adImpressions: 800, adClicks: 80, revenue: 800, profit: 80 },
  { date: '2026-06-22', sku: 'ES035BK', uniqueKey: '2026-06-22__ES035BK', totalOrders: 9, adSpend: 90, adImpressions: 900, adClicks: 90, revenue: 900, profit: 90 },
  { date: '2026-07-19', sku: 'ES035BK', uniqueKey: '2026-07-19__ES035BK', totalOrders: 99, adSpend: 990, adImpressions: 9900, adClicks: 990, revenue: 9900, profit: 990 },
];

const singleDay = buildComparison(rangeRecords, { startDate: '2026-06-17', endDate: '2026-06-17', sku: 'ES035BK' });
assert.equal(singleDay.current.totalOrders, 3);
assert.equal(singleDay.previousRange.startDate, '2026-06-16');
assert.equal(singleDay.previousRange.endDate, '2026-06-16');
assert.equal(singleDay.previous.totalOrders, 2);

const quickTodayRows = filterRecords(rangeRecords, buildQuickRange('today', realToday));
assert.equal(quickTodayRows.some((row) => row.date === '2026-07-19'), false);
assert.equal(quickTodayRows.reduce((sum, row) => sum + row.totalOrders, 0), 9);

const realSevenDayRange = buildQuickRange('7', realToday);
assert.deepEqual(realSevenDayRange, { allDates: false, startDate: '2026-06-16', endDate: '2026-06-22' });
const sevenDays = buildComparison(rangeRecords, { ...realSevenDayRange, sku: 'ES035BK' });
assert.equal(sevenDays.previousRange.startDate, '2026-06-09');
assert.equal(sevenDays.previousRange.endDate, '2026-06-15');
assert.equal(sevenDays.previous.totalOrders, 1);

const allSkuRows = filterRecords(rangeRecords, { startDate: '2026-06-17', endDate: '2026-06-17' });
assert.equal(new Set(allSkuRows.map((row) => row.sku)).size, 2);

const singleSkuRows = filterRecords(rangeRecords, { startDate: '2026-06-17', endDate: '2026-06-17', sku: 'ES040BK' });
assert.equal(singleSkuRows.length, 1);
assert.equal(singleSkuRows[0].sku, 'ES040BK');

const noPrevious = buildComparison(rangeRecords, { startDate: '2026-06-09', endDate: '2026-06-09' });
assert.equal(noPrevious.hasPreviousData, false);

assert.deepEqual(resolveDateRange(rangeRecords, { startDate: '2026-06-21', endDate: '2026-06-17' }), { allDates: false, startDate: '2026-06-17', endDate: '2026-06-21' });
assert.equal(filterRecords(rangeRecords, { startDate: '2026-06-17', endDate: '2026-06-17' }).reduce((sum, row) => sum + row.totalOrders, 0), 7);
assert.equal(filterRecords(rangeRecords, buildQuickRange('today', '2026-06-23')).length, 0);
assert.equal(hasEffectiveDailyData({ date: '2026-07-19', sku: 'ES035BK', totalOrders: 0, adSpend: 0, revenue: 0, profit: 0, operationAction: '' }), false);
assert.equal(hasEffectiveDailyData({ date: '2026-07-19', sku: 'ES035BK', totalOrders: 1, adSpend: 0, revenue: 0, profit: 0, operationAction: '' }), true);

console.log('core tests passed');
