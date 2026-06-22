import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { toIsoDate } from '../src/utils/excel.js';
import { parseOperationActionText, normalizeAction, actionToSummary } from '../src/utils/actions.js';
import { buildEffectAnalysis } from '../src/utils/effectAnalysis.js';
import { buildPeriodComparison, filterRecords, getPreviousDateRange, countDaysInclusive, addDays } from '../src/utils/history.js';

const XLSX = { SSF: { parse_date_code(value) { if (value === 46190) return { y: 2026, m: 6, d: 17 }; if (value === 46191) return { y: 2026, m: 6, d: 18 }; return null; } } };
assert.equal(toIsoDate(46190, XLSX), '2026-06-17');
assert.equal(toIsoDate(46191, XLSX), '2026-06-18');
assert.equal(toIsoDate(new Date(2026, 5, 17), XLSX), '2026-06-17');
assert.equal(toIsoDate('2026/06/17', XLSX), '2026-06-17');
assert.match(toIsoDate('6月17日', XLSX), /^\d{4}-06-17$/);

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
  { date: '2026-06-08', sku: 'ES100BK', uniqueKey: '2026-06-08__ES100BK', totalOrders: 5, revenue: 500, adSpend: 180, profit: 40, adImpressions: 5000, adClicks: 100, stock: 80 },
  { date: '2026-06-09', sku: 'ES100BK', uniqueKey: '2026-06-09__ES100BK', totalOrders: 5, revenue: 500, adSpend: 180, profit: 50, adImpressions: 5200, adClicks: 105, stock: 75 },
  { date: '2026-06-10', sku: 'ES100BK', uniqueKey: '2026-06-10__ES100BK', totalOrders: 5, revenue: 500, adSpend: 180, profit: 60, adImpressions: 5300, adClicks: 110, stock: 70 },
  { date: '2026-06-11', sku: 'ES100BK', uniqueKey: '2026-06-11__ES100BK', totalOrders: 5, revenue: 500, adSpend: 180, profit: 70, adImpressions: 5400, adClicks: 115, stock: 65 },
  { date: '2026-06-12', sku: 'ES100BK', uniqueKey: '2026-06-12__ES100BK', totalOrders: 5, revenue: 500, adSpend: 180, profit: 80, adImpressions: 5500, adClicks: 120, stock: 60 },
  { date: '2026-06-13', sku: 'ES100BK', uniqueKey: '2026-06-13__ES100BK', totalOrders: 5, revenue: 500, adSpend: 180, profit: 90, adImpressions: 5600, adClicks: 125, stock: 55 },
  { date: '2026-06-14', sku: 'ES100BK', uniqueKey: '2026-06-14__ES100BK', totalOrders: 5, revenue: 500, adSpend: 180, profit: 100, adImpressions: 5700, adClicks: 130, stock: 50 },
  { date: '2026-06-15', sku: 'ES100BK', uniqueKey: '2026-06-15__ES100BK', totalOrders: 6, revenue: 650, adSpend: 150, profit: 120, adImpressions: 6000, adClicks: 150, stock: 45 },
  { date: '2026-06-16', sku: 'ES100BK', uniqueKey: '2026-06-16__ES100BK', totalOrders: 6, revenue: 650, adSpend: 150, profit: 130, adImpressions: 6100, adClicks: 155, stock: 40 },
  { date: '2026-06-17', sku: 'ES100BK', uniqueKey: '2026-06-17__ES100BK', totalOrders: 7, revenue: 700, adSpend: 140, profit: 150, adImpressions: 6200, adClicks: 160, stock: 35 },
  { date: '2026-06-18', sku: 'ES100BK', uniqueKey: '2026-06-18__ES100BK', totalOrders: 7, revenue: 720, adSpend: 140, profit: 160, adImpressions: 6300, adClicks: 165, stock: 32 },
  { date: '2026-06-19', sku: 'ES100BK', uniqueKey: '2026-06-19__ES100BK', totalOrders: 7, revenue: 720, adSpend: 140, profit: 170, adImpressions: 6400, adClicks: 170, stock: 31 },
  { date: '2026-06-20', sku: 'ES100BK', uniqueKey: '2026-06-20__ES100BK', totalOrders: 8, revenue: 800, adSpend: 130, profit: 190, adImpressions: 6500, adClicks: 180, stock: 29 },
  { date: '2026-06-21', sku: 'ES100BK', uniqueKey: '2026-06-21__ES100BK', totalOrders: 8, revenue: 820, adSpend: 130, profit: 200, adImpressions: 6600, adClicks: 185, stock: 28 },
  { date: '2026-06-17', sku: 'ES200BK', uniqueKey: '2026-06-17__ES200BK', totalOrders: 3, revenue: 240, adSpend: 120, profit: -20, adImpressions: 4000, adClicks: 80, stock: 20 },
  { date: '2026-06-21', sku: 'ES200BK', uniqueKey: '2026-06-21__ES200BK', totalOrders: 2, revenue: 180, adSpend: 180, profit: -80, adImpressions: 5000, adClicks: 90, stock: 18 },
];

assert.equal(addDays('2026-06-17', 1), '2026-06-18');
assert.equal(addDays('2026-06-17', -1), '2026-06-16');
assert.equal(countDaysInclusive('2026-06-15', '2026-06-21'), 7);
assert.deepEqual(getPreviousDateRange('2026-06-15', '2026-06-21'), { startDate: '2026-06-08', endDate: '2026-06-14', dayCount: 7 });
assert.deepEqual(getPreviousDateRange('2026-06-20', '2026-06-21'), { startDate: '2026-06-18', endDate: '2026-06-19', dayCount: 2 });

const singleDayComparison = buildPeriodComparison(rangeRecords, { startDate: '2026-06-17', endDate: '2026-06-17' });
assert.equal(singleDayComparison.previousRange.startDate, '2026-06-16');
assert.equal(singleDayComparison.previousRange.endDate, '2026-06-16');
assert.equal(singleDayComparison.hasComparison, true);
assert.equal(singleDayComparison.currentSummary.totalOrders, 10);

const sevenDayComparison = buildPeriodComparison(rangeRecords, { startDate: '2026-06-15', endDate: '2026-06-21' });
assert.equal(sevenDayComparison.previousRange.startDate, '2026-06-08');
assert.equal(sevenDayComparison.previousRange.endDate, '2026-06-14');
assert.equal(sevenDayComparison.currentSummary.dateCount, 7);
assert.ok(sevenDayComparison.skuRows.length >= 2);
assert.ok(sevenDayComparison.skuRows.some((row) => row.sku === 'ES200BK' && row.judgment === '库存风险'));

const singleSkuComparison = buildPeriodComparison(rangeRecords, { startDate: '2026-06-15', endDate: '2026-06-21', sku: 'ES100BK' });
assert.equal(singleSkuComparison.currentSummary.skuCount, 1);
assert.equal(filterRecords(rangeRecords, { startDate: '2026-06-15', endDate: '2026-06-21', sku: 'ES100BK' }).every((row) => row.sku === 'ES100BK'), true);

const noPreviousComparison = buildPeriodComparison(rangeRecords, { startDate: '2026-06-08', endDate: '2026-06-08' });
assert.equal(noPreviousComparison.hasComparison, false);
assert.match(noPreviousComparison.metrics[0].trend, /上升|持平|下降/);

const june17Rows = filterRecords(rangeRecords, { startDate: '2026-06-17', endDate: '2026-06-17' });
assert.equal(june17Rows.length, 2);
assert.equal(buildPeriodComparison(rangeRecords, { startDate: '2026-06-17', endDate: '2026-06-17' }).currentSummary.totalOrders, 10);

const june17To21 = buildPeriodComparison(rangeRecords, { startDate: '2026-06-17', endDate: '2026-06-21' });
assert.equal(june17To21.currentSummary.totalOrders, 42);
assert.equal(june17To21.currentRecords.length, 7);
assert.equal(june17To21.previousRange.startDate, '2026-06-12');
assert.equal(june17To21.previousRange.endDate, '2026-06-16');

const allDatesComparison = buildPeriodComparison(rangeRecords, { quickRange: '全部日期' });
assert.equal(allDatesComparison.currentRange.startDate, '2026-06-08');
assert.equal(allDatesComparison.currentRange.endDate, '2026-06-21');
assert.equal(allDatesComparison.hasComparison, false);


const mainSource = readFileSync(new URL('../src/main.js', import.meta.url), 'utf8');
assert.match(mainSource, /id="start-date-filter" type="date"/);
assert.match(mainSource, /id="end-date-filter" type="date"/);
assert.doesNotMatch(mainSource, /id="date-filter"/);
assert.doesNotMatch(mainSource, /quick-range-filter/);
assert.match(mainSource, /data-quick-range/);
assert.match(mainSource, /全部日期/);

const periodLinked = buildEffectAnalysis(rangeRecords, [], { startDate: '2026-06-15', endDate: '2026-06-21', sku: 'ES100BK' }, singleSkuComparison)[0];
assert.ok(periodLinked.recommendations.some((item) => item.reason.includes('当前选择时间段')));

console.log('core tests passed');
