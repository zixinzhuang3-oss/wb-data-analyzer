import assert from 'node:assert/strict';
import fs from 'node:fs';
import { buildSheetDiagnostics, hasEffectiveDailyData, normalizeSheetRow, toIsoDate, toNumber } from '../src/utils/excel.js';
import { parseOperationActionText, normalizeAction, actionToSummary, mergeActionRecords, shouldReplaceAction, validateCpmMinBids } from '../src/utils/actions.js';
import { buildEffectAnalysis } from '../src/utils/effectAnalysis.js';
import { buildComparison, filterRecords, resolveDateRange } from '../src/utils/history.js';
import { addDays, buildQuickRange, formatDateInTimeZone, getTodayDate } from '../src/utils/date.js';
import { formatPercent, formatRuble, formatYuan } from '../src/utils/analysis.js';
import { CNY_TO_RUB, toProfitRub } from '../src/utils/currency.js';

const XLSX = { SSF: { parse_date_code(value) { if (value === 46190) return { y: 2026, m: 6, d: 17 }; if (value === 46191) return { y: 2026, m: 6, d: 18 }; return null; } } };
assert.equal(toIsoDate(46190, XLSX), '2026-06-17');
assert.equal(toIsoDate(46191, XLSX), '2026-06-18');
assert.equal(toIsoDate(new Date(2026, 5, 17), XLSX), '2026-06-17');
assert.equal(toIsoDate('2026/06/17', XLSX), '2026-06-17');
assert.match(toIsoDate('6月17日', XLSX), /^\d{4}-06-17$/);

assert.equal(toNumber('1,512 ₽'), 1512);
assert.equal(toNumber('1512'), 1512);
assert.equal(toNumber('1,512.50 ₽'), 1512.5);
assert.equal(toNumber('0 ₽'), 0);
assert.equal(toNumber('  '), 0);

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
  assert.equal(parsed.source, 'excel_auto');
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
const actions = [parseOperationActionText(scenarios[5][1], '2026-06-17', 'ES035BK')];
const analysis = buildEffectAnalysis(records, actions, { date: '2026-06-18', sku: 'ES035BK' })[0];
assert.equal(analysis.actionMeta.analysisDate, '2026-06-18');
assert.equal(analysis.actionMeta.comparisonDate, '2026-06-17');
assert.equal(analysis.actionMeta.requiredActionDate, '2026-06-17');
assert.equal(analysis.actionMeta.usedActionDate, '2026-06-17');
assert.equal(analysis.actionMeta.sku, 'ES035BK');
assert.equal(analysis.actionMeta.found, true);
assert.ok(analysis.effects.some((item) => item.text.includes('CPC搜索、CPM搜索、CPM推荐')));
assert.ok(analysis.recommendations.some((item) => item.reason.includes('CPC') || item.reason.includes('CPM')));



const noAdPreviousAction = normalizeAction({ date: '2026-06-21', sku: 'ES030BK', cpcEnabled: '关闭', cpmEnabled: '关闭', source: '手动填写' });
const noAdAnalysis = buildEffectAnalysis([
  { date: '2026-06-21', sku: 'ES030BK', uniqueKey: '2026-06-21__ES030BK', totalOrders: 2, adSpend: 0, adImpressions: 0, adClicks: 0, revenue: 200, profit: 60, stock: 30 },
  { date: '2026-06-22', sku: 'ES030BK', uniqueKey: '2026-06-22__ES030BK', totalOrders: 3, adSpend: 0, adImpressions: 0, adClicks: 0, revenue: 300, profit: 90, stock: 30 },
], [noAdPreviousAction], { date: '2026-06-22', sku: 'ES030BK' })[0];
assert.equal(noAdAnalysis.latestAction.adStatus, '无广告');
assert.equal(noAdAnalysis.latestAction.cpcEnabled, '关闭');
assert.equal(noAdAnalysis.latestAction.cpmEnabled, '关闭');
assert.doesNotMatch(actionToSummary(noAdAnalysis.latestAction), /CPC搜索出价：10/);

const staleActionOnly = normalizeAction({ date: '2026-06-20', sku: 'ES030BK', cpcEnabled: '开启', cpcSearchBid: 10, cpmEnabled: '关闭', source: '手动填写' });
const missingPreviousActionAnalysis = buildEffectAnalysis([
  { date: '2026-06-21', sku: 'ES030BK', uniqueKey: '2026-06-21__ES030BK', totalOrders: 2, adSpend: 0, revenue: 200, profit: 60 },
  { date: '2026-06-22', sku: 'ES030BK', uniqueKey: '2026-06-22__ES030BK', totalOrders: 3, adSpend: 0, revenue: 300, profit: 90 },
], [staleActionOnly], { date: '2026-06-22', sku: 'ES030BK' })[0];
assert.equal(missingPreviousActionAnalysis.latestAction, null);
assert.equal(missingPreviousActionAnalysis.actionMeta.requiredActionDate, '2026-06-21');
assert.equal(missingPreviousActionAnalysis.actionMeta.found, false);
assert.match(missingPreviousActionAnalysis.primaryRecommendation.reason, /未找到 2026-06-21 \/ ES030BK 的动作记录，无法判断动作效果/);



// Import / merge order invariance and action priority rules.
const yesterdayJsonAction = normalizeAction({ date: '2026/06/22', sku: 'ES920E', cpcEnabled: '开启', cpcSearchBid: 20, cpmEnabled: '关闭', source: 'json_import', updatedAt: '2026-06-22T20:00:00.000Z' });
const todayExcelAction = parseOperationActionText('CPC=开启；CPC搜索出价=10；CPC预算=100；CPM=关闭', '2026-06-23', 'ES920E');
const sequenceA = mergeActionRecords([], [yesterdayJsonAction]);
const sequenceAFinal = mergeActionRecords(sequenceA.actions, [todayExcelAction]);
const sequenceB = mergeActionRecords([], [todayExcelAction]);
const sequenceBFinal = mergeActionRecords(sequenceB.actions, [yesterdayJsonAction]);
assert.equal(sequenceAFinal.actions.find((item) => item.uniqueKey === '2026-06-22__ES920E')?.source, 'json_import');
assert.equal(sequenceBFinal.actions.find((item) => item.uniqueKey === '2026-06-22__ES920E')?.source, 'json_import');
assert.deepEqual(new Set(sequenceAFinal.actions.map((item) => item.uniqueKey)), new Set(sequenceBFinal.actions.map((item) => item.uniqueKey)));

const manualAction = normalizeAction({ date: '2026-06-23', sku: 'ES920E', cpcEnabled: '开启', cpcSearchBid: 35, cpmEnabled: '关闭', source: 'manual', updatedAt: '2026-06-23T10:00:00.000Z' });
const blankExcelAction = parseOperationActionText('', '2026-06-23', 'ES920E');
assert.equal(blankExcelAction, null);
assert.equal(mergeActionRecords([manualAction], [todayExcelAction]).actions[0].source, 'manual');
assert.equal(mergeActionRecords([manualAction], [todayExcelAction]).stats.keptLocal, 1);

const manualModifiedAction = normalizeAction({ date: '2026-06-23', sku: 'ES920E', cpcEnabled: '开启', cpcSearchBid: 50, cpmEnabled: '关闭', source: 'manual_modified', updatedAt: '2026-06-23T09:00:00.000Z' });
const newerJsonAction = normalizeAction({ date: '2026-06-23', sku: 'ES920E', cpcEnabled: '开启', cpcSearchBid: 45, cpmEnabled: '关闭', source: 'json_import', updatedAt: '2026-06-23T12:00:00.000Z' });
assert.equal(shouldReplaceAction(newerJsonAction, manualModifiedAction), false);
assert.equal(mergeActionRecords([manualModifiedAction], [newerJsonAction]).actions[0].cpcSearchBid, 50);

const olderJsonAction = normalizeAction({ date: '2026-06-23', sku: 'ES920E', cpcEnabled: '开启', cpcSearchBid: 11, cpmEnabled: '关闭', source: 'json_import', updatedAt: '2026-06-23T08:00:00.000Z' });
const newerSamePriorityJsonAction = normalizeAction({ date: '2026-06-23', sku: 'ES920E', cpcEnabled: '开启', cpcSearchBid: 12, cpmEnabled: '关闭', source: 'json_import', updatedAt: '2026-06-23T09:00:00.000Z' });
assert.equal(shouldReplaceAction(newerSamePriorityJsonAction, olderJsonAction), true);
assert.equal(mergeActionRecords([olderJsonAction], [newerSamePriorityJsonAction]).actions[0].cpcSearchBid, 12);

const localOnlyAction = normalizeAction({ date: '2026-06-21', sku: 'LOCAL', cpcEnabled: '关闭', cpmEnabled: '关闭', source: 'manual' });
assert.equal(mergeActionRecords([localOnlyAction], [yesterdayJsonAction]).actions.some((item) => item.uniqueKey === '2026-06-21__LOCAL'), true);

const exactPreviousActionAnalysis = buildEffectAnalysis([
  { date: '2026-06-22', sku: 'ES920E', uniqueKey: '2026-06-22__ES920E', totalOrders: 1, adSpend: 10, revenue: 100, profit: 10 },
  { date: '2026-06-23', sku: 'ES920E', uniqueKey: '2026-06-23__ES920E', totalOrders: 2, adSpend: 20, revenue: 200, profit: 20 },
], [yesterdayJsonAction], { date: '2026-06-23', sku: 'ES920E' })[0];
assert.equal(exactPreviousActionAnalysis.actionMeta.requiredActionDate, '2026-06-22');
assert.equal(exactPreviousActionAnalysis.latestAction.uniqueKey, '2026-06-22__ES920E');

const exactMissingActionAnalysis = buildEffectAnalysis([
  { date: '2026-06-22', sku: 'ES921E', uniqueKey: '2026-06-22__ES921E', totalOrders: 1, adSpend: 0, revenue: 100, profit: 10 },
  { date: '2026-06-23', sku: 'ES921E', uniqueKey: '2026-06-23__ES921E', totalOrders: 2, adSpend: 0, revenue: 200, profit: 20 },
], [], { date: '2026-06-23', sku: 'ES921E' })[0];
assert.equal(exactMissingActionAnalysis.latestAction, null);
assert.match(exactMissingActionAnalysis.primaryRecommendation.reason, /未找到 2026-06-22 \/ ES921E 的动作记录/);
assert.doesNotMatch(exactMissingActionAnalysis.primaryRecommendation.reason, /CPC|CPM|无广告/);

const slashDateAction = normalizeAction({ date: '2026/06/22', sku: 'ES922E', cpcEnabled: '关闭', cpmEnabled: '关闭', source: 'json_import' });
assert.equal(slashDateAction.uniqueKey, '2026-06-22__ES922E');
const serialDateAction = normalizeAction({ date: toIsoDate(46191, XLSX), sku: 'ES923E', cpcEnabled: '关闭', cpmEnabled: '关闭', source: 'json_import' });
assert.equal(serialDateAction.uniqueKey, '2026-06-18__ES923E');
const exportedLike = { records: [{ date: '2026-06-23', sku: 'ES920E', uniqueKey: '2026-06-23__ES920E' }], actionRecords: [yesterdayJsonAction], recommendationHistory: [{ date: '2026-06-23', sku: 'ES920E', recommendationType: '观察1天' }], settings: {}, exportedAt: '2026-06-23T00:00:00.000Z' };
assert.equal(exportedLike.actionRecords.length, 1);
assert.equal(mergeActionRecords([], exportedLike.actionRecords).actions[0].uniqueKey, '2026-06-22__ES920E');

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


const noDataTodayAction = normalizeAction({ date: '2026-06-21', sku: 'ES030BK', cpcEnabled: '关闭', cpmEnabled: '关闭', source: '手动填写' });
const noDataTodayAnalysis = buildEffectAnalysis([
  { date: '2026-06-21', sku: 'ES030BK', uniqueKey: '2026-06-21__ES030BK', totalOrders: 2, adSpend: 0, revenue: 200, profit: 60, stock: 30 },
  { date: '2026-06-22', sku: 'ES030BK', uniqueKey: '2026-06-22__ES030BK', totalOrders: 0, adSpend: 0, revenue: 0, profit: 0, hasValidBusinessData: false },
], [noDataTodayAction], { date: '2026-06-22', sku: 'ES030BK' })[0];
assert.equal(noDataTodayAnalysis.noValidData, true);
assert.equal(noDataTodayAnalysis.actionMeta.found, true);
assert.match(noDataTodayAnalysis.primaryRecommendation.reason, /当前日期暂无有效数据/);
assert.equal(noDataTodayAnalysis.recommendations.some((item) => /恢复|暂停广告|加大预算|降低/.test(item.type + item.reason)), false);
assert.equal(noDataTodayAnalysis.metrics.totalOrders.todayVsYesterday.rate, 0);
assert.doesNotMatch(noDataTodayAnalysis.primaryRecommendation.reason, /下降 100|订单下降|恢复低预算|恢复小预算/);

const missingTodayWithoutBlankRow = buildEffectAnalysis([
  { date: '2026-06-21', sku: 'ES031BK', uniqueKey: '2026-06-21__ES031BK', totalOrders: 4, adSpend: 0, revenue: 400, profit: 80, stock: 30 },
], [normalizeAction({ date: '2026-06-21', sku: 'ES031BK', cpcEnabled: '关闭', cpmEnabled: '关闭', source: '手动填写' })], { date: '2026-06-22', sku: 'ES031BK' })[0];
assert.equal(missingTodayWithoutBlankRow.noValidData, true);
assert.match(missingTodayWithoutBlankRow.primaryRecommendation.reason, /已找到上一日动作记录，但当前日期暂无有效数据/);
assert.equal(missingTodayWithoutBlankRow.recommendations.some((item) => /恢复/.test(item.type + item.reason)), false);

const realZeroOrderAnalysis = buildEffectAnalysis([
  { date: '2026-06-21', sku: 'ES032BK', uniqueKey: '2026-06-21__ES032BK', totalOrders: 2, adSpend: 50, revenue: 200, profit: 20, stock: 30 },
  { date: '2026-06-22', sku: 'ES032BK', uniqueKey: '2026-06-22__ES032BK', totalOrders: 0, adSpend: 0, revenue: 0, profit: 0, stock: 30, hasValidBusinessData: true },
], [normalizeAction({ date: '2026-06-21', sku: 'ES032BK', cpcEnabled: '关闭', cpmEnabled: '关闭', source: '手动填写' })], { date: '2026-06-22', sku: 'ES032BK' })[0];
assert.equal(realZeroOrderAnalysis.noValidData, undefined);
assert.equal(realZeroOrderAnalysis.metrics.totalOrders.today, 0);
assert.equal(realZeroOrderAnalysis.metrics.totalOrders.todayVsYesterday.rate, -1);

const futureBlankRows = [
  { date: '2026-06-21', sku: 'ES033BK', uniqueKey: '2026-06-21__ES033BK', totalOrders: 5, adSpend: 10, revenue: 500, profit: 50 },
  { date: '2026-07-19', sku: 'ES033BK', uniqueKey: '2026-07-19__ES033BK', totalOrders: 0, adSpend: 0, revenue: 0, profit: 0, hasValidBusinessData: false },
];
assert.equal(filterRecords(futureBlankRows, { allDates: true }).length, 1);
assert.equal(buildComparison(futureBlankRows, { allDates: true }).current.totalOrders, 5);
assert.equal(buildEffectAnalysis(futureBlankRows, [], { date: '2026-07-19', sku: 'ES033BK' })[0].noValidData, true);

const fixedRow = [];
fixedRow[0] = '2026-06-17';
fixedRow[11] = 20; // totalOrders via fallback header name below
fixedRow[13] = 1000;
fixedRow[14] = 50;
fixedRow[15] = 0;
fixedRow[16] = 10;
fixedRow[17] = 0;
fixedRow[26] = 1646.87;
fixedRow[27] = 1872.06;
fixedRow[28] = 6;
fixedRow[29] = 0;
fixedRow[30] = 0;
fixedRow[31] = 4000;
fixedRow[32] = 200;
fixedRow[33] = 0;
fixedRow[34] = 40;
fixedRow[35] = 0;
fixedRow[36] = 0;
fixedRow[25] = 999999; // wrong neighboring column must not be treated as revenue
const fixedHeaders = ['日期', '', '', '', '', '', '', '', '', '', '', '总订单', '', '错误曝光标题', '错误点击标题', '', '', '', '', '', '', '', '', '', '', '错误销售额标题', 'AA销售额'];
const fixedRecord = normalizeSheetRow('ES068BK', fixedHeaders, fixedRow, XLSX);
assert.equal(fixedRecord.impressions, 1000);
assert.equal(fixedRecord.clicks, 50);
assert.equal(fixedRecord.ctr, 0.05);
assert.equal(fixedRecord.addToCart, 10);
assert.equal(fixedRecord.conversionRate, 0.2);
assert.equal(fixedRecord.orderConversionRate, 0.4);
assert.equal(fixedRecord.revenue, 1646.87);
assert.equal(fixedRecord.adSpend, 1872.06);
assert.equal(fixedRecord.adOrders, 6);
assert.equal(fixedRecord.adImpressions, 4000);
assert.equal(fixedRecord.adClicks, 200);
assert.equal(fixedRecord.adCtr, 0.05);
assert.equal(fixedRecord.adAddToCart, 40);
assert.equal(fixedRecord.adClickAddToCartRate, 0.2);
assert.equal(fixedRecord.adCostPerOrder, 312.01);
assert.equal(Number(fixedRecord.adAvgClickCost.toFixed(2)), 9.36);
assert.equal(formatRuble(fixedRecord.revenue), '₽1,646.87');
assert.equal(formatRuble(fixedRecord.adSpend), '₽1,872.06');
assert.doesNotMatch(formatRuble(fixedRecord.revenue), /¥/);
assert.equal(formatYuan(67.27), '¥67.27');
assert.equal(formatPercent(fixedRecord.adSpend / fixedRecord.revenue), '113.7%');
assert.equal(Number((fixedRecord.revenue / fixedRecord.adSpend).toFixed(2)), 0.88);
assert.equal(CNY_TO_RUB, 11.5);
assert.equal(toProfitRub({ profitCny: 70.80 }), 814.2);
assert.equal(formatRuble(toProfitRub({ profitCny: 70.80 })), '₽814.20');
assert.equal(formatPercent(toProfitRub({ profitCny: 70.80 }) / 20853.92), '3.9%');

const fixedCurrencyRow = [...fixedRow];
fixedCurrencyRow[26] = '1,512.50 ₽';
fixedCurrencyRow[27] = '1,512 ₽';
const fixedCurrencyRecord = normalizeSheetRow('ES069BK', fixedHeaders, fixedCurrencyRow, XLSX);
assert.equal(fixedCurrencyRecord.revenue, 1512.5);
assert.equal(fixedCurrencyRecord.adSpend, 1512);
assert.equal(fixedCurrencyRecord.adShare, fixedCurrencyRecord.adSpend / fixedCurrencyRecord.revenue);

const diagnostics = buildSheetDiagnostics('ES068BK', [fixedHeaders, fixedRow], 0, XLSX);
assert.equal(diagnostics.fields['总订单销售额（不含刷单）'], 'AA 列，单位 ₽');
assert.equal(diagnostics.fields['总广告费'], 'AB 列，单位 ₽');

const blankAdRow = [...fixedRow];
for (let i = 27; i <= 36; i += 1) blankAdRow[i] = '';
blankAdRow[13] = 800;
blankAdRow[14] = 40;
blankAdRow[16] = 8;
const blankAdRecord = normalizeSheetRow('ES068BK', fixedHeaders, blankAdRow, XLSX);
assert.equal(blankAdRecord.adStatus, '无广告数据');
assert.equal(blankAdRecord.impressions, 800);
assert.equal(blankAdRecord.clicks, 40);
assert.equal(blankAdRecord.conversionRate, 0.2);

const trafficRecords = [
  { date: '2026-06-17', sku: 'ES068BK', uniqueKey: '2026-06-17__ES068BK', totalOrders: 2, impressions: 100, clicks: 10, addToCart: 4, adImpressions: 1000, adClicks: 100, adAddToCart: 20, adOrders: 1, adSpend: 30, revenue: 100, profit: 10, adStatus: '开启' },
  { date: '2026-06-18', sku: 'ES068BK', uniqueKey: '2026-06-18__ES068BK', totalOrders: 4, impressions: 300, clicks: 30, addToCart: 6, adImpressions: 2000, adClicks: 200, adAddToCart: 40, adOrders: 3, adSpend: 90, revenue: 300, profit: 20, adStatus: '开启' },
];
const trafficSummary = buildComparison(trafficRecords, { startDate: '2026-06-17', endDate: '2026-06-18', sku: 'ES068BK' }).current;
assert.equal(trafficSummary.impressions, 400);
assert.equal(trafficSummary.clicks, 40);
assert.equal(trafficSummary.ctr, 0.1);
assert.equal(trafficSummary.addToCart, 10);
assert.equal(trafficSummary.cvr, 0.25);
assert.equal(trafficSummary.orderConversionRate, 0.15);
assert.equal(trafficSummary.adImpressions, 3000);
assert.equal(trafficSummary.adClicks, 300);
assert.equal(trafficSummary.adCtr, 0.1);
assert.equal(trafficSummary.adAddToCart, 60);
assert.equal(trafficSummary.adClickAddToCartRate, 0.2);
assert.equal(trafficSummary.adOrders, 4);
assert.equal(trafficSummary.totalAdSpend, 120);
assert.equal(trafficSummary.adShare, 0.3);
assert.equal(trafficSummary.adCostPerOrder, 30);
assert.equal(trafficSummary.adAvgClickCost, 0.4);
assert.equal(trafficSummary.acos, 0.3);
assert.equal(trafficSummary.roi, 10 / 3);
assert.equal(trafficSummary.totalProfitCny, 30);
assert.equal(trafficSummary.totalProfitRub, 345);
assert.equal(trafficSummary.margin, trafficSummary.totalProfitRub / trafficSummary.totalRevenue);

assert.equal(trafficSummary.totalRevenue, 400);
assert.equal(formatRuble(trafficSummary.totalRevenue), '₽400.00');
assert.equal(formatRuble(trafficSummary.totalAdSpend), '₽120.00');
assert.equal(formatRuble(trafficSummary.totalProfit), '₽345.00');
assert.equal(formatYuan(trafficSummary.totalProfitCny), '¥30.00');
assert.doesNotMatch(formatRuble(trafficSummary.totalRevenue), /¥/);
assert.doesNotMatch(formatRuble(trafficSummary.totalAdSpend), /¥/);
assert.equal(trafficSummary.acos, trafficSummary.totalAdSpend / trafficSummary.totalRevenue);
assert.equal(trafficSummary.roi, trafficSummary.totalRevenue / trafficSummary.totalAdSpend);

const mainSource = fs.readFileSync(new URL('../src/main.js', import.meta.url), 'utf8');
assert.match(mainSource, /总销售额 ₽/);
assert.match(mainSource, /总广告费 ₽/);
assert.match(mainSource, /总利润 ₽/);
assert.match(mainSource, /总销售额：\$\{formatRuble/);
assert.match(mainSource, /总广告费：\$\{formatRuble/);
assert.match(mainSource, /总利润：\$\{formatRuble/);
assert.match(mainSource, /原始 \$\{formatYuan/);
assert.doesNotMatch(mainSource, /总销售额（人民币|总广告费（人民币|总销售额 ¥|总广告费 ¥/);

const buildPoorCpmAnalysis = (actionOverrides, sku = `CPM${Math.random().toString(36).slice(2, 8)}`) => buildEffectAnalysis([
  { date: '2026-06-21', sku, uniqueKey: `2026-06-21__${sku}`, totalOrders: 3, adSpend: 100, adImpressions: 1000, adClicks: 50, revenue: 600, profitRub: 80, stock: 50 },
  { date: '2026-06-22', sku, uniqueKey: `2026-06-22__${sku}`, totalOrders: 1, adSpend: 220, adImpressions: 4000, adClicks: 40, revenue: 120, profitRub: -180, stock: 50 },
], [normalizeAction({ date: '2026-06-21', sku, cpcEnabled: '关闭', cpmEnabled: '开启', cpmDailyBudget: 500, source: '手动填写', ...actionOverrides })], { date: '2026-06-22', sku })[0];

const combinedText = (analysisResult) => analysisResult.recommendations.map((item) => `${item.type} ${item.reason}`).join('\n');

const minSearchAnalysis = buildPoorCpmAnalysis({ cpmPosition: '仅搜索', cpmBidType: '手动出价', cpmSearchBid: 450 }, 'CPM_SEARCH_MIN');
assert.doesNotMatch(combinedText(minSearchAnalysis), /降低搜索出价 10%|降至 3|低于 450/);
assert.match(combinedText(minSearchAnalysis), /搜索出价已是最低 450|无法继续降低出价/);

const minRecommendAnalysis = buildPoorCpmAnalysis({ cpmPosition: '仅推荐', cpmBidType: '手动出价', cpmRecommendBid: 200 }, 'CPM_RECO_MIN');
assert.doesNotMatch(combinedText(minRecommendAnalysis), /降低推荐出价|低于 200/);
assert.match(combinedText(minRecommendAnalysis), /推荐出价已是最低 200|无法继续降低出价/);

const highSearchAnalysis = buildPoorCpmAnalysis({ cpmPosition: '仅搜索', cpmBidType: '手动出价', cpmSearchBid: 600 }, 'CPM_SEARCH_HIGH');
assert.match(combinedText(highSearchAnalysis), /搜索出价 600 高于最低 450/);
assert.doesNotMatch(combinedText(highSearchAnalysis), /降至 [0-3]\d\d/);

const highRecommendAnalysis = buildPoorCpmAnalysis({ cpmPosition: '仅推荐', cpmBidType: '手动出价', cpmRecommendBid: 300 }, 'CPM_RECO_HIGH');
assert.match(combinedText(highRecommendAnalysis), /推荐出价 300 高于最低 200/);
assert.doesNotMatch(combinedText(highRecommendAnalysis), /降至 1\d\d/);

const bothMinAnalysis = buildPoorCpmAnalysis({ cpmPosition: '搜索+推荐', cpmBidType: '手动出价', cpmSearchBid: 450, cpmRecommendBid: 200 }, 'CPM_BOTH_MIN');
assert.match(combinedText(bothMinAnalysis), /搜索和推荐均已在最低出价|降低整体预算|暂停表现更差的位置/);
assert.doesNotMatch(combinedText(bothMinAnalysis), /降低搜索出价 10%|降至 [0-3]\d\d|低于 450|低于 200/);

const unifiedSearchRecommend = normalizeAction({ date: '2026-06-21', sku: 'CPM_UNIFIED_BOTH', cpcEnabled: '关闭', cpmEnabled: '开启', cpmPosition: '搜索+推荐', cpmBidType: '统一出价', cpmUnifiedBid: 449 });
assert.match(validateCpmMinBids(unifiedSearchRecommend).join('；'), /不能低于 450/);
const unifiedMinAnalysis = buildPoorCpmAnalysis({ cpmPosition: '搜索+推荐', cpmBidType: '统一出价', cpmUnifiedBid: 450 }, 'CPM_UNIFIED_MIN');
assert.match(combinedText(unifiedMinAnalysis), /统一出价已是搜索最低 450|无法继续降低统一出价/);

const recommendOnlyUnified = normalizeAction({ date: '2026-06-21', sku: 'CPM_UNIFIED_RECO', cpcEnabled: '关闭', cpmEnabled: '开启', cpmPosition: '仅推荐', cpmBidType: '统一出价', cpmUnifiedBid: 199 });
assert.match(validateCpmMinBids(recommendOnlyUnified).join('；'), /不能低于 200/);
assert.deepEqual(validateCpmMinBids({ ...recommendOnlyUnified, cpmUnifiedBid: 200 }), []);

console.log('core tests passed');
