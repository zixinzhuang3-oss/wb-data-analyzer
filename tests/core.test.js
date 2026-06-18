import assert from 'node:assert/strict';
import { toIsoDate } from '../src/utils/excel.js';
import { parseOperationActionText, normalizeAction } from '../src/utils/actions.js';

const XLSX = {
  SSF: {
    parse_date_code(value) {
      // SheetJS-compatible date objects for regression cases used by real WB files.
      if (value === 46190) return { y: 2026, m: 6, d: 17 };
      if (value === 46191) return { y: 2026, m: 6, d: 18 };
      return null;
    },
  },
};

assert.equal(toIsoDate(46190, XLSX), '2026-06-17');
assert.equal(toIsoDate(46191, XLSX), '2026-06-18');
assert.equal(toIsoDate(new Date(2026, 5, 17), XLSX), '2026-06-17');
assert.equal(toIsoDate('2026/06/17', XLSX), '2026-06-17');
assert.match(toIsoDate('6月17日', XLSX), /^\d{4}-06-17$/);

const cpmManual = parseOperationActionText('广告状态=开启；广告模式=CPM；广告位置=搜索+推荐；出价方式=手动出价；搜索出价=730；推荐出价=200；每日预算=500；预算动作=降低预算；价格动作=保持价格；主图动作=未更换；关键词动作=未调整；库存动作=库存正常；备注=广告支出过高，控费观察', '2026-06-17', 'ES030BK');
assert.equal(cpmManual.source, 'Excel 自动识别');
assert.equal(cpmManual.adMode, 'CPM');
assert.equal(cpmManual.adPosition, '搜索+推荐');
assert.equal(cpmManual.cpmSearchBid, 730);
assert.equal(cpmManual.cpmRecommendBid, 200);

const cpc = parseOperationActionText('广告状态=开启；广告模式=CPC；广告位置=搜索；搜索出价=25；每日预算=300；预算动作=保持预算；备注=搜索位保量观察', '2026-06-17', 'ES031BK');
assert.equal(cpc.adMode, 'CPC');
assert.equal(cpc.adPosition, '搜索');
assert.equal(cpc.cpcSearchBid, 25);
assert.equal(cpc.cpmRecommendBid, '');

const closed = parseOperationActionText('广告状态=关闭；广告模式=无广告；预算动作=暂停广告；备注=广告关闭观察自然流量', '2026-06-17', 'ES032BK');
assert.equal(closed.adStatus, '关闭');
assert.equal(closed.adMode, '无广告');
assert.equal(closed.cpcSearchBid, '');

assert.equal(parseOperationActionText('今天先观察自然流量，不要强行解析', '2026-06-17', 'ES033BK'), null);

const unified = normalizeAction({ date: '2026-06-17', sku: 'ES034BK', adMode: 'CPM', adPosition: '搜索+推荐', cpmBidType: '统一出价', cpmUnifiedBid: '360', cpmSearchBid: '700', cpmRecommendBid: '200' });
assert.equal(unified.cpmUnifiedBid, 360);
assert.equal(unified.cpmSearchBid, '');
assert.equal(unified.cpmRecommendBid, '');

console.log('core tests passed');
