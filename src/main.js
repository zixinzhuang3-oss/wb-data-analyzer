import { parseExcelWorkbook } from './utils/excel.js';
import { deleteAction, exportBackup, getAllActions, getAllRecords, importBackup, saveAction, saveRecords } from './utils/storage.js';
import { buildComparison, filterRecords, getDateOptions, getSkuOptions, summarizeByDate } from './utils/history.js';
import { fieldLabels } from './utils/fields.js';
import { formatMoney, formatPercent } from './utils/analysis.js';
import { ACTION_FIELDS, actionToSummary, applyAdRules, createEmptyAction, normalizeAction } from './utils/actions.js';
import { buildEffectAnalysis, metricLabels } from './utils/effectAnalysis.js';
import { buildQuickRange, getTodayDate } from './utils/date.js';

const root = document.getElementById('root');
const state = {
  records: [],
  actions: [],
  actionDraft: createEmptyAction(),
  filters: { startDate: '', endDate: '', allDates: false, sku: '' },
  today: { date: '', source: '浏览器本地时间', timeZone: 'Asia/Shanghai' },
  defaultRangeApplied: false,
  lastImport: null,
  selectedDetailKey: '',
  status: '请上传包含 SKU sheet 的 Excel 文件，系统会把“日期 + SKU”作为唯一键保存到 IndexedDB。',
  busy: false,
};

const html = (value) => String(value ?? '').replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[char]));
const formatNumber = (value) => new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 2 }).format(Number(value) || 0);

const displayFields = [
  'date', 'sku', 'adStatus', 'linkId', 'operationAction', 'price', 'reviews', 'rating', 'stock', 'reviewOrders', 'actualOrders',
  'totalOrders', 'impressions', 'clicks', 'ctr', 'addToCart', 'conversionRate', 'organicImpressions', 'organicClicks',
  'organicOrders', 'adSpend', 'adOrders', 'adShare', 'adCtr', 'adImpressions', 'adClicks', 'adClickAddToCartRate',
  'adAddToCart', 'adCostPerOrder', 'adAvgClickCost', 'revenue', 'commission', 'russiaCost',
  'deliveryFee', 'acquiringFee', 'storageFee', 'remittanceFee', 'profit', 'keywordRank',
];

const moneyFields = new Set(['price', 'adSpend', 'adCostPerOrder', 'adAvgClickCost', 'revenue', 'commission', 'russiaCost', 'deliveryFee', 'acquiringFee', 'storageFee', 'remittanceFee', 'profit']);
const percentFields = new Set(['ctr', 'conversionRate', 'adShare', 'adCtr', 'adClickAddToCartRate', 'orderConversionRate']);

const renderValue = (key, value) => {
  if (moneyFields.has(key)) return formatMoney(value);
  if (percentFields.has(key)) return formatPercent(value);
  if (key === 'operationAction') return html(value || '-');
  return html(value || value === 0 ? formatNumber(value) : '-');
};

const downloadJson = (filename, data) => {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
};

async function refreshRecords() {
  const [records, actions] = await Promise.all([getAllRecords(), getAllActions()]);
  state.records = records;
  state.actions = actions;
}

async function handleExcelUpload(file) {
  if (!file) return;
  state.busy = true;
  state.status = `正在解析 ${file.name} ...`;
  render();
  try {
    const result = await parseExcelWorkbook(file);
    if (!result.records.length) throw new Error('没有识别到有效 SKU sheet 或有效每日数据行。');
    await saveRecords(result.records);
    if (result.actions?.length) await Promise.all(result.actions.map(saveAction));
    await refreshRecords();
    applyDefaultQuickRange();
    state.lastImport = {
      fileName: file.name,
      savedCount: result.records.length,
      actionCount: result.actions?.length || 0,
      skuSheets: result.skuSheets,
      skippedSheets: result.skippedSheets,
      diagnostics: result.diagnostics || [],
    };
    state.status = `导入完成：保存/覆盖 ${result.records.length} 行，自动识别动作 ${result.actions?.length || 0} 条，识别 SKU sheet ${result.skuSheets.length} 个。`;
  } catch (error) {
    state.status = `导入失败：${error.message || error}`;
  } finally {
    state.busy = false;
    render();
  }
}

function renderFieldDiagnostics() {
  const diagnostics = state.lastImport?.diagnostics || [];
  if (!diagnostics.length) return '';
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">⌕</span><div><h2>字段识别诊断</h2><p>整体流量与广告流量优先按固定列位读取，避免字段名变化导致读取为 0。</p></div></div>
    <div class="record-list">${diagnostics.map((item) => `<article class="record-card">
      <strong>${html(item.sheetName)}</strong>
      ${Object.entries(item.fields).map(([label, column]) => `<span>${html(label)}：${html(column)}</span>`).join('')}
      ${item.blankAdDates?.length ? `<span>广告空白日期：${item.blankAdDates.map(html).join('、')}，该日期广告数据为空，已识别为广告关闭/无广告数据</span>` : '<span>未发现广告列 AB-AK 全空的数据行</span>'}
    </article>`).join('')}</div></section>`;
}

async function handleBackupImport(file) {
  if (!file) return;
  state.busy = true;
  state.status = `正在导入备份 ${file.name} ...`;
  render();
  try {
    const backup = JSON.parse(await file.text());
    const result = await importBackup(backup);
    await refreshRecords();
    applyDefaultQuickRange();
    state.status = `备份导入完成：恢复 ${result.records} 条历史记录、${result.actions} 条动作记录。`; 
  } catch (error) {
    state.status = `备份导入失败：${error.message || error}`;
  } finally {
    state.busy = false;
    render();
  }
}

async function handleBackupExport() {
  const backup = await exportBackup();
  downloadJson(`wb-data-backup-${new Date().toISOString().slice(0, 10)}.json`, backup);
  state.status = `已导出备份 JSON：${backup.records.length} 条历史记录、${backup.actions.length} 条动作记录。`;
  render();
}

const renderOptions = (options, current, placeholder) => [`<option value="">${placeholder}</option>`, ...options.map((option) => `<option value="${html(option)}" ${option === current ? 'selected' : ''}>${html(option)}</option>`)].join('');

function renderImportSummary() {
  if (!state.lastImport) return '<p class="empty-state">尚未导入 Excel。本阶段会自动跳过 wb利润定价表、ozon利润定价表、Sheet10 等辅助 sheet。</p>';
  const { fileName, savedCount, actionCount = 0, skuSheets, skippedSheets } = state.lastImport;
  return `<div class="import-result">
    <strong>${html(fileName)}</strong>
    <span>保存/覆盖 ${savedCount} 行</span>
    <span>Excel 自动识别动作：${actionCount} 条</span>
    <span>识别 SKU：${skuSheets.map(html).join('、') || '-'}</span>
    <span>跳过辅助 sheet：${skippedSheets.map(html).join('、') || '-'}</span>
  </div>`;
}

function renderHistoryCards(records) {
  const rows = summarizeByDate(records);
  if (!rows.length) return '<p class="empty-state">暂无导入数据记录。</p>';
  return `<div class="record-list">${rows.map((row) => `<article class="record-card">
    <strong>${row.date}</strong>
    <span>SKU 数量：${row.skuCount}</span>
    <span>总订单：${formatNumber(row.totalOrders)}</span>
    <span>总广告费：${formatMoney(row.totalAdSpend)}</span>
    <span>总利润：${formatMoney(row.totalProfit)}</span>
  </article>`).join('')}</div>`;
}

function renderTable(records) {
  if (!records.length) return '<p class="empty-state">当前时间段无数据</p>';
  const actions = actionMap();
  const header = `<th>详情/动作</th>${displayFields.map((key) => `<th>${fieldLabels[key]}</th>`).join('')}<th>结构化动作</th>`;
  const rows = records.map((record) => {
    const action = actions.get(record.uniqueKey);
    return `<tr>
      <td><button class="table-action" data-detail-key="${html(record.uniqueKey)}" type="button">查看详情</button></td>
      ${displayFields.map((key) => `<td>${renderValue(key, record[key])}</td>`).join('')}
      <td>${html(actionToSummary(action))}</td>
    </tr>`;
  }).join('');
  return `<div class="table-wrap history-table"><table><thead><tr>${header}</tr></thead><tbody>${rows}</tbody></table></div>`;
}


const actionMap = () => new Map(state.actions.map((action) => [action.uniqueKey, action]));

const getCurrentActionKey = () => `${state.actionDraft.date || ''}__${state.actionDraft.sku || ''}`;

const setActionDraftFromKey = (uniqueKey) => {
  const action = actionMap().get(uniqueKey);
  const record = state.records.find((item) => item.uniqueKey === uniqueKey);
  const [date = '', sku = ''] = uniqueKey.split('__');
  state.actionDraft = action ? { ...action } : createEmptyAction(record?.date || date, record?.sku || sku);
};

async function handleActionSave() {
  const action = normalizeAction(state.actionDraft);
  if (!action.date || !action.sku) {
    state.status = '动作保存失败：请先选择日期和 SKU。';
    render();
    return;
  }
  await saveAction(action);
  await refreshRecords();
  state.selectedDetailKey = action.uniqueKey;
  state.actionDraft = { ...action };
  state.status = `动作记录已保存：${action.date} ${action.sku}。`;
  render();
}

async function handleActionDelete(uniqueKey = getCurrentActionKey()) {
  if (!uniqueKey || uniqueKey === '__') {
    state.status = '动作删除失败：请先选择日期和 SKU。';
    render();
    return;
  }
  await deleteAction(uniqueKey);
  await refreshRecords();
  const [date, sku] = uniqueKey.split('__');
  state.actionDraft = createEmptyAction(date, sku);
  state.status = `动作记录已删除：${date} ${sku}。`;
  render();
}

function actionFieldVisible(field, draft) {
  if (['cpcSearchBid', 'cpcDailyBudget'].includes(field.key)) return draft.cpcEnabled === '开启';
  if (['cpmPosition', 'cpmBidType', 'cpmDailyBudget'].includes(field.key)) return draft.cpmEnabled === '开启';
  if (field.key === 'cpmSearchBid') return draft.cpmEnabled === '开启' && draft.cpmBidType === '手动出价' && ['仅搜索', '搜索+推荐'].includes(draft.cpmPosition);
  if (field.key === 'cpmRecommendBid') return draft.cpmEnabled === '开启' && draft.cpmBidType === '手动出价' && ['仅推荐', '搜索+推荐'].includes(draft.cpmPosition);
  if (field.key === 'cpmUnifiedBid') return draft.cpmEnabled === '开启' && draft.cpmBidType === '统一出价';
  return true;
}

function renderActionField(field) {
  const draft = applyAdRules(state.actionDraft);
  if (!actionFieldVisible(field, draft)) return '';
  const value = draft[field.key] ?? '';
  if (field.type === 'textarea') {
    return `<label class="form-field wide-field"><span>${field.label}</span><textarea data-action-field="${field.key}" rows="3">${html(value)}</textarea></label>`;
  }
  if (field.type === 'number') {
    return `<label class="form-field"><span>${field.label}</span><input data-action-field="${field.key}" type="number" step="0.01" value="${html(value)}" /></label>`;
  }
  const disabled = field.disabled ? 'disabled' : '';
  return `<label class="form-field"><span>${field.label}</span><select data-action-field="${field.key}" ${disabled}><option value="">请选择</option>${field.options.map((option) => `<option value="${html(option)}" ${option === value ? 'selected' : ''}>${html(option)}</option>`).join('')}</select></label>`;
}

function renderActionModule(dates, skus) {
  const currentKey = getCurrentActionKey();
  const existing = actionMap().get(currentKey);
  const groups = ['基础字段', 'CPC 模块', 'CPM 模块'].map((group) => `<fieldset class="action-fieldset"><legend>${group}</legend><div class="action-form-grid">${ACTION_FIELDS.filter((field) => field.group === group).map(renderActionField).join('')}</div></fieldset>`).join('');
  return `<section class="panel action-panel">
    <div class="panel-heading"><span class="panel-icon">✎</span><div><h2>每日动作记录</h2><p>按“日期 + SKU”记录基础动作，并把 CPC 模块与 CPM 模块独立维护；整体广告状态会根据 CPC/CPM 开关自动计算。</p></div></div>
    <div class="action-form-grid">
      <label class="form-field"><span>日期</span><select id="action-date"><option value="">选择日期</option>${dates.map((date) => `<option value="${html(date)}" ${date === state.actionDraft.date ? 'selected' : ''}>${html(date)}</option>`).join('')}</select></label>
      <label class="form-field"><span>SKU</span><select id="action-sku"><option value="">选择 SKU</option>${skus.map((sku) => `<option value="${html(sku)}" ${sku === state.actionDraft.sku ? 'selected' : ''}>${html(sku)}</option>`).join('')}</select></label>
    </div>
    ${groups}
    <div class="action-row form-actions">
      <button id="save-action" type="button">${existing ? '更新动作记录' : '保存动作记录'}</button>
      <button id="delete-action" type="button" ${existing ? '' : 'disabled'}>删除动作记录</button>
      <button id="clear-action" type="button">清空表单</button>
    </div>
  </section>`;
}

function renderSkuDetail(record, action) {
  if (!record) return '<p class="empty-state">点击历史数据明细中的“查看详情”，即可同时查看每日数据和当天动作。</p>';
  return `<div class="detail-box">
    <h3>${html(record.date)} · ${html(record.sku)}</h3>
    <div class="detail-grid">
      <div><span>总订单</span><strong>${formatNumber(record.totalOrders)}</strong></div>
      <div><span>广告费</span><strong>${formatMoney(record.adSpend)}</strong></div>
      <div><span>销售额</span><strong>${formatMoney(record.revenue)}</strong></div>
      <div><span>利润</span><strong>${formatMoney(record.profit)}</strong></div>
      <div><span>整体广告状态</span><strong>${html(action?.adStatus || record.adStatus || '-')}</strong></div>
      <div><span>CPC 状态</span><strong>${html(action?.cpcEnabled || '-')}</strong></div>
      <div><span>CPC 搜索出价</span><strong>${html(action?.cpcSearchBid ?? '-')}</strong></div>
      <div><span>CPC 预算</span><strong>${html(action?.cpcDailyBudget ?? '-')}</strong></div>
      <div><span>CPM 状态</span><strong>${html(action?.cpmEnabled || '-')}</strong></div>
      <div><span>CPM 投放位置</span><strong>${html(action?.cpmPosition || '-')}</strong></div>
      <div><span>CPM 出价方式</span><strong>${html(action?.cpmBidType || '-')}</strong></div>
      <div><span>CPM 搜索出价</span><strong>${html(action?.cpmSearchBid ?? '-')}</strong></div>
      <div><span>CPM 推荐出价</span><strong>${html(action?.cpmRecommendBid ?? '-')}</strong></div>
      <div><span>CPM 统一出价</span><strong>${html(action?.cpmUnifiedBid ?? '-')}</strong></div>
      <div><span>CPM 预算</span><strong>${html(action?.cpmDailyBudget ?? '-')}</strong></div>
    </div>
    <section class="detail-section"><h4>Excel 运营动作原文</h4><p>${html(record.operationAction || 'Excel 中未填写运营动作原文')}</p></section>
    <section class="detail-section"><h4>系统结构化动作</h4><p>${html(actionToSummary(action))}</p></section>
  </div>`;
}


function renderStrategyBoard(analyses, comparison) {
  if (!analyses.length) return '<section class="panel"><h2>明日策略建议看板</h2><p class="empty-state">导入历史数据后将自动生成明日策略建议。</p></section>';
  const context = comparison?.hasPreviousData ? `当前区间 ${rangeText(comparison.currentRange)} 对比 ${rangeText(comparison.previousRange)}：订单变化 ${formatPercent(deltaInfo(comparison.current.totalOrders, comparison.previous.totalOrders).rate)}，广告费变化 ${formatPercent(deltaInfo(comparison.current.totalAdSpend, comparison.previous.totalAdSpend).rate)}，利润变化 ${formatMoney(deltaInfo(comparison.current.totalProfit, comparison.previous.totalProfit).value)}。明日策略会结合该区间趋势与每日动作。` : '当前筛选区间无对比数据，明日策略主要结合当前区间、最近日期和每日动作生成。';
  return `<section class="panel strategy-board"><div class="panel-heading"><span class="panel-icon">★</span><div><h2>明日策略建议看板</h2><p>结合当前选择时间段、上一同长度时间段和每日动作，判断动作是否有效，并输出明日广告与运营建议。</p></div></div><p class="strategy-context">${html(context)}</p><div class="recommendation-grid">${analyses.map((item) => `<article class="recommendation-card priority-${item.primaryRecommendation.priority}">
    <div class="recommendation-head"><strong>${html(item.sku)}</strong><span>${html(item.date)}</span></div>
    <h3>${html(item.primaryRecommendation.type)}</h3>
    <p>${html(item.primaryRecommendation.reason)}</p>
  </article>`).join('')}</div></section>`;
}

function renderEffectCards(analyses) {
  if (!analyses.length) return '';
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">↗</span><div><h2>动作效果分析卡片</h2><p>对比今天 vs 昨天、近 3 天/7 天均值，以及动作前 3 天 vs 动作后 3 天。</p></div></div><div class="effect-grid">${analyses.map((item) => {
    const effectText = item.effects.length ? item.effects.map((effect) => `${effect.level}：${effect.text}`).join(' ') : '暂无明确动作效果，建议继续观察。';
    const yesterdayAction = item.previousAction || item.latestAction;
    const resultText = `昨天动作 → 今天结果：${actionToSummary(yesterdayAction)}；今天订单 ${formatNumber(item.metrics.totalOrders.today)}，广告费 ${formatMoney(item.metrics.adSpend.today)}，利润 ${formatMoney(item.metrics.profit.today)}。`;
    return `<article class="effect-card"><div class="recommendation-head"><strong>${html(item.sku)}</strong><span>${html(actionToSummary(item.latestAction))}</span></div>
      <p>${html(resultText)}</p>
      <p>${html(effectText)}</p>
      <div class="mini-metrics">
        <span>订单 ${formatNumber(item.metrics.totalOrders.today)} / 昨日 ${formatNumber(item.metrics.totalOrders.yesterday)}</span>
        <span>广告费 ${formatMoney(item.metrics.adSpend.today)} / 近3天 ${formatMoney(item.metrics.adSpend.last3Avg)}</span>
        <span>利润 ${formatMoney(item.metrics.profit.today)} / 动作前后 ${formatMoney(item.metrics.profit.actionWindowDelta.value)}</span>
        <span>ROI ${formatNumber(item.metrics.roi.today)} · ACOS ${formatPercent(item.metrics.acos.today)}</span>
      </div>
    </article>`;
  }).join('')}</div></section>`;
}

function renderRiskPanel(analyses) {
  const risks = analyses.flatMap((item) => item.risks.map((risk) => ({ sku: item.sku, risk })));
  if (!risks.length) return `<section class="panel"><div class="panel-heading compact"><span class="panel-icon">!</span><h2>SKU 风险提示</h2></div><p class="empty-state">当前筛选范围内暂无高风险提示。</p></section>`;
  return `<section class="panel risk-panel"><div class="panel-heading compact"><span class="panel-icon">!</span><h2>SKU 风险提示</h2></div><div class="risk-list">${risks.map((item) => `<div class="risk-item"><strong>${html(item.sku)}</strong><span>${html(item.risk)}</span></div>`).join('')}</div></section>`;
}

function renderSuggestionHistory(analyses, dates, skus) {
  const rows = analyses.flatMap((item) => item.recommendations.map((rec) => ({ date: item.date, sku: item.sku, ...rec })));
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">☷</span><div><h2>建议历史</h2><p>建议历史跟随上方日期和 SKU 筛选，可查看每个 SKU 的建议类型和原因。</p></div></div>${rows.length ? `<div class="table-wrap"><table><thead><tr><th>日期</th><th>SKU</th><th>建议类型</th><th>优先级</th><th>原因</th></tr></thead><tbody>${rows.map((row) => `<tr><td>${html(row.date)}</td><td>${html(row.sku)}</td><td>${html(row.type)}</td><td>${html(row.priority)}</td><td>${html(row.reason)}</td></tr>`).join('')}</tbody></table></div>` : '<p class="empty-state">当前筛选范围暂无建议。</p>'}</section>`;
}

const metricConfig = [
  ['总订单', 'totalOrders', formatNumber],
  ['总销售额', 'totalRevenue', formatMoney],
  ['总利润', 'totalProfit', formatMoney],
  ['利润率', 'margin', formatPercent],
  ['ROI', 'roi', formatNumber],
  ['ACOS', 'acos', formatPercent],
  ['总曝光', 'impressions', formatNumber],
  ['总点击', 'clicks', formatNumber],
  ['整体 CTR', 'ctr', formatPercent],
  ['总加购', 'addToCart', formatNumber],
  ['整体加购转化率', 'cvr', formatPercent],
  ['订单转化率', 'orderConversionRate', formatPercent],
  ['广告状态', 'adStatus', html],
  ['总广告曝光', 'adImpressions', formatNumber],
  ['总广告点击', 'adClicks', formatNumber],
  ['广告 CTR', 'adCtr', formatPercent],
  ['总广告加购', 'adAddToCart', formatNumber],
  ['广告点击转加购率', 'adClickAddToCartRate', formatPercent],
  ['广告订单', 'adOrders', formatNumber],
  ['广告费', 'totalAdSpend', formatMoney],
  ['广告费占比', 'adShare', formatPercent],
  ['每单费用', 'adCostPerOrder', formatMoney],
  ['广告平均点击费', 'adAvgClickCost', formatMoney],
  ['SKU 数量', 'skuCount', formatNumber],
  ['日期数量', 'dateCount', formatNumber],
];

const deltaInfo = (current, previous) => {
  const value = (Number(current) || 0) - (Number(previous) || 0);
  const rate = previous ? value / Math.abs(previous) : 0;
  return { value, rate, trend: Math.abs(value) < 0.000001 ? '持平' : value > 0 ? '上升' : '下降' };
};

const rangeText = (range) => range?.allDates ? '全部日期' : `${range.startDate || '-'} 至 ${range.endDate || '-'}`;

function applyDefaultQuickRange() {
  if (!state.today.date) return;
  const nextRange = buildQuickRange('7', state.today.date);
  state.filters = { ...state.filters, ...nextRange };
  state.defaultRangeApplied = true;
}

function renderDateContext(comparison) {
  return `<div class="date-context">
    <strong>当前日期基准：${html(state.today.date || '-')}</strong>
    <strong>日期来源：${html(state.today.source || '浏览器本地时间')}</strong>
    <strong>当前时间段：${html(rangeText(comparison.currentRange))}</strong>
    <strong>对比时间段：${comparison.previousRange ? html(rangeText(comparison.previousRange)) : '全部日期暂无上期对比'}</strong>
  </div>`;
}

function renderComparisonMetrics(comparison) {
  return `${renderDateContext(comparison)}
  <div class="metrics-grid comparison-grid">${metricConfig.map(([label, key, formatter]) => {
    const current = comparison.current[key];
    const previous = comparison.previous[key];
    const diff = deltaInfo(current, previous);
    const noPrevious = !comparison.hasPreviousData;
    return `<div class="metric-card comparison-card"><span>${label}</span><strong>${formatter(current)}</strong><small>上期：${noPrevious ? '无对比数据' : formatter(previous)}</small><small>变化：${noPrevious ? '无对比数据' : `${formatter(diff.value)} / ${formatPercent(diff.rate)}`}</small><small>趋势：${noPrevious ? '无对比数据' : diff.trend}</small></div>`;
  }).join('')}</div>`;
}

function judgeSku(current, previous) {
  const order = deltaInfo(current.totalOrders, previous.totalOrders);
  const spend = deltaInfo(current.totalAdSpend, previous.totalAdSpend);
  const profit = deltaInfo(current.totalProfit, previous.totalProfit);
  if ((current.stock || 0) > 0 && current.stock < Math.max(5, current.totalOrders * 2)) return '库存风险';
  if (spend.value > 0 && order.value <= 0) return '广告费上升但订单未增长';
  if (profit.value < 0) return '利润下降';
  if (spend.value < 0 && order.value >= 0 && profit.value > 0) return '控费有效';
  if (order.value > 0 && profit.value > 0 && current.roi > 1.5) return '建议加预算';
  if (current.totalAdSpend > 0 && current.totalOrders === 0 && current.totalProfit < 0) return '建议暂停广告';
  if (order.value > 0 || profit.value > 0) return '表现改善';
  return '需要观察';
}

function buildSkuRows(comparison) {
  const skus = [...new Set([...comparison.currentRecords, ...comparison.previousRecords].map((row) => row.sku).filter(Boolean))].sort();
  return skus.map((sku) => {
    const current = comparison.currentRecords.filter((row) => row.sku === sku);
    const previous = comparison.previousRecords.filter((row) => row.sku === sku);
    const currentSummary = summarizeSku(current);
    const previousSummary = summarizeSku(previous);
    return { sku, current: currentSummary, previous: previousSummary, judge: comparison.hasPreviousData ? judgeSku(currentSummary, previousSummary) : '无对比数据' };
  });
}

function summarizeSku(rows) {
  const totalOrders = rows.reduce((sum, row) => sum + (Number(row.totalOrders) || 0), 0);
  const totalRevenue = rows.reduce((sum, row) => sum + (Number(row.revenue) || 0), 0);
  const totalAdSpend = rows.reduce((sum, row) => sum + (Number(row.adSpend) || 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + (Number(row.profit) || 0), 0);
  const stock = rows.at(-1)?.stock || 0;
  return { totalOrders, totalRevenue, totalAdSpend, totalProfit, stock, roi: totalAdSpend ? totalRevenue / totalAdSpend : 0, acos: totalRevenue ? totalAdSpend / totalRevenue : 0 };
}

function renderSkuComparison(comparison) {
  const rows = buildSkuRows(comparison);
  if (!rows.length) return '<section class="panel"><h2>SKU 区间对比</h2><p class="empty-state">当前筛选范围暂无 SKU 数据。</p></section>';
  const change = (a, b, formatter = formatNumber) => formatter((Number(a) || 0) - (Number(b) || 0));
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">⇄</span><div><h2>SKU 区间对比</h2><p>${state.filters.sku ? '当前为单个 SKU 区间趋势。' : '全部 SKU 按当前区间与上一同长度区间逐项对比。'}</p></div></div><div class="table-wrap"><table><thead><tr><th>SKU</th><th>当前订单</th><th>上期订单</th><th>订单变化</th><th>当前销售额</th><th>上期销售额</th><th>销售额变化</th><th>当前广告费</th><th>上期广告费</th><th>广告费变化</th><th>当前利润</th><th>上期利润</th><th>利润变化</th><th>当前 ROI</th><th>上期 ROI</th><th>当前 ACOS</th><th>上期 ACOS</th><th>系统判断</th></tr></thead><tbody>${rows.map((row) => `<tr><td>${html(row.sku)}</td><td>${formatNumber(row.current.totalOrders)}</td><td>${comparison.hasPreviousData ? formatNumber(row.previous.totalOrders) : '无对比数据'}</td><td>${comparison.hasPreviousData ? change(row.current.totalOrders, row.previous.totalOrders) : '无对比数据'}</td><td>${formatMoney(row.current.totalRevenue)}</td><td>${comparison.hasPreviousData ? formatMoney(row.previous.totalRevenue) : '无对比数据'}</td><td>${comparison.hasPreviousData ? change(row.current.totalRevenue, row.previous.totalRevenue, formatMoney) : '无对比数据'}</td><td>${formatMoney(row.current.totalAdSpend)}</td><td>${comparison.hasPreviousData ? formatMoney(row.previous.totalAdSpend) : '无对比数据'}</td><td>${comparison.hasPreviousData ? change(row.current.totalAdSpend, row.previous.totalAdSpend, formatMoney) : '无对比数据'}</td><td>${formatMoney(row.current.totalProfit)}</td><td>${comparison.hasPreviousData ? formatMoney(row.previous.totalProfit) : '无对比数据'}</td><td>${comparison.hasPreviousData ? change(row.current.totalProfit, row.previous.totalProfit, formatMoney) : '无对比数据'}</td><td>${formatNumber(row.current.roi)}</td><td>${comparison.hasPreviousData ? formatNumber(row.previous.roi) : '无对比数据'}</td><td>${formatPercent(row.current.acos)}</td><td>${comparison.hasPreviousData ? formatPercent(row.previous.acos) : '无对比数据'}</td><td>${html(row.judge)}</td></tr>`).join('')}</tbody></table></div></section>`;
}

function renderIntervalSummary(comparison) {
  if (!comparison.hasPreviousData) return `<section class="panel"><h2>区间对比总结</h2><p class="empty-state">当前时间段 ${html(rangeText(comparison.currentRange))}，无对比数据。</p></section>`;
  const order = deltaInfo(comparison.current.totalOrders, comparison.previous.totalOrders);
  const spend = deltaInfo(comparison.current.totalAdSpend, comparison.previous.totalAdSpend);
  const profit = deltaInfo(comparison.current.totalProfit, comparison.previous.totalProfit);
  let advice = '建议继续观察核心 SKU 的订单、广告费和利润变化。';
  if (spend.value < 0 && Math.abs(order.rate) <= 0.1 && profit.value > 0) advice = '说明控费效果较好，建议继续保持当前预算策略。';
  else if (spend.rate > 0.2 && order.rate < 0.1 && profit.value < 0) advice = '说明广告放量效率较差，建议重点检查高花费 SKU，并降低 ROI 差的广告预算。';
  else if (order.value > 0 && profit.value > 0 && comparison.current.roi > 1.5) advice = '说明增长质量较好，库存充足 SKU 可适当加预算。';
  return `<section class="panel"><h2>区间对比总结</h2><p>当前时间段 ${html(rangeText(comparison.currentRange))}，较上一时间段订单${order.value >= 0 ? '增加' : '减少'} ${formatPercent(Math.abs(order.rate))}，广告费${spend.value >= 0 ? '增加' : '下降'} ${formatPercent(Math.abs(spend.rate))}，利润${profit.value >= 0 ? '增加' : '下降'} ${formatMoney(Math.abs(profit.value))}，${advice}</p></section>`;
}

function render() {
  if (!state.defaultRangeApplied && state.today.date && !state.filters.startDate && !state.filters.endDate && !state.filters.allDates) applyDefaultQuickRange();
  const filtered = filterRecords(state.records, state.filters);
  const comparison = buildComparison(state.records, state.filters);
  const dates = getDateOptions(state.records);
  const skus = getSkuOptions(state.records);
  const actions = actionMap();
  const selectedRecord = state.records.find((record) => record.uniqueKey === state.selectedDetailKey) || filtered[0];
  const selectedAction = selectedRecord ? actions.get(selectedRecord.uniqueKey) : null;
  const effectAnalyses = buildEffectAnalysis(state.records, state.actions, state.filters);

  root.innerHTML = `<main class="app-shell">
    <section class="hero">
      <div>
        <p class="eyebrow">WB Daily Operations Review</p>
        <h1>WB 每日运营复盘与广告策略决策系统</h1>
        <p>第一阶段已支持 Excel 模板解析、SKU sheet 自动识别、历史数据 IndexedDB 保存、重复日期 + SKU 覆盖更新，以及 JSON 备份恢复。</p>
      </div>
      <div class="hero-badge"><span class="hero-icon">▦</span><span>Excel 导入 · 历史保存 · 备份恢复</span></div>
    </section>

    <section class="panel import-panel">
      <div class="panel-heading"><span class="panel-icon">⇧</span><div><h2>导入 Excel 每日数据</h2><p>支持 .xlsx/.xls；只读取形如 ES032BK 的 SKU sheet，并自动过滤利润定价表、Sheet10 等辅助 sheet。</p></div></div>
      <div class="action-row">
        <label class="primary-upload"><input id="excel-upload" type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" />选择 Excel 文件</label>
        <button id="export-backup" type="button" ${state.records.length ? '' : 'disabled'}>导出备份 JSON</button>
        <label class="secondary-upload"><input id="backup-upload" type="file" accept=".json,application/json" />导入备份 JSON</label>
      </div>
      <p class="status-line ${state.status.includes('失败') ? 'error-text' : ''}">${html(state.status)}</p>
      ${renderImportSummary()}
    </section>
    ${renderFieldDiagnostics()}

    <section class="panel">
      <div class="panel-heading"><span class="panel-icon">▤</span><div><h2>筛选与汇总</h2><p>快捷时间段基于真实当前日期（优先网络时间，失败后使用浏览器本地时间），不会使用 Excel 最大日期作为今天。</p></div></div>
      <div class="filter-row">
        <label class="form-field"><span>开始日期</span><input id="start-date-filter" type="date" value="${html(state.filters.startDate)}" /></label>
        <label class="form-field"><span>结束日期</span><input id="end-date-filter" type="date" value="${html(state.filters.endDate)}" /></label>
        <select id="sku-filter">${renderOptions(skus, state.filters.sku, '全部 SKU')}</select>
        <div class="quick-range-row">
          <button data-range-preset="today" type="button">今天</button>
          <button data-range-preset="yesterday" type="button">昨天</button>
          <button data-range-preset="3" type="button">最近 3 天</button>
          <button data-range-preset="7" type="button">最近 7 天</button>
          <button data-range-preset="14" type="button">最近 14 天</button>
          <button data-range-preset="30" type="button">最近 30 天</button>
          <button data-range-preset="all" type="button">全部日期</button>
        </div>
      </div>
      ${renderComparisonMetrics(comparison)}
    </section>

    ${renderIntervalSummary(comparison)}
    ${renderSkuComparison(comparison)}
    ${renderActionModule(dates, skus)}
    ${renderStrategyBoard(effectAnalyses, comparison)}
    ${renderEffectCards(effectAnalyses)}
    ${renderRiskPanel(effectAnalyses)}
    ${renderSuggestionHistory(effectAnalyses, dates, skus)}

    <section class="content-grid">
      <div class="panel wide"><h2>历史数据明细</h2>${renderTable(filtered)}</div>
      <aside class="panel strategy-panel"><div class="panel-heading compact"><span class="panel-icon">◎</span><h2>导入数据记录</h2></div>${renderHistoryCards(state.records)}<div class="sku-detail"><h2>SKU 详情</h2>${renderSkuDetail(selectedRecord, selectedAction)}</div></aside>
    </section>
  </main>`;

  document.getElementById('excel-upload')?.addEventListener('change', (event) => handleExcelUpload(event.target.files?.[0]));
  document.getElementById('backup-upload')?.addEventListener('change', (event) => handleBackupImport(event.target.files?.[0]));
  document.getElementById('export-backup')?.addEventListener('click', handleBackupExport);
  document.getElementById('start-date-filter')?.addEventListener('change', (event) => { state.filters.startDate = event.target.value; state.filters.allDates = false; render(); });
  document.getElementById('end-date-filter')?.addEventListener('change', (event) => { state.filters.endDate = event.target.value; state.filters.allDates = false; render(); });
  document.getElementById('sku-filter')?.addEventListener('change', (event) => { state.filters.sku = event.target.value; render(); });
  document.querySelectorAll('[data-range-preset]').forEach((button) => button.addEventListener('click', (event) => {
    const preset = event.target.dataset.rangePreset;
    state.filters = { ...state.filters, ...buildQuickRange(preset, state.today.date) };
    state.defaultRangeApplied = true;
    render();
  }));
  document.getElementById('action-date')?.addEventListener('change', (event) => { state.actionDraft.date = event.target.value; setActionDraftFromKey(getCurrentActionKey()); render(); });
  document.getElementById('action-sku')?.addEventListener('change', (event) => { state.actionDraft.sku = event.target.value; setActionDraftFromKey(getCurrentActionKey()); render(); });
  document.querySelectorAll('[data-action-field]').forEach((input) => {
    input.addEventListener('input', (event) => { state.actionDraft[event.target.dataset.actionField] = event.target.value; });
    input.addEventListener('change', (event) => { state.actionDraft[event.target.dataset.actionField] = event.target.value; state.actionDraft.source = state.actionDraft.source === 'Excel 自动识别' ? '手动修改' : state.actionDraft.source; state.actionDraft = applyAdRules(state.actionDraft); render(); });
  });
  document.getElementById('save-action')?.addEventListener('click', handleActionSave);
  document.getElementById('delete-action')?.addEventListener('click', () => handleActionDelete());
  document.getElementById('clear-action')?.addEventListener('click', () => { state.actionDraft = createEmptyAction(state.actionDraft.date, state.actionDraft.sku); render(); });
  document.querySelectorAll('[data-detail-key]').forEach((button) => button.addEventListener('click', (event) => {
    state.selectedDetailKey = event.target.dataset.detailKey;
    setActionDraftFromKey(state.selectedDetailKey);
    render();
  }));
}

Promise.all([refreshRecords(), getTodayDate()]).then(([, today]) => {
  state.today = today;
  applyDefaultQuickRange();
  render();
}).catch((error) => {
  state.status = `初始化失败：${error.message || error}`;
  render();
});
