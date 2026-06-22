import { parseExcelWorkbook } from './utils/excel.js';
import { deleteAction, exportBackup, getAllActions, getAllRecords, importBackup, saveAction, saveRecords } from './utils/storage.js';
import { buildPeriodComparison, buildPeriodSummaryText, filterRecords, getDateOptions, getSkuOptions, summarizeByDate } from './utils/history.js';
import { fieldLabels } from './utils/fields.js';
import { formatMoney, formatPercent } from './utils/analysis.js';
import { ACTION_FIELDS, actionToSummary, applyAdRules, createEmptyAction, normalizeAction } from './utils/actions.js';
import { buildEffectAnalysis, metricLabels } from './utils/effectAnalysis.js';

const root = document.getElementById('root');
const state = {
  records: [],
  actions: [],
  actionDraft: createEmptyAction(),
  filters: { startDate: '', endDate: '', sku: '', quickRange: '今天' },
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
  'organicOrders', 'adSpend', 'adOrders', 'adShare', 'adImpressions', 'adClicks', 'revenue', 'commission', 'russiaCost',
  'deliveryFee', 'acquiringFee', 'storageFee', 'remittanceFee', 'profit', 'keywordRank',
];

const moneyFields = new Set(['price', 'adSpend', 'revenue', 'commission', 'russiaCost', 'deliveryFee', 'acquiringFee', 'storageFee', 'remittanceFee', 'profit']);
const percentFields = new Set(['ctr', 'conversionRate', 'adShare']);

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
    state.lastImport = {
      fileName: file.name,
      savedCount: result.records.length,
      actionCount: result.actions?.length || 0,
      skuSheets: result.skuSheets,
      skippedSheets: result.skippedSheets,
    };
    state.status = `导入完成：保存/覆盖 ${result.records.length} 行，自动识别动作 ${result.actions?.length || 0} 条，识别 SKU sheet ${result.skuSheets.length} 个。`;
  } catch (error) {
    state.status = `导入失败：${error.message || error}`;
  } finally {
    state.busy = false;
    render();
  }
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
const QUICK_RANGES = ['今天', '昨天', '最近 3 天', '最近 7 天', '最近 14 天', '最近 30 天', '全部日期'];

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


const formatByType = (value, type) => {
  if (type === 'money') return formatMoney(value);
  if (type === 'percent') return formatPercent(value);
  return formatNumber(value);
};

const formatDelta = (value, type) => {
  const prefix = Number(value) > 0 ? '+' : '';
  if (type === 'money') return `${prefix}${formatMoney(value)}`;
  if (type === 'percent') return `${prefix}${formatPercent(value)}`;
  return `${prefix}${formatNumber(value)}`;
};

function applyQuickRange(range) {
  state.filters.quickRange = range;
  const dates = getDateOptions(state.records).sort();
  const latest = dates.at(-1) || new Date().toISOString().slice(0, 10);
  if (range === '全部日期') {
    state.filters.startDate = dates[0] || latest;
    state.filters.endDate = latest;
    render();
    return;
  }
  const endDate = range === '昨天' ? addDateForUi(latest, -1) : latest;
  const lengthMap = { 今天: 1, 昨天: 1, '最近 3 天': 3, '最近 7 天': 7, '最近 14 天': 14, '最近 30 天': 30 };
  const days = lengthMap[range] || 7;
  state.filters.endDate = endDate;
  state.filters.startDate = addDateForUi(endDate, -(days - 1));
  render();
}

function addDateForUi(date, days) {
  const [year, month, day] = String(date || '').split('-').map(Number);
  if (!year || !month || !day) return '';
  const parsed = new Date(Date.UTC(year, month - 1, day));
  parsed.setUTCDate(parsed.getUTCDate() + days);
  return parsed.toISOString().slice(0, 10);
}

function ensureDefaultDateRange(dates) {
  if (state.filters.startDate || state.filters.endDate || !dates.length) return;
  const sorted = [...dates].sort();
  const latest = sorted.at(-1);
  state.filters.endDate = latest;
  state.filters.startDate = latest;
}

function renderMetricCard(metric, hasComparison) {
  const qualityLabel = { good: '利好', bad: '需关注', risk: '库存风险', neutral: '中性' }[metric.quality] || '中性';
  return `<div class="metric-card trend-${metric.quality}">
    <span>${html(metric.label)}</span>
    <strong>${formatByType(metric.current, metric.type)}</strong>
    <small>${hasComparison ? `上期 ${formatByType(metric.previous, metric.type)} · ${metric.trend} ${formatDelta(metric.delta, metric.type)} · ${formatPercent(metric.percent)}` : '上期：无对比数据'}</small>
    <em>${qualityLabel}</em>
  </div>`;
}

function renderPeriodSummary(comparison) {
  return `<section class="period-summary"><h3>区间对比总结</h3><p>${html(buildPeriodSummaryText(comparison))}</p></section>`;
}

function renderSkuComparisonTable(comparison) {
  if (state.filters.sku) return '';
  if (!comparison.skuRows.length) return '<section class="period-summary"><h3>SKU 对比明细</h3><p class="empty-state">当前时间段暂无 SKU 对比数据。</p></section>';
  const rows = comparison.skuRows.map((row) => `<tr>
    <td>${html(row.sku)}</td><td>${formatNumber(row.currentOrders)}</td><td>${formatNumber(row.previousOrders)}</td><td>${formatDelta(row.orderDelta, 'number')}</td>
    <td>${formatMoney(row.currentAdSpend)}</td><td>${formatMoney(row.previousAdSpend)}</td><td>${formatDelta(row.adSpendDelta, 'money')}</td>
    <td>${formatMoney(row.currentProfit)}</td><td>${formatMoney(row.previousProfit)}</td><td>${formatDelta(row.profitDelta, 'money')}</td>
    <td>${formatNumber(row.currentRoi)}</td><td>${formatNumber(row.previousRoi)}</td><td>${formatPercent(row.currentAcos)}</td><td>${formatPercent(row.previousAcos)}</td><td>${html(row.judgment)}</td>
  </tr>`).join('');
  return `<section class="period-summary"><h3>SKU 对比明细</h3><div class="table-wrap"><table><thead><tr><th>SKU</th><th>当前订单</th><th>上期订单</th><th>订单变化</th><th>当前广告费</th><th>上期广告费</th><th>广告费变化</th><th>当前利润</th><th>上期利润</th><th>利润变化</th><th>当前 ROI</th><th>上期 ROI</th><th>当前 ACOS</th><th>上期 ACOS</th><th>系统判断</th></tr></thead><tbody>${rows}</tbody></table></div></section>`;
}

function renderTable(records) {
  if (!records.length) return '<p class="empty-state">当前筛选条件下没有历史数据。</p>';
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


function renderStrategyBoard(analyses) {
  if (!analyses.length) return '<section class="panel"><h2>明日策略建议看板</h2><p class="empty-state">导入历史数据后将自动生成明日策略建议。</p></section>';
  return `<section class="panel strategy-board"><div class="panel-heading"><span class="panel-icon">★</span><div><h2>明日策略建议看板</h2><p>结合历史数据和每日动作，判断动作是否有效，并输出明日广告与运营建议。</p></div></div><div class="recommendation-grid">${analyses.map((item) => `<article class="recommendation-card priority-${item.primaryRecommendation.priority}">
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

function render() {
  const dates = getDateOptions(state.records);
  ensureDefaultDateRange(dates);
  const filtered = filterRecords(state.records, state.filters);
  const periodComparison = buildPeriodComparison(state.records, state.filters);
  const skus = getSkuOptions(state.records);
  const actions = actionMap();
  const selectedRecord = state.records.find((record) => record.uniqueKey === state.selectedDetailKey) || filtered[0];
  const selectedAction = selectedRecord ? actions.get(selectedRecord.uniqueKey) : null;
  const effectAnalyses = buildEffectAnalysis(state.records, state.actions, state.filters, periodComparison);

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

    <section class="panel summary-panel">
      <div class="panel-heading"><span class="panel-icon">▤</span><div><h2>筛选与汇总</h2><p>按任意时间段和 SKU 查看汇总，并自动对比上一个同长度时间段。</p></div></div>
      <div class="filter-row period-filter-row">
        <label class="form-field"><span>开始日期</span><input id="start-date-filter" type="date" value="${html(state.filters.startDate)}" /></label>
        <label class="form-field"><span>结束日期</span><input id="end-date-filter" type="date" value="${html(state.filters.endDate)}" /></label>
        <div class="quick-range-buttons"><span>快捷时间段</span>${QUICK_RANGES.map((option) => `<button type="button" class="quick-range-button ${option === state.filters.quickRange ? 'active' : ''}" data-quick-range="${html(option)}">${html(option)}</button>`).join('')}</div>
        <label class="form-field"><span>SKU 筛选</span><select id="sku-filter">${renderOptions(skus, state.filters.sku, '全部 SKU')}</select></label>
      </div>
      <div class="period-range-row">
        <div><span>当前时间段</span><strong>${html(periodComparison.currentRange.startDate || '-')} 至 ${html(periodComparison.currentRange.endDate || '-')}</strong></div>
        <div><span>对比时间段</span><strong>${periodComparison.currentRange.isAllDates ? '全部日期暂无上期对比' : (periodComparison.hasComparison ? `${html(periodComparison.previousRange.startDate)} 至 ${html(periodComparison.previousRange.endDate)}` : '无对比数据')}</strong></div>
      </div>
      <div class="metrics-grid extended-metrics">${periodComparison.metrics.map((metric) => renderMetricCard(metric, periodComparison.hasComparison)).join('')}</div>
      ${renderPeriodSummary(periodComparison)}
      ${renderSkuComparisonTable(periodComparison)}
    </section>

    ${renderActionModule(dates, skus)}
    ${renderStrategyBoard(effectAnalyses)}
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
  document.getElementById('start-date-filter')?.addEventListener('change', (event) => { state.filters.startDate = event.target.value; state.filters.quickRange = '自定义时间段'; render(); });
  document.getElementById('end-date-filter')?.addEventListener('change', (event) => { state.filters.endDate = event.target.value; state.filters.quickRange = '自定义时间段'; render(); });
  document.querySelectorAll('[data-quick-range]').forEach((button) => button.addEventListener('click', (event) => applyQuickRange(event.currentTarget.dataset.quickRange)));
  document.getElementById('sku-filter')?.addEventListener('change', (event) => { state.filters.sku = event.target.value; render(); });
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

refreshRecords().then(render).catch((error) => {
  state.status = `IndexedDB 初始化失败：${error.message || error}`;
  render();
});
