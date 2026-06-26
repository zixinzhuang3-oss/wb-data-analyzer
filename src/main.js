import { parseExcelWorkbook } from './utils/excel.js';
import { deleteAction, exportBackup, getAllActions, getAllRecords, importBackup, saveAction, saveExcelActions, saveRecords } from './utils/storage.js';
import { buildComparison, filterRecords, getDateOptions, getSkuOptions, summarizeByDate } from './utils/history.js';
import { fieldLabels } from './utils/fields.js';
import { formatPercent, formatRuble, formatYuan } from './utils/analysis.js';
import { ACTION_FIELDS, ACTION_SOURCE_LABELS, CPM_RECOMMEND_MIN_BID, CPM_SEARCH_MIN_BID, actionToSummary, applyAdRules, buildActionKey, createEmptyAction, getActionRecord, getEffectiveAction, getCpmMinBidForAction, isActionContentEqual, normalizeAction, validateCpmMinBids } from './utils/actions.js';
import { buildEffectAnalysis, buildSkuActionHistory, metricLabels } from './utils/effectAnalysis.js';
import { buildQuickRange, formatLocalDateKey, getTodayDate } from './utils/date.js';
import { toProfitCny, toProfitRub } from './utils/currency.js';

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
  'deliveryFee', 'acquiringFee', 'storageFee', 'remittanceFee', 'profitRub', 'profitCny', 'keywordRank',
];

const rubleFields = new Set(['price', 'adSpend', 'adCostPerOrder', 'adAvgClickCost', 'revenue', 'commission', 'russiaCost', 'deliveryFee', 'acquiringFee', 'storageFee', 'remittanceFee', 'profitRub']);
const yuanFields = new Set(['profitCny']);
const percentFields = new Set(['ctr', 'conversionRate', 'adShare', 'adCtr', 'adClickAddToCartRate', 'orderConversionRate']);

const renderValue = (key, value) => {
  if (rubleFields.has(key)) return formatRuble(value);
  if (yuanFields.has(key)) return formatYuan(value);
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
  try {
    const [records, actions] = await Promise.all([getAllRecords(), getAllActions()]);
    state.records = Array.isArray(records) ? records : [];
    state.actions = Array.isArray(actions) ? actions : [];
  } catch (error) {
    console.warn('读取 IndexedDB 失败，已使用空数据继续初始化。', error);
    state.records = [];
    state.actions = [];
    state.status = 'IndexedDB 暂无可用数据或存在旧版本脏数据，页面已使用空状态打开。';
  }
}

async function handleExcelUpload(file) {
  if (!file) return;
  state.busy = true;
  state.status = `正在解析 ${file.name} ...`;
  render();
  try {
    const result = await parseExcelWorkbook(file);
    if (!result.records.length) throw new Error('没有识别到有效 SKU sheet 或有效每日数据行。');
    const recordStats = await saveRecords(result.records);
    const actionStats = await saveExcelActions(result.actions || []);
    await refreshRecords();
    applyDefaultQuickRange();
    state.lastImport = {
      fileName: file.name,
      savedCount: result.records.length,
      recordsAdded: recordStats.added,
      recordsOverwritten: recordStats.overwritten,
      actionCount: actionStats.autoActionAdded,
      keptManualActions: actionStats.keptManualActions,
      skuSheets: result.skuSheets,
      skippedSheets: result.skippedSheets,
      diagnostics: result.diagnostics || [],
    };
    state.status = `导入完成：新增经营数据 ${recordStats.added} 条，覆盖经营数据 ${recordStats.overwritten} 条，自动识别动作记录 ${actionStats.autoActionAdded} 条，被保留的手动动作记录 ${actionStats.keptManualActions} 条。`;
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
    state.status = `备份导入完成：新增经营数据 ${result.recordsAdded} 条，覆盖经营数据 ${result.recordsOverwritten} 条，新增动作记录 ${result.actionsAdded} 条，覆盖动作记录 ${result.actionsOverwritten} 条，保留本地较新动作记录 ${result.actionsKeptLocal} 条，当前动作记录总数 ${result.currentActionTotal} 条。`;
  } catch (error) {
    state.status = `备份导入失败：${error.message || error}`;
  } finally {
    state.busy = false;
    render();
  }
}

async function handleBackupExport() {
  const backup = await exportBackup();
  downloadJson(`wb-data-backup-${formatLocalDateKey(new Date())}.json`, backup);
  state.status = `已导出：经营数据 ${backup.records.length} 条，动作记录 ${backup.actionRecords.length} 条，建议历史 ${backup.recommendationHistory.length} 条。`;
  render();
}

async function handleEffectRefresh() {
  await refreshRecords();
  state.status = '已从 IndexedDB 刷新动作分析数据。';
  render();
}

const renderOptions = (options, current, placeholder) => [`<option value="">${placeholder}</option>`, ...options.map((option) => `<option value="${html(option)}" ${option === current ? 'selected' : ''}>${html(option)}</option>`)].join('');

function renderImportSummary() {
  if (!state.lastImport) return '<p class="empty-state">尚未导入 Excel。本阶段会自动跳过 wb利润定价表、ozon利润定价表、Sheet10 等辅助 sheet。</p>';
  const { fileName, savedCount, recordsAdded = 0, recordsOverwritten = 0, actionCount = 0, keptManualActions = 0, skuSheets, skippedSheets } = state.lastImport;
  return `<div class="import-result">
    <strong>${html(fileName)}</strong>
    <span>保存/覆盖 ${savedCount} 行</span>
    <span>新增经营数据：${recordsAdded} 条</span>
    <span>覆盖经营数据：${recordsOverwritten} 条</span>
    <span>Excel 自动识别动作：${actionCount} 条</span>
    <span>保留手动动作：${keptManualActions} 条</span>
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
    <span>总销售额：${formatRuble(row.totalRevenue)}</span>
    <span>总广告费：${formatRuble(row.totalAdSpend)}</span>
    <span>总利润：${formatRuble(row.totalProfitRub)}</span>
    <small>原始 ${formatYuan(row.totalProfitCny)}</small>
  </article>`).join('')}</div>`;
}

function renderTable(records) {
  if (!records.length) return '<p class="empty-state">当前时间段无数据</p>';
  const header = `<th>详情/动作</th>${displayFields.map((key) => `<th>${fieldLabels[key]}</th>`).join('')}<th>结构化动作</th>`;
  const rows = records.map((record) => {
    const action = getActionRecord(state.actions, record.date, record.sku);
    return `<tr>
      <td><button class="table-action" data-detail-key="${html(record.uniqueKey)}" type="button">查看详情</button></td>
      ${displayFields.map((key) => `<td>${renderValue(key, record[key])}</td>`).join('')}
      <td>${html(actionToSummary(action))}</td>
    </tr>`;
  }).join('');
  return `<div class="table-wrap history-table"><table><thead><tr>${header}</tr></thead><tbody>${rows}</tbody></table></div>`;
}


const actionMap = () => new Map(state.actions.map((action) => [action.uniqueKey, action]));
const findAction = (date, sku) => getActionRecord(state.actions, date, sku);
const findEffectiveAction = (date, sku) => getEffectiveAction(state.actions, date, sku);

const getCurrentActionKey = () => buildActionKey(state.actionDraft.date || '', state.actionDraft.sku || '');

const setActionDraftFromKey = (uniqueKey) => {
  const record = state.records.find((item) => item.uniqueKey === uniqueKey);
  const separator = uniqueKey.indexOf('__') >= 0 ? '__' : '_';
  const [date = '', sku = ''] = uniqueKey.split(separator);
  const lookupDate = record?.date || date;
  const lookupSku = record?.sku || sku;
  const effective = findEffectiveAction(lookupDate, lookupSku);
  state.actionDraft = effective.action ? { ...effective.action, date: lookupDate, sku: lookupSku } : createEmptyAction(lookupDate, lookupSku);
};

async function handleActionSave() {
  const currentKey = getCurrentActionKey();
  const existing = getActionRecord(state.actions, state.actionDraft.date, state.actionDraft.sku);
  const effective = getEffectiveAction(state.actions, state.actionDraft.date, state.actionDraft.sku);
  const inheritedUnchanged = !existing && effective.isInherited && isActionContentEqual(state.actionDraft, effective.action);
  const action = normalizeAction({ ...state.actionDraft, source: existing ? 'manual_modified' : inheritedUnchanged ? 'inherited_saved' : 'manual_modified', originalActionDate: inheritedUnchanged ? effective.sourceActionDate : state.actionDraft.originalActionDate, sourceActionDate: inheritedUnchanged ? effective.sourceActionDate : state.actionDraft.sourceActionDate, effectiveFromDate: state.actionDraft.date, isInherited: false });
  if (!action.date || !action.sku) {
    state.status = '动作保存失败：请先选择日期和 SKU。';
    render();
    return;
  }
  const bidErrors = validateCpmMinBids(action);
  if (bidErrors.length) {
    state.status = `动作保存失败：${bidErrors.join('；')}。CPM 搜索最低出价为 ${CPM_SEARCH_MIN_BID}，推荐最低出价为 ${CPM_RECOMMEND_MIN_BID}。`;
    render();
    return;
  }
  await saveAction(action);
  await refreshRecords();
  state.selectedDetailKey = action.uniqueKey;
  state.actionDraft = { ...action };
  state.status = inheritedUnchanged ? `已将继承动作保存为当前日期动作：${action.date} ${action.sku}。` : `已保存新动作，从当前日期开始生效：${action.date} ${action.sku}。`;
  render();
}

async function handleActionDelete(uniqueKey = getCurrentActionKey()) {
  if (!uniqueKey || uniqueKey === '_' || uniqueKey === '__') {
    state.status = '动作删除失败：请先选择日期和 SKU。';
    render();
    return;
  }
  await deleteAction(uniqueKey);
  await refreshRecords();
  const separator = uniqueKey.indexOf('__') >= 0 ? '__' : '_';
  const [date, sku] = uniqueKey.split(separator);
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
    const cpmBidField = ['cpmSearchBid', 'cpmRecommendBid', 'cpmUnifiedBid'].includes(field.key);
    const min = cpmBidField ? getCpmMinBidForAction(draft, field.key) : '';
    const hint = cpmBidField ? `<small class="field-hint">CPM 搜索最低出价为 ${CPM_SEARCH_MIN_BID}，推荐最低出价为 ${CPM_RECOMMEND_MIN_BID}。</small>` : '';
    return `<label class="form-field"><span>${field.label}</span><input data-action-field="${field.key}" type="number" step="0.01" ${min ? `min="${min}"` : ''} value="${html(value)}" />${hint}</label>`;
  }
  const disabled = field.disabled ? 'disabled' : '';
  return `<label class="form-field"><span>${field.label}</span><select data-action-field="${field.key}" ${disabled}><option value="">请选择</option>${field.options.map((option) => `<option value="${html(option)}" ${option === value ? 'selected' : ''}>${html(ACTION_SOURCE_LABELS[option] || option)}</option>`).join('')}</select></label>`;
}

function renderActionDiagnostics() {
  const date = state.actionDraft.date || state.selectedDetailKey.split(state.selectedDetailKey.includes('__') ? '__' : '_')[0] || '';
  const sku = state.actionDraft.sku || state.selectedDetailKey.split(state.selectedDetailKey.includes('__') ? '__' : '_')[1] || '';
  const previousDate = date ? buildQuickRange('yesterday', date).startDate : '';
  const currentKey = buildActionKey(date, sku);
  const previousKey = buildActionKey(previousDate, sku);
  const currentAction = date && sku ? findAction(date, sku) : null;
  const previousAction = previousDate && sku ? findAction(previousDate, sku) : null;
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">?</span><div><h2>动作记录诊断</h2><p>按统一日期 key 检查当前日期和上一日的动作记录是否存在。</p></div></div>
    <div class="metrics-grid">
      <div class="metric-card"><span>当前选择</span><strong>${html(date || '-')} / ${html(sku || '-')}</strong><small>${currentAction ? '有动作记录' : '无动作记录'} · key ${html(currentKey || '-')}</small></div>
      <div class="metric-card"><span>上一日</span><strong>${html(previousDate || '-')} / ${html(sku || '-')}</strong><small>${previousAction ? '有动作记录' : '无动作记录'} · key ${html(previousKey || '-')}</small></div>
      <div class="metric-card"><span>当前记录来源</span><strong>${html(ACTION_SOURCE_LABELS[currentAction?.source] || currentAction?.source || '-')}</strong><small>更新时间：${html(currentAction?.updatedAt || '-')}</small></div>
      <div class="metric-card"><span>上一日记录来源</span><strong>${html(ACTION_SOURCE_LABELS[previousAction?.source] || previousAction?.source || '-')}</strong><small>更新时间：${html(previousAction?.updatedAt || '-')}</small></div>
      <div class="metric-card"><span>IndexedDB 动作记录总数</span><strong>${formatNumber(state.actions.length)}</strong></div>
      <div class="metric-card"><span>IndexedDB 经营数据总数</span><strong>${formatNumber(state.records.length)}</strong></div>
    </div></section>`;
}

function renderActionModule(dates, skus) {
  const currentKey = getCurrentActionKey();
  const existing = getActionRecord(state.actions, state.actionDraft.date, state.actionDraft.sku);
  const effective = getEffectiveAction(state.actions, state.actionDraft.date, state.actionDraft.sku);
  const sourceText = existing ? `当前日期已有保存动作。来源：${ACTION_SOURCE_LABELS[existing.source] || existing.source || '-'}。` : effective.isInherited ? `当前日期未保存新动作，正在沿用 ${effective.sourceActionDate} 的动作。如需从今天开始改变策略，请修改后点击保存。` : '该 SKU 暂无历史动作记录，请填写初始动作。';
  const groups = ['基础字段', 'CPC 模块', 'CPM 模块'].map((group) => `<fieldset class="action-fieldset"><legend>${group}</legend><div class="action-form-grid">${ACTION_FIELDS.filter((field) => field.group === group).map(renderActionField).join('')}</div></fieldset>`).join('');
  return `<section class="panel action-panel">
    <div class="panel-heading"><span class="panel-icon">✎</span><div><h2>每日动作记录</h2><p>按“日期 + SKU”记录基础动作，并把 CPC 模块与 CPM 模块独立维护；整体广告状态会根据 CPC/CPM 开关自动计算。CPM 搜索出价最低：${CPM_SEARCH_MIN_BID}；CPM 推荐出价最低：${CPM_RECOMMEND_MIN_BID}。</p></div></div>
    <div class="action-form-grid">
      <label class="form-field"><span>日期</span><select id="action-date"><option value="">选择日期</option>${dates.map((date) => `<option value="${html(date)}" ${date === state.actionDraft.date ? 'selected' : ''}>${html(date)}</option>`).join('')}</select></label>
      <label class="form-field"><span>SKU</span><select id="action-sku"><option value="">选择 SKU</option>${skus.map((sku) => `<option value="${html(sku)}" ${sku === state.actionDraft.sku ? 'selected' : ''}>${html(sku)}</option>`).join('')}</select></label>
    </div>
    <p class="status-line">${html(sourceText)}</p>
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
      <div><span>广告费 ₽</span><strong>${formatRuble(record.adSpend)}</strong></div>
      <div><span>销售额 ₽</span><strong>${formatRuble(record.revenue)}</strong></div>
      <div><span>利润 ₽</span><strong>${formatRuble(toProfitRub(record))}</strong><small>原始 ${formatYuan(toProfitCny(record))}</small></div>
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
  const allNoValidData = analyses.every((item) => item.noValidData);
  const context = allNoValidData ? '当前时间段暂无有效数据，请先导入数据。' : comparison?.hasPreviousData ? `当前区间 ${rangeText(comparison.currentRange)} 对比 ${rangeText(comparison.previousRange)}：订单变化 ${formatPercent(deltaInfo(comparison.current.totalOrders, comparison.previous.totalOrders).rate)}，广告费变化 ${formatPercent(deltaInfo(comparison.current.totalAdSpend, comparison.previous.totalAdSpend).rate)}，利润变化 ${formatRuble(deltaInfo(comparison.current.totalProfitRub, comparison.previous.totalProfitRub).value)}。明日策略会结合该区间趋势与每日动作。` : '当前筛选区间无对比数据，明日策略主要结合当前区间、最近日期和每日动作生成。';
  return `<section class="panel strategy-board"><div class="panel-heading"><span class="panel-icon">★</span><div><h2>明日策略建议看板</h2><p>结合当前选择时间段、7 天内最近动作、动作后 1/3/7 天窗口和每日动作历史，判断动作是否有效，并输出明日广告与运营建议。CPM 搜索出价最低：${CPM_SEARCH_MIN_BID}；CPM 推荐出价最低：${CPM_RECOMMEND_MIN_BID}。</p></div></div><p class="strategy-context">${html(context)}</p><div class="recommendation-grid">${analyses.map((item) => `<article class="recommendation-card priority-${item.primaryRecommendation.priority}">
    <div class="recommendation-head"><strong>${html(item.sku)}</strong><span>${html(item.date)}</span></div>
    <h3>${html(item.noValidData ? '暂无当前日期数据' : item.primaryRecommendation.type)}</h3>
    <div class="mini-metrics action-meta">
      <span>当前分析日期：${html(item.actionMeta?.analysisDate || item.date)}</span>
      <span>SKU：${html(item.sku)}</span>
      <span>生效动作日期：${html(item.actionMeta?.effectiveActionDate || item.actionMeta?.usedActionDate || '-')}</span>
      <span>当天新动作：${item.actionMeta?.isTodayNewAction ? '是' : '否'}</span>
      <span>继承动作：${item.actionMeta?.isInherited ? '是' : '否'}</span>
      <span>继承来源日期：${html(item.actionMeta?.isInherited ? item.actionMeta?.sourceActionDate : '-')}</span>
      <span>动作已连续执行：${item.actionMeta?.daysSinceAction ?? '-'} 天</span>
      <span>动作距今天数：${item.actionMeta?.daysSinceAction ?? '-'}</span>
      <span>动作后 1 天：${html(item.actionWindows?.after1?.summary || '-')}</span>
      <span>动作后 3 天：${html(item.actionWindows?.after3?.summary || '-')}</span>
      <span>动作后 7 天：${html(item.actionWindows?.after7?.summary || '-')}</span>
    </div>
    <p><strong>最近动作内容：</strong>${html(item.latestAction ? actionToSummary(item.latestAction) : item.actionMeta?.missingMessage || '最近 7 天未找到动作记录。')}</p>
    <p><strong>系统判断：</strong>${html(item.effects?.[0]?.text || (item.noValidData ? '暂无当前日期数据，无法判断。' : item.primaryRecommendation.reason))}</p>
    <p><strong>明日建议：</strong>${html(item.primaryRecommendation.reason)}</p>
  </article>`).join('')}</div></section>`;
}

function renderEffectCards(analyses) {
  if (!analyses.length) return '';
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">↗</span><div><h2>动作效果分析卡片</h2><p>不再只查上一日动作；默认向前 7 天查找当前 SKU 最近动作，并展示动作后 1/3/7 天结果。</p></div><button id="refresh-effect-analysis" type="button">刷新动作分析</button></div><div class="effect-grid">${analyses.map((item) => {
    const effectText = item.noValidData ? (item.actionMeta?.found ? `${item.actionMeta?.missingMessage || '已找到最近动作'} 当前日期暂无有效数据，暂时无法判断动作效果。` : '当前时间段暂无有效数据，无法生成策略建议。') : item.effects.length ? item.effects.map((effect) => `${effect.level}：${effect.text}`).join(' ') : '暂无明确动作效果，建议继续观察。';
    const missingActionText = item.actionMeta?.missingMessage || `最近 7 天未找到动作记录。`;
    const actionSummary = item.actionMeta?.found ? actionToSummary(item.latestAction) : missingActionText;
    const resultText = item.noValidData
      ? (item.actionMeta?.found ? `${item.actionMeta?.missingMessage || '已找到最近动作'} 当前日期暂无有效数据，暂时无法判断动作效果。` : '当前时间段暂无有效数据，无法生成策略建议。')
      : item.actionMeta?.found
        ? `最近动作 → 当前结果：${actionSummary}；今天订单 ${formatNumber(item.metrics.totalOrders.today)}，广告费 ${formatRuble(item.metrics.adSpend.today)}，利润 ${formatRuble(item.metrics.profit.today)}。`
        : `${missingActionText} 今天订单 ${formatNumber(item.metrics.totalOrders.today)}，广告费 ${formatRuble(item.metrics.adSpend.today)}，利润 ${formatRuble(item.metrics.profit.today)}。`;
    return `<article class="effect-card"><div class="recommendation-head"><strong>${html(item.sku)}</strong><span>${html(actionSummary)}</span></div>
      <div class="mini-metrics action-meta">
        <span>分析日期：${html(item.actionMeta?.analysisDate || item.date)}</span>
        <span>对比日期：${html(item.actionMeta?.comparisonDate || '-')}</span>
        <span>当前 SKU：${html(item.actionMeta?.sku || item.sku)}</span>
        <span>上一日动作日期：${html(item.actionMeta?.requiredActionDate || '-')}</span>
        <span>生效动作日期：${html(item.actionMeta?.effectiveActionDate || item.actionMeta?.usedActionDate || '-')}</span>
        <span>当天新动作：${item.actionMeta?.isTodayNewAction ? '是' : '否'}</span>
        <span>继承动作：${item.actionMeta?.isInherited ? '是' : '否'}</span>
        <span>继承来源日期：${html(item.actionMeta?.isInherited ? item.actionMeta?.sourceActionDate : '-')}</span>
        <span>动作距今天数：${item.actionMeta?.daysSinceAction ?? '-'}</span>
        <span>查找 key：${html(item.actionMeta?.lookupKey || '-')}</span>
        <span>是否找到动作：${item.actionMeta?.found ? '是' : '否'}</span>
        <span>找到的动作来源 source：${html(item.actionMeta?.found ? item.actionMeta.source : '未找到 IndexedDB 动作记录')}</span>
      </div>
      <p>${html(resultText)}</p>
      <p>${html(effectText)}</p>
      ${item.noValidData ? '<div class="mini-metrics"><span>暂无当前日期数据，无法判断。</span></div>' : `<div class="mini-metrics">
        <span>订单 ${formatNumber(item.metrics.totalOrders.today)} / 昨日 ${formatNumber(item.metrics.totalOrders.yesterday)}</span>
        <span>广告费 ${formatRuble(item.metrics.adSpend.today)} / 近3天 ${formatRuble(item.metrics.adSpend.last3Avg)}</span>
        <span>利润 ${formatRuble(item.metrics.profit.today)} / 动作前后 ${formatRuble(item.metrics.profit.actionWindowDelta.value)}</span>
        <span>ROI ${formatNumber(item.metrics.roi.today)} · ACOS ${formatPercent(item.metrics.acos.today)}</span>
        <span>动作后1天 ${html(item.actionWindows?.after1?.summary || '-')}</span>
        <span>动作后3天 ${html(item.actionWindows?.after3?.summary || '-')}</span>
        <span>动作后7天 ${html(item.actionWindows?.after7?.summary || '-')}</span>
      </div>`}
    </article>`;
  }).join('')}</div></section>`;
}

function renderRiskPanel(analyses) {
  const risks = analyses.flatMap((item) => item.risks.map((risk) => ({ sku: item.sku, risk })));
  if (!risks.length) return `<section class="panel"><div class="panel-heading compact"><span class="panel-icon">!</span><h2>SKU 风险提示</h2></div><p class="empty-state">当前筛选范围内暂无高风险提示。</p></section>`;
  return `<section class="panel risk-panel"><div class="panel-heading compact"><span class="panel-icon">!</span><h2>SKU 风险提示</h2></div><div class="risk-list">${risks.map((item) => `<div class="risk-item"><strong>${html(item.sku)}</strong><span>${html(item.risk)}</span></div>`).join('')}</div></section>`;
}


function renderSkuActionHistoryPanel(records, actions) {
  const selectedSku = state.filters.sku || state.actionDraft.sku || state.records.find((record) => record.uniqueKey === state.selectedDetailKey)?.sku || '';
  if (!selectedSku) return '<section class="panel"><div class="panel-heading"><span class="panel-icon">☰</span><div><h2>SKU 动作历史</h2><p>选择某个 SKU 后显示最近 30 天动作记录。</p></div></div><p class="empty-state">请选择 SKU 查看动作历史。</p></section>';
  const endDate = state.filters.endDate || state.today.date || records.map((record) => record.date).sort().at(-1) || '';
  const rows = buildSkuActionHistory(records, actions, selectedSku, endDate);
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">☰</span><div><h2>SKU 动作历史</h2><p>${html(selectedSku)} 最近 30 天动作记录，包含动作后 1/3/7 天结果和系统判断。</p></div></div>${rows.length ? `<div class="table-wrap"><table><thead><tr><th>日期</th><th>记录类型</th><th>动作摘要</th><th>CPC 状态</th><th>CPM 状态</th><th>价格动作</th><th>主图动作</th><th>关键词动作</th><th>库存动作</th><th>备注</th><th>动作后 1 天结果</th><th>动作后 3 天结果</th><th>动作后 7 天结果</th><th>系统判断</th></tr></thead><tbody>${rows.map((row) => `<tr><td>${html(row.date)}</td><td>${html(row.type === 'inherited' ? `继承 ${row.sourceActionDate}` : (ACTION_SOURCE_LABELS[row.action.source] || row.action.source || '当天保存'))}</td><td>${html(row.summary)}</td><td>${html(row.action.cpcEnabled || '-')}</td><td>${html(row.action.cpmEnabled || '-')}</td><td>${html(row.action.priceAction || '-')}</td><td>${html(row.action.imageAction || '-')}</td><td>${html(row.action.keywordAction || '-')}</td><td>${html(row.action.stockAction || '-')}</td><td>${html(row.action.note || '-')}</td><td>${html(row.windows.after1.summary)}</td><td>${html(row.windows.after3.summary)}</td><td>${html(row.windows.after7.summary)}</td><td>${html(row.judgement)}</td></tr>`).join('')}</tbody></table></div>` : '<p class="empty-state">最近 30 天未找到动作记录。</p>'}</section>`;
}

function renderSuggestionHistory(analyses, dates, skus) {
  const rows = analyses.flatMap((item) => item.recommendations.map((rec) => ({ date: item.date, sku: item.sku, ...rec })));
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">☷</span><div><h2>建议历史</h2><p>建议历史跟随上方日期和 SKU 筛选，可查看每个 SKU 的建议类型和原因。</p></div></div>${rows.length ? `<div class="table-wrap"><table><thead><tr><th>日期</th><th>SKU</th><th>建议类型</th><th>优先级</th><th>原因</th></tr></thead><tbody>${rows.map((row) => `<tr><td>${html(row.date)}</td><td>${html(row.sku)}</td><td>${html(row.type)}</td><td>${html(row.priority)}</td><td>${html(row.reason)}</td></tr>`).join('')}</tbody></table></div>` : '<p class="empty-state">当前筛选范围暂无建议。</p>'}</section>`;
}

const metricConfig = [
  ['总订单', 'totalOrders', formatNumber],
  ['总销售额 ₽', 'totalRevenue', formatRuble],
  ['总利润 ₽', 'totalProfitRub', formatRuble],
  ['原始利润 ¥', 'totalProfitCny', formatYuan],
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
  ['总广告费 ₽', 'totalAdSpend', formatRuble],
  ['广告费占比', 'adShare', formatPercent],
  ['每单费用 ₽', 'adCostPerOrder', formatRuble],
  ['广告平均点击费 ₽', 'adAvgClickCost', formatRuble],
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
  const profit = deltaInfo(current.totalProfitRub, previous.totalProfitRub);
  if ((current.stock || 0) > 0 && current.stock < Math.max(5, current.totalOrders * 2)) return '库存风险';
  if (spend.value > 0 && order.value <= 0) return '广告费上升但订单未增长';
  if (profit.value < 0) return '利润下降';
  if (spend.value < 0 && order.value >= 0 && profit.value > 0) return '控费有效';
  if (order.value > 0 && profit.value > 0 && current.roi > 1.5) return '建议加预算';
  if (current.totalAdSpend > 0 && current.totalOrders === 0 && current.totalProfitRub < 0) return '建议暂停广告';
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
  const totalProfitCny = rows.reduce((sum, row) => sum + toProfitCny(row), 0);
  const totalProfitRub = rows.reduce((sum, row) => sum + toProfitRub(row), 0);
  const totalProfit = totalProfitRub;
  const stock = rows.at(-1)?.stock || 0;
  return { totalOrders, totalRevenue, totalAdSpend, totalProfit, totalProfitRub, totalProfitCny, stock, roi: totalAdSpend ? totalRevenue / totalAdSpend : 0, acos: totalRevenue ? totalAdSpend / totalRevenue : 0 };
}

function renderSkuComparison(comparison) {
  const rows = buildSkuRows(comparison);
  if (!rows.length) return '<section class="panel"><h2>SKU 区间对比</h2><p class="empty-state">当前筛选范围暂无 SKU 数据。</p></section>';
  const change = (a, b, formatter = formatNumber) => formatter((Number(a) || 0) - (Number(b) || 0));
  return `<section class="panel"><div class="panel-heading"><span class="panel-icon">⇄</span><div><h2>SKU 区间对比</h2><p>${state.filters.sku ? '当前为单个 SKU 区间趋势。' : '全部 SKU 按当前区间与上一同长度区间逐项对比。'}</p></div></div><div class="table-wrap"><table><thead><tr><th>SKU</th><th>当前订单</th><th>上期订单</th><th>订单变化</th><th>当前销售额</th><th>上期销售额</th><th>销售额变化</th><th>当前广告费</th><th>上期广告费</th><th>广告费变化</th><th>当前利润</th><th>上期利润</th><th>利润变化</th><th>当前 ROI</th><th>上期 ROI</th><th>当前 ACOS</th><th>上期 ACOS</th><th>系统判断</th></tr></thead><tbody>${rows.map((row) => `<tr><td>${html(row.sku)}</td><td>${formatNumber(row.current.totalOrders)}</td><td>${comparison.hasPreviousData ? formatNumber(row.previous.totalOrders) : '无对比数据'}</td><td>${comparison.hasPreviousData ? change(row.current.totalOrders, row.previous.totalOrders) : '无对比数据'}</td><td>${formatRuble(row.current.totalRevenue)}</td><td>${comparison.hasPreviousData ? formatRuble(row.previous.totalRevenue) : '无对比数据'}</td><td>${comparison.hasPreviousData ? change(row.current.totalRevenue, row.previous.totalRevenue, formatRuble) : '无对比数据'}</td><td>${formatRuble(row.current.totalAdSpend)}</td><td>${comparison.hasPreviousData ? formatRuble(row.previous.totalAdSpend) : '无对比数据'}</td><td>${comparison.hasPreviousData ? change(row.current.totalAdSpend, row.previous.totalAdSpend, formatRuble) : '无对比数据'}</td><td>${formatRuble(row.current.totalProfitRub)}<br><small>原始 ${formatYuan(row.current.totalProfitCny)}</small></td><td>${comparison.hasPreviousData ? `${formatRuble(row.previous.totalProfitRub)}<br><small>原始 ${formatYuan(row.previous.totalProfitCny)}</small>` : '无对比数据'}</td><td>${comparison.hasPreviousData ? change(row.current.totalProfitRub, row.previous.totalProfitRub, formatRuble) : '无对比数据'}</td><td>${formatNumber(row.current.roi)}</td><td>${comparison.hasPreviousData ? formatNumber(row.previous.roi) : '无对比数据'}</td><td>${formatPercent(row.current.acos)}</td><td>${comparison.hasPreviousData ? formatPercent(row.previous.acos) : '无对比数据'}</td><td>${html(row.judge)}</td></tr>`).join('')}</tbody></table></div></section>`;
}

function renderIntervalSummary(comparison) {
  if (!comparison.hasPreviousData) return `<section class="panel"><h2>区间对比总结</h2><p class="empty-state">当前时间段 ${html(rangeText(comparison.currentRange))}，无对比数据。</p></section>`;
  const order = deltaInfo(comparison.current.totalOrders, comparison.previous.totalOrders);
  const spend = deltaInfo(comparison.current.totalAdSpend, comparison.previous.totalAdSpend);
  const profit = deltaInfo(comparison.current.totalProfitRub, comparison.previous.totalProfitRub);
  let advice = '建议继续观察核心 SKU 的订单、广告费和利润变化。';
  if (spend.value < 0 && Math.abs(order.rate) <= 0.1 && profit.value > 0) advice = '说明控费效果较好，建议继续保持当前预算策略。';
  else if (spend.rate > 0.2 && order.rate < 0.1 && profit.value < 0) advice = '说明广告放量效率较差，建议重点检查高花费 SKU，并降低 ROI 差的广告预算。';
  else if (order.value > 0 && profit.value > 0 && comparison.current.roi > 1.5) advice = '说明增长质量较好，库存充足 SKU 可适当加预算。';
  return `<section class="panel"><h2>区间对比总结</h2><p>当前时间段 ${html(rangeText(comparison.currentRange))}，较上一时间段订单${order.value >= 0 ? '增加' : '减少'} ${formatPercent(Math.abs(order.rate))}，广告费${spend.value >= 0 ? '增加' : '下降'} ${formatPercent(Math.abs(spend.rate))}，利润${profit.value >= 0 ? '增加' : '下降'} ${formatRuble(Math.abs(profit.value))}，${advice}</p></section>`;
}

const renderModule = (name, renderer) => {
  try {
    return renderer();
  } catch (error) {
    console.error(`${name} 加载失败`, error);
    return `<section class="panel"><h2>${html(name)}</h2><p class="empty-state error-text">该模块加载失败，请检查数据或刷新页面。</p></section>`;
  }
};

function render() {
  try {
  if (!state.defaultRangeApplied && state.today.date && !state.filters.startDate && !state.filters.endDate && !state.filters.allDates) applyDefaultQuickRange();
  const safeRecords = Array.isArray(state.records) ? state.records : [];
  const safeActions = Array.isArray(state.actions) ? state.actions : [];
  const filtered = filterRecords(safeRecords, state.filters);
  const comparison = buildComparison(safeRecords, state.filters);
  const dates = getDateOptions(safeRecords);
  const skus = getSkuOptions(safeRecords);
  const selectedRecord = safeRecords.find((record) => record.uniqueKey === state.selectedDetailKey) || filtered[0];
  const selectedAction = selectedRecord ? findAction(selectedRecord.date, selectedRecord.sku) : null;
  const effectAnalyses = buildEffectAnalysis(safeRecords, safeActions, state.filters);

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
      ${renderModule('导入摘要', renderImportSummary)}
    </section>
    ${renderModule('字段识别诊断', renderFieldDiagnostics)}

    <section class="panel">
      <div class="panel-heading"><span class="panel-icon">▤</span><div><h2>筛选与汇总</h2><p>快捷时间段基于真实当前日期；销售额、广告费和页面主利润均为卢布 ₽；原始利润保留人民币 ¥，按固定汇率换算。</p></div></div>
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
      ${renderModule('筛选汇总指标', () => renderComparisonMetrics(comparison))}
    </section>

    ${renderModule('区间汇总', () => renderIntervalSummary(comparison))}
    ${renderModule('SKU 对比', () => renderSkuComparison(comparison))}
    ${renderModule('每日动作记录', () => renderActionModule(dates, skus))}
    ${renderModule('动作记录诊断', renderActionDiagnostics)}
    ${renderModule('明日策略建议看板', () => renderStrategyBoard(effectAnalyses, comparison))}
    ${renderModule('动作效果分析卡片', () => renderEffectCards(effectAnalyses))}
    ${renderModule('SKU 风险提示', () => renderRiskPanel(effectAnalyses))}
    ${renderModule('建议历史', () => renderSuggestionHistory(effectAnalyses, dates, skus))}
    ${renderModule('SKU 动作历史', () => renderSkuActionHistoryPanel(safeRecords, safeActions))}

    <section class="content-grid">
      <div class="panel wide"><h2>历史数据明细</h2>${renderModule('历史数据明细', () => renderTable(filtered))}</div>
      <aside class="panel strategy-panel"><div class="panel-heading compact"><span class="panel-icon">◎</span><h2>导入数据记录</h2></div>${renderModule('导入数据记录', () => renderHistoryCards(safeRecords))}<div class="sku-detail"><h2>SKU 详情</h2>${renderModule('SKU 详情', () => renderSkuDetail(selectedRecord, selectedAction))}</div></aside>
    </section>
  </main>`;

  document.getElementById('excel-upload')?.addEventListener('change', (event) => handleExcelUpload(event.target.files?.[0]));
  document.getElementById('backup-upload')?.addEventListener('change', (event) => handleBackupImport(event.target.files?.[0]));
  document.getElementById('export-backup')?.addEventListener('click', handleBackupExport);
  document.getElementById('refresh-effect-analysis')?.addEventListener('click', handleEffectRefresh);
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
    input.addEventListener('change', (event) => { state.actionDraft[event.target.dataset.actionField] = event.target.value; state.actionDraft.source = ['excel_auto', 'json_import'].includes(state.actionDraft.source) ? 'manual_modified' : state.actionDraft.source; state.actionDraft = applyAdRules(state.actionDraft); render(); });
  });
  document.getElementById('save-action')?.addEventListener('click', handleActionSave);
  document.getElementById('delete-action')?.addEventListener('click', () => handleActionDelete());
  document.getElementById('clear-action')?.addEventListener('click', () => { state.actionDraft = createEmptyAction(state.actionDraft.date, state.actionDraft.sku); render(); });
  document.querySelectorAll('[data-detail-key]').forEach((button) => button.addEventListener('click', (event) => {
    state.selectedDetailKey = event.target.dataset.detailKey;
    setActionDraftFromKey(state.selectedDetailKey);
    render();
  }));
  } catch (error) {
    console.error('主渲染入口加载失败', error);
    root.innerHTML = '<main class="app-shell"><section class="panel"><h2>页面加载失败</h2><p class="empty-state error-text">该模块加载失败，请检查数据或刷新页面。</p></section></main>';
  }
}

Promise.all([refreshRecords(), getTodayDate()]).then(([, today]) => {
  state.today = today?.date ? today : { date: formatLocalDateKey(new Date()), source: '浏览器本地时间', timeZone: 'Asia/Shanghai' };
  applyDefaultQuickRange();
  render();
}).catch((error) => {
  state.status = `初始化失败：${error.message || error}`;
  render();
});
