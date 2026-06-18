import { parseExcelWorkbook } from './utils/excel.js';
import { deleteAction, exportBackup, getAllActions, getAllRecords, importBackup, saveAction, saveRecords } from './utils/storage.js';
import { filterRecords, getDateOptions, getSkuOptions, summarizeByDate, summarizeRecords } from './utils/history.js';
import { fieldLabels } from './utils/fields.js';
import { formatMoney, formatPercent } from './utils/analysis.js';
import { ACTION_FIELDS, actionToSummary, createEmptyAction, normalizeAction } from './utils/actions.js';
import { buildEffectAnalysis, metricLabels } from './utils/effectAnalysis.js';

const root = document.getElementById('root');
const state = {
  records: [],
  actions: [],
  actionDraft: createEmptyAction(),
  filters: { date: '', sku: '' },
  lastImport: null,
  selectedDetailKey: '',
  status: '请上传包含 SKU sheet 的 Excel 文件，系统会把“日期 + SKU”作为唯一键保存到 IndexedDB。',
  busy: false,
};

const html = (value) => String(value ?? '').replace(/[&<>"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[char]));
const formatNumber = (value) => new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 2 }).format(Number(value) || 0);

const displayFields = [
  'date', 'sku', 'linkId', 'operationAction', 'price', 'reviews', 'rating', 'stock', 'reviewOrders', 'actualOrders',
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
    await refreshRecords();
    state.lastImport = {
      fileName: file.name,
      savedCount: result.records.length,
      skuSheets: result.skuSheets,
      skippedSheets: result.skippedSheets,
    };
    state.status = `导入完成：保存/覆盖 ${result.records.length} 行，识别 SKU sheet ${result.skuSheets.length} 个。`;
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

function renderImportSummary() {
  if (!state.lastImport) return '<p class="empty-state">尚未导入 Excel。本阶段会自动跳过 wb利润定价表、ozon利润定价表、Sheet10 等辅助 sheet。</p>';
  const { fileName, savedCount, skuSheets, skippedSheets } = state.lastImport;
  return `<div class="import-result">
    <strong>${html(fileName)}</strong>
    <span>保存/覆盖 ${savedCount} 行</span>
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

function renderActionField(field) {
  const value = state.actionDraft[field.key] ?? '';
  if (field.type === 'textarea') {
    return `<label class="form-field wide-field"><span>${field.label}</span><textarea data-action-field="${field.key}" rows="3">${html(value)}</textarea></label>`;
  }
  if (field.type === 'number') {
    return `<label class="form-field"><span>${field.label}</span><input data-action-field="${field.key}" type="number" step="0.01" value="${html(value)}" /></label>`;
  }
  return `<label class="form-field"><span>${field.label}</span><select data-action-field="${field.key}"><option value="">请选择</option>${field.options.map((option) => `<option value="${html(option)}" ${option === value ? 'selected' : ''}>${html(option)}</option>`).join('')}</select></label>`;
}

function renderActionModule(dates, skus) {
  const currentKey = getCurrentActionKey();
  const existing = actionMap().get(currentKey);
  return `<section class="panel action-panel">
    <div class="panel-heading"><span class="panel-icon">✎</span><div><h2>每日动作记录</h2><p>按“日期 + SKU”记录当天运营动作和广告策略调整；Excel B 列原文会保留，结构化动作以此表单为准。</p></div></div>
    <div class="action-form-grid">
      <label class="form-field"><span>日期</span><select id="action-date"><option value="">选择日期</option>${dates.map((date) => `<option value="${html(date)}" ${date === state.actionDraft.date ? 'selected' : ''}>${html(date)}</option>`).join('')}</select></label>
      <label class="form-field"><span>SKU</span><select id="action-sku"><option value="">选择 SKU</option>${skus.map((sku) => `<option value="${html(sku)}" ${sku === state.actionDraft.sku ? 'selected' : ''}>${html(sku)}</option>`).join('')}</select></label>
      ${ACTION_FIELDS.map(renderActionField).join('')}
    </div>
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
    return `<article class="effect-card"><div class="recommendation-head"><strong>${html(item.sku)}</strong><span>${html(actionToSummary(item.latestAction))}</span></div>
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
  const filtered = filterRecords(state.records, state.filters);
  const totals = summarizeRecords(filtered);
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

    <section class="panel">
      <div class="panel-heading"><span class="panel-icon">▤</span><div><h2>筛选与汇总</h2><p>按日期和 SKU 查看已导入历史数据，所有数据保存在当前浏览器 IndexedDB 中。</p></div></div>
      <div class="filter-row">
        <select id="date-filter">${renderOptions(dates, state.filters.date, '全部日期')}</select>
        <select id="sku-filter">${renderOptions(skus, state.filters.sku, '全部 SKU')}</select>
      </div>
      <div class="metrics-grid">
        <div class="metric-card"><span>已选日期数</span><strong>${totals.dateCount}</strong><small>历史总记录 ${state.records.length} 行</small></div>
        <div class="metric-card"><span>SKU 数量</span><strong>${totals.skuCount}</strong><small>当前筛选范围</small></div>
        <div class="metric-card"><span>总订单</span><strong>${formatNumber(totals.totalOrders)}</strong><small>销售额 ${formatMoney(totals.totalRevenue)}</small></div>
        <div class="metric-card"><span>总广告费</span><strong>${formatMoney(totals.totalAdSpend)}</strong><small>广告投入</small></div>
        <div class="metric-card"><span>总利润</span><strong>${formatMoney(totals.totalProfit)}</strong><small>当前筛选范围</small></div>
      </div>
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
  document.getElementById('date-filter')?.addEventListener('change', (event) => { state.filters.date = event.target.value; render(); });
  document.getElementById('sku-filter')?.addEventListener('change', (event) => { state.filters.sku = event.target.value; render(); });
  document.getElementById('action-date')?.addEventListener('change', (event) => { state.actionDraft.date = event.target.value; setActionDraftFromKey(getCurrentActionKey()); render(); });
  document.getElementById('action-sku')?.addEventListener('change', (event) => { state.actionDraft.sku = event.target.value; setActionDraftFromKey(getCurrentActionKey()); render(); });
  document.querySelectorAll('[data-action-field]').forEach((input) => {
    input.addEventListener('input', (event) => { state.actionDraft[event.target.dataset.actionField] = event.target.value; });
    input.addEventListener('change', (event) => { state.actionDraft[event.target.dataset.actionField] = event.target.value; });
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
