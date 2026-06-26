import { normalizeDateKey } from './date.js';

export const normalizeDate = (date) => normalizeDateKey(date);
export const normalizeSku = (sku) => String(sku || '').trim().toUpperCase();
export const buildActionKey = (date, sku) => `${normalizeDate(date)}_${normalizeSku(sku)}`;
export const getActionRecord = (actions = [], date = '', sku = '') => {
  const key = buildActionKey(date, sku);
  return actions.find((action) => action?.uniqueKey === key || buildActionKey(action?.date, action?.sku) === key) || null;
};

export const ACTION_INHERIT_LOOKBACK_DAYS = 30;
export const ACTION_LOOKBACK_DAYS = ACTION_INHERIT_LOOKBACK_DAYS;
export const ACTION_HISTORY_DAYS = 30;

export const getSkuActionTimeline = (actions = [], sku = '', { beforeDate = '', fromDate = '', toDate = '' } = {}) => {
  const targetSku = normalizeSku(sku);
  return actions
    .map((action) => normalizeAction(action))
    .filter((action) => action.sku && (!targetSku || action.sku === targetSku))
    .filter((action) => !beforeDate || action.date < normalizeDate(beforeDate))
    .filter((action) => !fromDate || action.date >= normalizeDate(fromDate))
    .filter((action) => !toDate || action.date <= normalizeDate(toDate))
    .sort((a, b) => a.date.localeCompare(b.date));
};

export const findRecentAction = (actions = [], date = '', sku = '', lookbackDays = ACTION_LOOKBACK_DAYS) => {
  const analysisDate = normalizeDate(date);
  const targetSku = normalizeSku(sku);
  const timeline = getSkuActionTimeline(actions, targetSku, { beforeDate: analysisDate });
  const earliest = normalizeDateKey(new Date(new Date(`${analysisDate}T00:00:00Z`).getTime() - lookbackDays * 86400000));
  const recent = timeline.filter((action) => action.date >= earliest).at(-1) || null;
  return {
    action: recent,
    timeline,
    found: Boolean(recent),
    lookbackDays,
    analysisDate,
    daysSinceAction: recent ? Math.round((new Date(`${analysisDate}T00:00:00Z`) - new Date(`${recent.date}T00:00:00Z`)) / 86400000) : null,
    previousDayHadAction: Boolean(getActionRecord(actions, new Date(new Date(`${analysisDate}T00:00:00Z`).getTime() - 86400000).toISOString().slice(0, 10), targetSku)),
  };
};


export const getEffectiveAction = (actions = [], date = '', sku = '', lookbackDays = ACTION_INHERIT_LOOKBACK_DAYS) => {
  const analysisDate = normalizeDate(date);
  const targetSku = normalizeSku(sku);
  const explicit = getActionRecord(actions, analysisDate, targetSku);
  if (explicit) {
    const action = normalizeAction(explicit);
    return {
      action: { ...action, effectiveFromDate: action.date, sourceActionDate: action.date, isInherited: false },
      explicitAction: action,
      found: true,
      isInherited: false,
      sourceActionDate: action.date,
      effectiveFromDate: action.date,
      daysSinceAction: 0,
      lookbackDays,
      message: '当前日期已有保存动作。',
    };
  }
  const recent = findRecentAction(actions, analysisDate, targetSku, lookbackDays);
  if (!recent.action) {
    return {
      action: null,
      explicitAction: null,
      found: false,
      isInherited: false,
      sourceActionDate: '',
      effectiveFromDate: '',
      daysSinceAction: null,
      lookbackDays,
      message: `最近 ${lookbackDays} 天未找到动作记录。`,
    };
  }
  const inherited = normalizeAction(recent.action);
  return {
    action: { ...inherited, date: analysisDate, uniqueKey: buildActionKey(analysisDate, targetSku), effectiveFromDate: inherited.date, sourceActionDate: inherited.date, originalActionDate: inherited.date, isInherited: true },
    explicitAction: null,
    found: true,
    isInherited: true,
    sourceActionDate: inherited.date,
    effectiveFromDate: inherited.date,
    daysSinceAction: recent.daysSinceAction,
    lookbackDays,
    message: `当前日期未保存新动作，正在沿用最近一次动作：${inherited.date}。`,
  };
};

export const getActionComparable = (action = {}) => {
  const ignored = new Set(['uniqueKey', 'updatedAt', 'source', 'effectiveFromDate', 'sourceActionDate', 'originalActionDate', 'isInherited']);
  const normalized = normalizeAction(action);
  return Object.fromEntries(Object.entries(normalized).filter(([key]) => !ignored.has(key)));
};

export const isActionContentEqual = (a = {}, b = {}) => JSON.stringify(getActionComparable(a)) === JSON.stringify(getActionComparable(b));

export const OVERALL_AD_STATUS_OPTIONS = ['无广告', '仅 CPC', '仅 CPM', 'CPC+CPM'];
export const AD_STATUS_OPTIONS = OVERALL_AD_STATUS_OPTIONS;
export const BOOLEAN_STATUS_OPTIONS = ['开启', '关闭'];
export const CPM_POSITION_OPTIONS = ['仅搜索', '仅推荐', '搜索+推荐'];
export const CPM_BID_TYPE_OPTIONS = ['手动出价', '统一出价'];
export const CPM_SEARCH_MIN_BID = 450;
export const CPM_RECOMMEND_MIN_BID = 200;
export const BUDGET_ACTION_OPTIONS = ['加大预算', '降低预算', '保持预算', '暂停广告', '恢复广告'];
export const PRICE_ACTION_OPTIONS = ['涨价', '降价', '保持价格', '参加活动'];
export const IMAGE_ACTION_OPTIONS = ['更换主图', '未更换'];
export const KEYWORD_ACTION_OPTIONS = ['调整标题', '调整关键词', '未调整'];
export const STOCK_ACTION_OPTIONS = ['补货', '库存不足', '库存正常'];
export const ACTION_SOURCE_OPTIONS = ['manual', 'manual_modified', 'inherited_saved', 'json_import', 'excel_auto'];
export const ACTION_SOURCE_LABELS = {
  manual: '手动填写',
  manual_modified: '手动修改',
  inherited_saved: '继承后保存',
  json_import: 'JSON 导入',
  excel_auto: 'Excel 自动识别',
};
export const ACTION_SOURCE_PRIORITY = { excel_auto: 1, json_import: 2, inherited_saved: 3, manual: 3, manual_modified: 4 };

export const ACTION_FIELDS = [
  { key: 'adStatus', label: '整体广告状态', options: OVERALL_AD_STATUS_OPTIONS, disabled: true, group: '基础字段' },
  { key: 'budgetAction', label: '预算动作', options: BUDGET_ACTION_OPTIONS, group: '基础字段' },
  { key: 'priceAction', label: '价格动作', options: PRICE_ACTION_OPTIONS, group: '基础字段' },
  { key: 'imageAction', label: '主图动作', options: IMAGE_ACTION_OPTIONS, group: '基础字段' },
  { key: 'keywordAction', label: '关键词动作', options: KEYWORD_ACTION_OPTIONS, group: '基础字段' },
  { key: 'stockAction', label: '库存动作', options: STOCK_ACTION_OPTIONS, group: '基础字段' },
  { key: 'source', label: '记录来源', options: ACTION_SOURCE_OPTIONS, group: '基础字段' },
  { key: 'note', label: '总备注', type: 'textarea', group: '基础字段' },
  { key: 'cpcEnabled', label: 'CPC 是否开启', options: BOOLEAN_STATUS_OPTIONS, group: 'CPC 模块' },
  { key: 'cpcSearchBid', label: 'CPC 搜索出价', type: 'number', group: 'CPC 模块' },
  { key: 'cpcDailyBudget', label: 'CPC 每日预算', type: 'number', group: 'CPC 模块' },
  { key: 'cpcNote', label: 'CPC 备注', type: 'textarea', group: 'CPC 模块' },
  { key: 'cpmEnabled', label: 'CPM 是否开启', options: BOOLEAN_STATUS_OPTIONS, group: 'CPM 模块' },
  { key: 'cpmPosition', label: 'CPM 投放位置', options: CPM_POSITION_OPTIONS, group: 'CPM 模块' },
  { key: 'cpmBidType', label: 'CPM 出价方式', options: CPM_BID_TYPE_OPTIONS, group: 'CPM 模块' },
  { key: 'cpmSearchBid', label: 'CPM 搜索出价', type: 'number', group: 'CPM 模块' },
  { key: 'cpmRecommendBid', label: 'CPM 推荐出价', type: 'number', group: 'CPM 模块' },
  { key: 'cpmUnifiedBid', label: 'CPM 统一出价', type: 'number', group: 'CPM 模块' },
  { key: 'cpmDailyBudget', label: 'CPM 每日预算', type: 'number', group: 'CPM 模块' },
  { key: 'cpmNote', label: 'CPM 备注', type: 'textarea', group: 'CPM 模块' },
  { key: 'rawOperationAction', label: '运营动作原文', type: 'textarea', group: '基础字段' },
];

export const createEmptyAction = (date = '', sku = '') => ({
  date, sku, adStatus: '无广告', cpcEnabled: '关闭', cpcSearchBid: '', cpcDailyBudget: '', cpcNote: '',
  cpmEnabled: '关闭', cpmPosition: '', cpmBidType: '', cpmSearchBid: '', cpmRecommendBid: '', cpmUnifiedBid: '', cpmDailyBudget: '', cpmNote: '',
  budgetAction: '', priceAction: '', imageAction: '', keywordAction: '', stockAction: '', source: 'manual', rawOperationAction: '', note: '',
  adMode: '无广告', adPosition: '', dailyBudget: '',
});

const numberFields = new Set(['cpcSearchBid', 'cpcDailyBudget', 'cpmSearchBid', 'cpmRecommendBid', 'cpmUnifiedBid', 'cpmDailyBudget', 'dailyBudget', 'searchBid', 'recommendBid']);
const toNumberOrEmpty = (value) => {
  if (value === '' || value === null || value === undefined) return '';
  const parsed = Number(String(value).replace(/[,￥₽$\s]/g, ''));
  return Number.isNaN(parsed) ? '' : parsed;
};

const enabledFrom = (value) => ['开启', '是', 'true', '1', true].includes(value);
const disabledFrom = (value) => ['关闭', '否', 'false', '0', false].includes(value);
const normalizePosition = (value) => ({ 搜索: '仅搜索', 推荐: '仅推荐', 仅搜索: '仅搜索', 仅推荐: '仅推荐', '搜索+推荐': '搜索+推荐' }[value] || value || '');

const normalizeActionAliases = (action = {}) => {
  const next = { ...action };
  if (next.adMode && !next.adStatus) next.adStatus = next.adMode;
  if (next.adPosition && !next.cpmPosition) next.cpmPosition = normalizePosition(next.adPosition);
  if (next.dailyBudget !== '' && next.dailyBudget !== undefined) {
    if ((next.adMode === 'CPC' || next.cpcEnabled === '开启') && !next.cpcDailyBudget) next.cpcDailyBudget = next.dailyBudget;
    if ((next.adMode === 'CPM' || next.cpmEnabled === '开启') && !next.cpmDailyBudget) next.cpmDailyBudget = next.dailyBudget;
  }
  if (next.adType) {
    if (next.adType === 'CPC') next.cpcEnabled = '开启';
    if (String(next.adType).includes('CPM') || ['统一出价', '手动出价'].includes(next.adType)) next.cpmEnabled = '开启';
    if (next.adType === 'CPM搜索') next.cpmPosition = '仅搜索';
    if (next.adType === 'CPM推荐') next.cpmPosition = '仅推荐';
    if (['统一出价', '手动出价'].includes(next.adType)) next.cpmBidType = next.adType;
  }
  if (next.adMode === 'CPC') next.cpcEnabled = '开启';
  if (next.adMode === 'CPM') next.cpmEnabled = '开启';
  if (next.adMode === '无广告') { next.cpcEnabled = '关闭'; next.cpmEnabled = '关闭'; }
  if (next.adStatus === '仅 CPC') { next.cpcEnabled = '开启'; next.cpmEnabled = '关闭'; }
  if (next.adStatus === '仅 CPM') { next.cpcEnabled = '关闭'; next.cpmEnabled = '开启'; }
  if (next.adStatus === 'CPC+CPM') { next.cpcEnabled = '开启'; next.cpmEnabled = '开启'; }
  if (next.adStatus === '开启' && !next.adMode) next.cpmEnabled = '开启';
  if (next.adStatus === '关闭' || next.adStatus === '无广告数据') { next.cpcEnabled = '关闭'; next.cpmEnabled = '关闭'; }
  if (next.searchBid !== '' && next.searchBid !== undefined) {
    if (next.cpcEnabled === '开启' && (next.cpcSearchBid === undefined || next.cpcSearchBid === '')) next.cpcSearchBid = next.searchBid;
    if (next.cpmEnabled === '开启' && (next.cpmSearchBid === undefined || next.cpmSearchBid === '')) next.cpmSearchBid = next.searchBid;
  }
  if (next.recommendBid !== '' && next.recommendBid !== undefined && !next.cpmRecommendBid) next.cpmRecommendBid = next.recommendBid;
  return next;
};

export const calculateOverallAdStatus = (action = {}) => {
  const cpc = action.cpcEnabled === '开启' || enabledFrom(action.cpcEnabled);
  const cpm = action.cpmEnabled === '开启' || enabledFrom(action.cpmEnabled);
  if (cpc && cpm) return 'CPC+CPM';
  if (cpc) return '仅 CPC';
  if (cpm) return '仅 CPM';
  return '无广告';
};

export const getCpmMinBidForAction = (action = {}, field = 'cpmUnifiedBid') => {
  const position = normalizePosition(action.cpmPosition || action.adPosition);
  if (field === 'cpmSearchBid') return CPM_SEARCH_MIN_BID;
  if (field === 'cpmRecommendBid') return CPM_RECOMMEND_MIN_BID;
  if (field === 'cpmUnifiedBid') return ['仅推荐', '推荐'].includes(position) ? CPM_RECOMMEND_MIN_BID : CPM_SEARCH_MIN_BID;
  return 0;
};

export const validateCpmMinBids = (action = {}) => {
  const errors = [];
  const enabled = enabledFrom(action.cpmEnabled) || action.cpmEnabled === '开启';
  if (!enabled) return errors;
  const position = normalizePosition(action.cpmPosition || action.adPosition);
  const bidType = action.cpmBidType || '手动出价';
  const check = (field, label) => {
    const value = toNumberOrEmpty(action[field]);
    const min = getCpmMinBidForAction({ ...action, cpmPosition: position }, field);
    if (value !== '' && value < min) errors.push(`${label}不能低于 ${min}`);
  };
  if (bidType === '统一出价') check('cpmUnifiedBid', position === '仅推荐' ? 'CPM 推荐统一出价' : 'CPM 搜索/推荐统一出价');
  else {
    if (['仅搜索', '搜索', '搜索+推荐'].includes(position)) check('cpmSearchBid', 'CPM 搜索出价');
    if (['仅推荐', '推荐', '搜索+推荐'].includes(position)) check('cpmRecommendBid', 'CPM 推荐出价');
  }
  return errors;
};

export const applyAdRules = (action) => {
  const next = { ...action };
  next.cpcEnabled = enabledFrom(next.cpcEnabled) ? '开启' : disabledFrom(next.cpcEnabled) ? '关闭' : (next.cpcEnabled || '关闭');
  next.cpmEnabled = enabledFrom(next.cpmEnabled) ? '开启' : disabledFrom(next.cpmEnabled) ? '关闭' : (next.cpmEnabled || '关闭');
  next.cpmPosition = normalizePosition(next.cpmPosition);
  if (next.cpcEnabled !== '开启') { next.cpcSearchBid = ''; next.cpcDailyBudget = ''; }
  if (next.cpmEnabled !== '开启') {
    next.cpmPosition = ''; next.cpmBidType = ''; next.cpmSearchBid = ''; next.cpmRecommendBid = ''; next.cpmUnifiedBid = ''; next.cpmDailyBudget = '';
  } else {
    next.cpmPosition = next.cpmPosition || '搜索+推荐';
    next.cpmBidType = next.cpmBidType || '手动出价';
    if (next.cpmBidType === '统一出价') { next.cpmSearchBid = ''; next.cpmRecommendBid = ''; }
    if (next.cpmBidType === '手动出价') {
      next.cpmUnifiedBid = '';
      if (next.cpmPosition === '仅搜索') next.cpmRecommendBid = '';
      if (next.cpmPosition === '仅推荐') next.cpmSearchBid = '';
    }
  }
  next.adStatus = calculateOverallAdStatus(next);
  next.adMode = next.adStatus;
  next.adPosition = next.cpmPosition;
  next.dailyBudget = next.cpcDailyBudget || next.cpmDailyBudget || '';
  return next;
};

export const normalizeAction = (action) => {
  const aliased = normalizeActionAliases(action);
  const date = normalizeDateKey(aliased.date);
  const sku = normalizeSku(aliased.sku);
  const sourceAlias = { 'Excel 自动识别': 'excel_auto', 手动填写: 'manual', 手动修改: 'manual_modified', 继承后保存: 'inherited_saved', 'JSON 导入': 'json_import' };
  const normalized = { ...createEmptyAction(), ...aliased, date, sku, source: sourceAlias[aliased.source] || aliased.source || 'manual' };
  numberFields.forEach((field) => { if (field in normalized) normalized[field] = toNumberOrEmpty(normalized[field]); });
  return { ...applyAdRules(normalized), uniqueKey: buildActionKey(date, sku), updatedAt: normalized.updatedAt || new Date().toISOString() };
};

export const shouldReplaceAction = (incoming = {}, existing = {}) => {
  if (!existing?.uniqueKey) return true;
  const incomingPriority = ACTION_SOURCE_PRIORITY[incoming.source] || 0;
  const existingPriority = ACTION_SOURCE_PRIORITY[existing.source] || 0;
  if (incomingPriority !== existingPriority) return incomingPriority > existingPriority;
  return String(incoming.updatedAt || '') >= String(existing.updatedAt || '');
};

export const mergeActionRecords = (localActions = [], incomingActions = []) => {
  const merged = new Map(localActions.map((action) => [action.uniqueKey, action]));
  const stats = { added: 0, overwritten: 0, keptLocal: 0 };
  incomingActions.forEach((action) => {
    const existing = merged.get(action.uniqueKey);
    if (!existing) {
      merged.set(action.uniqueKey, action);
      stats.added += 1;
    } else if (shouldReplaceAction(action, existing)) {
      merged.set(action.uniqueKey, action);
      stats.overwritten += 1;
    } else {
      stats.keptLocal += 1;
    }
  });
  return { actions: [...merged.values()], stats };
};

const keyMap = new Map([
  ['广告状态', 'adStatus'], ['整体广告状态', 'adStatus'], ['广告模式', 'adMode'], ['广告位置', 'cpmPosition'], ['投放位置', 'cpmPosition'],
  ['CPC', 'cpcEnabled'], ['CPC是否开启', 'cpcEnabled'], ['CPC搜索出价', 'cpcSearchBid'], ['CPC预算', 'cpcDailyBudget'], ['CPC每日预算', 'cpcDailyBudget'], ['CPC备注', 'cpcNote'],
  ['CPM', 'cpmEnabled'], ['CPM是否开启', 'cpmEnabled'], ['CPM位置', 'cpmPosition'], ['CPM投放位置', 'cpmPosition'], ['CPM出价方式', 'cpmBidType'], ['出价方式', 'cpmBidType'],
  ['搜索出价', 'searchBid'], ['CPM搜索出价', 'cpmSearchBid'], ['推荐出价', 'cpmRecommendBid'], ['推荐位出价', 'cpmRecommendBid'], ['CPM推荐出价', 'cpmRecommendBid'], ['统一出价', 'cpmUnifiedBid'], ['CPM统一出价', 'cpmUnifiedBid'],
  ['每日预算', 'dailyBudget'], ['每日广告预算', 'dailyBudget'], ['CPM预算', 'cpmDailyBudget'], ['CPM每日预算', 'cpmDailyBudget'], ['CPM备注', 'cpmNote'],
  ['预算动作', 'budgetAction'], ['价格动作', 'priceAction'], ['主图动作', 'imageAction'], ['关键词动作', 'keywordAction'], ['库存动作', 'stockAction'], ['备注', 'note'], ['总备注', 'note'],
]);

export const parseOperationActionText = (text, date = '', sku = '') => {
  const raw = String(text || '').trim();
  if (!raw || !raw.includes('=')) return null;
  const parsed = { date, sku };
  let matched = 0;
  let unknown = 0;
  raw.split(/[；;]/).map((part) => part.trim()).filter(Boolean).forEach((part) => {
    const index = part.indexOf('=');
    if (index <= 0) { unknown += 1; return; }
    const key = part.slice(0, index).trim().replace(/\s+/g, '');
    const field = keyMap.get(key);
    if (!field) { unknown += 1; return; }
    parsed[field] = part.slice(index + 1).trim(); matched += 1;
  });
  if (!matched || unknown > 0) return null;
  parsed.rawOperationAction = raw; parsed.source = 'excel_auto';
  return normalizeAction(parsed);
};

export const actionToSummary = (action) => {
  if (!action) return '未填写结构化动作';
  const normalized = normalizeAction(action);
  const parts = [
    normalized.source && `来源：${ACTION_SOURCE_LABELS[normalized.source] || normalized.source}`,
    `整体广告状态：${normalized.adStatus}`,
    `CPC状态：${normalized.cpcEnabled}`,
    normalized.cpcSearchBid !== '' && `CPC搜索出价：${normalized.cpcSearchBid}`,
    normalized.cpcDailyBudget !== '' && `CPC预算：${normalized.cpcDailyBudget}`,
    normalized.cpcNote && `CPC备注：${normalized.cpcNote}`,
    `CPM状态：${normalized.cpmEnabled}`,
    normalized.cpmPosition && `CPM投放位置：${normalized.cpmPosition}`,
    normalized.cpmBidType && `CPM出价方式：${normalized.cpmBidType}`,
    normalized.cpmSearchBid !== '' && `CPM搜索出价：${normalized.cpmSearchBid}`,
    normalized.cpmRecommendBid !== '' && `CPM推荐出价：${normalized.cpmRecommendBid}`,
    normalized.cpmUnifiedBid !== '' && `CPM统一出价：${normalized.cpmUnifiedBid}`,
    normalized.cpmEnabled === '开启' && `最低出价提示：CPM搜索最低 ${CPM_SEARCH_MIN_BID}，CPM推荐最低 ${CPM_RECOMMEND_MIN_BID}`,
    normalized.cpmDailyBudget !== '' && `CPM预算：${normalized.cpmDailyBudget}`,
    normalized.cpmNote && `CPM备注：${normalized.cpmNote}`,
    normalized.budgetAction && `预算动作：${normalized.budgetAction}`,
    normalized.priceAction && `价格：${normalized.priceAction}`,
    normalized.imageAction && `主图：${normalized.imageAction}`,
    normalized.keywordAction && `关键词：${normalized.keywordAction}`,
    normalized.stockAction && `库存：${normalized.stockAction}`,
    normalized.note && `总备注：${normalized.note}`,
  ].filter(Boolean);
  return parts.length ? parts.join('；') : '未填写结构化动作';
};
