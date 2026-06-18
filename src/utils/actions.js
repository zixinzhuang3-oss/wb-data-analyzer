export const AD_STATUS_OPTIONS = ['开启', '关闭', '无广告数据'];
export const AD_MODE_OPTIONS = ['无广告', 'CPC', 'CPM'];
export const AD_POSITION_OPTIONS = ['搜索', '推荐', '搜索+推荐'];
export const CPM_BID_TYPE_OPTIONS = ['手动出价', '统一出价'];
export const BUDGET_ACTION_OPTIONS = ['加大预算', '降低预算', '保持预算', '暂停广告', '恢复广告'];
export const PRICE_ACTION_OPTIONS = ['涨价', '降价', '保持价格', '参加活动'];
export const IMAGE_ACTION_OPTIONS = ['更换主图', '未更换'];
export const KEYWORD_ACTION_OPTIONS = ['调整标题', '调整关键词', '未调整'];
export const STOCK_ACTION_OPTIONS = ['补货', '库存不足', '库存正常'];
export const ACTION_SOURCE_OPTIONS = ['Excel 自动识别', '手动填写', '手动修改'];

export const ACTION_FIELDS = [
  { key: 'adStatus', label: '广告状态', options: AD_STATUS_OPTIONS },
  { key: 'adMode', label: '广告模式', options: AD_MODE_OPTIONS },
  { key: 'adPosition', label: '广告位置', options: AD_POSITION_OPTIONS },
  { key: 'cpmBidType', label: 'CPM 出价方式', options: CPM_BID_TYPE_OPTIONS },
  { key: 'cpcSearchBid', label: 'CPC 搜索出价', type: 'number' },
  { key: 'cpmSearchBid', label: 'CPM 搜索出价', type: 'number' },
  { key: 'cpmRecommendBid', label: 'CPM 推荐出价', type: 'number' },
  { key: 'cpmUnifiedBid', label: 'CPM 统一出价', type: 'number' },
  { key: 'dailyBudget', label: '每日广告预算', type: 'number' },
  { key: 'budgetAction', label: '预算动作', options: BUDGET_ACTION_OPTIONS },
  { key: 'priceAction', label: '价格动作', options: PRICE_ACTION_OPTIONS },
  { key: 'imageAction', label: '主图动作', options: IMAGE_ACTION_OPTIONS },
  { key: 'keywordAction', label: '关键词动作', options: KEYWORD_ACTION_OPTIONS },
  { key: 'stockAction', label: '库存动作', options: STOCK_ACTION_OPTIONS },
  { key: 'source', label: '记录来源', options: ACTION_SOURCE_OPTIONS },
  { key: 'rawOperationAction', label: '运营动作原文', type: 'textarea' },
  { key: 'note', label: '备注', type: 'textarea' },
];

export const createEmptyAction = (date = '', sku = '') => ({
  date, sku, adStatus: '', adMode: '', adPosition: '', cpmBidType: '', cpcSearchBid: '', cpmSearchBid: '', cpmRecommendBid: '', cpmUnifiedBid: '', dailyBudget: '',
  budgetAction: '', priceAction: '', imageAction: '', keywordAction: '', stockAction: '', source: '手动填写', rawOperationAction: '', note: '',
});

const numberFields = new Set(['cpcSearchBid', 'cpmSearchBid', 'cpmRecommendBid', 'cpmUnifiedBid', 'dailyBudget', 'searchBid', 'recommendBid']);
const toNumberOrEmpty = (value) => {
  if (value === '' || value === null || value === undefined) return '';
  const parsed = Number(String(value).replace(/[,￥₽$\s]/g, ''));
  return Number.isNaN(parsed) ? '' : parsed;
};

const normalizeActionAliases = (action = {}) => {
  const next = { ...action };
  if (!next.adMode && next.adType) {
    if (next.adType === 'CPC') next.adMode = 'CPC';
    else if (String(next.adType).includes('CPM') || ['统一出价', '手动出价'].includes(next.adType)) next.adMode = 'CPM';
  }
  if (!next.adPosition && next.adType === 'CPM搜索') next.adPosition = '搜索';
  if (!next.adPosition && next.adType === 'CPM推荐') next.adPosition = '推荐';
  if (!next.cpmBidType && ['统一出价', '手动出价'].includes(next.adType)) next.cpmBidType = next.adType;
  if (next.searchBid !== '' && next.searchBid !== undefined) {
    if (next.adMode === 'CPC' && (next.cpcSearchBid === undefined || next.cpcSearchBid === '')) next.cpcSearchBid = next.searchBid;
    if (next.adMode === 'CPM' && (next.cpmSearchBid === undefined || next.cpmSearchBid === '')) next.cpmSearchBid = next.searchBid;
  }
  if (next.recommendBid !== '' && next.recommendBid !== undefined && (next.cpmRecommendBid === undefined || next.cpmRecommendBid === '')) next.cpmRecommendBid = next.recommendBid;
  return next;
};

export const applyAdRules = (action) => {
  const next = { ...action };
  if (next.adMode === '无广告') {
    next.adStatus = next.adStatus || '关闭'; next.adPosition = ''; next.cpmBidType = ''; next.cpcSearchBid = ''; next.cpmSearchBid = ''; next.cpmRecommendBid = ''; next.cpmUnifiedBid = '';
  } else if (next.adMode === 'CPC') {
    next.adStatus = next.adStatus || '开启'; next.adPosition = '搜索'; next.cpmBidType = ''; next.cpmSearchBid = ''; next.cpmRecommendBid = ''; next.cpmUnifiedBid = '';
  } else if (next.adMode === 'CPM') {
    next.adStatus = next.adStatus || '开启'; next.cpcSearchBid = '';
    if (next.cpmBidType === '统一出价') { next.cpmSearchBid = ''; next.cpmRecommendBid = ''; }
    if (next.cpmBidType === '手动出价') {
      next.cpmUnifiedBid = '';
      if (next.adPosition === '搜索') next.cpmRecommendBid = '';
      if (next.adPosition === '推荐') next.cpmSearchBid = '';
    }
  }
  return next;
};

export const normalizeAction = (action) => {
  const aliased = normalizeActionAliases(action);
  const date = String(aliased.date || '').trim();
  const sku = String(aliased.sku || '').trim();
  const normalized = { ...createEmptyAction(), ...aliased, date, sku };
  numberFields.forEach((field) => { if (field in normalized) normalized[field] = toNumberOrEmpty(normalized[field]); });
  return { ...applyAdRules(normalized), uniqueKey: `${date}__${sku}`, updatedAt: new Date().toISOString() };
};

const keyMap = new Map([
  ['广告状态', 'adStatus'], ['广告模式', 'adMode'], ['广告位置', 'adPosition'], ['出价方式', 'cpmBidType'], ['CPM出价方式', 'cpmBidType'],
  ['搜索出价', 'searchBid'], ['CPC搜索出价', 'cpcSearchBid'], ['CPM搜索出价', 'cpmSearchBid'], ['推荐出价', 'cpmRecommendBid'], ['推荐位出价', 'cpmRecommendBid'], ['CPM推荐出价', 'cpmRecommendBid'], ['统一出价', 'cpmUnifiedBid'],
  ['每日预算', 'dailyBudget'], ['每日广告预算', 'dailyBudget'], ['预算动作', 'budgetAction'], ['价格动作', 'priceAction'], ['主图动作', 'imageAction'], ['关键词动作', 'keywordAction'], ['库存动作', 'stockAction'], ['备注', 'note'],
]);

export const parseOperationActionText = (text, date = '', sku = '') => {
  const raw = String(text || '').trim();
  if (!raw || !raw.includes('=')) return null;
  const parsed = createEmptyAction(date, sku);
  let matched = 0;
  raw.split(/[；;]/).map((part) => part.trim()).filter(Boolean).forEach((part) => {
    const index = part.indexOf('=');
    if (index <= 0) return;
    const key = part.slice(0, index).trim().replace(/\s+/g, '');
    const field = keyMap.get(key);
    if (!field) return;
    parsed[field] = part.slice(index + 1).trim(); matched += 1;
  });
  if (!matched) return null;
  parsed.rawOperationAction = raw; parsed.source = 'Excel 自动识别';
  return normalizeAction(parsed);
};

export const actionToSummary = (action) => {
  if (!action) return '未填写结构化动作';
  const normalized = normalizeAction(action);
  const parts = [
    normalized.source && `来源：${normalized.source}`,
    normalized.adStatus && `广告状态：${normalized.adStatus}`,
    normalized.adMode && `广告模式：${normalized.adMode}`,
    normalized.adPosition && `广告位置：${normalized.adPosition}`,
    normalized.cpmBidType && `出价方式：${normalized.cpmBidType}`,
    normalized.cpcSearchBid !== '' && `CPC搜索出价：${normalized.cpcSearchBid}`,
    normalized.cpmSearchBid !== '' && `CPM搜索出价：${normalized.cpmSearchBid}`,
    normalized.cpmRecommendBid !== '' && `CPM推荐出价：${normalized.cpmRecommendBid}`,
    normalized.cpmUnifiedBid !== '' && `统一出价：${normalized.cpmUnifiedBid}`,
    normalized.dailyBudget !== '' && `预算：${normalized.dailyBudget}`,
    normalized.budgetAction && `预算动作：${normalized.budgetAction}`,
    normalized.priceAction && `价格：${normalized.priceAction}`,
    normalized.imageAction && `主图：${normalized.imageAction}`,
    normalized.keywordAction && `关键词：${normalized.keywordAction}`,
    normalized.stockAction && `库存：${normalized.stockAction}`,
    normalized.note && `备注：${normalized.note}`,
  ].filter(Boolean);
  return parts.length ? parts.join('；') : '未填写结构化动作';
};
