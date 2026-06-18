export const AD_TYPE_OPTIONS = ['CPC', 'CPM搜索', 'CPM推荐', '统一出价', '手动出价'];
export const BUDGET_ACTION_OPTIONS = ['加大预算', '降低预算', '保持预算', '暂停广告'];
export const PRICE_ACTION_OPTIONS = ['涨价', '降价', '保持价格', '参加活动'];
export const IMAGE_ACTION_OPTIONS = ['更换主图', '未更换'];
export const KEYWORD_ACTION_OPTIONS = ['调整标题', '调整关键词', '未调整'];
export const STOCK_ACTION_OPTIONS = ['补货', '库存不足', '库存正常'];

export const ACTION_FIELDS = [
  { key: 'adType', label: '广告类型', options: AD_TYPE_OPTIONS },
  { key: 'searchBid', label: '搜索出价', type: 'number' },
  { key: 'recommendBid', label: '推荐位出价', type: 'number' },
  { key: 'dailyBudget', label: '每日广告预算', type: 'number' },
  { key: 'budgetAction', label: '预算动作', options: BUDGET_ACTION_OPTIONS },
  { key: 'priceAction', label: '价格动作', options: PRICE_ACTION_OPTIONS },
  { key: 'imageAction', label: '主图动作', options: IMAGE_ACTION_OPTIONS },
  { key: 'keywordAction', label: '关键词动作', options: KEYWORD_ACTION_OPTIONS },
  { key: 'stockAction', label: '库存动作', options: STOCK_ACTION_OPTIONS },
  { key: 'note', label: '备注', type: 'textarea' },
];

export const createEmptyAction = (date = '', sku = '') => ({
  date,
  sku,
  adType: '',
  searchBid: '',
  recommendBid: '',
  dailyBudget: '',
  budgetAction: '',
  priceAction: '',
  imageAction: '',
  keywordAction: '',
  stockAction: '',
  note: '',
});

const toNumberOrEmpty = (value) => {
  if (value === '' || value === null || value === undefined) return '';
  const parsed = Number(String(value).replace(/[,￥₽$\s]/g, ''));
  return Number.isNaN(parsed) ? '' : parsed;
};

export const normalizeAction = (action) => {
  const date = String(action.date || '').trim();
  const sku = String(action.sku || '').trim();
  return {
    ...createEmptyAction(),
    ...action,
    date,
    sku,
    searchBid: toNumberOrEmpty(action.searchBid),
    recommendBid: toNumberOrEmpty(action.recommendBid),
    dailyBudget: toNumberOrEmpty(action.dailyBudget),
    uniqueKey: `${date}__${sku}`,
    updatedAt: new Date().toISOString(),
  };
};

export const actionToSummary = (action) => {
  if (!action) return '未填写结构化动作';
  const parts = [
    action.adType && `广告：${action.adType}`,
    action.searchBid !== '' && action.searchBid !== undefined && `搜索出价：${action.searchBid}`,
    action.recommendBid !== '' && action.recommendBid !== undefined && `推荐出价：${action.recommendBid}`,
    action.dailyBudget !== '' && action.dailyBudget !== undefined && `预算：${action.dailyBudget}`,
    action.budgetAction && `预算动作：${action.budgetAction}`,
    action.priceAction && `价格：${action.priceAction}`,
    action.imageAction && `主图：${action.imageAction}`,
    action.keywordAction && `关键词：${action.keywordAction}`,
    action.stockAction && `库存：${action.stockAction}`,
    action.note && `备注：${action.note}`,
  ].filter(Boolean);
  return parts.length ? parts.join('；') : '未填写结构化动作';
};
