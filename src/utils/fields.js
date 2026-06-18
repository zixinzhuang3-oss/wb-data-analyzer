export const AUXILIARY_SHEET_PATTERNS = [
  /利润定价表/i,
  /ozon/i,
  /wb利润/i,
  /^sheet\s*\d+$/i,
  /^sheet\d+$/i,
  /汇总/i,
  /说明/i,
  /模板/i,
];

export const SKU_SHEET_PATTERN = /^[A-Z]{2}\d{3,}[A-Z0-9]*$/i;

export const DAILY_FIELDS = [
  { key: 'date', label: '日期', aliases: ['日期', 'date', '统计日期', '时间'] },
  { key: 'sku', label: 'SKU', aliases: ['SKU', 'sku', '产品', '商品编码'] },
  { key: 'adStatus', label: '广告状态', aliases: ['广告状态'] },
  { key: 'linkId', label: '链接ID', aliases: ['链接ID', '链接id', '链接', 'linkId', 'link', 'nmId', 'NM ID'] },
  { key: 'operationAction', label: '运营动作原文', aliases: ['运营动作原文', '运营动作', '动作', '备注', '操作记录'] },
  { key: 'price', label: '价格', aliases: ['价格', '售价', 'price'] },
  { key: 'reviews', label: '评论数', aliases: ['评论数', '评论', 'reviews'] },
  { key: 'rating', label: '评分', aliases: ['评分', 'rating'] },
  { key: 'stock', label: '库存', aliases: ['库存', 'stock'] },
  { key: 'reviewOrders', label: '测评单', aliases: ['测评单', '测评订单'] },
  { key: 'actualOrders', label: '实际订单', aliases: ['实际订单', '真实订单'] },
  { key: 'totalOrders', label: '总订单', aliases: ['总订单', '订单', '订单数', 'totalOrders'] },
  { key: 'impressions', label: '曝光', aliases: ['曝光', '总曝光', 'impressions'] },
  { key: 'clicks', label: '点击', aliases: ['点击', '点击数', 'clicks'] },
  { key: 'ctr', label: '点击率', aliases: ['点击率', 'CTR', 'ctr'] },
  { key: 'addToCart', label: '加购', aliases: ['加购', '加购数', '购物车'] },
  { key: 'conversionRate', label: '转化率', aliases: ['转化率', 'CVR', 'cvr'] },
  { key: 'organicImpressions', label: '自然曝光', aliases: ['自然曝光', '自然展示'] },
  { key: 'organicClicks', label: '自然点击', aliases: ['自然点击'] },
  { key: 'organicOrders', label: '自然订单', aliases: ['自然订单'] },
  { key: 'adSpend', label: '广告费', aliases: ['广告费', '广告花费', 'adSpend', 'spend'] },
  { key: 'adOrders', label: '广告订单', aliases: ['广告订单'] },
  { key: 'adShare', label: '广告占比', aliases: ['广告占比'] },
  { key: 'adImpressions', label: '广告曝光', aliases: ['广告曝光'] },
  { key: 'adClicks', label: '广告点击', aliases: ['广告点击'] },
  { key: 'revenue', label: '销售额', aliases: ['销售额', '成交额', 'GMV', 'revenue'] },
  { key: 'commission', label: '佣金', aliases: ['佣金', '平台佣金'] },
  { key: 'russiaCost', label: '到俄成本', aliases: ['到俄成本', '到俄费用'] },
  { key: 'deliveryFee', label: '配送费', aliases: ['配送费', '物流费'] },
  { key: 'acquiringFee', label: '收单费', aliases: ['收单费'] },
  { key: 'storageFee', label: '仓储费', aliases: ['仓储费'] },
  { key: 'remittanceFee', label: '回款费', aliases: ['回款费'] },
  { key: 'profit', label: '利润', aliases: ['利润', 'profit'] },
  { key: 'keywordRank', label: '关键词排名', aliases: ['关键词排名', '排名', '关键词位置'] },
];

export const NUMERIC_FIELD_KEYS = new Set([
  'price', 'reviews', 'rating', 'stock', 'reviewOrders', 'actualOrders', 'totalOrders', 'impressions', 'clicks',
  'ctr', 'addToCart', 'conversionRate', 'organicImpressions', 'organicClicks', 'organicOrders', 'adSpend',
  'adOrders', 'adShare', 'adImpressions', 'adClicks', 'revenue', 'commission', 'russiaCost', 'deliveryFee',
  'acquiringFee', 'storageFee', 'remittanceFee', 'profit', 'keywordRank',
]);

export const isAuxiliarySheet = (sheetName) => AUXILIARY_SHEET_PATTERNS.some((pattern) => pattern.test(String(sheetName).trim()));

export const isSkuSheet = (sheetName) => {
  const name = String(sheetName).trim();
  return !isAuxiliarySheet(name) && SKU_SHEET_PATTERN.test(name);
};

export const fieldLabels = Object.fromEntries(DAILY_FIELDS.map((field) => [field.key, field.label]));
