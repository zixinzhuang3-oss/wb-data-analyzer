const FIELD_ALIASES = {
  date: ['date', '日期', 'day', '统计日期'],
  sku: ['sku', 'SKU', '产品', '商品编码', '商品'],
  link: ['link', '链接', 'url', 'URL', '链接ID', 'listing'],
  orders: ['orders', '订单', '订单数', 'sales_orders', '销量'],
  units: ['units', '件数', '销售件数', '数量'],
  revenue: ['revenue', '销售额', '成交额', 'GMV', 'gmv'],
  productCost: ['productCost', '产品成本', '货品成本', '采购成本'],
  shippingCost: ['shippingCost', '物流成本', '运费', '头程', '尾程'],
  fee: ['fee', '平台费', '佣金', '手续费'],
  spend: ['spend', '广告花费', '广告费', 'ad_spend', '花费'],
  impressions: ['impressions', '曝光', '展示'],
  clicks: ['clicks', '点击', '点击数'],
  adOrders: ['adOrders', '广告订单', '广告订单数'],
  adRevenue: ['adRevenue', '广告销售额', '广告成交额'],
  visitors: ['visitors', '访客', '访客数', 'UV', 'uv'],
  sessions: ['sessions', '会话', 'sessions', '浏览量'],
  conversionRate: ['conversionRate', '转化率', 'CVR', 'cvr'],
};

const parseCsvText = (text) => {
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (char === '"' && inQuotes && next === '"') {
      cell += '"';
      i += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      row.push(cell.trim());
      cell = '';
    } else if ((char === '\n' || char === '\r') && !inQuotes) {
      if (char === '\r' && next === '\n') i += 1;
      row.push(cell.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += char;
    }
  }

  row.push(cell.trim());
  if (row.some(Boolean)) rows.push(row);
  return rows;
};

const toNumber = (value) => {
  if (value === undefined || value === null || value === '') return 0;
  const cleaned = String(value).replace(/[%,$￥,\s]/g, '');
  const parsed = Number(cleaned);
  if (Number.isNaN(parsed)) return 0;
  return String(value).includes('%') ? parsed / 100 : parsed;
};

const findField = (row, canonical) => FIELD_ALIASES[canonical].find((field) => Object.prototype.hasOwnProperty.call(row, field));

export const parseCsvFile = async (file) => {
  const text = await file.text();
  const [headers, ...records] = parseCsvText(text);
  if (!headers?.length) return [];
  return records.map((record) => {
    const row = Object.fromEntries(headers.map((header, index) => [header.trim(), record[index] ?? '']));
    return normalizeRow(row);
  });
};

export const normalizeRow = (row) => {
  const normalized = {};
  Object.keys(FIELD_ALIASES).forEach((canonical) => {
    const field = findField(row, canonical);
    const value = field ? row[field] : undefined;
    normalized[canonical] = ['date', 'sku', 'link'].includes(canonical) ? (value ? String(value).trim() : '') : toNumber(value);
  });
  normalized.id = `${normalized.date}-${normalized.sku}-${normalized.link}-${Math.random().toString(36).slice(2)}`;
  return normalized;
};
