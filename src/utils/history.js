const sum = (rows, key) => rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);

export const getSkuOptions = (records) => [...new Set(records.map((record) => record.sku).filter(Boolean))].sort();
export const getDateOptions = (records) => [...new Set(records.map((record) => record.date).filter(Boolean))].sort().reverse();

export const filterRecords = (records, filters) => records.filter((record) => {
  const dateMatched = !filters.date || record.date === filters.date;
  const skuMatched = !filters.sku || record.sku === filters.sku;
  return dateMatched && skuMatched;
});

export const summarizeRecords = (records) => ({
  dateCount: new Set(records.map((record) => record.date).filter(Boolean)).size,
  skuCount: new Set(records.map((record) => record.sku).filter(Boolean)).size,
  totalOrders: sum(records, 'totalOrders'),
  totalAdSpend: sum(records, 'adSpend'),
  totalProfit: sum(records, 'profit'),
  totalRevenue: sum(records, 'revenue'),
});

export const summarizeByDate = (records) => {
  const grouped = new Map();
  records.forEach((record) => {
    if (!grouped.has(record.date)) grouped.set(record.date, []);
    grouped.get(record.date).push(record);
  });
  return [...grouped.entries()].sort((a, b) => b[0].localeCompare(a[0])).map(([date, rows]) => ({
    date,
    skuCount: new Set(rows.map((row) => row.sku)).size,
    totalOrders: sum(rows, 'totalOrders'),
    totalAdSpend: sum(rows, 'adSpend'),
    totalProfit: sum(rows, 'profit'),
  }));
};
