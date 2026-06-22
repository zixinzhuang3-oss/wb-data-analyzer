export const CNY_TO_RUB = 11.5;

export const toProfitCny = (record = {}) => Number(record.profitCny ?? record.profit) || 0;
const roundCurrency = (value) => Math.round((Number(value) || 0) * 100) / 100;

export const toProfitRub = (record = {}) => {
  const explicit = Number(record.profitRub);
  return Number.isFinite(explicit) && record.profitRub !== undefined && record.profitRub !== null ? explicit : roundCurrency(toProfitCny(record) * CNY_TO_RUB);
};
