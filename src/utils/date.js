const DAY_MS = 86400000;
export const DEFAULT_TIME_ZONE = 'Asia/Shanghai';

const pad2 = (value) => String(value).padStart(2, '0');
const isIsoDate = (value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value || ''));

export const formatDateInTimeZone = (date = new Date(), timeZone = DEFAULT_TIME_ZONE) => {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date).reduce((acc, part) => ({ ...acc, [part.type]: part.value }), {});
  return `${parts.year}-${parts.month}-${parts.day}`;
};

export const addDays = (date, days) => {
  if (!isIsoDate(date)) return '';
  const [year, month, day] = date.split('-').map(Number);
  const next = new Date(Date.UTC(year, month - 1, day) + days * DAY_MS);
  return `${next.getUTCFullYear()}-${pad2(next.getUTCMonth() + 1)}-${pad2(next.getUTCDate())}`;
};

export const buildQuickRange = (preset, todayDate) => {
  if (preset === 'all') return { allDates: true, startDate: '', endDate: '' };
  if (!isIsoDate(todayDate)) return { allDates: false, startDate: '', endDate: '' };
  if (preset === 'yesterday') {
    const yesterday = addDays(todayDate, -1);
    return { allDates: false, startDate: yesterday, endDate: yesterday };
  }
  const days = preset === 'today' ? 1 : Number(preset);
  const safeDays = Number.isFinite(days) && days > 0 ? days : 1;
  return { allDates: false, startDate: addDays(todayDate, -(safeDays - 1)), endDate: todayDate };
};

const parseWorldTimeApiDate = (payload, timeZone = DEFAULT_TIME_ZONE) => {
  const datetime = payload?.datetime || payload?.utc_datetime;
  if (!datetime) return '';
  return formatDateInTimeZone(new Date(datetime), timeZone);
};

export const getBrowserTodayDate = (timeZone = DEFAULT_TIME_ZONE) => ({
  date: formatDateInTimeZone(new Date(), timeZone),
  source: '浏览器本地时间',
  timeZone,
});

export const getTodayDate = async ({ timeZone = DEFAULT_TIME_ZONE, fetchImpl = globalThis.fetch } = {}) => {
  if (typeof fetchImpl === 'function') {
    try {
      const response = await fetchImpl(`https://worldtimeapi.org/api/timezone/${encodeURIComponent(timeZone)}`, { cache: 'no-store' });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const date = parseWorldTimeApiDate(await response.json(), timeZone);
      if (isIsoDate(date)) return { date, source: '网络时间', timeZone };
    } catch {
      // Network time is best-effort. Fall back to the browser clock below.
    }
  }
  return getBrowserTodayDate(timeZone);
};
