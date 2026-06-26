export const DEFAULT_TIME_ZONE = 'Asia/Shanghai';

const pad2 = (value) => String(value).padStart(2, '0');
export const isIsoDate = (value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value || ''));

export const isValidDateKey = (value) => {
  if (!isIsoDate(value)) return false;
  const [year, month, day] = String(value).split('-').map(Number);
  const parsed = new Date(year, month - 1, day);
  return parsed.getFullYear() === year && parsed.getMonth() === month - 1 && parsed.getDate() === day;
};

export const formatLocalDateKey = (date = new Date()) => {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
};

export const normalizeDateKey = (value, XLSX = globalThis.XLSX) => {
  if (value === undefined || value === null || value === '') return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return formatLocalDateKey(value);
  if (typeof value === 'number') {
    const parsed = XLSX?.SSF?.parse_date_code?.(value);
    if (!parsed) return '';
    const key = `${parsed.y}-${pad2(parsed.m)}-${pad2(parsed.d)}`;
    return isValidDateKey(key) ? key : '';
  }
  const text = String(value).trim();
  let match = text.match(/^(\d{4})[\-/\.年](\d{1,2})[\-/\.月](\d{1,2})/);
  if (match) {
    const key = `${match[1]}-${pad2(match[2])}-${pad2(match[3])}`;
    return isValidDateKey(key) ? key : '';
  }
  match = text.match(/^(\d{1,2})\s*月\s*(\d{1,2})\s*日?$/);
  if (match) {
    const key = `${new Date().getFullYear()}-${pad2(match[1])}-${pad2(match[2])}`;
    return isValidDateKey(key) ? key : '';
  }
  match = text.match(/^(\d{1,2})[\-/\.](\d{1,2})$/);
  if (match) {
    const key = `${new Date().getFullYear()}-${pad2(match[1])}-${pad2(match[2])}`;
    return isValidDateKey(key) ? key : '';
  }
  return isValidDateKey(text) ? text : '';
};

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
  if (!isValidDateKey(date)) return '';
  const [year, month, day] = date.split('-').map(Number);
  const next = new Date(year, month - 1, day);
  next.setDate(next.getDate() + days);
  return formatLocalDateKey(next);
};

export const buildQuickRange = (preset, todayDate) => {
  if (preset === 'all') return { allDates: true, startDate: '', endDate: '' };
  if (!isValidDateKey(todayDate)) return { allDates: false, startDate: '', endDate: '' };
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
  const parsed = new Date(datetime);
  if (Number.isNaN(parsed.getTime())) return '';
  return formatDateInTimeZone(parsed, timeZone);
};

export const getBrowserTodayDate = (timeZone = DEFAULT_TIME_ZONE) => ({
  date: formatDateInTimeZone(new Date(), timeZone) || formatLocalDateKey(new Date()),
  source: '浏览器本地时间',
  timeZone,
});

export const getTodayDate = async ({ timeZone = DEFAULT_TIME_ZONE, fetchImpl = globalThis.fetch } = {}) => {
  if (typeof fetchImpl === 'function') {
    try {
      const response = await fetchImpl(`https://worldtimeapi.org/api/timezone/${encodeURIComponent(timeZone)}`, { cache: 'no-store' });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const date = parseWorldTimeApiDate(await response.json(), timeZone);
      if (isValidDateKey(date)) return { date, source: '网络时间', timeZone };
    } catch (error) {
      console.warn('获取网络时间失败，已回退到浏览器本地时间。', error);
    }
  }
  return getBrowserTodayDate(timeZone);
};
