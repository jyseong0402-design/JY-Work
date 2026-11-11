// 반복 일정: 매일 / 매주 / 격주 / 매월 / 분기 / 매년
// 일정 하나(e)에 repeat, repeatUntil(종료일), skip(건너뛴 날짜들), doneDates(완료한 날짜들)를 저장하고
// 화면에 보여줄 때 기간 안의 날짜들로 펼칩니다.
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const z = (n) => String(n).padStart(2, '0');
const ymd = (d) => `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
const parse = (s) => new Date(s + 'T00:00:00');

export const REPEATS = [
  ['', '반복 안 함'], ['daily', '매일'], ['weekly', '매주'], ['biweekly', '격주 (2주마다)'],
  ['monthly', '매월'], ['quarterly', '분기 (3개월마다)'], ['yearly', '매년'],
];

export function repeatLabel(e) {
  if (!e.repeat) return '';
  const d = parse(e.date);
  return {
    daily: '매일',
    weekly: `매주 ${DOW[d.getDay()]}요일`,
    biweekly: `격주 ${DOW[d.getDay()]}요일`,
    monthly: `매월 ${d.getDate()}일`,
    quarterly: `분기마다 ${d.getDate()}일`,
    yearly: `매년 ${d.getMonth() + 1}월 ${d.getDate()}일`,
  }[e.repeat] || '';
}

// n번째 반복 날짜 (월 단위는 31일이 없는 달이면 그 달 마지막 날)
function nth(e, n) {
  const d = parse(e.date);
  const addMonths = (m) => {
    const day = d.getDate();
    const t = new Date(d.getFullYear(), d.getMonth() + m, 1);
    t.setDate(Math.min(day, new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate()));
    return t;
  };
  switch (e.repeat) {
    case 'daily': d.setDate(d.getDate() + n); return d;
    case 'weekly': d.setDate(d.getDate() + 7 * n); return d;
    case 'biweekly': d.setDate(d.getDate() + 14 * n); return d;
    case 'monthly': return addMonths(n);
    case 'quarterly': return addMonths(3 * n);
    case 'yearly': return addMonths(12 * n);
    default: return d;
  }
}

// from~to 안에 들어오는 이 일정의 날짜들
export function occurrences(e, from, to) {
  if (!e.date) return [];
  if (!e.repeat) return e.date >= from && e.date <= to ? [e.date] : [];
  const end = e.repeatUntil && e.repeatUntil < to ? e.repeatUntil : to;
  const skip = new Set(e.skip || []);
  const out = [];
  // 시작 근처로 빠르게 건너뛰기 (매일·매주는 계산으로)
  let n = 0;
  const span = (parse(from) - parse(e.date)) / 864e5;
  if (span > 0) {
    if (e.repeat === 'daily') n = Math.floor(span);
    else if (e.repeat === 'weekly') n = Math.floor(span / 7);
    else if (e.repeat === 'biweekly') n = Math.floor(span / 14);
    else if (e.repeat === 'monthly') n = Math.max(0, Math.floor(span / 31));
    else if (e.repeat === 'quarterly') n = Math.max(0, Math.floor(span / 93));
    else if (e.repeat === 'yearly') n = Math.max(0, Math.floor(span / 366));
  }
  for (let guard = 0; guard < 2000; guard++, n++) {
    const s = ymd(nth(e, n));
    if (s > end) break;
    if (s >= from && !skip.has(s)) out.push(s);
  }
  return out;
}

// 일정 목록을 기간 안의 날짜별 항목으로 펼침 (각 항목: 원래 일정 + date=그 날짜, occ=그 날짜)
export function expandEvents(events, from, to) {
  const out = [];
  for (const e of events) {
    for (const d of occurrences(e, from, to)) {
      out.push({ ...e, date: d, occ: d, startDate: e.date, done: e.repeat ? (e.doneDates || []).includes(d) : !!e.done });
    }
  }
  return out;
}
