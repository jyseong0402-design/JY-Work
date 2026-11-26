// 📦 수불재고 보고 구역 (1.39.0~): 점포별 월·주간 매출, 직납 발주 대비 출고(미납), 로스(유형별·주간·품목), 증정·시식, 행사
// stock.js의 parseStockFile이 readExtra(read)를 불러 숫자만 저장하고, 화면은 extraHtml(s)로 그림
//  - '1-1 점포별 월매출' : 'N월' 합계 줄 × 채널(롯데 특정·직납, 이마트, 케이마켓, 카메레오, KOREA MART, 송캉)·점포 열 (원)
//  - '1-2 점포별 주간매출' : B열 주차·C열 빈칸 = 주 합계 줄
//  - '2-1 (직납)발주-입고량 차이' : C열 '주차' 블록(롯데·케이마켓·이마트) × wk 3칸(발주·출고·차이 kg), total 줄 + SKU 줄
//  - '2-2 특정 로스 유형별 분류' : 월별 합계(직납+창고+특정 폐기) + 유형별(직납·창고·특정 폐기·특정 1+1) × 품목, 아래 '주차별' 표
//  - '2-2 직납로스(일, 주, 월)발주일기준' : 주별 직납 로스율·매출물량(목표 3%), 월별 합계(1+1 포함)
//  - '매장별 증정품 매입' : 주차 블록(롯데·이마트 krw·kg, 점포, SKU)
//  - '행사' : 오른쪽 달력(1행 'N월', 2행 일, 품목 줄 메모) + 'Event/Start/End/Product' 목록

export const EXTRA_SHEETS = {
  storeM: /^1-1\s*점포별\s*월매출/,
  storeW: /^1-2\s*점포별\s*주간매출/,
  orders: /^2-1\s*\(직납\)\s*발주/,
  lossType: /^2-2\s*특정\s*로스\s*유형별/,
  lossDirect: /^2-2\s*직납로스/,
  gifts: /^매장별\s*증정품/,
  events: /^행사$/,
};

const num = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
const isInt = (v) => typeof v === 'number' && Number.isInteger(v);
const serialDate = (v) => (typeof v === 'number' && v > 40000 && v < 60000 ? new Date(Math.round((v - 25569) * 864e5)) : null);
const ymd = (d) => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
const r1 = (v) => Math.round(num(v) * 10) / 10;
const r3 = (v) => Math.round(num(v) * 1000) / 1000;

/* ================= 읽기 ================= */
// 점포 열: 채널 묶음(total·(ind)·(direct)·KOREA MART·송캉) 아래 점포
const GRP_RE = /total|\(ind\)|\(direct\)|^(korea mart|송캉|이마트|케이마켓|카메레오)$/i; // 주간 시트는 이마트 합계 열 이름이 그냥 '이마트'
const grpName = (n) => n.replace(/^롯데\s*\(ind\)$/i, '롯데 특정').replace(/^롯데\s*\(direct\)$/i, '롯데 직납').replace(/\s*total$/i, '').trim() || '전체';
function storeCols(g, hr) {
  const cols = [];
  let grp = '';
  for (let c = 4; c <= g.maxC; c++) {
    const n = g.txt(hr, c);
    if (!n) continue;
    if (/^(week|month|year|입고일기준|품목)$/i.test(n)) continue;
    if (GRP_RE.test(n)) {
      const name = grpName(n);
      const top = /^total$/i.test(n.trim()) || /^롯데 total$/i.test(n.trim());
      cols.push({ c, n: name, k: top ? 'top' : 'grp' });
      if (!top) grp = name;
    } else cols.push({ c, n, k: 'store', g: grp });
  }
  return cols;
}
function readStoreMonthly(g) {
  const h = g.find(/^롯데 total$/);
  if (!h) return null;
  const cols = storeCols(g, h.r);
  const months = {};
  for (let r = h.r + 1; r <= g.maxR; r++) {
    const m = /^(\d{1,2})월$/.exec(g.txt(r, 2));
    if (!m) continue;
    const v = cols.map((x) => Math.round(num(g.at(r, x.c))));
    if (v.some((x) => x)) months[+m[1]] = v;
  }
  return { cols: cols.map(({ n, k, g: gg }) => ({ n, k, g: gg || '' })), months };
}
function readStoreWeekly(g, keep = 10) {
  const h = g.find(/^롯데 total$/);
  if (!h) return null;
  const cols = storeCols(g, h.r);
  const weeks = [];
  for (let r = h.r + 1; r <= g.maxR; r++) {
    const wk = g.at(r, 2);
    if (!isInt(wk) || g.txt(r, 3)) continue;
    const v = cols.map((x) => Math.round(num(g.at(r, x.c))));
    if (v.some((x) => x)) weeks.push({ wk, v });
  }
  return { cols: cols.map(({ n, k, g: gg }) => ({ n, k, g: gg || '' })), weeks: weeks.slice(-keep) };
}
// 직납 발주 대비 출고: 'wkN' 3칸(발주·출고·차이)
function readOrders(g, keepW = 10, keepSku = 4) {
  const out = [];
  for (let r = 1; r <= g.maxR; r++) {
    if (g.txt(r, 3) !== '주차') continue;
    const name = g.txt(r, 1) || `직납 ${out.length + 1}`;
    const byWk = new Map();
    const colWk = [];
    for (let c = 5; c <= g.maxC; c++) {
      const m = /^wk\s*(\d+)/i.exec(g.txt(r, c));
      if (!m || !/발주/.test(g.txt(r + 1, c))) continue;
      colWk.push({ c, wk: +m[1] });
    }
    const tr = r + 2;
    for (const { c, wk } of colWk) {
      const o = byWk.get(wk) || { wk, order: 0, ship: 0 };
      o.order += num(g.at(tr, c)); o.ship += num(g.at(tr, c + 1));
      byWk.set(wk, o);
    }
    const weeks = [...byWk.values()].filter((x) => x.order > 0).sort((a, b) => a.wk - b.wk).slice(-keepW).map((x) => ({ wk: x.wk, order: r1(x.order), ship: r1(x.ship) }));
    if (!weeks.length) continue;
    const lastWks = new Set(weeks.slice(-keepSku).map((x) => x.wk));
    const skus = [];
    for (let rr = tr + 1; rr <= g.maxR; rr++) {
      const n = g.txt(rr, 3);
      if (!n || n === '주차') break;
      let short = 0, order = 0;
      for (const { c, wk } of colWk) if (lastWks.has(wk)) { order += num(g.at(rr, c)); short += Math.max(0, num(g.at(rr, c)) - num(g.at(rr, c + 1))); }
      if (order > 0) skus.push({ n, order: r1(order), short: r1(short) });
    }
    out.push({ name, weeks, skus: skus.sort((a, b) => b.short - a.short).slice(0, 5) });
  }
  return out.length ? out : null;
}
// 로스 유형별 (월별·주별)
const LOSS_TYPES = [['direct', /^직납로스/, '직납'], ['wh', /^창고로스/, '창고'], ['spec', /^특정\s*폐기/, '특정 폐기'], ['oneplus', /1\+1/, '특정 1+1']];
function readLossType(g) {
  const h = g.find(/^직납로스$/);
  if (!h) return null;
  const hr = h.r, hr4 = hr + 1;
  const mc = g.find(/^Month$/, null, [hr4, hr4]);
  if (!mc) return null;
  const items = []; for (let c = mc.c + 4; c < mc.c + 11; c++) { const n = g.txt(hr4, c); if (n) items.push(n); }
  const blocks = LOSS_TYPES.map(([key, re]) => { for (let c = 1; c <= g.maxC; c++) if (re.test(g.txt(hr, c))) return { key, c }; return null; }).filter(Boolean);
  const months = {};
  for (let r = hr4 + 1; r <= hr4 + 14; r++) {
    const m = g.at(r, mc.c);
    if (!isInt(m) || m < 1 || m > 12 || g.at(r, mc.c + 1) === undefined) continue;
    const o = { total: r1(g.at(r, mc.c + 1)), qty: r1(g.at(r, mc.c + 2)), rate: r3(g.at(r, mc.c + 3)), items: items.map((_, i) => r1(g.at(r, mc.c + 4 + i))) };
    for (const b of blocks) {
      const hasQty = /Loss/i.test(g.txt(hr4, b.c));
      o[b.key] = { loss: r1(g.at(r, b.c)), qty: hasQty ? r1(g.at(r, b.c - 2)) : 0, rate: hasQty ? r3(g.at(r, b.c - 1)) : 0, items: items.map((_, i) => r1(g.at(r, b.c + 1 + i))) };
    }
    if (o.total || o.qty) months[m] = o;
  }
  // 주차별 표
  const weeks = [];
  const w = g.find(/^주차별$/, [2]);
  if (w) {
    for (let r = w.r + 2; r <= g.maxR; r++) {
      const wk = g.at(r, 2);
      if (!isInt(wk)) { if (weeks.length) break; continue; }
      const qty = num(g.at(r, 4));
      if (!qty) continue;
      const d = serialDate(g.at(r, 1));
      weeks.push({ wk, start: d ? ymd(d) : '', total: r1(g.at(r, 3)), qty: r1(qty), rate: r3(g.at(r, 5)) });
    }
  }
  return { items, months, weeks: weeks.slice(-10) };
}
// 직납로스 시트: 주별 직납 로스율(목표) + 월 합계(1+1 포함)
function readLossDirect(g) {
  // 주별 블록: 'Week' 바로 오른쪽이 'loss%'인 칸 (왼쪽 일별 표의 'Week'와 구분)
  let h = null;
  for (let r = 1; r <= 8 && !h; r++) for (let c = 1; c <= g.maxC; c++) if (/^Week$/i.test(g.txt(r, c)) && /loss/i.test(g.txt(r, c + 1))) { h = { r, c }; break; }
  const out = { weeks: [], target: 0.03, months: {} };
  if (h) {
    const tt = g.txt(h.r - 1, h.c); const tm = /(\d+(?:\.\d+)?)\s*%/.exec(tt); if (tm) out.target = +tm[1] / 100;
    for (let r = h.r + 1; r <= g.maxR; r++) {
      const wk = g.at(r, h.c);
      if (!isInt(wk)) { if (out.weeks.length) break; continue; }
      const qty = num(g.at(r, h.c + 2));
      if (qty) out.weeks.push({ wk, rate: r3(g.at(r, h.c + 1)), qty: r1(qty), loss: r1(g.at(r, h.c + 3)) });
    }
    out.weeks = out.weeks.slice(-10);
  }
  const t = g.find(/^Total Loss$/);
  if (t) for (let r = t.r + 1; r <= t.r + 13; r++) { const m = g.at(r, t.c - 1); if (isInt(m) && m >= 1 && m <= 12 && num(g.at(r, t.c))) out.months[m] = { total: r1(g.at(r, t.c)), qty: r1(g.at(r, t.c + 1)), rate: r3(g.at(r, t.c + 2)) }; }
  return out;
}
// 증정품: '롯데(krw)' 머리 줄마다 한 주
function readGifts(g, keep = 6) {
  const out = [];
  for (let r = 1; r <= g.maxR; r++) {
    if (!/^롯데\(krw\)$/.test(g.txt(r, 5))) continue;
    const wk = g.at(r, 3);
    const pos = {}; for (let c = 1; c <= g.maxC; c++) { const t = g.txt(r, c); if (/^(롯데|이마트)\((krw|kg)\)$/.test(t)) pos[t] = c; }
    const stores = (key, kgKey) => { const res = []; const c0 = pos[key], k0 = pos[kgKey]; if (!c0) return res; for (let c = c0 + 1; c < c0 + 8; c++) { const n = g.txt(r, c); if (!n || /\(kg\)|단가/.test(n)) break; res.push({ n, krw: Math.round(num(g.at(r + 1, c))), kg: k0 ? r1(g.at(r + 1, k0 + (c - c0))) : 0 }); } return res; };
    const side = (key, kgKey) => (pos[key] ? { krw: Math.round(num(g.at(r + 1, pos[key]))), kg: pos[kgKey] ? r1(g.at(r + 1, pos[kgKey])) : 0, stores: stores(key, kgKey) } : null);
    const skus = [];
    for (let rr = r + 2; rr <= g.maxR; rr++) {
      const n = g.txt(rr, 4);
      if (!n || /^롯데\(krw\)$/.test(g.txt(rr, 5))) break;
      const krw = num(g.at(rr, pos['롯데(krw)'])) + (pos['이마트(krw)'] ? num(g.at(rr, pos['이마트(krw)'])) : 0);
      const kg = num(g.at(rr, pos['롯데(kg)'])) + (pos['이마트(kg)'] ? num(g.at(rr, pos['이마트(kg)'])) : 0);
      if (krw || kg) skus.push({ n, krw: Math.round(krw), kg: r1(kg) });
    }
    if (isInt(wk)) out.push({ wk, lotte: side('롯데(krw)', '롯데(kg)'), emart: side('이마트(krw)', '이마트(kg)'), skus: skus.sort((a, b) => b.krw - a.krw).slice(0, 5) });
  }
  return out.length ? out.sort((a, b) => a.wk - b.wk).slice(-keep) : null;
}
// 행사: 달력 메모 + Event 목록
function readEvents(g, baseDate) {
  const out = [];
  const by = +baseDate.slice(0, 4), bm = +baseDate.slice(5, 7);
  const yOf = (m) => (m < bm - 6 ? by + 1 : m > bm + 6 ? by - 1 : by);
  const p = g.find(/^품목명$/);
  if (p) {
    let monRow = 0, dayRow = 0;
    for (let r = 1; r < p.r; r++) for (let c = p.c + 1; c <= g.maxC; c++) { if (/^\d{1,2}월$/.test(g.txt(r, c))) monRow = monRow || r; if (g.at(r, c) === 1 && g.at(r, c + 1) === 2) dayRow = dayRow || r; }
    if (monRow && dayRow) {
      const monAt = (c) => { for (let x = c; x > p.c; x--) { const m = /^(\d{1,2})월$/.exec(g.txt(monRow, x)); if (m) return +m[1]; } return 0; };
      let prod = '';
      for (let r = p.r + 1; r <= p.r + 25; r++) {
        if (/^STT$/i.test(g.txt(r, p.c))) break;
        prod = g.txt(r, p.c) || prod;
        for (let c = p.c + 3; c <= g.maxC; c++) {
          const note = g.txt(r, c), day = g.at(dayRow, c), m = monAt(c);
          if (!note || !isInt(day) || !m) continue;
          const d = `${yOf(m)}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
          out.push({ from: d, to: d, product: prod, note });
        }
      }
    }
  }
  const e = g.find(/^Event$/);
  if (e && /^Start$/i.test(g.txt(e.r, e.c + 1))) {
    let blank = 0;
    for (let r = e.r + 1; r <= e.r + 60 && blank < 3; r++) {
      const name = g.txt(r, e.c), s = serialDate(g.at(r, e.c + 1)), t = serialDate(g.at(r, e.c + 2));
      if (!name) { blank++; continue; }
      blank = 0;
      if (!s && !t) continue;
      const price = [num(g.at(r, e.c + 4)) && `직납 ${Math.round(num(g.at(r, e.c + 4))).toLocaleString('ko-KR')}`, num(g.at(r, e.c + 6)) && `특정 ${Math.round(num(g.at(r, e.c + 6))).toLocaleString('ko-KR')}`].filter(Boolean).join(' / ');
      out.push({ from: ymd(s || t), to: ymd(t || s), product: g.txt(r, e.c + 3), note: `${name}${price ? ` (${price} VND)` : ''}` });
    }
  }
  return out.sort((a, b) => a.from.localeCompare(b.from));
}

export async function readExtra(read, baseDate) {
  const x = {};
  const tryRead = async (key, fn) => { try { const g = await read(EXTRA_SHEETS[key]); if (g) x[key] = fn(g); } catch (err) { console.warn('수불재고', key, err); } };
  await tryRead('storeM', readStoreMonthly);
  await tryRead('storeW', readStoreWeekly);
  await tryRead('orders', readOrders);
  await tryRead('lossType', readLossType);
  await tryRead('lossDirect', readLossDirect);
  await tryRead('gifts', readGifts);
  await tryRead('events', (g) => readEvents(g, baseDate));
  return x;
}

/* ================= 분석 ================= */
const f0 = (v) => Math.round(num(v)).toLocaleString('ko-KR');
const f1 = (v) => (Math.round(num(v) * 10) / 10).toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const mil = (w) => f1(num(w) / 1e6);
const pc1 = (r) => `${(Math.round(num(r) * 1000) / 10).toFixed(1)}%`;
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const chg = (a, b) => (b ? (a / b - 1) : null);
const chgTxt = (r) => (r === null ? '-' : `${r >= 0 ? '+' : '−'}${Math.abs(Math.round(r * 100))}%`);
const md = (s) => { const [, m, d] = String(s).split('-'); return `${+m}/${+d}`; };
const addDays = (s, n) => { const d = new Date(s + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return ymd(d); };

// 주차 → 시작일 (로스 주차별 표의 '주 시작일'로 맞춤, 없으면 기준일에서 거꾸로)
// 시작일이 비어 있는 주(진행 중인 주 등)는 알고 있는 주에서 7일씩 계산
function wkStart(x, wk) {
  const known = (x.lossType?.weeks || []).filter((w) => w.start);
  const hit = known.find((w) => w.wk === wk);
  if (hit) return hit.start;
  const a = known[known.length - 1];
  return a ? addDays(a.start, (wk - a.wk) * 7) : '';
}
const wkLabel = (x, wk) => { const s = wkStart(x, wk); return s ? `${md(s)}~` : `wk${wk}`; };
// 가장 최근 주차 (다른 표들의 마지막 주)
const curWk = (x) => Math.max(0, ...(x.storeW?.weeks || []).map((w) => w.wk), ...(x.lossType?.weeks || []).map((w) => w.wk));
// 최근 4주 안에 발주가 있는 직납 채널만 (예: 케이마켓은 wk23 이후 없음)
const liveOrders = (x) => (x.orders || []).filter((o) => o.weeks.length && o.weeks[o.weeks.length - 1].wk >= curWk(x) - 3);
// 기준일이 그 주 토요일 전이면 진행 중인 주
const wkDone = (x, wk, base) => { const s = wkStart(x, wk); return s ? addDays(s, 6) <= base : true; };

// 로스 숫자 (유형별 시트 기준, 없으면 비교시트)
export function lossNow(s) {
  const x = s.x || {}, m = +s.date.slice(5, 7);
  const L = x.lossType?.months || {};
  if (L[m]) return { m, cur: L[m], prev: L[m - 1] || null, src: 'type', incl: x.lossDirect?.months?.[m] || null };
  const lm = s.monthly?.loss || {};
  if (lm[m] !== undefined) { const b = s.monthly.lossKg?.[m] || s.monthly.kg?.[m]; return { m, cur: { total: lm[m], qty: b || 0, rate: b ? lm[m] / b : 0 }, prev: null, src: 'cmp' }; }
  return null;
}

// 점포 계산: 이번 달은 진행 중이라 월말 예상(MTD ÷ 진도율)으로 비교
function storeRows(s) {
  const M = s.x?.storeM; if (!M) return null;
  const m = +s.date.slice(5, 7), prog = s.kpi?.progress || 0;
  const ms = Object.keys(M.months).map(Number).filter((x) => x <= m).sort((a, b) => a - b);
  const show = ms.slice(-4);
  const rows = M.cols.map((c, i) => {
    const v = (mm) => M.months[mm]?.[i] ?? 0;
    const cur = v(m), proj = prog ? cur / prog : cur, prev = v(m - 1);
    const avg3 = [m - 1, m - 2, m - 3].map(v).reduce((a, b) => a + b, 0) / 3;
    return { ...c, vals: show.map(v), cur, proj, prev, avg3, dPrev: chg(proj, prev), dAvg: chg(proj, avg3) };
  });
  return { show, rows, m };
}

export function extraFindings(s, add) {
  const x = s.x || {};
  // 점포 월
  const S = storeRows(s);
  if (S) {
    const st = S.rows.filter((r) => r.k === 'store');
    const stop = st.filter((r) => r.prev >= 300000 && r.cur === 0);
    if (stop.length) add('bad', `이번 달 매출 없는 점포: ${stop.map((r) => `${r.g} ${r.n}(전월 ${mil(r.prev)})`).join(', ')}`);
    const fresh = st.filter((r) => r.cur > 0 && r.prev === 0 && r.avg3 === 0);
    if (fresh.length) add('good', `새로 매출 생긴 점포: ${fresh.map((r) => `${r.g} ${r.n} ${mil(r.cur)}`).join(', ')}`);
    const drop = st.filter((r) => r.cur > 0 && r.prev >= 1000000 && r.dPrev !== null && r.dPrev <= -0.3).sort((a, b) => a.dPrev - b.dPrev).slice(0, 4);
    if (drop.length) add('bad', `월말 예상 기준 전월보다 30% 넘게 줄 점포: ${drop.map((r) => `${r.g} ${r.n} ${chgTxt(r.dPrev)}`).join(', ')}`);
    const top = [...st].sort((a, b) => b.cur - a.cur).slice(0, 3);
    if (top.length && top[0].cur) add('info', `이번 달 매출 상위 점포: ${top.map((r) => `${r.g} ${r.n} ${mil(r.cur)}`).join(' · ')}`);
  }
  // 점포 주간: 다 끝난 주 기준 2주 연속 감소
  const W = x.storeW;
  if (W?.weeks?.length >= 3) {
    const done = W.weeks.filter((w) => wkDone(x, w.wk, s.date));
    if (done.length >= 3) {
      const [a, b, c] = done.slice(-3);
      const dec = W.cols.map((col, i) => ({ col, a: a.v[i], b: b.v[i], c: c.v[i] })).filter((r) => r.col.k === 'store' && r.a >= 300000 && r.b < r.a && r.c < r.b);
      if (dec.length) add('bad', `2주 연속 매출 감소 점포 (${wkLabel(x, a.wk)}→${wkLabel(x, c.wk)}): ${dec.slice(0, 5).map((r) => `${r.col.g} ${r.col.n} ${mil(r.a)}→${mil(r.c)}`).join(', ')}`);
    }
  }
  // 미납
  for (const o of liveOrders(x)) {
    const done = o.weeks.filter((w) => wkDone(x, w.wk, s.date));
    const last = done[done.length - 1];
    if (!last || !last.order) continue;
    const rate = last.ship / last.order;
    const low = done.slice(-4).filter((w) => w.order && w.ship / w.order < 0.9).length;
    add(rate < 0.9 ? 'bad' : 'good', `${o.name} 직납 ${wkLabel(x, last.wk)} 주 발주 ${f0(last.order)}kg 중 출고 ${f0(last.ship)}kg (이행률 ${pc1(rate)}${rate < 1 ? `, 미납 ${f0(last.order - last.ship)}kg` : ''})${low >= 2 ? ` · 최근 4주 중 ${low}주 90% 미만` : ''}${o.skus[0]?.short ? ` · 미납 큰 SKU ${o.skus[0].n} ${f0(o.skus[0].short)}kg` : ''}`);
  }
  // 로스
  const L = lossNow(s);
  if (L && L.src === 'type') {
    const c = L.cur, p = L.prev;
    add(p && c.rate > p.rate ? 'bad' : 'good', `로스 ${L.m}월 ${f0(c.total)}kg · 물량 대비 ${pc1(c.rate)}${p ? ` (전월 ${pc1(p.rate)})` : ''} — 직납 ${f0(c.direct?.loss)} · 창고 ${f0(c.wh?.loss)} · 특정 폐기 ${f0(c.spec?.loss)}${c.oneplus?.loss ? ` · 1+1 ${f0(c.oneplus.loss)}kg 별도` : ''}`);
    if (c.spec && p?.spec && c.spec.rate > p.spec.rate + 0.01) add('bad', `특정 폐기율 ${pc1(c.spec.rate)} — 전월 ${pc1(p.spec.rate)}보다 높음`);
    const items = x.lossType.items || [];
    const tot = items.map((n, i) => ({ n, v: num(c.direct?.items?.[i]) + num(c.wh?.items?.[i]) + num(c.spec?.items?.[i]) + num(c.oneplus?.items?.[i]) })).sort((a, b) => b.v - a.v);
    const all = tot.reduce((a, b) => a + b.v, 0);
    if (all && tot[0].v / all >= 0.35) add('info', `로스가 가장 큰 품목: ${tot[0].n} ${f0(tot[0].v)}kg (${Math.round((tot[0].v / all) * 100)}%)`);
    const lw = (x.lossType.weeks || []).filter((w) => wkDone(x, w.wk, s.date));
    let streak = 0; for (let i = lw.length - 1; i >= 0 && lw[i].rate > 0.07; i--) streak++;
    if (streak >= 3) add('bad', `주간 로스율 7% 초과 ${streak}주 연속 (최근 ${pc1(lw[lw.length - 1].rate)})`);
    const dw = (x.lossDirect?.weeks || []).filter((w) => wkDone(x, w.wk, s.date));
    const tg = x.lossDirect?.target || 0.03, over = dw.slice(-4).filter((w) => w.rate > tg).length;
    if (dw.length) add(over ? 'bad' : 'good', `직납 로스율 최근 주 ${pc1(dw[dw.length - 1].rate)} · 목표 ${pc1(tg)}${over ? ` · 최근 4주 중 ${over}주 초과` : ' 이내'}`);
  }
  // 증정
  const G = x.gifts;
  if (G?.length) {
    const g = G[G.length - 1], prev = G[G.length - 2];
    const krw = num(g.lotte?.krw) + num(g.emart?.krw), kg = num(g.lotte?.kg) + num(g.emart?.kg);
    const wsale = W?.weeks?.find((w) => w.wk === g.wk)?.v?.[W.cols.findIndex((c) => c.k === 'top')] || 0;
    const pk = prev ? num(prev.lotte?.krw) + num(prev.emart?.krw) : 0;
    add('info', `증정·시식 ${wkLabel(x, g.wk)} 주 ${mil(krw)}백만원 · ${f1(kg)}kg${wsale ? ` (주간 매출 대비 ${pc1(krw / wsale)})` : ''}${pk ? ` · 전주 ${mil(pk)}` : ''}`);
  }
  // 행사
  const E = (x.events || []).filter((e) => e.to >= addDays(s.date, -7) && e.from <= addDays(s.date, 21));
  if (E.length) add('info', `행사 (기준일 전후): ${E.slice(0, 4).map((e) => `${md(e.from)}${e.to !== e.from ? `~${md(e.to)}` : ''} ${e.product} ${e.note}`).join(' / ')}`);
}

export function extraReport(s) {
  const x = s.x || {}, lines = [];
  const S = storeRows(s);
  if (S) {
    const st = S.rows.filter((r) => r.k === 'store' && r.cur > 0).sort((a, b) => b.cur - a.cur);
    const dn = S.rows.filter((r) => r.k === 'store' && r.prev >= 1000000 && r.dPrev !== null && r.dPrev <= -0.3);
    lines.push(`6) 점포: 상위 ${st.slice(0, 3).map((r) => `${r.n}(${r.g}) ${mil(r.cur)}`).join(', ')}${dn.length ? ` / 부진(월말 예상 −30% 이상) ${dn.slice(0, 3).map((r) => `${r.n} ${chgTxt(r.dPrev)}`).join(', ')}` : ''}`);
  }
  const ol = liveOrders(x).map((o) => { const d = o.weeks.filter((w) => wkDone(x, w.wk, s.date)); const l = d[d.length - 1]; return l ? `${o.name} ${pc1(l.ship / l.order)}(미납 ${f0(l.order - l.ship)}kg)` : ''; }).filter(Boolean);
  if (ol.length) lines.push(`7) 직납 이행률(지난주): ${ol.join(', ')}`);
  const L = lossNow(s);
  if (L?.src === 'type') lines.push(`8) 로스: ${L.m}월 ${f0(L.cur.total)}kg (${pc1(L.cur.rate)}${L.prev ? `, 전월 ${pc1(L.prev.rate)}` : ''}) — 직납 ${f0(L.cur.direct?.loss)} / 창고 ${f0(L.cur.wh?.loss)} / 특정 폐기 ${f0(L.cur.spec?.loss)}, 1+1 ${f0(L.cur.oneplus?.loss)}kg 별도`);
  const G = x.gifts;
  if (G?.length) { const g = G[G.length - 1]; lines.push(`9) 증정·시식(${wkLabel(x, g.wk)} 주): ${mil(num(g.lotte?.krw) + num(g.emart?.krw))}백만원, ${f1(num(g.lotte?.kg) + num(g.emart?.kg))}kg`); }
  return lines;
}

/* ================= 그리기 ================= */
const tdChg = (r) => `<td class="${r === null ? '' : r < -0.05 ? 'st-dn' : r > 0.05 ? 'st-up' : ''}">${chgTxt(r)}</td>`;
function heat(v, max) { if (!max || !v) return ''; const a = Math.round((v / max) * 26); return ` style="background:color-mix(in srgb, var(--primary) ${a}%, transparent)"`; }

function storeMonthHtml(s) {
  const S = storeRows(s);
  if (!S) return '';
  const head = `<tr><th>채널 · 점포</th>${S.show.map((mm) => `<th>${mm}월${mm === S.m ? '<br><small>(진행 중)</small>' : ''}</th>`).join('')}<th>${S.m}월 월말 예상</th><th>전월 대비</th><th>3개월 평균 대비</th></tr>`;
  const body = S.rows.filter((r) => r.k !== 'top' || /^전체$/.test(r.n)).filter((r) => r.k !== 'store' || r.vals.some((v) => v)).map((r) => {
    const mx = Math.max(...r.vals);
    const cls = r.k === 'store' ? '' : 'tot';
    const flag = r.k === 'store' && r.prev >= 300000 && r.cur === 0 ? ' st-out' : '';
    return `<tr class="${cls}${flag}"><th>${r.k === 'store' ? `<span class="muted sm">${esc(r.g)}</span> ` : ''}${esc(r.n)}</th>${r.vals.map((v) => `<td${r.k === 'store' ? heat(v, mx) : ''}>${v ? mil(v) : '-'}</td>`).join('')}<td>${r.proj ? mil(r.proj) : '-'}</td>${tdChg(r.dPrev)}${tdChg(r.dAvg)}</tr>`;
  }).join('');
  return `<div class="card"><h3>🏬 점포별 월 추이 <small class="muted">1-1 점포별 월매출 · 백만원 · 진할수록 그 점포의 큰 달 · 이번 달은 월말 예상(누계 ÷ 영업일 진도)으로 비교</small></h3><div class="tbl-wrap"><table class="kt st-store">${head}${body}</table></div></div>`;
}

function storeWeekHtml(s) {
  const x = s.x || {}, W = x.storeW;
  if (!W?.weeks?.length) return '';
  const ws = W.weeks.slice(-6);
  const done = ws.filter((w) => wkDone(x, w.wk, s.date));
  const a = done[done.length - 1], b = done[done.length - 2];
  const head = `<tr><th>채널 · 점포</th>${ws.map((w) => `<th>${wkLabel(x, w.wk)}${wkDone(x, w.wk, s.date) ? '' : '<br><small>(진행 중)</small>'}</th>`).join('')}<th>지난주 대비</th></tr>`;
  const body = W.cols.map((c, i) => ({ c, i })).filter(({ c, i }) => (c.k !== 'top' || /^전체$/.test(c.n)) && (c.k !== 'store' || ws.some((w) => w.v[i]))).map(({ c, i }) => {
    const vals = ws.map((w) => w.v[i]);
    const r = a && b ? chg(a.v[i], b.v[i]) : null;
    const dec = c.k === 'store' && done.length >= 3 && (() => { const [p, q, t] = done.slice(-3).map((w) => w.v[i]); return p >= 300000 && q < p && t < q; })();
    return `<tr class="${c.k === 'store' ? '' : 'tot'}${dec ? ' st-low' : ''}"><th>${c.k === 'store' ? `<span class="muted sm">${esc(c.g)}</span> ` : ''}${esc(c.n)}${dec ? ' <small>⚠ 2주 연속 감소</small>' : ''}</th>${vals.map((v) => `<td>${v ? mil(v) : '-'}</td>`).join('')}${tdChg(r)}</tr>`;
  }).join('');
  return `<div class="card"><h3>📅 점포별 주간 매출 <small class="muted">1-2 점포별 주간매출 · 백만원 · 주 = 일~토</small></h3><div class="tbl-wrap"><table class="kt st-store">${head}${body}</table></div></div>`;
}

function ordersHtml(s) {
  const x = s.x || {}, O = liveOrders(s.x || {});
  if (!O?.length) return '';
  const blocks = O.map((o) => {
    const ws = o.weeks.slice(-8);
    const done = ws.filter((w) => wkDone(x, w.wk, s.date));
    const last = done[done.length - 1];
    const card = last ? `<div class="kpi-cards st-mini"><div class="card kpi-card st-card ${last.ship / last.order < 0.9 ? 'warn' : ''}"><div class="kc-t">지난주 이행률</div><div class="kc-big">${pc1(last.ship / last.order)}</div><div class="kc-sub muted sm">${wkLabel(x, last.wk)} 주</div></div><div class="card kpi-card st-card"><div class="kc-t">지난주 미납</div><div class="kc-big">${f0(Math.max(0, last.order - last.ship))}<small>kg</small></div><div class="kc-sub muted sm">발주 ${f0(last.order)} · 출고 ${f0(last.ship)}</div></div>${o.skus[0]?.short ? `<div class="card kpi-card st-card"><div class="kc-t">미납 큰 SKU (최근 4주)</div><div class="kc-big" style="font-size:16px">${esc(o.skus[0].n)}</div><div class="kc-sub muted sm">${f0(o.skus[0].short)}kg${o.skus[1]?.short ? ` · 다음 ${esc(o.skus[1].n)} ${f0(o.skus[1].short)}` : ''}</div></div>` : ''}</div>` : '';
    const row = (lab, fn, cls = () => '') => `<tr><th>${lab}</th>${ws.map((w) => `<td class="${cls(w)}">${fn(w)}</td>`).join('')}</tr>`;
    return `<h4 class="st-h4">${esc(o.name)} 직납</h4>${card}<div class="tbl-wrap"><table class="kt"><tr><th>kg</th>${ws.map((w) => `<th>${wkLabel(x, w.wk)}${wkDone(x, w.wk, s.date) ? '' : '<br><small>(진행 중)</small>'}</th>`).join('')}</tr>
      ${row('발주', (w) => f0(w.order))}${row('출고', (w) => f0(w.ship))}${row('미납', (w) => f0(Math.max(0, w.order - w.ship)), (w) => (w.order - w.ship > w.order * 0.1 ? 'st-dn' : ''))}${row('이행률', (w) => pc1(w.ship / w.order), (w) => (w.ship / w.order < 0.9 ? 'st-dn' : 'st-up'))}</table></div>`;
  }).join('');
  return `<div class="card"><h3>📦 직납 발주 대비 출고 (미납) <small class="muted">2-1 (직납)발주-입고량 차이 · 최근 8주 · 이행률 = 출고 ÷ 발주, 90% 미만 빨강</small></h3>${blocks}</div>`;
}

const LT = [['direct', '직납', '#378ADD'], ['wh', '창고', '#888780'], ['spec', '특정 폐기', '#D85A30'], ['oneplus', '1+1', '#EF9F27']];
function lossHtml(s) {
  const x = s.x || {}, T = x.lossType;
  const L = lossNow(s);
  if (!T || !L || L.src !== 'type') return '';
  const c = L.cur, p = L.prev || {};
  const pcard = (k, lab) => { const v = c[k] || {}, pv = p[k] || {}; const hasRate = k !== 'wh'; return `<div class="card kpi-card st-card ${hasRate && pv.rate !== undefined && v.rate > pv.rate + 0.005 ? 'warn' : ''}"><div class="kc-t">${lab}</div><div class="kc-big">${f0(v.loss)}<small>kg</small></div><div class="kc-sub muted sm">${hasRate ? `${pc1(v.rate)}${pv.rate !== undefined ? ` · 전월 ${pc1(pv.rate)}` : ''}` : `전월 ${f0(pv.loss)}kg`}</div></div>`; };
  const ms = Object.keys(T.months).map(Number).sort((a, b) => a - b);
  const maxM = Math.max(...ms.map((mm) => LT.reduce((t, [k]) => t + num(T.months[mm][k]?.loss), 0)), 1);
  const stack = `<div class="st-stack">${ms.map((mm) => `<div class="col" title="${mm}월 ${LT.map(([k, n]) => `${n} ${f0(T.months[mm][k]?.loss)}`).join(' · ')}">${LT.map(([k, , col]) => `<i style="height:${(num(T.months[mm][k]?.loss) / maxM) * 100}%;background:${col}"></i>`).join('')}</div>`).join('')}</div><div class="st-stack-x">${ms.map((mm) => `<span${mm === L.m ? ' class="on"' : ''}>${mm}월</span>`).join('')}</div>
    <div class="st-legend">${LT.map(([, n, col]) => `<span><b style="background:${col}"></b>${n}</span>`).join('')}<span class="muted">높이 = 로스 kg</span></div>`;
  const last5 = ms.slice(-5);
  const mt = `<table class="kt"><tr><th>유형 (kg)</th>${last5.map((mm) => `<th>${mm}월</th>`).join('')}</tr>
    ${LT.slice(0, 3).map(([k, n]) => `<tr><th>${n}</th>${last5.map((mm) => `<td>${f0(T.months[mm][k]?.loss)}</td>`).join('')}</tr>`).join('')}
    <tr class="tot"><th>합계 · 물량 대비</th>${last5.map((mm) => `<td>${f0(T.months[mm].total)} · ${pc1(T.months[mm].rate)}</td>`).join('')}</tr>
    <tr><th>특정 1+1 (별도)</th>${last5.map((mm) => `<td>${T.months[mm].oneplus?.loss ? f0(T.months[mm].oneplus.loss) : '-'}</td>`).join('')}</tr></table>`;
  // 주간
  const lw = (T.weeks || []).slice(-8), dw = x.lossDirect?.weeks || [], tg = x.lossDirect?.target || 0.03;
  const dOf = (wk) => dw.find((w) => w.wk === wk);
  const wt = lw.length ? `<table class="kt"><tr><th>주</th>${lw.map((w) => `<th>${wkLabel(x, w.wk)}${wkDone(x, w.wk, s.date) ? '' : '<br><small>(진행 중)</small>'}</th>`).join('')}</tr>
    <tr><th>전체 로스 kg</th>${lw.map((w) => `<td>${f0(w.total)}</td>`).join('')}</tr>
    <tr><th>전체 로스율</th>${lw.map((w) => `<td class="${w.rate > 0.07 ? 'st-dn' : w.rate < 0.04 ? 'st-up' : ''}">${pc1(w.rate)}</td>`).join('')}</tr>
    <tr><th>직납 로스율 <small class="muted">(목표 ${pc1(tg)})</small></th>${lw.map((w) => { const d = dOf(w.wk); return `<td class="${d && d.rate > tg ? 'st-dn' : 'st-up'}">${d ? pc1(d.rate) : '-'}</td>`; }).join('')}</tr></table>` : '';
  // 품목
  const items = T.items || [];
  const it = items.map((n, i) => { const v = (k) => num(c[k]?.items?.[i]); return { n, d: v('direct'), w: v('wh'), sp: v('spec'), op: v('oneplus'), t: v('direct') + v('wh') + v('spec') + v('oneplus') }; }).filter((r) => r.t).sort((a, b) => b.t - a.t);
  const itAll = it.reduce((a, b) => a + b.t, 0);
  const itt = it.length ? `<table class="kt"><tr><th>품목 (kg)</th><th>직납</th><th>창고</th><th>특정 폐기</th><th>1+1</th><th>합계</th><th>비중</th></tr>${it.map((r) => `<tr><th>${esc(r.n)}</th><td>${f1(r.d)}</td><td>${f1(r.w)}</td><td>${f1(r.sp)}</td><td>${f1(r.op)}</td><td><b>${f0(r.t)}</b></td><td>${Math.round((r.t / itAll) * 100)}%</td></tr>`).join('')}</table>` : '';
  return `<div class="card"><h3>🗑 로스 <small class="muted">2-2 특정 로스 유형별 분류 + 2-2 직납로스 · ${L.m}월 누계</small></h3>
    <div class="kpi-cards st-mini"><div class="card kpi-card st-card ${p.rate !== undefined && c.rate > p.rate ? 'warn' : ''}"><div class="kc-t">전체 (직납+창고+특정 폐기)</div><div class="kc-big">${f0(c.total)}<small>kg</small></div><div class="kc-sub muted sm">${pc1(c.rate)}${p.rate !== undefined ? ` · 전월 ${pc1(p.rate)}` : ''}</div></div>${pcard('direct', '직납')}${pcard('wh', '창고')}${pcard('spec', '특정 폐기')}${pcard('oneplus', '특정 1+1 (별도)')}</div>
    ${stack}<div class="tbl-wrap">${mt}</div>
    <h4 class="st-h4">주간 로스</h4><div class="tbl-wrap">${wt}</div>
    <h4 class="st-h4">품목별 로스 · ${L.m}월</h4><div class="tbl-wrap">${itt}</div>
    <p class="muted sm">합계 기준이 시트마다 달라요: 유형별 시트 합계 ${f0(c.total)}kg (1+1 제외)${L.incl ? ` · 직납로스 시트 월 합계 ${f0(L.incl.total)}kg (1+1 포함, ${pc1(L.incl.rate)})` : ''}</p></div>`;
}

function giftsHtml(s) {
  const x = s.x || {}, G = x.gifts;
  if (!G?.length) return '';
  const W = x.storeW, ti = W ? W.cols.findIndex((c) => c.k === 'top') : -1;
  const sale = (wk) => (W && ti >= 0 ? W.weeks.find((w) => w.wk === wk)?.v?.[ti] || 0 : 0);
  const g = G[G.length - 1];
  const krw = (z) => num(z.lotte?.krw) + num(z.emart?.krw), kg = (z) => num(z.lotte?.kg) + num(z.emart?.kg);
  const stores = [...(g.lotte?.stores || []).map((z) => ({ ...z, g: '롯데' })), ...(g.emart?.stores || []).map((z) => ({ ...z, g: '이마트' }))].filter((z) => z.krw).sort((a, b) => b.krw - a.krw);
  return `<div class="card"><h3>🎁 증정·시식 <small class="muted">매장별 증정품 매입 · 원 · 주차별</small></h3>
    <div class="kpi-cards st-mini"><div class="card kpi-card st-card"><div class="kc-t">${wkLabel(x, g.wk)} 주 증정</div><div class="kc-big">${mil(krw(g))}<small>백만원</small></div><div class="kc-sub muted sm">롯데 ${mil(g.lotte?.krw)} · 이마트 ${mil(g.emart?.krw)}</div></div><div class="card kpi-card st-card"><div class="kc-t">증정 물량</div><div class="kc-big">${f1(kg(g))}<small>kg</small></div><div class="kc-sub muted sm">${sale(g.wk) ? `주간 매출 대비 ${pc1(krw(g) / sale(g.wk))}` : ''}</div></div></div>
    <div class="tbl-wrap"><table class="kt"><tr><th>주</th>${G.map((z) => `<th>${wkLabel(x, z.wk)}</th>`).join('')}</tr>
      <tr><th>증정 금액(원)</th>${G.map((z) => `<td>${f0(krw(z))}</td>`).join('')}</tr><tr><th>물량(kg)</th>${G.map((z) => `<td>${f1(kg(z))}</td>`).join('')}</tr>
      <tr><th>주간 매출 대비</th>${G.map((z) => `<td>${sale(z.wk) ? pc1(krw(z) / sale(z.wk)) : '-'}</td>`).join('')}</tr></table></div>
    <div class="st-two"><div><h4 class="st-h4">SKU별 (${wkLabel(x, g.wk)} 주)</h4><table class="kt"><tr><th>SKU</th><th>원</th><th>kg</th></tr>${g.skus.map((z) => `<tr><th>${esc(z.n)}</th><td>${f0(z.krw)}</td><td>${f1(z.kg)}</td></tr>`).join('')}</table></div>
      <div><h4 class="st-h4">점포별</h4><table class="kt"><tr><th>점포</th><th>원</th><th>kg</th></tr>${stores.slice(0, 6).map((z) => `<tr><th><span class="muted sm">${z.g}</span> ${esc(z.n)}</th><td>${f0(z.krw)}</td><td>${f1(z.kg)}</td></tr>`).join('')}</table></div></div></div>`;
}

function eventsHtml(s) {
  const E = (s.x?.events || []).filter((e) => e.to >= addDays(s.date, -31) && e.from <= addDays(s.date, 31));
  if (!E.length) return s.x?.events ? '<div class="card"><h3>🏷 행사</h3><p class="muted sm">기준일 전후 한 달 안의 행사 메모가 없습니다 (행사 시트 달력·Event 목록 기준).</p></div>' : '';
  const st = (e) => (e.to < s.date ? '<span class="muted">지남</span>' : e.from > s.date ? '<span class="tag">예정</span>' : '<b class="st-up">진행 중</b>');
  return `<div class="card"><h3>🏷 행사 <small class="muted">행사 시트 달력 메모·Event 목록 · 기준일 전후 한 달 · 일자별 매출 그래프에 ▲로 표시</small></h3><div class="tbl-wrap"><table class="kt"><tr><th>날짜</th><th>품목</th><th>내용</th><th>상태</th></tr>${E.map((e) => `<tr><th>${md(e.from)}${e.to !== e.from ? `~${md(e.to)}` : ''}</th><td style="text-align:left">${esc(e.product)}</td><td style="text-align:left">${esc(e.note)}</td><td>${st(e)}</td></tr>`).join('')}</table></div></div>`;
}

export function extraHtml(s) {
  if (!s.x) return '<div class="card"><p class="muted sm">📌 점포·미납·로스·증정·행사 구역은 이 파일을 한 번 다시 불러오면 나타납니다 (1.39.0 이전에 불러온 자료).</p></div>';
  return storeMonthHtml(s) + storeWeekHtml(s) + ordersHtml(s) + lossHtml(s) + giftsHtml(s) + eventsHtml(s);
}
// 일자별 그래프에 행사 날 표시용
export function eventDays(s) {
  const set = new Set();
  for (const e of s.x?.events || []) for (let d = e.from; d <= e.to; d = addDays(d, 1)) { if (d.slice(0, 7) === s.date.slice(0, 7)) set.add(+d.slice(8)); if (d > e.to || set.size > 40) break; }
  return set;
}
