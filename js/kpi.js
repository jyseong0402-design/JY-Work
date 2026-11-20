// 📈 KPI 현황: KPI 엑셀의 Summary 시트를 읽어 요약 카드·그래프·특이사항·전월 비교를 보여줌
// 파일은 이 기기에서만 읽고, 뽑아낸 숫자(월별 실적·예산)만 store 'kpi' 항목으로 저장 → 탭·PC 동기화
// 항목 { id, title, year, month(파일의 결산월), file, loadedAt, rows:[{label, top, grp, a:[12], b:[12]}], notes:{ '10': '메모' } }
import { store } from './store.js';
import { loadXLSX } from './excel.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ui = { id: null, m: null, mode: (() => { try { return localStorage.getItem('jy-kpimode') || 'month'; } catch { return 'month'; } })() };
let toastFn = () => {};

/* ---------- 엑셀 읽기 ---------- */
const num = (v) => (typeof v === 'number' && isFinite(v) ? Math.round(v * 10000) / 10000 : 0);
const hash = (s) => { let h = 0; for (const c of String(s)) h = (h * 31 + c.charCodeAt(0)) | 0; return (h >>> 0).toString(36); };

export function parseKpi(XLSX, wb, file = '') {
  const sn = wb.SheetNames.find((n) => /^summary/i.test(n.trim()));
  if (!sn) throw new Error("이 파일에서 'Summary' 시트를 찾지 못했습니다. (KPI 파일인지 확인해주세요)");
  const ws = wb.Sheets[sn];
  // 시트 범위가 B1 등에서 시작해도 열 번호가 엑셀과 같도록 A1부터 읽음
  if (ws['!ref']) { const r = XLSX.utils.decode_range(ws['!ref']); r.s = { r: 0, c: 0 }; ws['!ref'] = XLSX.utils.encode_range(r); }
  const g = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  // 연도 줄(2026A / 2026B)과 그 아래 월 줄을 찾음
  let hr = -1;
  for (let r = 0; r < Math.min(g.length, 15); r++) if ((g[r] || []).filter((v) => /^\d{4}A$/.test(String(v ?? '').trim())).length >= 6) { hr = r; break; }
  if (hr < 0) throw new Error('Summary 시트에서 월별 실적 열(예: 2026A)을 찾지 못했습니다.');
  const A = [], B = [];
  let year = 0;
  (g[hr] || []).forEach((v, c) => {
    const m = /^(\d{4})([AB])$/.exec(String(v ?? '').trim());
    const mo = Number((g[hr + 1] || [])[c]);
    if (!m || !(mo >= 1 && mo <= 12)) return;
    const arr = m[2] === 'A' ? A : B;
    if (arr[mo - 1] === undefined) arr[mo - 1] = c; // 처음 나온 열만 (뒤쪽 '당월' 열 제외)
    if (m[2] === 'A' && !year) year = +m[1];
  });
  const month = Number((g[2] || [])[1]) || A.reduce((mx, c, i) => (c !== undefined ? i + 1 : mx), 0);
  const title = String((g[1] || [])[1] || '').trim() || 'KPI 실적';
  const rows = [];
  let grp = -1;
  for (let r = hr + 2; r < g.length; r++) {
    const row = g[r] || [];
    const top = String(row[3] ?? '').trim(), sub = String(row[4] ?? '').trim();
    const label = top || sub;
    if (!label || /^[-\d.,%\s]+$/.test(label)) continue;
    if (top) grp++;
    rows.push({ label, top: !!top, grp, a: A.map((c) => num(row[c])), b: B.map((c) => num(row[c])) });
  }
  if (!rows.length) throw new Error('Summary 시트에서 항목(매출액 등)을 찾지 못했습니다.');
  return { id: `k${year}_${hash(title.replace(/\s+/g, ''))}`, title, year, month: Math.min(12, Math.max(1, month)), file, rows };
}

async function readFile(f) {
  const XLSX = await loadXLSX();
  const buf = await f.arrayBuffer();
  const names = XLSX.read(buf, { bookSheets: true }).SheetNames;
  const sn = names.find((n) => /^summary/i.test(n.trim()));
  const wb = XLSX.read(buf, sn ? { sheets: [sn] } : {});
  return parseKpi(XLSX, wb, f.name);
}

async function loadFiles(files) {
  const f = [...files].find((x) => /\.xls[xmb]?$/i.test(x.name));
  if (!f) return toastFn('엑셀 파일(.xlsx)을 넣어주세요.');
  const msg = $('#kpiMsg');
  msg.textContent = `「${f.name}」 읽는 중…`;
  try {
    const k = await readFile(f);
    const prev = store.find('kpi', k.id);
    store.upsert('kpi', { ...k, loadedAt: new Date().toISOString(), notes: prev?.notes || {} });
    ui.id = k.id; ui.m = k.month;
    msg.textContent = '';
    toastFn(`${k.year}년 ${k.month}월 KPI를 불러왔습니다.`);
    render();
  } catch (e) {
    msg.textContent = '⚠️ ' + (e.message || e);
  }
}

/* ---------- 숫자 꺼내기 ---------- */
const sum = (arr, n) => arr.slice(0, n).reduce((s, v) => s + (v || 0), 0);
function model(k) {
  const tops = k.rows.filter((r) => r.top);
  const find = (re, nth = 0) => tops.filter((r) => re.test(r.label.replace(/\s/g, '')))[nth];
  const kids = (r) => (r ? k.rows.filter((x) => !x.top && x.grp === r.grp) : []);
  const sales = find(/^매출액/), chan = find(/^매출액/, 1), qty = find(/매출물량/), price = find(/매출단가/);
  return {
    sales, qty, price, chan,
    items: kids(sales), chans: kids(chan), qtys: kids(qty), prices: kids(price),
    cogs: find(/^매출원가/), gp: find(/^GP\(/i), gpp: find(/^GP%/i), sell: find(/^판매비/), cm: find(/^CM\(/i), cmp: find(/^CM%/i),
    admin: find(/^관리비/), ebit: find(/^EBIT\(/i), ebitp: find(/^EBIT%/i), adj: find(/^조정사항/), ebitda: find(/조정EBITDA/i),
  };
}

/* ---------- 표시 형식 ---------- */
const f1 = (v) => (Math.abs(v) < 0.05 ? '-' : v.toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }));
const f0 = (v) => (Math.abs(v) < 0.5 ? '-' : Math.round(v).toLocaleString('ko-KR'));
const pc = (v) => `${Math.round(v * 100)}%`;
const sg = (v, f) => (Math.abs(v) < 1e-9 ? '±0' : (v > 0 ? '+' : '') + f(v));
const rate = (cur, base) => (Math.abs(base) < 1e-9 ? null : (cur - base) / Math.abs(base));
// kind: 'amt'(백만원, 클수록 좋음) 'cost'(비용, 작을수록 좋음) 'pct'(비율) 'won'(원/kg) 'ton'
const fmtOf = (kind) => (kind === 'pct' ? pc : kind === 'won' ? f0 : f1);
function delta(cur, base, kind) {
  const d = cur - base;
  if (Math.abs(d) < (kind === 'pct' ? 0.0005 : kind === 'won' ? 0.5 : 0.05)) return '<span class="kd">±0</span>';
  const good = kind === 'cost' ? d < 0 : d > 0;
  const txt = kind === 'pct' ? `${d > 0 ? '+' : ''}${(d * 100).toFixed(1)}%p` : sg(d, fmtOf(kind));
  const r = kind === 'pct' || base <= 0 ? null : rate(cur, base);
  return `<span class="kd ${good ? 'up' : 'dn'}">${d > 0 ? '▲' : '▼'} ${txt}${r !== null && Math.abs(r) < 10 ? ` (${r > 0 ? '+' : ''}${Math.round(r * 100)}%)` : ''}</span>`;
}

/* ---------- 특이사항 자동 찾기 ---------- */
function findings(M, m) {
  const out = [];
  const i = m - 1, p = m - 2;
  const add = (lv, txt) => out.push({ lv, txt });
  const S = M.sales;
  if (!S) return out;
  const cur = S.a[i], prev = p >= 0 ? S.a[p] : 0, bud = S.b[i] || 0;
  // 파일 점검: 품목별 합계와 거래처별 합계가 다르면 KPI 파일 쪽 누락 가능성
  if (M.chan && M.items.length && M.chans.length) {
    const d = M.sales.a[i] - M.chan.a[i];
    if (Math.abs(d) >= 0.1) add('bad', `⚠️ 파일 점검: 품목별 매출 합계 ${f1(M.sales.a[i])}와 거래처별 합계 ${f1(M.chan.a[i])}가 ${f1(Math.abs(d))}백만원 다릅니다 (Summary 거래처 행 확인 필요)`);
  }
  if (p >= 0) {
    const r = rate(cur, prev);
    if (r !== null && Math.abs(r) >= 0.1) add(r > 0 ? 'good' : 'bad', `매출 전월 대비 ${r > 0 ? '증가' : '감소'}: ${f1(prev)} → ${f1(cur)}백만원 (${r > 0 ? '+' : ''}${Math.round(r * 100)}%)`);
    const lists = [['품목', M.items], ['거래처', M.chans]];
    for (const [nm, list] of lists) for (const x of list) {
      const c = x.a[i], b = x.a[p];
      if (b < 0.05 && c >= 0.05) add('info', `${nm} 신규 매출: ${x.label} ${f1(c)}백만원`);
      else if (b >= 0.05 && c < 0.05) add('bad', `${nm} 매출 없음: ${x.label} (전월 ${f1(b)}백만원)`);
      else {
        const rr = rate(c, b);
        if (rr !== null && Math.abs(c - b) >= 1 && Math.abs(rr) >= 0.2) add(rr > 0 ? 'good' : 'bad', `${x.label} 매출 ${rr > 0 ? '증가' : '감소'}: ${f1(b)} → ${f1(c)}백만원 (${rr > 0 ? '+' : ''}${Math.round(rr * 100)}%)`);
      }
    }
    for (const x of M.prices) {
      const c = x.a[i], b = x.a[p], rr = rate(c, b);
      if (c > 0 && b > 0 && rr !== null && Math.abs(rr) >= 0.07) add(rr > 0 ? 'good' : 'bad', `${x.label} 단가 ${rr > 0 ? '상승' : '하락'}: ${f0(b)} → ${f0(c)}원/kg (${rr > 0 ? '+' : ''}${Math.round(rr * 100)}%)`);
    }
    if (M.gpp) { const d = M.gpp.a[i] - M.gpp.a[p]; if (Math.abs(d) >= 0.05) add(d > 0 ? 'good' : 'bad', `GP% ${d > 0 ? '개선' : '악화'}: ${pc(M.gpp.a[p])} → ${pc(M.gpp.a[i])} (${d > 0 ? '+' : ''}${(d * 100).toFixed(1)}%p)`); }
    for (const x of [M.sell, M.admin]) if (x) {
      const c = x.a[i], b = x.a[p], rr = rate(c, b);
      if (rr !== null && Math.abs(c - b) >= 1 && Math.abs(rr) >= 0.15) add(rr > 0 ? 'bad' : 'good', `${x.label.replace(/\(.*\)/, '')} ${rr > 0 ? '증가' : '감소'}: ${f1(b)} → ${f1(c)}백만원 (${rr > 0 ? '+' : ''}${Math.round(rr * 100)}%)`);
    }
    for (const x of [M.cm, M.ebit]) if (x) {
      const c = x.a[i], b = x.a[p], nm = x.label.replace(/\(.*\)/, '');
      if (b < 0 && c >= 0) add('good', `${nm} 흑자 전환: ${f1(b)} → ${f1(c)}백만원`);
      else if (b >= 0 && c < 0) add('bad', `${nm} 적자 전환: ${f1(b)} → ${f1(c)}백만원`);
      else if (x === M.ebit && Math.abs(c - b) >= 3) add(c > b ? 'good' : 'bad', `EBIT ${c > b ? '개선' : '악화'}: ${f1(b)} → ${f1(c)}백만원`);
    }
  }
  if (bud > 0) {
    const ach = cur / bud;
    if (ach < 0.9 || ach > 1.1) add(ach >= 1 ? 'good' : 'bad', `매출 예산 달성률 ${Math.round(ach * 100)}% (예산 ${f1(bud)} / 실적 ${f1(cur)}백만원)`);
  }
  if (M.ebit && M.ebit.b[i]) { const d = M.ebit.a[i] - M.ebit.b[i]; if (Math.abs(d) >= 3) add(d > 0 ? 'good' : 'bad', `EBIT 예산 대비 ${d > 0 ? '+' : ''}${f1(d)}백만원 (예산 ${f1(M.ebit.b[i])} / 실적 ${f1(M.ebit.a[i])})`); }
  if (m >= 3 && cur > 0 && cur >= Math.max(...S.a.slice(0, i))) add('good', `올해 최고 월 매출 (${f1(cur)}백만원)`);
  const order = { bad: 0, good: 1, info: 2 };
  return out.sort((a, b) => order[a.lv] - order[b.lv]);
}

/* ---------- 그래프 (SVG) ---------- */
function niceMax(v) { if (v <= 0) return 1; const p = 10 ** Math.floor(Math.log10(v)); return Math.ceil(v / p / (v / p > 5 ? 2 : 1)) * p * (v / p > 5 ? 2 : 1); }
// 월별 막대(실적) + 예산 표시 + (선택) 비율 선
function barChart({ a, b, m, line, lineLabel, sel, title, cnt = 12 }) {
  const W = 640, H = 230, L = 40, R = line ? 40 : 12, T = 14, Bm = 26;
  const vals = [...a.slice(0, cnt), ...(b || []).slice(0, cnt)];
  // 눈금: 0을 지나는 깔끔한 간격(1·2·5×10ⁿ)
  const vmax = Math.max(0, ...vals), vmin = Math.min(0, ...vals);
  const raw = (vmax - vmin) / 4 || 1, p10 = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((f) => f * p10).find((v) => v >= raw);
  const hi = Math.ceil(vmax / step) * step || (vmin < 0 ? 0 : step), lo = Math.floor(vmin / step) * step;
  const y = (v) => T + (hi - v) / (hi - lo) * (H - T - Bm);
  const bw = (W - L - R) / cnt;
  let s = '';
  for (let v = lo; v <= hi + step / 2; v += step) { s += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="kg-grid"/><text x="${L - 6}" y="${y(v) + 4}" class="kg-ax" text-anchor="end">${Math.round(v * 100) / 100}</text>`; }
  if (lo < 0) s += `<line x1="${L}" x2="${W - R}" y1="${y(0)}" y2="${y(0)}" class="kg-zero"/>`;
  for (let k = 0; k < cnt; k++) {
    const x = L + k * bw, v = a[k] || 0, on = k < m;
    if (on) s += `<rect x="${x + bw * .18}" y="${Math.min(y(v), y(0))}" width="${bw * .64}" height="${Math.max(1, Math.abs(y(v) - y(0)))}" rx="3" class="kg-bar ${k === sel - 1 ? 'sel' : ''} ${v < 0 ? 'neg' : ''}"><title>${k + 1}월 ${f1(v)}</title></rect>`;
    if (b && b[k]) s += `<line x1="${x + bw * .1}" x2="${x + bw * .9}" y1="${y(b[k])}" y2="${y(b[k])}" class="kg-bud"><title>${k + 1}월 예산 ${f1(b[k])}</title></line>`;
    s += `<text x="${x + bw / 2}" y="${H - 8}" class="kg-ax ${k === sel - 1 ? 'sel' : ''}" text-anchor="middle">${k + 1}월</text>`;
  }
  if (line) {
    const lv = line.slice(0, m), lh = Math.max(0.05, ...lv.map(Math.abs));
    const ly = (v) => T + (1 - (v + lh) / (2 * lh)) * (H - T - Bm);
    s += `<polyline class="kg-line" points="${lv.map((v, k) => `${L + k * bw + bw / 2},${ly(v)}`).join(' ')}"/>`;
    s += lv.map((v, k) => `<circle cx="${L + k * bw + bw / 2}" cy="${ly(v)}" r="3.5" class="kg-dot"><title>${k + 1}월 ${lineLabel} ${pc(v)}</title></circle>`).join('');
    s += `<text x="${W - R + 6}" y="${ly(lv[lv.length - 1] || 0) + 4}" class="kg-ax kg-lt">${pc(lv[lv.length - 1] || 0)}</text>`;
  }
  const legend = `<span class="kl"><i class="kl-bar"></i>실적</span>${b ? '<span class="kl"><i class="kl-bud"></i>예산</span>' : ''}${line ? `<span class="kl"><i class="kl-line"></i>${lineLabel}</span>` : ''}`;
  return `<div class="card kg"><div class="kg-h"><b>${title}</b><span class="kg-leg">${legend}</span></div><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(title)}">${s}</svg></div>`;
}
// 항목별 당월 vs 전월 가로 막대
function compareBars(title, list, m) {
  const i = m - 1, p = m - 2;
  const rows = list.filter((x) => Math.abs(x.a[i]) >= 0.05 || (p >= 0 && Math.abs(x.a[p]) >= 0.05));
  if (!rows.length) return '';
  const hi = Math.max(0.1, ...rows.flatMap((x) => [x.a[i], p >= 0 ? x.a[p] : 0]));
  const total = rows.reduce((s, x) => s + x.a[i], 0) || 1;
  const body = rows.sort((x, y) => y.a[i] - x.a[i]).map((x) => {
    const c = x.a[i], b = p >= 0 ? x.a[p] : 0;
    return `<div class="kc-row"><span class="kc-n">${esc(x.label)}</span><span class="kc-bars"><i class="kc-cur" style="width:${Math.max(0, c / hi * 100)}%"></i>${p >= 0 ? `<i class="kc-prev" style="width:${Math.max(0, b / hi * 100)}%"></i>` : ''}</span><span class="kc-v"><b>${f1(c)}</b> <small class="muted">${Math.round(c / total * 100)}%</small>${p >= 0 ? `<br>${delta(c, b, 'amt')}` : ''}</span></div>`;
  }).join('');
  return `<div class="card kg"><div class="kg-h"><b>${title}</b><span class="kg-leg"><span class="kl"><i class="kl-bar"></i>${m}월</span>${p >= 0 ? `<span class="kl"><i class="kl-prev"></i>${m - 1}월</span>` : ''}</span></div>${body}</div>`;
}

/* ---------- 📑 누계 보고 (1~N월) ---------- */
const sumR = (arr, s, e) => arr.slice(s - 1, e).reduce((t, v) => t + (v || 0), 0); // s~e월 합
const MONTH_COLORS = ['#2b5ce6', '#1f9d62', '#d98a00', '#8e44ad', '#d64545', '#16a2b8', '#7f8c8d', '#c0392b', '#2c3e50'];
function quarters(n) {
  const out = [];
  for (let q = 0; q * 3 < n; q++) { const s = q * 3 + 1, e = Math.min(q * 3 + 3, n); out.push({ s, e, label: `${q + 1}Q`, part: e - s < 2 }); }
  return out;
}
// 월별 누적 막대(품목·거래처 구성)
function stackChart(title, list, n, total) {
  const rows = list.filter((x) => x.a.slice(0, n).some((v) => Math.abs(v) >= 0.05));
  if (!rows.length) return '';
  const W = 640, H = 240, L = 40, R = 12, T = 18, Bm = 26;
  const tot = Array.from({ length: n }, (_, k) => rows.reduce((s, x) => s + Math.max(0, x.a[k]), 0));
  const hi = niceMax(Math.max(...tot, ...(total ? total.slice(0, n) : [0])));
  const y = (v) => T + (hi - v) / hi * (H - T - Bm), bw = (W - L - R) / n;
  let s = '';
  for (let k = 0; k <= 4; k++) { const v = hi * k / 4; s += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="kg-grid"/><text x="${L - 6}" y="${y(v) + 4}" class="kg-ax" text-anchor="end">${Math.round(v)}</text>`; }
  for (let k = 0; k < n; k++) {
    let acc = 0;
    const x = L + k * bw + bw * .15;
    rows.forEach((r, j) => {
      const v = Math.max(0, r.a[k]);
      if (v < 0.01) return;
      s += `<rect x="${x}" y="${y(acc + v)}" width="${bw * .7}" height="${Math.max(0.5, y(acc) - y(acc + v))}" fill="${MONTH_COLORS[j % MONTH_COLORS.length]}"><title>${k + 1}월 ${esc(r.label)} ${f1(r.a[k])}</title></rect>`;
      acc += v;
    });
    s += `<text x="${x + bw * .35}" y="${y(acc) - 4}" class="kg-ax" text-anchor="middle">${f1(acc)}</text><text x="${x + bw * .35}" y="${H - 8}" class="kg-ax" text-anchor="middle">${k + 1}월</text>`;
  }
  const legend = rows.map((r, j) => `<span class="kl"><i style="background:${MONTH_COLORS[j % MONTH_COLORS.length]}"></i>${esc(r.label)}</span>`).join('');
  return `<div class="card kg"><div class="kg-h"><b>${title}</b></div><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(title)}">${s}</svg><div class="kg-leg wrap">${legend}</div></div>`;
}
// 비율(GP%·CM%·EBIT%) 월별 선
function ratioChart(title, series, n) {
  const W = 640, H = 220, L = 44, R = 48, T = 14, Bm = 26;
  const vals = series.flatMap((x) => x.v.slice(0, n));
  const hi = Math.max(0.1, Math.ceil(Math.max(...vals) * 10) / 10), lo = Math.min(0, Math.floor(Math.max(-2, Math.min(...vals)) * 10) / 10);
  const y = (v) => T + (hi - Math.max(lo, Math.min(hi, v))) / (hi - lo) * (H - T - Bm), bw = (W - L - R) / n;
  let s = '';
  const labels = [];
  for (let k = 0; k <= 4; k++) { const v = lo + (hi - lo) * k / 4; s += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="kg-grid"/><text x="${L - 6}" y="${y(v) + 4}" class="kg-ax" text-anchor="end">${pc(v)}</text>`; }
  if (lo < 0) s += `<line x1="${L}" x2="${W - R}" y1="${y(0)}" y2="${y(0)}" class="kg-zero"/>`;
  for (let k = 0; k < n; k++) s += `<text x="${L + k * bw + bw / 2}" y="${H - 8}" class="kg-ax" text-anchor="middle">${k + 1}월</text>`;
  series.forEach((x, j) => {
    const pts = x.v.slice(0, n).map((v, k) => [L + k * bw + bw / 2, y(v)]);
    s += `<polyline fill="none" stroke="${x.c}" stroke-width="2.2" points="${pts.map((p) => p.join(',')).join(' ')}"/>`;
    s += pts.map(([px, py], k) => `<circle cx="${px}" cy="${py}" r="3.2" fill="var(--panel)" stroke="${x.c}" stroke-width="2"><title>${k + 1}월 ${x.name} ${pc(x.v[k])}</title></circle>`).join('');
    labels.push({ y: pts[n - 1][1], c: x.c, t: pc(x.v[n - 1]) });
  });
  // 끝 값 표시: 오른쪽에 세로로 겹치지 않게
  labels.sort((p, q) => p.y - q.y).forEach((lb, j, arr) => { if (j && lb.y - arr[j - 1].y < 13) lb.y = arr[j - 1].y + 13; });
  s += labels.map((lb) => `<text x="${W - R + 8}" y="${lb.y + 4}" class="kg-ax" style="fill:${lb.c};font-weight:700">${lb.t}</text>`).join('');
  const legend = series.map((x) => `<span class="kl"><i style="height:0;border-top:2px solid ${x.c}"></i>${x.name}</span>`).join('');
  return `<div class="card kg"><div class="kg-h"><b>${title}</b><span class="kg-leg">${legend}</span></div><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(title)}">${s}</svg><p class="muted sm" style="margin:4px 0 0">-200% 아래 값은 그래프 아래끝에 맞춰 표시</p></div>`;
}

function periodFindings(M, n) {
  const out = [];
  const add = (lv, txt) => out.push({ lv, txt });
  const S = M.sales.a, Q = quarters(n);
  const lastQ = Q[Q.length - 1], prevQ = Q[Q.length - 2];
  const avg = (arr, q) => sumR(arr, q.s, q.e) / (q.e - q.s + 1);
  // 매출 흐름
  if (n >= 2) add(S[n - 1] >= S[0] ? 'good' : 'bad', `월 매출 1월 ${f1(S[0])} → ${n}월 ${f1(S[n - 1])}백만원 (${S[0] > 0 ? `${(S[n - 1] / S[0]).toFixed(1)}배` : '증가'})`);
  if (prevQ) {
    const a = avg(S, prevQ), b = avg(S, lastQ), r = rate(b, a);
    if (r !== null) add(r >= 0 ? 'good' : 'bad', `${lastQ.label}${lastQ.part ? `(${lastQ.s}~${lastQ.e}월)` : ''} 월평균 매출 ${f1(b)}백만원, ${prevQ.label} 대비 ${r > 0 ? '+' : ''}${Math.round(r * 100)}%`);
  }
  const mx = S.slice(0, n).indexOf(Math.max(...S.slice(0, n))), mn = S.slice(0, n).indexOf(Math.min(...S.slice(0, n)));
  add('info', `최고 매출 ${mx + 1}월 ${f1(S[mx])}백만원 · 최저 ${mn + 1}월 ${f1(S[mn])}백만원`);
  // 예산
  const yb = sumR(M.sales.b, 1, n), ya = sumR(S, 1, n);
  if (yb > 0) add(ya >= yb ? 'good' : 'bad', `누계 매출 예산 달성률 ${Math.round(ya / yb * 100)}% (예산 ${f1(yb)} / 실적 ${f1(ya)}백만원)`);
  if (M.ebit) { const ea = sumR(M.ebit.a, 1, n), eb = sumR(M.ebit.b, 1, n); if (eb) add(ea >= eb ? 'good' : 'bad', `누계 EBIT ${f1(ea)}백만원, 예산(${f1(eb)}) 대비 ${ea - eb > 0 ? '+' : ''}${f1(ea - eb)}`); }
  // 수익성
  if (M.gp && Q.length >= 2) {
    const g = (q) => sumR(M.gp.a, q.s, q.e) / (sumR(S, q.s, q.e) || 1);
    const d = g(lastQ) - g(Q[0]);
    add(d >= 0 ? 'good' : 'bad', `GP% ${Q[0].label} ${pc(g(Q[0]))} → ${lastQ.label} ${pc(g(lastQ))} (${d > 0 ? '+' : ''}${(d * 100).toFixed(1)}%p)`);
  }
  for (const [r, nm] of [[M.cm, 'CM'], [M.ebit, 'EBIT']]) if (r) {
    const a = r.a.slice(0, n);
    let k = n - 1; while (k >= 0 && a[k] >= 0) k--;
    if (k < n - 1 && k >= 0) add('good', `${nm} ${k + 2}월부터 흑자 유지 (${n}월 ${f1(a[n - 1])}백만원)`);
    else if (a[n - 1] < 0) { const w = a.indexOf(Math.min(...a)); if (w < n - 1) add('good', `${nm} 적자 폭 축소: 최대 ${w + 1}월 ${f1(a[w])} → ${n}월 ${f1(a[n - 1])}백만원`); }
  }
  if (M.admin && prevQ) { const a = avg(M.admin.a, Q[0]), b = avg(M.admin.a, lastQ), r = rate(b, a); if (r !== null && Math.abs(r) >= 0.1) add(r < 0 ? 'good' : 'bad', `관리비 월평균 ${Q[0].label} ${f1(a)} → ${lastQ.label} ${f1(b)}백만원 (${r > 0 ? '+' : ''}${Math.round(r * 100)}%)`); }
  if (M.sell && prevQ) { const a = sumR(M.sell.a, Q[0].s, Q[0].e) / (sumR(S, Q[0].s, Q[0].e) || 1), b = sumR(M.sell.a, lastQ.s, lastQ.e) / (sumR(S, lastQ.s, lastQ.e) || 1); if (Math.abs(b - a) >= 0.03) add(b < a ? 'good' : 'bad', `매출 대비 판매비 비율 ${Q[0].label} ${pc(a)} → ${lastQ.label} ${pc(b)}`); }
  // 품목 비중 변화
  if (Q.length >= 2) {
    const share = (x, q) => sumR(x.a, q.s, q.e) / (sumR(S, q.s, q.e) || 1);
    for (const x of M.items) { const a = share(x, Q[0]), b = share(x, lastQ); if (Math.abs(b - a) >= 0.04) add('info', `${x.label} 매출 비중 ${Q[0].label} ${pc(a)} → ${lastQ.label} ${pc(b)}`); }
    const ups = [], downs = [];
    for (const x of M.prices) {
      const qv = (q) => { const qq = M.qtys.find((y) => y.label === x.label), it = M.items.find((y) => y.label === x.label); const t = qq ? sumR(qq.a, q.s, q.e) : 0; return t && it ? sumR(it.a, q.s, q.e) / t * 1000 : 0; };
      const a = qv(Q[0]), b = qv(lastQ), r = rate(b, a);
      if (a && b && r !== null && Math.abs(r) >= 0.05) (r > 0 ? ups : downs).push(`${x.label} ${f0(a)}→${f0(b)}(${r > 0 ? '+' : ''}${Math.round(r * 100)}%)`);
    }
    if (ups.length) add('good', `평균 단가 상승 ${Q[0].label}→${lastQ.label}: ${ups.join(', ')} 원/kg`);
    if (downs.length) add('bad', `평균 단가 하락 ${Q[0].label}→${lastQ.label}: ${downs.join(', ')} 원/kg`);
  }
  // 거래처 시작·중단
  for (const x of M.chans) {
    const a = x.a.slice(0, n), on = a.map((v) => v >= 0.05);
    const first = on.indexOf(true), last = on.lastIndexOf(true);
    if (sumR(a, 1, n) < 3) continue; // 누계 3백만원 미만 소액 거래처는 생략
    if (first > 0) add('info', `${x.label} ${first + 1}월부터 거래 시작 (누계 ${f1(sumR(a, 1, n))}백만원)`);
    if (last >= 0 && last < n - 1) add('bad', `${x.label} ${last + 2}월부터 매출 없음 (마지막 ${last + 1}월, 누계 ${f1(sumR(a, 1, n))}백만원)`);
  }
  // 파일 점검
  if (M.chan && M.chans.length) {
    const bad = [];
    for (let k = 0; k < n; k++) if (Math.abs(S[k] - M.chan.a[k]) >= 0.1) bad.push(`${k + 1}월(${f1(S[k] - M.chan.a[k])})`);
    if (bad.length) add('bad', `⚠️ 파일 점검: 품목별 합계와 거래처별 합계가 다른 달 — ${bad.join(', ')}`);
  }
  const order = { bad: 0, good: 1, info: 2 };
  return out.sort((a, b) => order[a.lv] - order[b.lv]);
}

function renderYtd(k, M, n) {
  const ytd = (r, w = 'a') => sumR(r[w], 1, n);
  const div = (a, b) => (b ? a / b : 0);
  // 카드
  const card = (r, name, sub) => {
    if (!r) return '';
    const a = ytd(r), b = ytd(r, 'b');
    return `<div class="card kpi-card"><div class="kc-t">${name} <small>1~${n}월 누계</small></div><div class="kc-big ${a < 0 ? 'neg' : ''}">${f1(a)}<small>백만원</small></div>
      ${sub ? `<div class="kc-sub">${sub}</div>` : ''}
      <div class="kc-cmp"><div><span class="muted">월평균</span> ${f1(a / n)}</div>${b ? `<div><span class="muted">예산</span> ${f1(b)} ${delta(a, b, 'amt')}${b > 0 ? ` <small class="muted">달성 ${Math.round(a / b * 100)}%</small>` : ''}</div>` : ''}</div></div>`;
  };
  const S = ytd(M.sales);
  const cards = card(M.sales, '매출액', M.qty ? `물량 <b>${f1(ytd(M.qty))}톤</b> · 단가 <b>${f0(div(S, ytd(M.qty)) * 1000)}원/kg</b>` : '')
    + card(M.gp, 'GP', M.gp ? `GP% <b>${pc(div(ytd(M.gp), S))}</b>` : '') + card(M.cm, 'CM', M.cm ? `CM% <b>${pc(div(ytd(M.cm), S))}</b>` : '')
    + card(M.ebit, 'EBIT', M.ebit ? `EBIT% <b>${pc(div(ytd(M.ebit), S))}</b>` : '') + card(M.ebitda, '조정 EBITDA', M.adj && ytd(M.adj) ? `조정사항 ${f1(ytd(M.adj))}` : '');

  // 특이사항 + 보고 메모
  const fs = periodFindings(M, n);
  const ic = { bad: '<b class="kd dn">▼</b>', good: '<b class="kd up">▲</b>', info: '<b class="kd" style="color:var(--primary)">●</b>' };
  const key = `ytd${n}`;
  const notes = `<div class="card kpi-notes"><div class="kg-h"><b>📌 1~${n}월 주요 내용</b><small class="muted">기간 흐름·분기 비교·예산으로 자동 정리</small></div>
    <ul class="kf">${fs.map((f) => `<li class="${f.lv}">${ic[f.lv]} ${esc(f.txt)}</li>`).join('')}</ul>
    <label class="kn"><span class="muted sm">✏️ 보고 메모 (1~${n}월 보고에 덧붙일 내용 · 자동 저장)</span><textarea id="kpiNote" data-id="${k.id}" data-m="${key}" rows="4" placeholder="예: 3Q 이마트 입점으로 매출 확대, 관리비 절감(법인장 급여 종료) 효과">${esc((k.notes || {})[key] || '')}</textarea></label></div>`;

  // 월별 손익표 (PPT 형식)
  const mcols = Array.from({ length: n }, (_, j) => j);
  const ratioOf = (num) => (num ? (w, s, e) => div(sumR(num[w], s, e), sumR(M.sales[w], s, e)) : null);
  const priceOf = (it, q) => (w, s, e) => { const t = q ? sumR(q[w], s, e) : 0; return t ? sumR(it[w], s, e) / t * 1000 : 0; };
  const L = [];
  const push = (r, name, kind, cls = '', agg) => r && L.push({ r, name, kind, cls, agg });
  push(M.sales, '매출액', 'amt', 'top');
  M.items.forEach((x) => push(x, x.label, 'amt', 'sub'));
  if (M.chans.some((x) => ytd(x))) { L.push({ head: '거래처별 매출' }); M.chans.filter((x) => ytd(x) || ytd(x, 'b')).forEach((x) => push(x, x.label, 'amt', 'sub')); }
  push(M.qty, '매출물량(톤)', 'ton', 'top');
  push(M.price, '매출단가(원/kg)', 'won', 'top', M.qty && priceOf(M.sales, M.qty));
  M.prices.forEach((x) => push(x, x.label, 'won', 'sub', priceOf(M.items.find((y) => y.label === x.label) || x, M.qtys.find((y) => y.label === x.label))));
  push(M.cogs, '매출원가', 'cost', 'top');
  push(M.gp, 'GP', 'amt', 'top hl');
  push(M.gpp, 'GP%', 'pct', 'sub', ratioOf(M.gp));
  push(M.sell, '판매비', 'cost', 'top');
  push(M.cm, 'CM', 'amt', 'top hl');
  push(M.cmp, 'CM%', 'pct', 'sub', ratioOf(M.cm));
  push(M.admin, '관리비', 'cost', 'top');
  push(M.ebit, 'EBIT', 'amt', 'top hl');
  push(M.ebitp, 'EBIT%', 'pct', 'sub', ratioOf(M.ebit));
  if (M.adj && ytd(M.adj)) push(M.adj, '조정사항', 'amt', 'sub');
  push(M.ebitda, '조정 EBITDA', 'amt', 'top hl');
  const aggV = (x, w, s, e) => (x.agg ? x.agg(w, s, e) : x.kind === 'won' || x.kind === 'pct' ? 0 : sumR(x.r[w], s, e));
  const tbody = L.map((x) => {
    if (x.head) return `<tr class="kh"><th colspan="${n + 5}">${x.head}</th></tr>`;
    const F = fmtOf(x.kind), ya = aggV(x, 'a', 1, n), yb = aggV(x, 'b', 1, n);
    const ach = x.kind === 'pct' ? (yb ? `${((ya - yb) * 100).toFixed(1)}%p` : '') : yb > 0 && ya > -1e9 ? `${Math.round(ya / yb * 100)}%` : '';
    return `<tr class="${x.cls}"><th>${esc(x.name)}</th>${mcols.map((j) => `<td class="${x.r.a[j] < 0 ? 'neg' : ''}">${F(x.r.a[j])}</td>`).join('')}<td class="b ytd ${ya < 0 ? 'neg' : ''}">${F(ya)}</td><td>${yb ? F(yb) : '-'}</td><td>${yb ? delta(ya, yb, x.kind) : ''}</td><td class="muted">${ach}</td></tr>`;
  }).join('');
  const table = `<div class="card kt-wrap"><div class="kg-h"><b>📋 월별 손익 (1~${n}월)</b><small class="muted">단위: 백만원 (물량 톤, 단가 원/kg) · 누계 비율·단가는 합계로 다시 계산</small></div><div class="kt-scroll"><table class="kt kt-m">
    <thead><tr><th></th>${mcols.map((j) => `<th>${j + 1}월</th>`).join('')}<th>누계</th><th>누계 예산</th><th>예산 대비</th><th>달성률</th></tr></thead><tbody>${tbody}</tbody></table></div></div>`;

  // 분기 비교
  const Q = quarters(n);
  const qrows = [[M.sales, '매출액', 'amt'], [M.qty, '물량(톤)', 'ton'], [M.gp, 'GP', 'amt'], [M.gpp, 'GP%', 'pct', ratioOf(M.gp)], [M.sell, '판매비', 'cost'], [M.cm, 'CM', 'amt'], [M.cmp, 'CM%', 'pct', ratioOf(M.cm)], [M.admin, '관리비', 'cost'], [M.ebit, 'EBIT', 'amt'], [M.ebitda, '조정 EBITDA', 'amt']].filter((x) => x[0]);
  const qv = (x, q) => (x[3] ? x[3]('a', q.s, q.e) : sumR(x[0].a, q.s, q.e) / (x[2] === 'pct' ? 1 : q.e - q.s + 1));
  const qtable = Q.length >= 2 ? `<div class="card kt-wrap"><div class="kg-h"><b>📆 분기별 비교 (월평균)</b><small class="muted">분기 금액은 월평균, 비율은 분기 합계 기준${Q[Q.length - 1].part ? ` · ${Q[Q.length - 1].label}는 ${Q[Q.length - 1].s}~${Q[Q.length - 1].e}월만` : ''}</small></div><div class="kt-scroll"><table class="kt">
    <thead><tr><th></th>${Q.map((q) => `<th>${q.label}<br><small>${q.s}~${q.e}월</small></th>`).join('')}<th>직전 분기 대비</th><th>${Q[0].label} 대비</th></tr></thead>
    <tbody>${qrows.map((x) => { const F = fmtOf(x[2]); const v = Q.map((q) => qv(x, q)); const last = v[v.length - 1]; return `<tr class="${x[2] === 'pct' || x[2] === 'ton' ? 'sub' : ''}"><th>${x[1]}</th>${v.map((t) => `<td class="${t < 0 ? 'neg' : ''}">${F(t)}</td>`).join('')}<td>${delta(last, v[v.length - 2], x[2])}</td><td>${delta(last, v[0], x[2])}</td></tr>`; }).join('')}</tbody></table></div></div>` : '';

  // 그래프
  const charts = barChart({ a: M.sales.a, b: M.sales.b, m: n, cnt: n, sel: 0, line: M.gpp?.a, lineLabel: 'GP%', title: `월별 매출액과 GP% (1~${n}월)` })
    + stackChart('품목별 매출 구성 (백만원)', M.items, n, M.sales.a)
    + ratioChart('수익성 추이', [M.gpp && { name: 'GP%', v: M.gpp.a, c: '#1f9d62' }, M.cmp && { name: 'CM%', v: M.cmp.a, c: '#2b5ce6' }, M.ebitp && { name: 'EBIT%', v: M.ebitp.a, c: '#d64545' }].filter(Boolean), n)
    + (M.ebit ? barChart({ a: M.ebit.a, b: M.ebit.b, m: n, cnt: n, sel: 0, title: 'EBIT 월별 (백만원)' }) : '')
    + (M.chans.filter((x) => ytd(x)).length > 1 ? stackChart('거래처별 매출 구성 (백만원)', M.chans, n, M.chan?.a) : '');

  return `<div class="kpi-report-h"><h2>${k.year}년 ${n > 1 ? `1~${n}월` : "1월"} 실적 보고</h2><p class="muted sm">${esc(k.title)} · 📄 ${esc(k.file || '')}</p></div>
    <div class="kpi-cards">${cards}</div>${notes}${table}${qtable}<div class="kpi-charts">${charts}</div>`;
}

/* ---------- 화면 ---------- */
function list() { return [...(store.get().kpi || [])].sort((a, b) => b.year - a.year || String(a.title).localeCompare(String(b.title))); }

function render() {
  const box = $('#kpiBody');
  if (!box) return;
  const all = list();
  const sel = $('#kpiSel');
  if (!all.length) {
    sel.hidden = true; $('#kpiMonths').innerHTML = ''; $('#kpiMode').hidden = true;
    box.innerHTML = `<div class="card kpi-empty"><div style="font-size:40px">📈</div><b>KPI 엑셀 파일을 불러와 주세요</b>
      <p class="muted">위의 「📂 KPI 파일 불러오기」를 누르거나 파일을 이 화면에 끌어다 놓으면, Summary 시트의 월별 실적·예산을 읽어 요약·그래프·특이사항·전월 비교를 보여줍니다.<br>파일은 이 기기 안에서만 읽고, 숫자만 저장되어 탭·PC에서 같이 볼 수 있습니다.</p></div>`;
    return;
  }
  const k = all.find((x) => x.id === ui.id) || all[0];
  if (ui.id !== k.id) { ui.id = k.id; ui.m = k.month; }
  if (!ui.m || ui.m > k.month) ui.m = k.month;
  sel.hidden = all.length < 2;
  $('#kpiMode').hidden = false;
  sel.innerHTML = all.map((x) => `<option value="${x.id}" ${x.id === k.id ? 'selected' : ''}>${x.year}년 · ${esc(x.title)}</option>`).join('');
  const m = ui.m, i = m - 1, p = m - 2;
  $('#kpiMonths').innerHTML = Array.from({ length: k.month }, (_, j) => `<button type="button" data-m="${j + 1}" class="${j + 1 === m ? 'on' : ''}">${ui.mode === 'ytd' && j ? '1~' : ''}${j + 1}월</button>`).join('');
  const M = model(k);
  if (!M.sales) { box.innerHTML = '<div class="card">⚠️ 이 파일에서 매출액 항목을 찾지 못했습니다.</div>'; return; }
  document.querySelectorAll('#kpiMode [data-mode]').forEach((b) => b.classList.toggle('on', b.dataset.mode === ui.mode));
  if (ui.mode === 'ytd') {
    box.innerHTML = renderYtd(k, M, m) + '<p class="kpi-foot"><button type="button" class="btn sm" id="kpiPrint">🖨 인쇄·PDF 저장</button> <button type="button" class="btn sm danger" id="kpiDel">🗑 이 KPI 자료 지우기</button></p>';
    return;
  }

  // 요약 카드
  const card = (row, name, kind = 'amt', sub) => {
    if (!row) return '';
    const c = row.a[i], b = p >= 0 ? row.a[p] : null, bud = row.b[i];
    return `<div class="card kpi-card"><div class="kc-t">${name}</div><div class="kc-big ${c < 0 ? 'neg' : ''}">${f1(c)}<small>백만원</small></div>
      ${sub ? `<div class="kc-sub">${sub}</div>` : ''}
      <div class="kc-cmp">${b !== null ? `<div><span class="muted">전월</span> ${f1(b)} ${delta(c, b, kind)}</div>` : ''}${bud ? `<div><span class="muted">예산</span> ${f1(bud)} ${delta(c, bud, kind)}</div>` : ''}</div></div>`;
  };
  const ratio = (r) => (r ? `${r.label.replace(/\(.*\)/, '')} <b>${pc(r.a[i])}</b>${p >= 0 ? ` <small class="muted">(전월 ${pc(r.a[p])})</small>` : ''}` : '');
  const cards = card(M.sales, '매출액', 'amt', M.qty ? `물량 <b>${f1(M.qty.a[i])}톤</b> · 단가 <b>${f0(M.price?.a[i] || 0)}원/kg</b>` : '')
    + card(M.gp, 'GP (매출총이익)', 'amt', ratio(M.gpp)) + card(M.cm, 'CM (공헌이익)', 'amt', ratio(M.cmp))
    + card(M.ebit, 'EBIT (영업이익)', 'amt', ratio(M.ebitp)) + card(M.ebitda, '조정 EBITDA', 'amt', M.adj && M.adj.a[i] ? `조정사항 ${f1(M.adj.a[i])}` : '');

  // 특이사항
  const fs = findings(M, m);
  const ic = { bad: '<b class="kd dn">▼</b>', good: '<b class="kd up">▲</b>', info: '<b class="kd" style="color:var(--primary)">●</b>' };
  const note = (k.notes || {})[m] || '';
  const notes = `<div class="card kpi-notes"><div class="kg-h"><b>📌 ${m}월 특이사항</b><small class="muted">전월·예산과 비교해 자동으로 찾은 내용</small></div>
    ${fs.length ? `<ul class="kf">${fs.map((f) => `<li class="${f.lv}">${ic[f.lv]} ${esc(f.txt)}</li>`).join('')}</ul>` : '<p class="muted">눈에 띄는 변동이 없습니다.</p>'}
    <label class="kn"><span class="muted sm">✏️ 내 메모 (보고할 때 덧붙일 내용 · 자동 저장)</span><textarea id="kpiNote" data-id="${k.id}" data-m="${m}" rows="3" placeholder="예: 이마트 신규 입점 효과, 10월부터 aT 인건비 지원">${esc(note)}</textarea></label></div>`;

  // 그래프
  const charts = barChart({ a: M.sales.a, b: M.sales.b, m: k.month, sel: m, line: M.gpp?.a, lineLabel: 'GP%', title: '월별 매출액 (백만원)과 GP%' })
    + (M.ebit ? barChart({ a: M.ebit.a, b: M.ebit.b, m: k.month, sel: m, title: 'EBIT 월별 추이 (백만원)' }) : '')
    + compareBars('품목별 매출 (백만원)', M.items, m) + compareBars('거래처별 매출 (백만원)', M.chans, m);

  // 전월 비교표
  const ytd = (r) => sum(r.a, m), ytdB = (r) => sum(r.b, m);
  const ratioRow = (num, den) => (arr, n) => { const d = sum(den[arr], n); return d ? sum(num[arr], n) / d : 0; };
  const lines = [
    [M.sales, '매출액', 'amt'], [M.qty, '매출물량(톤)', 'ton'], [M.price, '매출단가(원/kg)', 'won'], [M.cogs, '매출원가', 'cost'], [M.gp, 'GP', 'amt'],
    [M.gpp, 'GP%', 'pct', M.gp && ratioRow(M.gp, M.sales)], [M.sell, '판매비', 'cost'], [M.cm, 'CM', 'amt'], [M.cmp, 'CM%', 'pct', M.cm && ratioRow(M.cm, M.sales)],
    [M.admin, '관리비', 'cost'], [M.ebit, 'EBIT', 'amt'], [M.ebitda, '조정 EBITDA', 'amt'],
  ].filter((x) => x[0]);
  const yv = (r, kind, rf, arr) => (kind === 'pct' ? (rf ? rf(arr, m) : 0) : kind === 'won' ? (sum(M.qty?.[arr] || [], m) ? sum(M.sales[arr], m) / sum(M.qty[arr], m) * 1000 : 0) : arr === 'a' ? ytd(r) : ytdB(r));
  const tr = ([r, name, kind, rf]) => {
    const F = fmtOf(kind), c = r.a[i], b = p >= 0 ? r.a[p] : null, bud = r.b[i], ya = yv(r, kind, rf, 'a'), yb = yv(r, kind, rf, 'b');
    return `<tr class="${kind === 'pct' || kind === 'ton' || kind === 'won' ? 'sub' : ''}"><th>${name}</th><td class="b">${F(c)}</td><td>${b !== null ? F(b) : ''}</td><td>${b !== null ? delta(c, b, kind) : ''}</td>
      <td>${bud ? F(bud) : '-'}</td><td>${bud ? delta(c, bud, kind) : ''}</td><td class="b">${F(ya)}</td><td>${yb ? F(yb) : '-'}</td><td>${yb ? delta(ya, yb, kind) : ''}</td></tr>`;
  };
  const table = `<div class="card kt-wrap"><div class="kg-h"><b>📊 ${m}월 전월·예산 비교</b><small class="muted">단위: 백만원 (물량 톤, 단가 원/kg)</small></div><div class="kt-scroll"><table class="kt">
    <thead><tr><th></th><th>${m}월</th><th>${p >= 0 ? `${m - 1}월` : '전월'}</th><th>전월 대비</th><th>예산</th><th>예산 대비</th><th>누계 1~${m}월</th><th>누계 예산</th><th>누계 예산 대비</th></tr></thead>
    <tbody>${lines.map(tr).join('')}</tbody></table></div></div>`;

  // 품목별 상세 (매출·물량·단가)
  const det = M.items.map((x) => {
    const q = M.qtys.find((y) => y.label === x.label), pr = M.prices.find((y) => y.label === x.label);
    const cell = (r, kind) => (r ? `<td class="b">${fmtOf(kind)(r.a[i])}</td><td>${p >= 0 ? delta(r.a[i], r.a[p], kind) : ''}</td>` : '<td></td><td></td>');
    return `<tr><th>${esc(x.label)}</th>${cell(x, 'amt')}${cell(q, 'ton')}${cell(pr, 'won')}</tr>`;
  }).join('');
  const detail = M.items.length ? `<div class="card kt-wrap"><div class="kg-h"><b>🍄 품목별 ${m}월 실적</b><small class="muted">전월 대비</small></div><div class="kt-scroll"><table class="kt">
    <thead><tr><th></th><th>매출(백만원)</th><th>전월 대비</th><th>물량(톤)</th><th>전월 대비</th><th>단가(원/kg)</th><th>전월 대비</th></tr></thead><tbody>${det}</tbody></table></div></div>` : '';

  const when = k.loadedAt ? new Date(k.loadedAt) : null;
  box.innerHTML = `<p class="muted sm kpi-src">📄 ${esc(k.file || '')}${when ? ` · ${when.getMonth() + 1}/${when.getDate()} 불러옴` : ''} · ${esc(k.title)}</p>
    <div class="kpi-cards">${cards}</div>${notes}<div class="kpi-charts">${charts}</div>${table}${detail}
    <p class="kpi-foot"><button type="button" class="btn sm danger" id="kpiDel">🗑 이 KPI 자료 지우기</button></p>`;
}

export function initKpi({ toast } = {}) {
  if (toast) toastFn = toast;
  const view = $('#view-kpi');
  if (!view) return;
  const inp = $('#kpiIn');
  $('#kpiLoad').addEventListener('click', () => inp.click());
  inp.addEventListener('change', () => { if (inp.files.length) loadFiles(inp.files); inp.value = ''; });
  view.addEventListener('dragover', (e) => { if ([...(e.dataTransfer?.types || [])].includes('Files')) { e.preventDefault(); view.classList.add('drop-on'); } });
  view.addEventListener('dragleave', (e) => { if (!view.contains(e.relatedTarget)) view.classList.remove('drop-on'); });
  view.addEventListener('drop', (e) => { if (!e.dataTransfer?.files?.length) return; e.preventDefault(); view.classList.remove('drop-on'); loadFiles(e.dataTransfer.files); });
  $('#kpiSel').addEventListener('change', (e) => { ui.id = e.target.value; ui.m = null; render(); });
  $('#kpiMonths').addEventListener('click', (e) => { const b = e.target.closest('[data-m]'); if (b) { ui.m = +b.dataset.m; render(); } });
  $('#kpiMode').addEventListener('click', (e) => {
    const b = e.target.closest('[data-mode]');
    if (!b) return;
    ui.mode = b.dataset.mode;
    try { localStorage.setItem('jy-kpimode', ui.mode); } catch { /* 저장 불가 */ }
    render();
  });
  let t;
  const saveNote = (el, id = el.dataset.id, m = el.dataset.m) => {
    const k = store.find('kpi', id);
    if (!k) return;
    const notes = { ...(k.notes || {}) };
    const v = el.value.trim();
    if ((notes[m] || '') === v) return;
    if (v) notes[m] = v; else delete notes[m];
    store.upsert('kpi', { id: k.id, notes });
  };
  view.addEventListener('input', (e) => { if (e.target.id === 'kpiNote') { clearTimeout(t); const el = e.target; t = setTimeout(() => saveNote(el), 800); } });
  view.addEventListener('focusout', (e) => { if (e.target.id === 'kpiNote') { clearTimeout(t); saveNote(e.target); } });
  view.addEventListener('click', (e) => {
    if (e.target.id === 'kpiPrint') { window.print(); return; }
    if (e.target.id !== 'kpiDel') return;
    const k = store.find('kpi', ui.id);
    if (k && confirm(`${k.year}년 「${k.title}」 KPI 자료를 지울까요?\n(엑셀 파일은 그대로이고, 사이트에 저장된 숫자와 메모만 지워집니다)`)) { store.remove('kpi', k.id); ui.id = null; render(); }
  });
  // 메모 입력 중에는 다시 그리지 않음 (다른 기기 동기화로 커서가 튀지 않게)
  store.subscribe(() => { if (document.activeElement?.id !== 'kpiNote') render(); });
  render();
}
