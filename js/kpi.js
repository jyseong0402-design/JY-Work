// 📈 KPI 현황: KPI 엑셀의 Summary 시트를 읽어 요약 카드·그래프·특이사항·전월 비교를 보여줌
// 파일은 이 기기에서만 읽고, 뽑아낸 숫자(월별 실적·예산)만 store 'kpi' 항목으로 저장 → 탭·PC 동기화
// 항목 { id, title, year, month(파일의 결산월), file, loadedAt, rows:[{label, top, grp, a:[12], b:[12]}], notes:{ '10': '메모' } }
import { store } from './store.js';
import { loadXLSX } from './excel.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ui = { id: null, m: null };
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
function barChart({ a, b, m, line, lineLabel, sel, title }) {
  const W = 640, H = 230, L = 40, R = line ? 40 : 12, T = 14, Bm = 26;
  const vals = [...a.slice(0, 12), ...(b || []).slice(0, 12)];
  const hi = niceMax(Math.max(0, ...vals)), lo = Math.min(0, ...vals) < 0 ? -niceMax(-Math.min(...vals)) : 0;
  const y = (v) => T + (hi - v) / (hi - lo) * (H - T - Bm);
  const bw = (W - L - R) / 12;
  let s = '';
  for (let k = 0; k <= 4; k++) { const v = lo + (hi - lo) * k / 4; s += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="kg-grid"/><text x="${L - 6}" y="${y(v) + 4}" class="kg-ax" text-anchor="end">${Math.round(v * 10) / 10}</text>`; }
  if (lo < 0) s += `<line x1="${L}" x2="${W - R}" y1="${y(0)}" y2="${y(0)}" class="kg-zero"/>`;
  for (let k = 0; k < 12; k++) {
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

/* ---------- 화면 ---------- */
function list() { return [...(store.get().kpi || [])].sort((a, b) => b.year - a.year || String(a.title).localeCompare(String(b.title))); }

function render() {
  const box = $('#kpiBody');
  if (!box) return;
  const all = list();
  const sel = $('#kpiSel');
  if (!all.length) {
    sel.hidden = true; $('#kpiMonths').innerHTML = '';
    box.innerHTML = `<div class="card kpi-empty"><div style="font-size:40px">📈</div><b>KPI 엑셀 파일을 불러와 주세요</b>
      <p class="muted">위의 「📂 KPI 파일 불러오기」를 누르거나 파일을 이 화면에 끌어다 놓으면, Summary 시트의 월별 실적·예산을 읽어 요약·그래프·특이사항·전월 비교를 보여줍니다.<br>파일은 이 기기 안에서만 읽고, 숫자만 저장되어 탭·PC에서 같이 볼 수 있습니다.</p></div>`;
    return;
  }
  const k = all.find((x) => x.id === ui.id) || all[0];
  if (ui.id !== k.id) { ui.id = k.id; ui.m = k.month; }
  if (!ui.m || ui.m > k.month) ui.m = k.month;
  sel.hidden = all.length < 2;
  sel.innerHTML = all.map((x) => `<option value="${x.id}" ${x.id === k.id ? 'selected' : ''}>${x.year}년 · ${esc(x.title)}</option>`).join('');
  const m = ui.m, i = m - 1, p = m - 2;
  $('#kpiMonths').innerHTML = Array.from({ length: k.month }, (_, j) => `<button type="button" data-m="${j + 1}" class="${j + 1 === m ? 'on' : ''}">${j + 1}월</button>`).join('');
  const M = model(k);
  if (!M.sales) { box.innerHTML = '<div class="card">⚠️ 이 파일에서 매출액 항목을 찾지 못했습니다.</div>'; return; }

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
    if (e.target.id !== 'kpiDel') return;
    const k = store.find('kpi', ui.id);
    if (k && confirm(`${k.year}년 「${k.title}」 KPI 자료를 지울까요?\n(엑셀 파일은 그대로이고, 사이트에 저장된 숫자와 메모만 지워집니다)`)) { store.remove('kpi', k.id); ui.id = null; render(); }
  });
  // 메모 입력 중에는 다시 그리지 않음 (다른 기기 동기화로 커서가 튀지 않게)
  store.subscribe(() => { if (document.activeElement?.id !== 'kpiNote') render(); });
  render();
}
