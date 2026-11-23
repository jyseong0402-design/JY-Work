// 📈 KPI 현황: KPI 엑셀의 Summary 시트를 읽어 요약 카드·그래프·특이사항·전월 비교를 보여줌
// 파일은 이 기기에서만 읽고, 뽑아낸 숫자(월별 실적·BGT)만 store 'kpi' 항목으로 저장 → 탭·PC 동기화
// 항목 { id, title, year, month(파일의 결산월), file, loadedAt, rows:[{label, top, grp, a:[12], b:[12]}], notes:{ '10': '메모' } }
import { store } from './store.js';
import { loadXLSX } from './excel.js';
import { aiChat, aiEnabled, aiJson } from './ai.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const lsGet = (k, d) => { try { return localStorage.getItem(k) ?? d; } catch { return d; } };
const ui = { id: null, m: null, mode: lsGet('jy-kpimode', 'month'), tgt: +lsGet('jy-kpitgt', 0) || 0 };
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
  const target = parseTarget(XLSX, wb);
  const ppt = parsePpt(XLSX, wb);
  const fm = Math.min(12, Math.max(1, month));
  const gl = parseGl(XLSX, wb, year, fm);
  const chanFix = gl ? fixChans(rows, gl, fm) : [];
  // 같은 KPI를 다시 불러오면 저장된 항목에 덮어써지므로(합치기), 이번 파일에 없는 값도 null/[]로 넣어 예전 값이 남지 않게 함
  const hub = parseHubChan(XLSX, wb);
  return { chanFix, id: `k${year}_${hash(title.replace(/\s+/g, ''))}`, title, year, month: fm, file, rows, target: target || null, ppt: ppt || null, gl: gl || null, chanQty: hub?.qty || null, chanPrice: hub?.price || null };
}

// 허브 시트('…결산기준'): 거래처별 매출물량(톤)·매출단가(원/kg) — D열 구역 제목(매출물량(톤)/매출단가) 아래 F열 거래처 행, 월 열은 위쪽 1~12 줄
const HUB_SHEET = (n) => /결산기준/.test(n);
function parseHubChan(XLSX, wb) {
  const sn = wb.SheetNames.find(HUB_SHEET);
  if (!sn || !wb.Sheets[sn]) return null;
  const g = fullRows(XLSX, wb.Sheets[sn]);
  const s = (r, c) => String((g[r] || [])[c] ?? '').trim();
  let mr = -1, mc = -1;
  for (let r = 0; r < 12 && mr < 0; r++) { const row = g[r] || []; for (let c = 0; c < row.length - 2; c++) if (row[c] === 1 && row[c + 1] === 2 && row[c + 2] === 3) { mr = r; mc = c; break; } }
  if (mr < 0) return null;
  // 거래처 행: D열=구역 이름(매출물량(톤)/매출단가…), E열 비고 F열=거래처 (품목 행은 E열)
  const out = { qty: {}, price: {} };
  for (let r = mr + 1; r < Math.min(g.length, 400); r++) {
    const d = s(r, 3), f = s(r, 5);
    if (!f || s(r, 4)) continue;
    const key = /^매출물량/.test(d) && !/비중/.test(d) ? 'qty' : /^매출단가/.test(d) ? 'price' : '';
    if (key && !out[key][f]) out[key][f] = Array.from({ length: 12 }, (_, j) => num((g[r] || [])[mc + j]));
  }
  return Object.keys(out.qty).length ? out : null;
}
const chanQtyOf = (k, label) => (k.chanQty || {})[label] || null;

// 거래처별 매출: Summary 거래처 칸이 비어 있는데(예: 10월 이마트 0) 거래처 합 < 매출액이면, 원장 511 거래처 금액으로 채움
//   원장 거래처명 → Summary 거래처 라벨 (THISO RETAIL = 베트남 이마트 운영사)
const CUST_MAP = [[/lotte|lot\b|trung tâm thương mại lot/i, /롯데/], [/thiso|e-?mart/i, /이마트/], [/co\.?op|saigon co/i, /coop|쿱/i], [/win ?mart|wincommerce|masan/i, /win|윈/i], [/k-?market/i, /케이마켓|k-?market/i], [/kamereo/i, /카메레오/]];
function fixChans(rows, gl, month) {
  const tops = rows.filter((r) => r.top && /^매출액/.test(r.label));
  if (tops.length < 2 || !gl.cust) return [];
  const total = tops[0], ctop = tops[1], chans = rows.filter((r) => !r.top && r.grp === ctop.grp);
  const fixes = [];
  for (let j = 0; j < month; j++) {
    const gap = (total.a[j] || 0) - chans.reduce((s, r) => s + (r.a[j] || 0), 0);
    if (gap < 0.5) continue;
    for (const r of chans) {
      if (Math.abs(r.a[j] || 0) >= 0.05) continue;
      const map = CUST_MAP.find(([, lab]) => lab.test(r.label));
      if (!map) continue;
      const v = Object.entries(gl.cust).filter(([o]) => map[0].test(o)).reduce((s, [, arr]) => s + (arr[j] || 0), 0) / VND_PER_KRW / 1e6;
      if (v < 0.05) continue;
      r.a[j] = v; ctop.a[j] = (ctop.a[j] || 0) + v;
      fixes.push({ label: r.label, m: j + 1, v: Math.round(v * 100) / 100 });
    }
  }
  return fixes;
}

const fullRows = (XLSX, ws) => {
  if (ws['!ref']) { const r = XLSX.utils.decode_range(ws['!ref']); r.s = { r: 0, c: 0 }; ws['!ref'] = XLSX.utils.encode_range(r); }
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
};

// 📝 경영회의 장표 표: 'PPT…' 시트에서 머리글 줄에 'Actual'·'Budget'·'Target'이 같이 있는 묶음
//   (9월 파일: PPT (2) 128행~, AJ=대항목·AL/AM=세부, AO~AW=1~9월, AX=Budget 당월, AY=%Ach, BC=Target 당월, BD=%Ach)
const PPT_SHEET = (n) => /^ppt/i.test(n.trim());
const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
function parsePpt(XLSX, wb) {
  for (const sn of wb.SheetNames.filter(PPT_SHEET)) {
    const ws = wb.Sheets[sn], hidden = (ws["!rows"] || []).map((x) => !!(x && x.hidden));
    const g = fullRows(XLSX, ws);
    const s = (r, c) => String((g[r] || [])[c] ?? '').trim();
    for (let r = 0; r < g.length - 2; r++) {
      const row = g[r] || [];
      const find = (re) => row.findIndex((v) => re.test(String(v ?? '').trim()));
      const cB = find(/^budget$|^bgt$/i), cT = find(/^target$/i);
      let cA = -1; for (let c = cB - 1; c >= 0; c--) if (/actual/i.test(String(row[c] ?? ''))) { cA = c; break; } // Budget 바로 왼쪽의 Actual (왼쪽 다른 표 제외)
      if (cA < 0 || cB < 0 || cT < 0) continue;
      const h = g[r + 1] || [];
      const months = [];
      for (let c = cA; c < cB; c++) {
        const raw = h[c];
        // 'Jan' 글자 또는 날짜 값(서식만 'mmm')
        const m = typeof raw === 'number' && raw > 20000 ? new Date(Math.round((raw - 25569) * 864e5)).getUTCMonth() + 1 : MON[s(r + 1, c).slice(0, 3).toLowerCase()];
        if (m) months.push({ c, m });
      }
      if (!months.length) continue;
      let lab0 = -1; for (let c = cA - 1; c >= 0; c--) if (/단위/.test(String(h[c] ?? ''))) { lab0 = c; break; } // 표 바로 왼쪽의 '(단위: 백만원)' 열
      const L = lab0 >= 0 ? lab0 : Math.max(0, cA - 5);
      const rows = [];
      let blank = 0, kind1 = 'amt';
      for (let rr = r + 2; rr < g.length && blank < 4; rr++) {
        let label = '', off = -1;
        for (let c = L; c < cA; c++) if (s(rr, c)) { label = s(rr, c); off = c - L; break; }
        if (!label) { blank++; continue; }
        blank = 0;
        if (/^dep$/i.test(label)) break;
        const lv = off === 0 ? 1 : 2;
        const pct = /%/.test(label);
        if (lv === 1 && !pct) kind1 = /원\/kg/.test(label) ? 'won' : /톤/.test(label) ? 'ton' : 'amt';
        const kind = pct ? 'pct' : lv === 1 ? kind1 : kind1;
        const v = (c) => num((g[rr] || [])[c]);
        rows.push({ label, lv, kind, a: months.map((x) => v(x.c)), bgt: v(cB), bAch: v(cB + 1), tgt: v(cT), tAch: v(cT + 1), ...(hidden[rr] ? { h: 1 } : {}) });
      }
      if (rows.length >= 5) return { sheet: sn, row: r + 1, months: months.map((x) => x.m), rows };
    }
  }
  return null;
}

// 📒 원장: 'GL(원장)' 시트 → 손익 계정(5·6·7·8xx) 월별 합계(VND) + 당월·전월 주요 거래(911 마감분개 제외)
const GL_SHEET = (n) => /원장|^gl\b/i.test(n.trim());
function parseGl(XLSX, wb, year, month) {
  const sn = wb.SheetNames.find(GL_SHEET);
  if (!sn) return null;
  const g = fullRows(XLSX, wb.Sheets[sn]);
  let hr = -1;
  for (let r = 0; r < 30 && hr < 0; r++) if ((g[r] || []).some((v) => /^account$/i.test(String(v ?? '').trim())) && (g[r] || []).some((v) => /debit/i.test(String(v ?? '')))) hr = r;
  if (hr < 0) return null;
  const H = g[hr].map((v) => String(v ?? '').trim().toLowerCase());
  const col = (re) => H.findIndex((v) => re.test(v));
  const cAcc = col(/^account$/), cCor = col(/corresponding/), cDr = col(/debit/), cCr = col(/credit/), cDesc = col(/general description/) >= 0 ? col(/general description/) : col(/description/);
  const cObj = col(/name of object/), cY = col(/^year$/), cM = col(/^month$/), cDate = col(/posting date/);
  const acc = {}, lines = [], cust = {}; // cust: 매출(511) 거래처별 월 합계 VND

  for (let r = hr + 1; r < g.length; r++) {
    const row = g[r] || [];
    const a = String(row[cAcc] ?? '').trim();
    if (!/^[5678]\d{2,}/.test(a) || /^911/.test(String(row[cCor] ?? ''))) continue;
    if (/profit and loss|lãi lỗ/i.test(String(row[cDesc] ?? ''))) continue; // 손익 마감 대체분개 제외
    let y = Number(row[cY]), m = Number(row[cM]);
    if (!(m >= 1) && cDate >= 0) { const d = row[cDate]; const dt = typeof d === 'number' ? new Date(Math.round((d - 25569) * 864e5)) : new Date(String(d)); y = dt.getFullYear(); m = dt.getMonth() + 1; }
    if (y !== year || !(m >= 1 && m <= 12)) continue;
    const dr = Number(row[cDr]) || 0, cr = Number(row[cCr]) || 0;
    const amt = /^[57]/.test(a) ? cr - dr : dr - cr; // 수익은 대변, 비용은 차변이 +
    if (!amt) continue;
    (acc[a] ||= Array(12).fill(0))[m - 1] += amt;
    if (/^511/.test(a)) { const o = String(row[cObj] ?? '').trim() || '(거래처 없음)'; (cust[o] ||= Array(12).fill(0))[m - 1] += amt; }
    if ((m === month || m === month - 1) && Math.abs(amt) >= 1e6) {
      const desc = String(row[cDesc] ?? '').split('/').map((x) => x.trim()).filter(Boolean);
      lines.push({ m, a, v: Math.round(amt), d: (desc[desc.length - 1] || '').slice(0, 90), o: String(row[cObj] ?? '').trim().slice(0, 40), c: String(row[cCor] ?? '').trim() });
    }
  }
  if (!Object.keys(acc).length) return null;
  lines.sort((p, q) => Math.abs(q.v) - Math.abs(p.v));
  return { sheet: sn, acc, cust, lines: lines.slice(0, 400) };
}

// 🎯 'Target … PL' 시트: 왼쪽은 월별 실적(3행 연도·4행 'N월'), 그 오른쪽 한 칸 띄고 Target 시나리오 열(월 기준 목표치)
//   B열=대항목(매출액·GP·판매비…), C열=중항목(판매단가·GP%…), D열=세부(품목·비용 계정·CM% 등)
const TARGET_SHEET = (n) => /^target.*pl\s*$/i.test(n.trim());
function parseTarget(XLSX, wb) {
  const sn = wb.SheetNames.find(TARGET_SHEET);
  const ws = sn && wb.Sheets[sn];
  if (!ws) return null;
  if (ws['!ref']) { const r = XLSX.utils.decode_range(ws['!ref']); r.s = { r: 0, c: 0 }; ws['!ref'] = XLSX.utils.encode_range(r); }
  const g = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  const s = (r, c) => String((g[r] || [])[c] ?? '').trim();
  // 월 열: 위 줄 4자리 연도 + 아래 줄 'N월'(또는 숫자)
  let hr = -1;
  for (let r = 0; r < 10 && hr < 0; r++) if ((g[r] || []).filter((v) => /^20\d\d$/.test(String(v ?? '').trim())).length >= 3) hr = r;
  if (hr < 0) return null;
  const A = [];
  let last = -1;
  (g[hr] || []).forEach((v, c) => {
    if (!/^20\d\d$/.test(String(v ?? '').trim())) return;
    const mo = parseInt(s(hr + 1, c), 10);
    if (mo >= 1 && mo <= 12 && A[mo - 1] === undefined) { A[mo - 1] = c; last = Math.max(last, c); }
  });
  if (last < 0) return null;
  // Target 열: 월 열 오른쪽에서 머리글(위·아래 줄)이 있는 첫 묶음
  const scen = [];
  let c = last + 1, gap = 0;
  while (c < last + 30) {
    const h = [s(hr, c), s(hr + 1, c)].filter(Boolean).join(' ');
    if (h && !/average|ytd|평균|누계/i.test(h)) { scen.push({ c, name: h }); gap = 0; } else if (scen.length && ++gap >= 1) break;
    c++;
  }
  if (!scen.length) return null;
  const rows = [];
  let p1 = '', p2 = '';
  for (let r = hr + 2; r < g.length; r++) {
    const b = s(r, 1), cc = s(r, 2), d = s(r, 3);
    const label = b || cc || d;
    if (!label || /^[-\d.,%\s]+$/.test(label)) continue;
    const lv = b ? 1 : cc ? 2 : 3;
    if (lv === 1) { p1 = label; p2 = ''; } else if (lv === 2) p2 = label;
    const row = g[r] || [];
    const a = Array.from({ length: 12 }, (_, k) => (A[k] !== undefined ? num(row[A[k]]) : 0));
    const t = scen.map((x) => num(row[x.c]));
    if (!a.some((v) => v) && !t.some((v) => v)) continue;
    rows.push({ label, lv, p: lv === 1 ? '' : lv === 2 ? p1 : (p2 || p1), a, t });
  }
  return { sheet: sn, scen: scen.map((x) => x.name), rows };
}

/* ---------- 🧮 양식 점검 · 다음 달 양식 만들기 (js/kpicheck.js) ---------- */
let lastFile = null; // { name, buf }
const form = { busy: '', res: null, month: 0, file: '', made: null };
const KIND = { value: ['수식 → 값', 'bad', '지난달은 수식이었는데 이번 달은 숫자를 직접 입력'], blank: ['수식 빠짐', 'bad', '지난달은 수식이었는데 이번 달은 비어 있음'], formula: ['수식 형태 다름', 'warn', '지난달 수식을 옆으로 옮긴 모양과 다름'], newformula: ['값 → 수식', 'info', '지난달은 숫자였는데 이번 달은 수식'], style: ['서식 다름', 'info', '표시 형식·글꼴·채우기 등 서식 번호가 다름'] };
function renderForm(k) {
  const m = form.month || k.month;
  let r = form.res;
  let body = '';
  if (r) {
    // '의도한 변경'으로 확인 처리한 칸은 경고에서 뺌 (notes['formOk{월}'] = {'시트!주소': 1})
    const okMap = (k.notes || {})[`formOk${m}`] || {};
    const isOk = (sh, i) => okMap[`${sh}!${i.addr}`];
    r = r.map((b) => ({ ...b, okCnt: b.issues.filter((i) => isOk(b.sheet, i)).length, issues: b.issues.filter((i) => !isOk(b.sheet, i)) }));
    const okTotal = r.reduce((s, b) => s + b.okCnt, 0);
    const all = r.flatMap((b) => b.issues.map((i) => ({ ...i, sheet: b.sheet, head: b.head })));
    const miss = r.filter((b) => b.missing);
    const cnt = (kd) => all.filter((i) => i.kind === kd).length;
    const need = miss.length + all.filter((i) => KIND[i.kind][1] !== 'info').length;
    const chips = [`<span class="fm-chip ${miss.length ? 'bad' : 'ok'}">당월 열 없음 ${miss.length}</span>`, ...Object.entries(KIND).map(([kd, [nm, lv]]) => `<span class="fm-chip ${cnt(kd) ? lv : 'ok'}">${nm} ${cnt(kd)}</span>`)].join('');
    const sheets = [...new Set(r.map((b) => b.sheet))].map((sh) => {
      const bs = r.filter((b) => b.sheet === sh), is = bs.flatMap((b) => b.issues.map((i) => ({ ...i, head: b.head }))), ms = bs.filter((b) => b.missing);
      const bad = is.filter((i) => KIND[i.kind][1] !== 'info').length + ms.length;
      const rows = bs.reduce((s, b) => s + b.stats.rows, 0);
      const list = [...ms.map((b) => `<li class="bad"><b>당월 열 없음</b> ${esc(b.head)} 표에 ${m}월 열이 없음</li>`),
        ...['value', 'blank', 'formula', 'newformula', 'style'].flatMap((kd) => { const xs = is.filter((i) => i.kind === kd); if (!xs.length) return []; const [nm, lv, ds] = KIND[kd];
          return [`<li class="${lv}"><b>${nm} ${xs.length}칸</b> <small class="muted">${ds}</small> <button type="button" class="btn sm ghost fm-okall" data-fm-okall="${esc(sh)}|${kd}">✔ 모두 의도한 변경</button><ul>${xs.slice(0, kd === 'style' ? 5 : 30).map((i) => `<li><label class="fm-ok" title="의도한 변경이면 체크 — 이번 달 경고에서 뺌"><input type="checkbox" data-fm-ok="${esc(sh)}!${esc(i.addr)}"></label> <code>${esc(i.addr)}</code> ${kd === 'style' ? `서식 ${i.prev} → ${i.cur}` : `${m - 1}월 <code>${esc(String(i.prev).slice(0, 120))}</code> → ${m}월 <code>${esc(String(i.cur).slice(0, 120)) || '(빈칸)'}</code>${i.want ? ` <small class="muted">예상 <code>${esc(i.want.slice(0, 120))}</code></small>` : ''}`}</li>`).join('')}${xs.length > (kd === 'style' ? 5 : 30) ? `<li class="muted">… 외 ${xs.length - (kd === 'style' ? 5 : 30)}칸</li>` : ''}</ul></li>`]; })];
      const okN = bs.reduce((s, b) => s + b.okCnt, 0);
      return `<details class="fm-sh" ${bad ? 'open' : ''}><summary>${bad ? '⚠️' : '✅'} <b>${esc(sh)}</b> <small class="muted">월 표 ${bs.length}개 · ${rows}줄 확인${bad ? ` · 확인 필요 ${bad}` : ''}${okN ? ` · 의도한 변경 ${okN}칸 확인됨` : ''}</small></summary>${list.length ? `<ul class="fm-list">${list.join('')}</ul>` : '<p class="muted sm">지난달과 같은 모양입니다.</p>'}</details>`;
    }).join('');
    body = `<p class="fm-verdict ${need ? 'bad' : 'ok'}">${need ? `⚠️ 확인 필요 ${need}건` : '✅ 지난달과 이어지는 양식에 문제 없음'}${okTotal ? ` <small class="muted">(의도한 변경으로 확인한 ${okTotal}칸 제외 · <button type="button" class="btn sm ghost" id="fmOkReset">확인 취소</button>)</small>` : ''}</p><div class="fm-chips">${chips}</div><p class="muted sm">「${esc(form.file)}」 ${m}월 열을 ${m - 1}월 열과 칸마다 비교 (원본 데이터 시트 GL·TB·Detail of sales·매출원가·폐기원가 등은 제외)</p>${sheets}`;
  }
  const made = form.made ? `<div class="card fm-made"><b>📄 ${form.made.next}월 양식을 만들었습니다</b> — 「${esc(form.made.name)}」 내려받기 됨
    <ul class="fm-list">${form.made.report.inserted.length ? `<li><b>열 추가</b> ${form.made.report.inserted.map(esc).join(', ')}</li>` : ''}
    <li><b>수식·서식 채움</b> ${form.made.report.filled.map(esc).join(', ') || '없음'}</li>
    ${form.made.report.pointers.length ? `<li><b>당월 참조 이동</b>(BGT·Target·% Ach.·누계 열) ${form.made.report.pointers.map(esc).join(', ')}</li>` : ''}
    ${form.made.report.months.length ? `<li><b>결산월</b> ${form.made.report.months.map(esc).join(', ')}</li>` : ''}
    ${form.made.report.manual.length ? `<li class="warn"><b>직접 입력 칸 ${form.made.report.manual.length}개</b> <small class="muted">이번 달에 숫자를 손으로 넣은 칸 — 다음 달도 직접 입력 (aT 지원금·정산 하드코딩 등)</small><br><small>${form.made.report.manual.slice(0, 40).map(esc).join(', ')}${form.made.report.manual.length > 40 ? ' …' : ''}</small></li>` : ''}
    ${form.made.report.noCol.length ? `<li class="warn"><b>열을 못 만든 표</b> ${form.made.report.noCol.map(esc).join(', ')}</li>` : ''}
    <li class="muted">다음 단계: 엑셀에서 열고(처음 1~2분 다시 계산) → 원본 데이터 시트(GL·TB·PL·Detail of sales·매출원가·폐기원가 등)에 ${form.made.next}월 로우데이터 붙여넣기 → 저장 후 여기서 다시 「양식 점검」</li></ul></div>` : '';
  return `<article class="rp fm"><header class="rp-head"><h2>🧮 베트남 KPI 양식 점검</h2>
    <p>KPI 파일의 <b>${m}월 열</b>이 <b>${m - 1}월 열</b>과 같은 수식·서식으로 이어졌는지 칸마다 확인합니다. 마감 뒤에는 <b>${m + 1}월 양식</b>(${m + 1}월 열 수식·서식을 미리 만든 파일)을 만들어 로우데이터만 붙이면 되게 합니다.</p></header>
    <div class="cm-tools"><button type="button" class="btn sm primary" id="fmRun" ${form.busy ? 'disabled' : ''}>🔍 양식 점검</button>
      <button type="button" class="btn sm" id="fmMake" ${form.busy || m >= 12 ? 'disabled' : ''}>📄 ${m + 1}월 양식 만들기</button>
      <button type="button" class="btn sm ghost" id="fmPick">📂 다른 KPI 파일</button><input type="file" id="fmIn" accept=".xlsx,.xlsm" hidden>
      <span class="muted sm">${form.busy ? esc(form.busy) : lastFile ? `대상: 「${esc(lastFile.name)}」` : '이번에 불러온 파일이 없어 📂로 KPI 파일을 골라 주세요'}</span></div>
    ${made}${body}</article>`;
}
async function formRun(k, make) {
  if (!lastFile) { $('#fmIn')?.click(); return; }
  const KC = await import('./kpicheck.js');
  const prog = (t) => { form.busy = t; const s = document.querySelector('.fm .cm-tools .muted'); if (s) s.textContent = t; };
  try {
    prog('파일 여는 중…');
    const book = await KC.openBook(lastFile.buf.slice(0));
    // 결산월: Summary B3 (없으면 불러온 KPI 월)
    const sm = book.sheets.find((s) => /^summary/i.test(s.name.trim()));
    let m = k.month;
    if (sm) { const c = KC.parseSheet(await book.zip.file(sm.path).async('string'), book.sst).get('3,2'); if (c && typeof c.v === 'number' && c.v >= 1 && c.v <= 12) m = c.v; }
    form.month = m; form.file = lastFile.name;
    if (!make) { form.res = await KC.checkBook(book, m, prog); form.made = null; }
    else {
      const { blob, report } = await KC.makeNextMonth(book, m, prog);
      const yymm = String(k.year).slice(2) + String(m + 1).padStart(2, '0');
      const name = /\d{4}/.test(lastFile.name) ? lastFile.name.replace(/(\d{2})(\d{2})(?=[^\d]*$|_)/, yymm).replace(/_v\w+(\.xls\w)$/i, '_양식$1') : `KPI_베트남_${yymm}_양식.xlsx`;
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name.endsWith('.xlsx') ? name : name.replace(/\.\w+$/, '') + '.xlsx'; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 60000);
      form.made = { name: a.download, next: m + 1, report };
    }
  } catch (e) { toastFn('양식 작업 실패: ' + (e.message || e)); }
  form.busy = '';
  render();
}

async function readFile(f) {
  const XLSX = await loadXLSX();
  const buf = await f.arrayBuffer();
  lastFile = { name: f.name, buf }; // 🧮 양식 점검·다음 달 양식에서 다시 씀(메모리에만)
  const names = XLSX.read(buf, { bookSheets: true }).SheetNames;
  const sn = names.find((n) => /^summary/i.test(n.trim()));
  const extra = names.filter((n) => TARGET_SHEET(n) || PPT_SHEET(n) || GL_SHEET(n) || HUB_SHEET(n));
  const wb = XLSX.read(buf, sn ? { sheets: [sn, ...extra], cellStyles: true } : {}); // cellStyles: 숨긴 행 정보
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
    let nb = 0; try { nb = kpiChecks(store.find('kpi', k.id)).filter((x) => x.lv === 'bad').length; } catch { /* 점검 실패 무시 */ }
    toastFn(`${k.year}년 ${k.month}월 KPI를 불러왔습니다.${nb ? ` 🔎 파일 점검: 확인 필요 ${nb}건` : ''}`);
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
    if (ach < 0.9 || ach > 1.1) add(ach >= 1 ? 'good' : 'bad', `매출 BGT 달성률 ${Math.round(ach * 100)}% (BGT ${f1(bud)} / 실적 ${f1(cur)}백만원)`);
  }
  if (M.ebit && M.ebit.b[i]) { const d = M.ebit.a[i] - M.ebit.b[i]; if (Math.abs(d) >= 3) add(d > 0 ? 'good' : 'bad', `EBIT BGT 대비 ${d > 0 ? '+' : ''}${f1(d)}백만원 (BGT ${f1(M.ebit.b[i])} / 실적 ${f1(M.ebit.a[i])})`); }
  if (m >= 3 && cur > 0 && cur >= Math.max(...S.a.slice(0, i))) add('good', `올해 최고 월 매출 (${f1(cur)}백만원)`);
  const order = { bad: 0, good: 1, info: 2 };
  return out.sort((a, b) => order[a.lv] - order[b.lv]);
}

/* ---------- 그래프 (SVG) ---------- */
function niceMax(v) { if (v <= 0) return 1; const p = 10 ** Math.floor(Math.log10(v)); return Math.ceil(v / p / (v / p > 5 ? 2 : 1)) * p * (v / p > 5 ? 2 : 1); }
// 월별 막대(실적) + BGT 표시 + (선택) 비율 선
function barChart({ a, b, m, line, lineLabel, sel, title, cnt = 12, vlabels = false }) {
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
    if (on && vlabels) s += `<text x="${x + bw / 2}" y="${v >= 0 ? y(v) - 5 : y(v) + 13}" class="kg-ax kg-val" text-anchor="middle">${f1(v)}</text>`;
    if (b && b[k]) s += `<line x1="${x + bw * .1}" x2="${x + bw * .9}" y1="${y(b[k])}" y2="${y(b[k])}" class="kg-bud"><title>${k + 1}월 BGT ${f1(b[k])}</title></line>`;
    s += `<text x="${x + bw / 2}" y="${H - 8}" class="kg-ax ${k === sel - 1 ? 'sel' : ''}" text-anchor="middle">${k + 1}월</text>`;
  }
  if (line) {
    const lv = line.slice(0, m), lh = Math.max(0.05, ...lv.map(Math.abs));
    const ly = (v) => T + (1 - (v + lh) / (2 * lh)) * (H - T - Bm);
    s += `<polyline class="kg-line" points="${lv.map((v, k) => `${L + k * bw + bw / 2},${ly(v)}`).join(' ')}"/>`;
    s += lv.map((v, k) => `<circle cx="${L + k * bw + bw / 2}" cy="${ly(v)}" r="3.5" class="kg-dot"><title>${k + 1}월 ${lineLabel} ${pc(v)}</title></circle>`).join('');
    s += `<text x="${W - R + 6}" y="${ly(lv[lv.length - 1] || 0) + 4}" class="kg-ax kg-lt">${pc(lv[lv.length - 1] || 0)}</text>`;
  }
  const legend = `<span class="kl"><i class="kl-bar"></i>실적</span>${b ? '<span class="kl"><i class="kl-bud"></i>BGT</span>' : ''}${line ? `<span class="kl"><i class="kl-line"></i>${lineLabel}</span>` : ''}`;
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

/* ---------- 🎯 Target 대비 ('Target … PL' 시트의 시나리오 = 월 기준 목표치) ---------- */
function tmodel(T) {
  if (!T || !T.rows?.length) return null;
  const key = (r) => r.label.replace(/\s/g, '');
  const f = (re, lv) => T.rows.find((r) => (!lv || r.lv === lv) && re.test(key(r)));
  const sales = f(/^매출액/, 1), sell = f(/^판매비/, 1), admin = f(/^관리비/, 1), at = f(/AT지원/i);
  const kids = (r) => (r ? T.rows.filter((x) => x.lv === 3 && x.p === r.label) : []);
  return {
    sales, qty: f(/^물량/, 1), price: f(/판매단가/), cogs: f(/^매출원가/, 1), gp: f(/^GP$/i, 1), gpp: f(/^GP%$/i),
    sell, at, cm: f(/^CM$/i, 1), cmp: f(/^CM%$/i), admin, ebit: f(/^EBIT$/i, 1), ebitp: f(/^EBIT%$/i),
    items: kids(sales), sellItems: kids(sell), atItems: kids(at), adminItems: kids(admin),
  };
}
const tgtIdx = (T) => Math.min(Math.max(0, ui.tgt || 0), (T?.scen?.length || 1) - 1);
// 최근 3개월(또는 m개월 미만이면 그만큼) 평균 — 비율·단가는 합계로 다시 계산
function recentAvg(TM, r, kind, m) {
  const s = Math.max(1, m - 2), cnt = m - s + 1;
  const sm = (x) => (x ? sumR(x.a, s, m) : 0);
  if (kind === 'pct') { const base = sm(TM.sales); const num = r === TM.gpp ? sm(TM.gp) : r === TM.cmp ? sm(TM.cm) : r === TM.ebitp ? sm(TM.ebit) : null; return num !== null && base ? num / base : sumR(r.a, s, m) / cnt; }
  if (kind === 'won' && r === TM.price) { const q = sm(TM.qty); return q ? sm(TM.sales) / q * 1000 : 0; }
  return sumR(r.a, s, m) / cnt;
}
function targetLines(TM) {
  const L = [];
  const push = (r, name, kind, cls = '') => r && L.push({ r, name: name || r.label, kind, cls });
  push(TM.sales, '매출액', 'amt', 'key');
  TM.items.forEach((x) => push(x, x.label, 'amt', 'sub'));
  push(TM.qty, '매출물량 (톤)', 'ton', 'sub');
  push(TM.price, '판매단가 (원/kg)', 'won', 'sub');
  push(TM.cogs, '매출원가', 'cost');
  push(TM.gp, 'GP (매출총이익)', 'amt', 'key');
  push(TM.gpp, 'GP%', 'pct', 'sub');
  push(TM.sell, '판매비', 'cost');
  TM.sellItems.forEach((x) => push(x, x.label, 'cost', 'sub'));
  push(TM.at, 'aT 지원사업 수익', 'amt', 'sub');
  push(TM.cm, 'CM (공헌이익)', 'amt', 'key');
  push(TM.cmp, 'CM%', 'pct', 'sub');
  push(TM.admin, '관리비', 'cost');
  TM.adminItems.forEach((x) => push(x, x.label, 'cost', 'sub'));
  push(TM.ebit, 'EBIT (영업이익)', 'amt', 'key');
  push(TM.ebitp, 'EBIT%', 'pct', 'sub');
  return L;
}
const tAch = (a, t, kind) => (kind === 'amt' && t > 0 && a >= 0 ? `${Math.round(a / t * 100)}%` : kind === 'cost' && t > 0 ? `${Math.round(a / t * 100)}%` : '');
function targetTable(k, m, { cls = 'rt' } = {}) {
  const T = k.target, TM = tmodel(T);
  if (!TM) return '';
  const si = tgtIdx(T), i = m - 1;
  const others = T.scen.map((n, j) => j).filter((j) => j !== si);
  const rows = targetLines(TM).map((x) => {
    const F = fmtOf(x.kind), a = x.r.a[i], t = x.r.t[si], r3 = recentAvg(TM, x.r, x.kind, m);
    return `<tr class="${x.cls}"><th>${esc(x.name)}</th><td class="${a < 0 ? 'neg' : ''} hl">${F(a)}</td><td class="${r3 < 0 ? 'neg' : ''}">${F(r3)}</td>
      <td class="tg">${F(t)}</td><td>${delta(a, t, x.kind)}</td><td class="muted">${tAch(a, t, x.kind)}</td>${others.map((j) => `<td class="muted">${F(x.r.t[j])}</td>`).join('')}</tr>`;
  }).join('');
  const seg = `<div class="seg kt-tsel">${T.scen.map((n, j) => `<button type="button" data-tgt="${j}" class="${j === si ? 'on' : ''}">${esc(n)}</button>`).join('')}</div>`;
  return `${seg}<div class="kt-scroll"><table class="${cls}"><thead><tr><th>항목</th><th>${m}월 실적</th><th>최근 3개월 평균<br><small>${Math.max(1, m - 2)}~${m}월</small></th><th>🎯 ${esc(T.scen[si])}</th><th>${m}월 − Target</th><th>달성률</th>${others.map((j) => `<th class="muted">${esc(T.scen[j])}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table></div>
    <p class="muted sm">Target은 '${esc(T.sheet)}' 시트의 월 기준 목표치입니다. 비용 항목은 Target보다 적을수록 초록(▼)으로 표시합니다.</p>`;
}
function targetFindings(k, m) {
  const TM = tmodel(k.target);
  if (!TM) return [];
  const si = tgtIdx(k.target), i = m - 1, nm = k.target.scen[si], out = [];
  const add = (lv, txt) => out.push({ lv, txt });
  if (TM.sales && TM.sales.t[si] > 0) { const a = TM.sales.a[i], t = TM.sales.t[si]; add(a >= t ? 'good' : 'bad', `🎯 ${nm} 매출 ${f1(t)} 대비 ${m}월 ${f1(a)}백만원 (달성 ${Math.round(a / t * 100)}%, ${a - t >= 0 ? '+' : ''}${f1(a - t)})`); }
  if (TM.gpp) { const d = TM.gpp.a[i] - TM.gpp.t[si]; if (Math.abs(d) >= 0.01) add(d >= 0 ? 'good' : 'bad', `🎯 GP% ${pc(TM.gpp.a[i])} — ${nm} ${pc(TM.gpp.t[si])} 대비 ${d > 0 ? '+' : ''}${(d * 100).toFixed(1)}%p`); }
  for (const [r, n] of [[TM.cm, 'CM'], [TM.ebit, 'EBIT']]) if (r) { const d = r.a[i] - r.t[si]; if (Math.abs(d) >= 0.5) add(d >= 0 ? 'good' : 'bad', `🎯 ${n} ${f1(r.a[i])} — ${nm} ${f1(r.t[si])} 대비 ${d > 0 ? '+' : ''}${f1(d)}백만원`); }
  const over = [...TM.sellItems.map((x) => ['판매비', x]), ...TM.adminItems.map((x) => ['관리비', x])].filter(([, x]) => x.a[i] - x.t[si] >= 0.5 && (x.t[si] <= 0 || x.a[i] / x.t[si] >= 1.2));
  if (over.length) add('bad', `🎯 Target 초과 비용: ${over.map(([g, x]) => `${g} ${x.label} ${f1(x.a[i])} (Target ${f1(x.t[si])})`).join(', ')}`);
  const under = TM.items.filter((x) => x.t[si] > 0.5 && x.a[i] / x.t[si] < 0.8);
  if (under.length) add('bad', `🎯 Target 미달 품목: ${under.map((x) => `${x.label} ${f1(x.a[i])}/${f1(x.t[si])}`).join(', ')}`);
  return out;
}

/* ---------- 📝 경영회의 코멘트 초안: 장표 표(PPT 시트) + A~E 코멘트(근거 숫자·원장·현지 문의·AI 초안) ---------- */
// 장표 표(PPT 시트)는 월 열을 손으로 추가해서, 새 달 KPI를 넣어도 표가 전월까지만 있을 수 있음
//   → 'Target … PL' 시트의 월 실적으로 빠진 달을 같은 규칙으로 채움 (9월 장표 숫자와 맞춰 확인한 규칙)
//   매출원가 상품대 = 총원물대 + 총 수출제비용, 판매비 항목 = Target 판매비 항목 − 같은 이름의 AT지원 항목,
//   관리비 인건비 = 인건비 + 식대 및 접대비, 기타지급수수료 = 회계/세무 + 사무용품 렌트비, '기타…' = 합계 − 나머지
const nzl = (s) => String(s ?? '').replace(/\(.*?\)/g, '').replace(/[\s/]+/g, '');
function pptMonth(k, rows, mm) {
  const T = k.target, j = mm - 1, M = model(k);
  const tr = (re, lv = 3) => T.rows.find((r) => r.lv <= lv && re.test(nzl(r.label)));
  const g = (r) => (r ? r.a[j] || 0 : 0);
  const kidsOf = (re) => T.rows.filter((r) => r.lv === 3 && re.test(nzl(r.p)));
  const kid = (re, label) => kidsOf(re).find((r) => nzl(r.label) === nzl(label));
  const items = kidsOf(/^매출액/).map((r) => r.label).filter((l) => nzl(l) !== '팽이');
  const other = (re) => items.reduce((s, l) => s + g(kid(re, l)), 0);
  const sales = g(tr(/^매출액/, 1)), qty = g(tr(/^물량/, 1));
  const cogsParts = { 상품대: g(tr(/^총원물대/)) + g(tr(/^총수출제비용/)), 수입제비용: g(tr(/^총수입제비용/)), 폐기원가: g(tr(/^총폐기금액/)) };
  const at = kidsOf(/지원/);
  const sell = (label) => { const L = nzl(label); const a = kidsOf(/^판매비/).filter((r) => nzl(r.label) === L || nzl(r.label).replace(/비$/, '') === L.replace(/비$/, '')); if (!a.length) return null; return a.reduce((s, r) => s + g(r), 0) - at.filter((r) => nzl(r.label) === L).reduce((s, r) => s + g(r), 0); };
  const ADM = [[/^인건비/, /^인건비|^식대/], [/^사무실/, /^사무실/], [/^숙소/, /^숙소/], [/^여비/, /^여비/], [/^기타지급수수료/, /^회계|^사무용품/]];
  const adm = (label) => { const L = nzl(label); const m = ADM.find(([re]) => re.test(L)); if (!m) return null; return kidsOf(/^관리비/).filter((r) => m[1].test(nzl(r.label))).reduce((s, r) => s + g(r), 0); };
  const top = (L) => (/^매출액/.test(L) ? sales : /^매출물량/.test(L) ? qty : /^매출단가/.test(L) ? g(tr(/^판매단가/)) : /^매출원가/.test(L) ? g(tr(/^매출원가/, 1))
    : /단위원가/.test(L) ? g(tr(/^단위원가/)) : /^GP$/i.test(L) ? g(tr(/^GP$/i, 1)) : /^판매비/.test(L) ? g(tr(/^판매비$/, 1)) : /^CM$/i.test(L) ? g(tr(/^CM$/i, 1))
      : /^관리비/.test(L) ? g(tr(/^관리비/, 1)) : /^조정EBITDA/i.test(L) ? (M.ebitda ? M.ebitda.a[j] || 0 : g(tr(/^조정EBIT$/i, 1))) : /^EBITDA/i.test(L) ? g(tr(/^EBIT$/i, 1)) : 0);
  let sec = '', secV = 0;
  const out = rows.map((r) => {
    const L = nzl(r.label);
    if (r.lv === 1 && r.kind !== 'pct') { sec = L; secV = top(L); return secV; }
    if (r.kind === 'pct') return sales ? secV / sales : 0;
    const pp = L === '팽이';
    if (/^매출액/.test(sec)) return L === '기타품목' ? other(/^매출액/) : g(kid(/^매출액/, r.label));
    if (/^매출물량/.test(sec)) return L === '기타품목' ? other(/^물량/) : g(kid(/^물량/, r.label));
    if (/^매출단가/.test(sec)) { if (L !== '기타품목') return g(kid(/^판매단가/, r.label)); const q = other(/^물량/); return q ? other(/^매출액/) / q * 1000 : 0; }
    if (/^매출원가/.test(sec)) return L in cogsParts ? cogsParts[L] : L === '기타품목' ? other(/^매출원가/) : g(kid(/^매출원가/, r.label));
    if (/^품목별단위원가/.test(sec)) { if (L !== '기타품목') return g(kid(/^단위원가/, r.label)); const q = other(/^물량/); return q ? other(/^매출원가/) / q * 1000 : 0; }
    if (/^단위원가/.test(sec)) return qty && L in cogsParts ? cogsParts[L] / qty * 1000 : 0;
    if (/^판매비/.test(sec)) return /^기타/.test(L) ? null : sell(r.label) ?? 0;
    if (/^관리비/.test(sec)) return /^기타$/.test(L) ? null : adm(r.label) ?? 0;
    return pp ? 0 : 0; // CM 품목별 등(장표에서 숨긴 행)
  });
  // '기타…' = 합계 − 같은 묶음의 나머지(숨긴 행 제외)
  let s0 = -1;
  rows.forEach((r, x) => { if (r.lv === 1 && r.kind !== 'pct') s0 = x; if (out[x] === null) { let rest = 0; for (let y = s0 + 1; y < rows.length && !(rows[y].lv === 1 && rows[y].kind !== 'pct'); y++) if (y !== x && !rows[y].h && rows[y].kind !== 'pct' && out[y] !== null) rest += out[y]; out[x] = out[s0] - rest; } });
  return out;
}
let pptCache = null;
function pptOf(k) {
  const P = k.ppt;
  if (!P || !k.target || !P.months.length) return P;
  const last = P.months[P.months.length - 1];
  if (last >= k.month || P.months.length !== last) return P;
  const ck = `${k.id}|${k.loadedAt}|${k.month}`;
  if (pptCache?.ck === ck) return pptCache.P;
  const M = model(k), add = [];
  for (let mm = last + 1; mm <= k.month; mm++) add.push(mm);
  const vals = add.map((mm) => pptMonth(k, P.rows, mm));
  const j = k.month - 1;
  const bOf = (L) => { const r = /^매출액/.test(L) ? M.sales : /^매출물량/.test(L) ? M.qty : /^매출단가/.test(L) ? M.price : /^매출원가/.test(L) ? M.cogs : /^GP$/i.test(L) ? M.gp : /^판매비/.test(L) ? M.sell : /^CM$/i.test(L) ? M.cm : /^관리비/.test(L) ? M.admin : /EBITDA/i.test(L) ? M.ebitda : null; return r ? r.b[j] || 0 : 0; };
  let secB = 0, salesB = bOf('매출액');
  const rows = P.rows.map((r, x) => {
    const a = [...r.a, ...vals.map((v) => v[x])], cur = a[a.length - 1], L = nzl(r.label);
    let bgt = 0;
    if (r.lv === 1 && r.kind !== 'pct') { bgt = /단위원가/.test(L) ? (M.qty?.b[j] ? (M.cogs?.b[j] || 0) / M.qty.b[j] * 1000 : 0) : bOf(L); secB = bgt; } else if (r.kind === 'pct') bgt = salesB ? secB / salesB : 0;
    const ach = (base) => (r.kind === 'pct' ? cur - base : base ? cur / base : 0);
    return { ...r, a, bgt, bAch: bgt ? ach(bgt) : 0, tAch: r.tgt ? ach(r.tgt) : 0 };
  });
  const out = { ...P, months: [...P.months, ...add], rows, auto: add };
  pptCache = { ck, P: out };
  return out;
}

const VND_PER_KRW = 18.5; // KPI 허브와 같은 환산(VND ÷ 18.5 = 원)
const ACC_NAME = { 511: '매출', 515: '금융수익', 632: '매출원가', 635: '금융비용', 6411: '판매 인건비', 6412: '판매 재료·포장', 6413: '판매 소모품', 6414: '판매 감가상각', 6415: '판매 보증', 6417: '판매 외주용역(운반·창고·라벨 등)', 6418: '판매 기타', 6421: '관리 인건비', 6422: '관리 소모품', 6423: '관리 사무용품', 6424: '관리 감가상각', 6425: '세금·공과', 6426: '충당금', 6427: '관리 외주용역(임차·통신 등)', 6428: '관리 기타', 711: '기타수익', 811: '기타비용', 821: '법인세' };
const accName = (a) => ACC_NAME[a.slice(0, 4)] || ACC_NAME[a.slice(0, 3)] || '';
const krwM = (vnd) => vnd / VND_PER_KRW / 1e6; // VND → 백만원
const fP = (v, kind) => { // 장표 숫자 형식: 음수는 괄호, 10 이상은 정수
  if (kind === 'pct') return Math.abs(v) < 5e-4 ? '-' : (v < 0 ? `(${Math.abs(v * 100).toFixed(1)}%)` : `${(v * 100).toFixed(1)}%`);
  if (Math.abs(v) < 0.05 && kind !== 'won') return '-';
  const a = Math.abs(v), t = kind === 'won' ? Math.round(a).toLocaleString('ko-KR') : a >= 10 ? Math.round(a).toLocaleString('ko-KR') : a.toFixed(1);
  return v < 0 ? `(${t})` : t;
};
// % Ach.는 시트 값이 아니라 당월 값으로 다시 계산 (10월 파일에서 시트 % Ach.가 9월 값으로 남아 있던 문제)
//   양수 기준: 실적÷기준 / 음수 기준(적자): (실적−기준)÷|기준| / 비율(%) 행: 실적−기준(%p) — 9월 장표 숫자와 같은 방식
//   EBITDA 행은 장표 원래 수식 그대로 1 − 실적÷기준 (=-IFERROR(AW/AY,)+1)
const achOf = (r, base) => { const a = r.a[r.a.length - 1]; if (!base || a == null) return 0; if (r.kind === 'pct') return a - base; if (/EBITDA/i.test(r.label)) return 1 - a / base; return base > 0 ? a / base : (a - base) / Math.abs(base); };
const fAch = (v, kind) => (kind === 'pct' ? (Math.abs(v) < 5e-4 ? '-' : `${v > 0 ? '' : '('}${Math.abs(v * 100).toFixed(1)}%p${v > 0 ? '' : ')'}`) : !v ? '-' : `${Math.round(v * 100)}%`);
const CM_GROUPS = [
  { id: 'A', name: '매출액', re: /^매출액/, unit: '백만원', acc: /^5/ },
  { id: 'B', name: '단위원가', re: /^단위원가/, unit: '원', acc: /^632|^1561/ },
  { id: 'C', name: '판매비', re: /^판매비/, unit: '백만원', acc: /^641/ },
  { id: 'D', name: '관리비', re: /^관리비/, unit: '백만원', acc: /^642/ },
  { id: 'E', name: '조정 EBITDA', re: /^조정\s*EBITDA/i, unit: '백만원', acc: null },
];
function pptBlock(P, re) { const i = P.rows.findIndex((r) => r.lv === 1 && re.test(r.label.replace(/\s+/g, ' '))); if (i < 0) return null; const kids = []; for (let j = i + 1; j < P.rows.length && (P.rows[j].lv === 2 || P.rows[j].kind === 'pct'); j++) if (P.rows[j].kind !== 'pct') kids.push(P.rows[j]); const seen = new Set(); // 같은 이름 세부 행이 두 번 있으면(숨긴 행 등) 보이는 행 하나만
  const uniq = [...kids.filter((x) => !x.h), ...kids.filter((x) => x.h)].filter((x) => !seen.has(x.label) && seen.add(x.label));
  return { row: P.rows[i], kids: kids.filter((x) => uniq.includes(x)) }; }
const sgn = (v, d = 0) => `${v > 0 ? '+' : v < 0 ? '-' : '±'}${Math.abs(v).toLocaleString('ko-KR', { maximumFractionDigits: d, minimumFractionDigits: d })}`;
// 제목줄: 9월 장표와 같은 문구 — "매출액 단기 Target 초과 달성 +15백만원, +37%," / "계획 대비 -67백만원, -55%"
//   (조정 EBITDA는 장표처럼 Target 대비 한 줄만)
const dec = (v) => (Math.abs(v) >= 1 ? 0 : 1);
function cmHeadline(G, B) {
  const r = B.row, cur = r.a[r.a.length - 1], won = G.unit === '원';
  // 장표처럼 표에 보이는(반올림된) 숫자끼리 뺌: 55 - 122 = -67
  const rd = (v) => (won || Math.abs(v) >= 10 ? Math.round(v) : Math.round(v * 10) / 10);
  const part = (base) => { const d = rd(cur) - rd(base); return `${sgn(d, won ? 0 : dec(d))}${G.unit}${base > 0 && cur >= 0 ? `, ${sgn((cur / base - 1) * 100)}%` : ''}`; };
  // 앵커사 보고용: 좋은 결과는 제목줄에서 바로 드러나게 (숫자는 그대로)
  const up = G.id === 'A' || G.id === 'E'; // 클수록 좋은 항목
  const good = r.tgt && (up ? cur >= r.tgt : cur <= r.tgt);
  let streak = 0;
  if (good && G.id === 'A') for (let j = r.a.length - 1; j >= 0 && r.a[j] >= r.tgt; j--) streak++;
  const lead = G.id === 'A' && r.tgt ? (good ? '단기 Target 초과 달성' : '단기 Target 미달') : '단기 Target 대비';
  const tail = !good ? '' : G.id === 'A' ? (streak >= 2 ? `, ${streak}개월 연속 Target 초과` : '') : G.id === 'E' ? ' 초과 달성' : '로 Target 이내 관리';
  const L = [`${G.name} ${lead} ${r.tgt ? part(r.tgt) : '-'}${tail}`];
  if (r.bgt && G.id !== 'E') L.push(`계획 대비 ${part(r.bgt)}`);
  return L.join(',\n');
}
const CHK = ' [현지 확인]';
const pctRow = (P, re) => { const i = P.rows.findIndex((r) => r.lv === 1 && r.kind !== 'pct' && re.test(r.label.replace(/\s+/g, ' '))); return i >= 0 && P.rows[i + 1]?.kind === 'pct' ? P.rows[i + 1] : null; };
// 세부 줄(장표 형식): "항목: [사유] 전월 대비 -85원/kg 감소" — [사유]는 현지 확인 후 채울 자리
const WHY = '[사유] ';
function cmBullets(k, G, B) {
  const out = [], n = B.row.a.length, i = n - 1, p = n - 2;
  if (p < 0) return out;
  const P = pptOf(k), mn = P.months[i];
  const vol = pptBlock(P, /^매출물량/), kg = vol ? vol.row.a[i] * 1000 : 0, kgp = vol ? vol.row.a[p] * 1000 : 0;
  const word = (d) => (d > 0 ? '증가' : '절감');
  // B(원가)는 장표처럼 부호를 붙이고(-85원/kg 감소), 판매·관리비는 절감은 부호 없이(2.8백만원 절감)·증가는 +로
  const amt = (d, u) => (G.id === 'B' || d > 0 ? sgn(d, u === '백만원' ? 1 : 0) : Math.abs(d).toLocaleString('ko-KR', { maximumFractionDigits: u === '백만원' ? 1 : 0, minimumFractionDigits: u === '백만원' ? 1 : 0 })) + u;
  // 앵커사 보고 문체: 좋은 점 먼저, 숫자는 "당월(전월)" 형식, 사유가 필요한 곳은 [현지 확인]
  const pm = P.months[p], pc0 = (c, b) => (b ? `${c >= b ? '+' : '-'}${Math.abs(Math.round((c / b - 1) * 100))}%` : '');
  const best = (arr) => arr.slice(0, i + 1).every((v, j) => j === i || v <= arr[i]); // 연중 최고(최소 손실)
  if (G.id === 'A') {
    const M = model(k), s = B.row, d = s.a[i] - s.a[p];
    out.push(`${mn}월 매출 ${f1(s.a[i])}백만원(${pm}월 ${f1(s.a[p])})${best(s.a) ? '으로 연중 최고 실적' : ''}, 전월 대비 ${sgn(d, 1)}백만원(${pc0(s.a[i], s.a[p])}) ${d >= 0 ? '성장' : '감소'}`);
    const ch = M.chans.filter((x) => Math.abs(x.a[i]) >= 0.05).sort((x, y) => (y.a[i] - y.a[p]) - (x.a[i] - x.a[p]));
    for (const x of ch) {
      const dd = x.a[i] - x.a[p], q = chanQtyOf(k, x.label);
      // 거래처별 판매물량(허브 시트) — 장표 형식 "매출 50백만원, 판매물량 8.6톤"
      const qs = q && Math.abs(q[mn - 1] || 0) >= 0.005 ? `, 판매물량 ${q[mn - 1].toFixed(1)}톤(${pm}월 ${(q[mn - 2] || 0).toFixed(1)})` : '';
      out.push(`${x.label}: ${mn}월 매출 ${f1(x.a[i])}백만원(${pm}월 ${f1(x.a[p])})${qs}, 전월 대비 매출 ${sgn(dd, 1)}백만원(${pc0(x.a[i], x.a[p])}) ${dd >= 0 ? '성장' : '감소'}${CHK}`);
    }
    if (vol) { const v = vol.row; out.push(`판매물량 ${v.a[i].toFixed(1)}톤(${pm}월 ${v.a[p].toFixed(1)})${v.tgt ? `, 단기 Target ${v.tgt.toFixed(1)}톤 대비 ${Math.round(v.a[i] / v.tgt * 100)}% 달성` : ''}`); }
    const pr = pptBlock(P, /^매출단가/);
    if (pr) {
      const r = pr.row, dd = r.a[i] - r.a[p], why = priceReason(M, mn - 1, mn - 2);
      const head = r.tgt && r.a[i] >= r.tgt ? `단기 Target 대비 ${sgn(r.a[i] - r.tgt)}원/kg 초과 달성` : Math.abs(dd / r.a[p]) < 0.01 ? '전월 수준 유지' : `전월 대비 ${sgn(dd)}원/kg`;
      out.push(`매출단가 ${fP(r.a[i], 'won')}원/kg(${pm}월 ${fP(r.a[p], 'won')}), ${head}${why.length ? ` (${why[0]})` : ''}`);
    }
    out.push(`신규 채널:${CHK.replace(']', ': 입점·첫 납품 진행 현황]')}`);
    return out;
  }
  if (G.id === 'E') {
    const r = B.row, d = r.a[i] - r.a[p], pr = pctRow(P, /^조정\s*EBITDA/i), gp = pptBlock(P, /^GP$/i);
    out.push(`${mn}월 ${G.name} ${f1(r.a[i])}백만원(${pm}월 ${f1(r.a[p])})${best(r.a) ? (r.a[i] < 0 ? '으로 연중 최소 손실' : '으로 연중 최고') : ''}, 전월 대비 ${sgn(d, 1)}백만원 ${d >= 0 ? '개선' : '감소'}`);
    if (pr) { const dp = (pr.a[i] - pr.a[p]) * 100; out.push(`매출 대비 ${(pr.a[i] * 100).toFixed(1)}%(${pm}월 ${(pr.a[p] * 100).toFixed(1)}%)로 ${Math.abs(dp).toFixed(1)}%p ${dp >= 0 ? '개선' : '하락'}`); }
    if (gp) { const g = gp.row; out.push(`GP ${f1(g.a[i])}백만원(${pm}월 ${f1(g.a[p])})${best(g.a) ? '으로 연중 최고' : ''}${g.tgt ? `, 단기 Target 대비 ${sgn(g.a[i] - g.tgt, 1)}백만원` : ''}`); }
    return out;
  }
  if (G.id === 'C') {
    // aT 지원금을 맨 앞에 (지원 전 → 지원 후 판매비)
    const mm = mn;
    for (const { tot, kids } of subsidyRows(k)) {
      const c = Math.abs(tot.a[mm - 1] || 0), b = Math.abs(tot.a[mm - 2] || 0);
      if (c < 0.05) continue;
      const sell = B.row.a[i];
      out.push(`aT 지원사업 지원금 ${f1(c)}백만원 확보(${pm}월 ${f1(b)}${b ? `, ${pc0(c, b)}` : ''}): ${kids.filter((x) => Math.abs(x.a[mm - 1]) >= 0.05).map((x) => `${x.label} ${f1(Math.abs(x.a[mm - 1]))}`).join('·')}백만원, 지원 전 판매비 ${f1(sell + c)} → 지원 후 ${f1(sell)}백만원(${Math.round(c / (sell + c) * 100)}% 절감)`);
    }
  }
  if (G.id === 'D') {
    const r = B.row, d = r.a[i] - r.a[p], pr = pctRow(P, /^관리비/);
    if (d < 0) out.push(`전월 ${f1(r.a[p])}백만원 대비 ${f1(-d)}백만원(${pc0(r.a[i], r.a[p])}) 절감${pr ? `, 매출 대비 관리비율 ${(pr.a[i] * 100).toFixed(1)}%(${pm}월 ${(pr.a[p] * 100).toFixed(1)}%)로 ${Math.abs((pr.a[i] - pr.a[p]) * 100).toFixed(1)}%p 개선` : ''}`);
  }
  // 세부 항목은 전월 대비 변동이 큰 순서로 최대 3개(작은 변동은 생략)
  const kids = B.kids.filter((x) => x.kind !== 'pct').map((x) => {
    const won = x.kind === 'won';
    // 운반비는 장표처럼 kg당 금액으로
    if (G.id === 'C' && /운반/.test(x.label) && kg && kgp) { const d = x.a[i] * 1e6 / kg - x.a[p] * 1e6 / kgp; return { x, d, u: '원/kg', big: Math.abs(d) >= 10 }; }
    const d = x.a[i] - x.a[p];
    return { x, d, u: won ? '원/kg' : '백만원', big: won ? Math.abs(d) >= 10 : Math.abs(d) >= 0.3 };
  }).filter((y) => y.big)
    // 절감(좋은 변화)을 먼저, 그 안에서는 변동 큰 순
    .sort((y, z) => (y.d < 0) === (z.d < 0) ? Math.abs(z.d) / (z.u === '원/kg' ? 100 : 1) - Math.abs(y.d) / (y.u === '원/kg' ? 100 : 1) : y.d < 0 ? -1 : 1).slice(0, 4);
  for (const { x, d, u } of kids) {
    const why = G.id === 'C' || G.id === 'D' ? glReason(k, G.id === 'C' ? '판매비' : '관리비', x.label, mn) : G.id === 'B' ? costReason(k, x.label, mn) : '';
    const v0 = (v) => (Math.abs(v) < 0.05 && x.kind !== 'won' ? '0' : fP(v, x.kind));
    const now = `${v0(x.a[i])}${x.kind === 'won' ? '원/kg' : '백만원'}(${pm}월 ${v0(x.a[p])}), `;
    out.push(`${x.label}: ${now}전월 대비 ${amt(d, u)} ${word(d)}${why ? ` (${G.id === 'B' ? '' : '원장: '}${why})` : ''}${d > 0 ? CHK : ''}`);
  }
  if (G.id === 'B') {
    const wl = wasteLine(k, mn);
    if (wl) out.push(`폐기: ${wl}${wasteOf(k, mn).rate > WASTE_GOAL ? CHK : ''}`);
  }
  return out;
}
function cmFacts(k, G, B) {
  const out = [], n = B.row.a.length, i = n - 1, p = n - 2;
  if (p < 0) return out;
  const P = pptOf(k), vol = pptBlock(P, /^매출물량/);
  const kg = vol ? vol.row.a[i] * 1000 : 0, kgp = vol ? vol.row.a[p] * 1000 : 0;
  for (const x of B.kids.filter((x) => x.kind !== 'pct')) {
    const d = x.a[i] - x.a[p];
    if (Math.abs(d) < (x.kind === 'won' ? 1 : 0.05)) continue;
    const u = x.kind === 'won' ? '원/kg' : x.kind === 'ton' ? '톤' : '백만원';
    let t = `${x.label}: ${fP(x.a[p], x.kind)} → ${fP(x.a[i], x.kind)}${u === '원/kg' ? '원/kg' : ''} (전월 대비 ${sgn(d, x.kind === 'won' ? 0 : 1)}${u})`;
    if (x.tgt) t += ` · Target ${fP(x.tgt, x.kind)}`;
    if (G.id === 'C' && /운반/.test(x.label) && kg && kgp) t += ` · kg당 ${Math.round(x.a[p] * 1e6 / kgp)} → ${Math.round(x.a[i] * 1e6 / kg)}원`;
    out.push(t);
  }
  if (G.id === 'A') {
    const M = model(k);
    for (const x of M.chans) { const c = x.a[i], b = x.a[p]; if (Math.abs(c) >= 0.05 || Math.abs(b) >= 0.05) out.push(`${x.label}: ${f1(b)} → ${f1(c)}백만원 (전월 대비 ${sgn(c - b, 1)})`); }
    if (vol) out.push(`매출물량: ${vol.row.a[p].toFixed(1)} → ${vol.row.a[i].toFixed(1)}톤 · Target ${vol.row.tgt.toFixed(1)}톤`);
    const pr = pptBlock(P, /^매출단가/); if (pr) out.push(`매출단가: ${fP(pr.row.a[p], 'won')} → ${fP(pr.row.a[i], 'won')}원/kg · Target ${fP(pr.row.tgt, 'won')}`);
  }
  if (G.id === 'B') { const w = pptBlock(P, /폐기물량|폐기율/); if (w) out.push(`폐기: ${w.row.label} ${fP(w.row.a[p], w.row.kind)} → ${fP(w.row.a[i], w.row.kind)}`); }
  return out;
}
function cmGl(k, G) {
  if (!k.gl || !G.acc) return { accs: [], lines: [] };
  const m = k.month, i = m - 1, p = m - 2;
  const accs = Object.entries(k.gl.acc).filter(([a]) => G.acc.test(a)).map(([a, v]) => ({ a, cur: v[i] || 0, prev: p >= 0 ? v[p] || 0 : 0 }))
    .filter((x) => Math.abs(x.cur) + Math.abs(x.prev) > 0).sort((x, y) => Math.abs(y.cur - y.prev) - Math.abs(x.cur - x.prev));
  const top = new Set(accs.slice(0, 4).map((x) => x.a));
  const lines = k.gl.lines.filter((l) => l.m === m && top.has(l.a)).slice(0, 8);
  return { accs: accs.slice(0, 8), lines };
}
function cmQuestions(k) {
  const q = [];
  for (const G of CM_GROUPS) {
    const { accs } = cmGl(k, G);
    for (const x of accs.slice(0, 3)) {
      const d = x.cur - x.prev;
      if (Math.abs(krwM(d)) < 0.5) continue;
      const desc = k.gl.lines.find((l) => l.m === k.month && l.a === x.a)?.d || '';
      q.push(`[${G.name}] 계정 ${x.a}${accName(x.a) ? `(${accName(x.a)})` : ''}: ${k.month - 1}월 ${(x.prev / 1e6).toFixed(1)}M → ${k.month}월 ${(x.cur / 1e6).toFixed(1)}M VND (${d > 0 ? '+' : ''}${(d / 1e6).toFixed(1)}M, ≈${sgn(krwM(d), 1)}백만원)${desc ? ` — 주요 거래 "${desc}"` : ''} → 변동 사유 확인 부탁드립니다.`);
    }
  }
  return q;
}
/* ---------- 🗑 폐기물량·폐기율 (KPI 파일에 톤 데이터가 없어 지난달 장표 표 + 이번 달 입력) ---------- */
// 폐기율 = 폐기물량 ÷ (판매물량 + 폐기물량) — 9월 장표(1월 2.4톤/39.8%, 2월 3.3톤/26.9%)로 확인한 식
const WASTE_GOAL = 0.03;
function wasteOf(k, m) {
  const w = (k.notes || {}).waste || {}, M = model(k), raw = w[m];
  if (raw == null || raw === '' || !M.qty) return null;
  const t = typeof raw === 'object' ? +raw.t : +raw;
  if (typeof raw === 'object' && Number.isFinite(raw.r)) return { t, rate: raw.r }; // 지난 장표 값 그대로
  const q = M.qty.a[m - 1] || 0;
  return { t, rate: q + t > 0 ? t / (q + t) : 0 };
}
const wasteT = (v) => (v == null ? '' : typeof v === 'object' ? v.t : v);
function wasteCard(k, mn) {
  const P = pptOf(k), b = P && pptBlock(P, /^매출원가/), kid = b?.kids.find((x) => /폐기/.test(x.label));
  const ms = Array.from({ length: mn }, (_, j) => j + 1), w = (k.notes || {}).waste || {};
  const cell = (m) => { const x = wasteOf(k, m); return x ? `${(x.rate * 100).toFixed(1)}%` : '-'; };
  const cur = wasteOf(k, mn);
  return `<div class="card waste"><div class="kg-h"><b>🗑 폐기물량·폐기율</b><small class="muted">폐기율 = 폐기물량 ÷ (판매물량 + 폐기물량), 목표 ${WASTE_GOAL * 100}%</small></div>
    <div class="kt-scroll"><table class="rt cm-t wt"><thead><tr><th></th>${ms.map((m) => `<th>${m}월</th>`).join('')}</tr></thead><tbody>
    <tr><th>폐기물량(톤)</th>${ms.map((m) => `<td><input class="wt-in" data-waste-m="${m}" inputmode="decimal" value="${wasteT(w[m])}" placeholder="${m === mn ? '입력' : ''}"></td>`).join('')}</tr>
    <tr class="key"><th>폐기율(%)</th>${ms.map((m) => { const x = wasteOf(k, m); return `<td class="${x && x.rate > WASTE_GOAL ? 'neg' : ''}">${cell(m)}</td>`; }).join('')}</tr>
    ${kid && P ? `<tr class="sub"><th>폐기원가(백만원)</th>${ms.map((m) => { const i = P.months.indexOf(m); return `<td>${i >= 0 ? fP(kid.a[i], 'amt') : '-'}</td>`; }).join('')}</tr>` : ''}
    </tbody></table></div>
    <div class="cm-tools"><button type="button" class="btn sm" id="wfLoad">📂 폐기 파일 불러오기</button><input type="file" id="wfIn" accept=".xlsx,.xlsm" hidden>
      <span class="muted sm">「YYMM_베트남 폐기.xlsx」의 '삼도 결산' 시트 '삼도 폐기물량(톤)'을 월별로 채웁니다${(k.notes || {}).wasteSrc ? ` · 마지막: 「${esc(k.notes.wasteSrc)}」` : ''}</span></div>
    <p class="muted sm">${cur ? `${mn}월 폐기율 ${(cur.rate * 100).toFixed(1)}% — 코멘트 B·메일 초안에 자동 반영` : `${mn}월 폐기물량(톤)을 입력하거나 폐기 파일을 불러오면 폐기율이 계산되어 코멘트 B·메일 초안에 들어갑니다.`}</p></div>`;
}
// 폐기 파일(YYMM_베트남 폐기.xlsx) 읽기: B열 '…폐기물량(톤)' 행 + 위쪽 월 머리글(1~12), 아래 C열 품목별 톤
async function readWasteFile(file) {
  const KC = await import('./kpicheck.js');
  const book = await KC.openBook(await file.arrayBuffer());
  const list = [...book.sheets.filter((s) => /삼도\s*결산/.test(s.name)), ...book.sheets.filter((s) => !/삼도\s*결산/.test(s.name))];
  for (const s of list) {
    const f = book.zip.file(s.path);
    if (!f || f._data?.uncompressedSize > 5e6) continue; // 큰 원본 시트는 건너뜀
    const cells = KC.parseSheet(await f.async('string'), book.sst);
    const lab = [...cells.values()].find((x) => x.c <= 3 && /폐기물량\s*\(톤\)/.test(String(x.v ?? '')));
    if (!lab) continue;
    const blk = KC.monthBlocks(cells).filter((b) => b.row < lab.r && b.first > lab.c).sort((a, b) => b.row - a.row || a.first - b.first)[0];
    if (!blk) continue;
    const tons = {}, items = {};
    for (const [m, c] of Object.entries(blk.cols)) {
      const v = cells.get(lab.r + ',' + c)?.v;
      if (typeof v === 'number') tons[m] = Math.round(v * 1000) / 1000;
      // 바로 아래 품목 행(C열 품목명)
      for (let r = lab.r + 1; r <= lab.r + 10; r++) {
        const nm = cells.get(r + ',' + (lab.c + 1))?.v, iv = cells.get(r + ',' + c)?.v;
        if (!nm || typeof nm !== 'string' || cells.get(r + ',' + lab.c)?.v) break;
        if (typeof iv === 'number') (items[m] ||= {})[nm] = Math.round(iv * 1000) / 1000;
      }
    }
    return { sheet: s.name, tons, items };
  }
  return null;
}
// 품목별 폐기물량 증감 상위 2개 (폐기 파일을 불러온 경우)
const wasteItemsLine = (k, mn) => { const it = (k.notes || {}).wasteItems || {}, c = it[mn], p = it[mn - 1] || {}; if (!c) return ''; const d = Object.keys(c).map((n) => ({ n, d: (c[n] || 0) - (p[n] || 0) })).filter((x) => Math.abs(x.d) >= 0.05).sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 2); return d.length ? ` (${d.map((x) => `${x.n} ${x.d > 0 ? '+' : '-'}${Math.abs(x.d).toFixed(2)}톤`).join(', ')})` : ''; };
const wasteLine = (k, mn) => { const c = wasteOf(k, mn), p = wasteOf(k, mn - 1); if (!c) return ''; return `폐기율 ${(c.rate * 100).toFixed(1)}%${p ? `(${mn - 1}월 ${(p.rate * 100).toFixed(1)}%)` : ''}, 폐기물량 ${c.t.toFixed(1)}톤${p ? `(${mn - 1}월 ${p.t.toFixed(1)})` : ''}${wasteItemsLine(k, mn)}${c.rate <= WASTE_GOAL ? ` — 목표 ${WASTE_GOAL * 100}% 이내 관리` : ` — 목표 ${WASTE_GOAL * 100}%까지 축소 추진`}`; };

/* ---------- 🗓 베트남 KPI 월 마감 체크리스트 (매달 손으로 하는 단계, 사용자 작업 방식 기준) ---------- */
const CLOSE_STEPS = [
  ['결산월 변경', 'Summary·허브(닥터애그베트남삼도결산기준) B3 결산월을 당월로'],
  ['로우데이터 붙여넣기', 'FS(GL·TB·PL·Detail of sales·import), 수불 원가계산용, 회차별 수입제비용, 폐기'],
  ['허브 당월 열 보완', '이마트 블록·판매비·관리비 당월 열 수식 이어 붙이기', 'chan'],
  ['aT 지원금 입력', '월별조정 운반·라벨·창고·인건비 (aT 정산서 RD 기준)', 'at'],
  ['정산 하드코딩 확인', '운반비·라벨비·창고·숙소·인원수, 시식급여(TB 월 열 참조)'],
  ['Target·detail 시트 월 열', "Target 3 PL·detail 품목별·비적격 64289 당월 열 추가", 'tgt'],
  ['장표 시트(PPT (2)) 월 열·% Ach.', '장표 표에 당월 열 추가, BGT·Target 당월 값과 % Ach. 수식 갱신', 'ppt'],
  ['GL 수기 태그', 'GL(원장) Q·R·T열 태그'],
  ['허브 Check = 0', '허브 시트 당월 Check 0 확인'],
];
function closeCard(k) {
  const m = k.month, done = (k.notes || {})[`close${m}`] || {}, C = kpiChecks(k);
  const has = (re) => C.find((x) => re.test(x.txt));
  const sub = subsidyRows(k)[0];
  const auto = {
    chan: () => { const x = has(/거래처별 매출 합계/); return x ? (x.lv === 'ok' ? '✅ 거래처 합계 = 매출액' : '⚠️ 거래처 칸 비어 있음') : ''; },
    at: () => (sub ? (Math.abs(sub.tot.a[m - 1] || 0) >= 0.05 ? `✅ ${m}월 aT 지원금 ${f1(Math.abs(sub.tot.a[m - 1]))}백만원 반영` : `⚠️ ${m}월 aT 지원금 0`) : ''),
    tgt: () => (has(/Target PL|Target … PL|실적이 비어/) ? '⚠️ 당월 실적 비어 있음' : '✅ 당월 실적 있음'),
    ppt: () => (has(/열이 없음|% Ach\.가 .*안 맞음/) ? '⚠️ 월 열 또는 % Ach. 확인 필요' : '✅ 당월 열·% Ach. 정상'),
  };
  const n = CLOSE_STEPS.filter((_, j) => done[j]).length;
  return `<details class="card kpi-chk kpi-close" ${ui.closeOpen ? 'open' : ''}><summary><b>🗓 ${m}월 KPI 마감 체크리스트</b> <small class="muted">${n}/${CLOSE_STEPS.length} 완료</small></summary>
    <ul class="kchk">${CLOSE_STEPS.map(([t, d, a], j) => { const s = a ? auto[a]() : ''; return `<li class="${done[j] ? 'done' : ''}"><label><input type="checkbox" data-close-i="${j}" ${done[j] ? 'checked' : ''}><span><b>${esc(t)}</b> <small class="muted">${esc(d)}</small>${s ? `<br><small>${esc(s)}</small>` : ''}</span></label></li>`; }).join('')}</ul></details>`;
}

/* ---------- 🔎 KPI 파일 점검: 불러온 파일에서 숫자가 안 맞거나 빠진 곳을 찾아 알려줌 ---------- */
function kpiChecks(k) {
  const out = [], m = k.month, j = m - 1, M = model(k);
  const add = (lv, txt) => out.push({ lv, txt });
  // ① 거래처별 매출 합계 = 매출액
  const tops = k.rows.filter((r) => r.top && /^매출액/.test(r.label));
  if (tops.length > 1) {
    const chans = k.rows.filter((r) => !r.top && r.grp === tops[1].grp);
    const fix = (k.chanFix || []).filter((x) => x.m === m);
    const orig = chans.reduce((s, r) => s + (r.a[j] || 0), 0) - fix.reduce((s, x) => s + x.v, 0);
    const gap = (tops[0].a[j] || 0) - orig;
    if (gap >= 0.5) add('bad', `Summary 거래처별 매출 합계가 매출액보다 ${f1(gap)}백만원 적음${fix.length ? ` — 비어 있던 ${fix.map((x) => x.label).join('·')} 칸을 원장(GL)으로 채워 사용 중` : ''}. Summary 거래처 칸(허브 시트 연결) 확인 필요`);
    else add('ok', `거래처별 매출 합계 = 매출액 (${f1(tops[0].a[j] || 0)}백만원)`);
  }
  // ② 품목별 매출 합계 = 매출액
  if (M.items.length && M.sales) { const g = (M.sales.a[j] || 0) - M.items.reduce((s, x) => s + (x.a[j] || 0), 0); if (Math.abs(g) >= 0.5) add('bad', `품목별 매출 합계가 매출액과 ${f1(g)}백만원 차이`); }
  // ③ 장표 표(PPT 시트)에 당월 열이 있는지
  const P = k.ppt;
  if (!P) add('warn', "경영회의 장표 표('PPT…' 시트)를 찾지 못함");
  else if (P.months[P.months.length - 1] < m) add('bad', `'${P.sheet}' 시트 장표 표에 ${m}월 열이 없음 (${P.months[P.months.length - 1]}월까지) — 사이트는 Target PL 시트로 채워 보여 주는 중. 시트에 ${m}월 열 추가 필요`);
  else {
    // ④ 장표 표 % Ach.가 당월 숫자로 계산돼 있는지
    // % 행(% of revenue·% margin)도 포함 — 시트에서 나눗셈 수식을 아래로 복사하면 %p 행이 망가짐
    let par = '';
    const bad = P.rows.filter((r) => !r.h).flatMap((r) => { if (r.lv === 1 && r.kind !== 'pct') par = r.label; return [['BGT', r.bgt, r.bAch], ['Target', r.tgt, r.tAch]].filter(([, base, sheet]) => base && Math.abs(achOf(r, base) - (sheet || 0)) > 0.015).map(([nm, base, sheet]) => ({ r, t: `${r.kind === 'pct' ? `${par} ${r.label}` : r.label} ${nm} ${fAch(sheet || 0, r.kind)} → ${fAch(achOf(r, base), r.kind)}` })); });
    if (bad.length) {
      const pct = bad.some((x) => x.r.kind === 'pct'), eb = bad.some((x) => /EBITDA/i.test(x.r.label) && x.r.kind !== 'pct'), plain = bad.some((x) => x.r.kind !== 'pct' && !/EBITDA/i.test(x.r.label));
      const how = [plain && `일반 행은 =IFERROR(당월실적/기준,)`, pct && `% of revenue·% margin 행은 =당월실적−기준 (%p)`, eb && `EBITDA·조정 EBITDA 행은 =-IFERROR(당월실적/기준,)+1`].filter(Boolean).join(', ');
      add('bad', `'${P.sheet}' 시트 % Ach.가 ${m}월 숫자와 안 맞음 ${bad.length}곳 (예: ${bad.slice(0, 3).map((x) => x.t).join(', ')}) — 사이트 표는 다시 계산한 값으로 표시. 시트 수식: ${how}`);
    }
    else add('ok', `장표 표 % Ach.가 ${m}월 숫자와 일치`);
    // ⑤ Summary ↔ 장표 표 주요 숫자
    const pi = P.months.indexOf(m);
    const pairs = [['매출액', M.sales, /^매출액/], ['매출물량', M.qty, /^매출물량/], ['매출원가', M.cogs, /^매출원가/], ['GP', M.gp, /^GP$/i], ['판매비', M.sell, /^판매비/], ['CM', M.cm, /^CM$/i], ['관리비', M.admin, /^관리비/], ['조정 EBITDA', M.ebitda, /^조정\s*EBITDA/i]];
    const diff = pairs.map(([nm, sr, re]) => { const b = pptBlock(P, re); if (!sr || !b || pi < 0) return null; const d = (b.row.a[pi] || 0) - (sr.a[j] || 0); return Math.abs(d) >= 0.05 ? `${nm} 장표 ${f1(b.row.a[pi])} / Summary ${f1(sr.a[j] || 0)}` : null; }).filter(Boolean);
    if (diff.length) add('bad', `장표 표와 Summary 숫자가 다름: ${diff.join(', ')}`);
    else add('ok', `장표 표 = Summary (매출·원가·GP·판매비·관리비·EBITDA)`);
  }
  // ⑥ Target PL 시트에 당월 실적이 있는지
  if (!k.target) add('warn', "'Target … PL' 시트를 찾지 못해 Target 대비를 볼 수 없음");
  else { const s = k.target.rows.find((r) => /^매출액/.test(r.label)); if (s && !(Math.abs(s.a[j] || 0) > 0)) add('bad', `'${k.target.sheet}' 시트에 ${m}월 실적이 비어 있음`); }
  // ⑦ 원장(GL) 매출 = Summary 매출액
  const gs = k.gl?.acc && Object.entries(k.gl.acc).filter(([a]) => /^511/.test(a)).reduce((s, [, v]) => s + (v[j] || 0), 0);
  if (gs && M.sales) { const d = krwM(gs) - (M.sales.a[j] || 0); if (Math.abs(d) >= 0.5) add('warn', `원장(GL) 매출 ${f1(krwM(gs))}백만원과 Summary 매출액 ${f1(M.sales.a[j] || 0)}백만원이 ${f1(Math.abs(d))} 차이`); }
  return out;
}
function checksCard(k) {
  const C = kpiChecks(k), bad = C.filter((x) => x.lv !== 'ok');
  const ic = { bad: '⚠️', warn: 'ℹ️', ok: '✅' };
  return `<details class="card kpi-chk${bad.length ? ' has-bad' : ''}" ${bad.length ? 'open' : ''}><summary><b>🔎 ${k.month}월 KPI 파일 점검</b> <small class="muted">${bad.length ? `확인 필요 ${bad.length}건` : '이상 없음'}</small></summary>
    <ul class="kchk">${[...bad, ...C.filter((x) => x.lv === 'ok')].map((x) => `<li class="${x.lv}">${ic[x.lv]} ${esc(x.txt)}</li>`).join('')}</ul></details>`;
}

/* ---------- 📌 지난달 보고 후속 확인: 지난달 경영회의 PPT 베트남 장표에서 "예정·목표·추진·협의…" 문장을 뽑아 체크리스트로 ---------- */
const unx = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
async function readPrevPpt(file) {
  const { default: JSZip } = await import('./vendor/jszip.mjs');
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const names = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => +a.match(/\d+/)[0] - +b.match(/\d+/)[0]);
  for (const n of names) {
    const xml = await zip.file(n).async('string');
    const plain = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('').replace(/\s/g, '');
    if (!/베트남법인/.test(plain) || !/실적Issue/i.test(plain)) continue;
    const body = xml.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g, '');
    const paras = [...body.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((m) => ({ t: unx([...m[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((x) => x[1]).join('')).replace(/\s+/g, ' ').trim(), lv: +((m[1].match(/<a:pPr[^>]*lvl="(\d)"/) || [])[1] || 0) })).filter((x) => x.t.length > 1);
    const mon = +((plain.match(/베트남법인(\d{1,2})월실적Issue/) || [])[1] || 0);
    // 장표 오른쪽 폐기 표(1월…N월 / 폐기물량(톤) / 폐기율(%))에서 월별 폐기물량을 읽음
    const waste = {};
    for (const tb of xml.match(/<a:tbl>[\s\S]*?<\/a:tbl>/g) || []) {
      const rows = [...tb.matchAll(/<a:tr[^>]*>([\s\S]*?)<\/a:tr>/g)].map((r) => [...r[1].matchAll(/<a:tc[^>]*>([\s\S]*?)<\/a:tc>/g)].map((c) => [...c[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((x) => x[1]).join('').trim()));
      const h = rows.find((r) => r.filter((c) => /^\d{1,2}월$/.test(c)).length >= 3), w = rows.find((r) => /폐기물량/.test(r[0] || '')), rt = rows.find((r) => /폐기율/.test(r[0] || ''));
      if (!h || !w) continue;
      h.forEach((c, i) => {
        const mm = /^(\d{1,2})월$/.exec(c); const v = parseFloat(String(w[i] || '').replace(/,/g, '')); const r = rt ? parseFloat(String(rt[i] || '').replace(/[,%]/g, '')) : NaN;
        if (mm && Number.isFinite(v)) waste[+mm[1]] = Number.isFinite(r) ? { t: v, r: r / 100 } : v; // 장표에 적힌 폐기율도 같이 (톤 반올림 오차 방지)
      });
    }
    return { slide: n, mon, paras, waste };
  }
  return null;
}
const FU_RE = /(예정|목표|전망|추진|검토|협의|진행\s*중|계획|연기|확대|도입|축소|일시)/;
function prevItems(paras) {
  const out = [];
  let parent = '';
  for (const p of paras) {
    if (p.lv === 0) { parent = ''; continue; } // 제목줄(Target·계획 대비)은 제외
    if (p.lv === 1) parent = /:\s*$/.test(p.t) ? p.t.replace(/\s*:\s*$/, '') : '';
    if (!FU_RE.test(p.t) || p.t.length < 6) continue;
    out.push({ t: (p.lv >= 2 && parent ? `${parent} › ` : '') + p.t, done: false, memo: '' });
  }
  return out;
}
// 이번 달 숫자로 자동 판정(가능한 것만): ✅ 이행 / ⚠️ 확인 필요
function followHint(k, text, mn) {
  const P = pptOf(k), M = model(k), i = mn - 1, p = mn - 2;
  if (!P || p < 0) return '';
  const pi = P.months.indexOf(mn), pp = P.months.indexOf(mn - 1);
  const kid = (re, lab) => { const b = pptBlock(P, re); return b?.kids.find((x) => lab.test(x.label)); };
  const kv = (x) => (x && pi >= 0 ? [x.a[pi], x.a[pp]] : null);
  const ch = (re) => M.chans.find((x) => re.test(x.label));
  const num = (s) => { const m = [...s.matchAll(/(\d+(?:\.\d+)?)\s*백만원/g)]; return m.length ? +m[m.length - 1][1] : null; };
  const chk = (ok, s) => `${ok ? '✅' : '⚠️'} ${s}`;
  for (const [re, nm] of [[/이마트/, '이마트'], [/롯데/, '롯데마트']]) {
    if (!re.test(text) || !/매출/.test(text)) continue;
    const c = ch(re); if (!c) continue;
    const goal = num(text), v = c.a[i] || 0;
    return chk(goal ? v >= goal : v >= (c.a[p] || 0), `${mn}월 ${nm} 매출 ${f1(v)}백만원${goal ? ` (목표 ${goal}, ${Math.round(v / goal * 100)}%)` : ` (전월 ${f1(c.a[p] || 0)})`}`);
  }
  if (/coop|쿱|윈마트|win/i.test(text)) {
    const c = ch(/coop|쿱|윈|win/i), v = c ? c.a[i] || 0 : 0;
    return v >= 0.05 ? chk(true, `${mn}월 ${c.label} 매출 ${f1(v)}백만원`) : chk(false, `${mn}월 신규 채널 매출 없음 — 납품 일정·협의 결과 확인`);
  }
  if (/사무실|법률|미청구|합산/.test(text)) {
    const a = kv(kid(/^관리비/, /사무실/)), b = kv(kid(/^관리비/, /^기타$/));
    if (a) return chk(a[0] < a[1], `사무실임차료 ${f1(a[0])}백만원 (전월 ${f1(a[1])})${b ? `, 기타(법률 등) ${f1(b[0])} (전월 ${f1(b[1])})` : ''} — 일시 증가분 해소 여부`);
  }
  if (/숙소/.test(text)) { const v = kv(kid(/^관리비/, /숙소/)); if (v) return chk(v[0] < 0.05, `숙소임차료 ${f1(v[0])}백만원 (전월 ${f1(v[1])})`); }
  if (/차량|여비/.test(text)) { const v = kv(kid(/^관리비/, /여비/)); if (v) return chk(v[0] < v[1], `여비교통비 ${f1(v[0])}백만원 (전월 ${f1(v[1])}, ${v[0] - v[1] >= 0 ? '+' : ''}${f1(v[0] - v[1])})`); }
  if (/폐기/.test(text)) { const v = kv(kid(/^단위원가/, /폐기/)); if (v) return chk(v[0] <= v[1], `kg당 폐기원가 ${Math.round(v[0])}원 (전월 ${Math.round(v[1])}) — 폐기율은 폐기 물량으로 확인`); }
  if (/운반|합배송|하차지|운송/.test(text)) {
    const x = kid(/^판매비/, /운반/), vol = pptBlock(P, /^매출물량/);
    if (x && vol && pi >= 0) { const c = x.a[pi] * 1000 / vol.row.a[pi], b = x.a[pp] * 1000 / vol.row.a[pp]; return chk(c <= b, `kg당 운반비 ${Math.round(c)}원 (전월 ${Math.round(b)})`); }
  }
  if (/인센티브|판매사원|시식|진열/.test(text)) { const v = kv(kid(/^판매비/, /인건비/)); if (v) return `ℹ️ 판매 인건비 ${f1(v[0])}백만원 (전월 ${f1(v[1])}) — 효과는 현지 확인`; }
  return '';
}
// 지난달 PPT를 안 넣어도: 사이트에 저장된 지난달 코멘트(A~E)의 "예정·목표…" 문장을 이번 달 후속 목록으로
function autoFollow(k, mn) {
  const texts = ['A', 'B', 'C', 'D', 'E'].map((g) => (k.notes || {})[`cm${mn - 1}${g}`]).filter(Boolean);
  if (!texts.length) return null;
  const paras = texts.flatMap((t) => t.split('\n').filter((l) => l.trim()).map((l) => ({ t: l.replace(/^\s*-\s*/, '').trim(), lv: /^\s*-/.test(l) ? 1 : 0 })));
  const items = prevItems(paras);
  return items.length ? { src: `업무노트 ${mn - 1}월 코멘트`, mon: mn - 1, items, auto: 1 } : null;
}
function followCard(k, mn) {
  const fu = (k.notes || {})[`fu${mn}`] || autoFollow(k, mn);
  const items = fu?.items || [];
  const left = items.filter((x) => !x.done).length;
  // 가로로 긴 표: 확인 | 지난달 보고 내용 | 이번 달 자동 판정 | 결과·현지 답변 메모
  const rows = items.map((x, j) => { const h = followHint(k, x.t, mn); return `<li class="fu-it${x.done ? ' done' : ''}"><label class="fu-c"><input type="checkbox" data-fu-i="${j}" ${x.done ? 'checked' : ''} aria-label="확인"></label><div class="fu-t">${esc(x.t)}</div><div class="fu-h">${h ? esc(h) : '<span class="muted">수치 없음 — 직접 확인</span>'}</div><input class="fu-memo" data-fu-m="${j}" value="${esc(x.memo || '')}" placeholder="${mn}월 결과·현지 답변 메모"></li>`; }).join('');
  const head = items.length ? `<li class="fu-it fu-head"><div class="fu-c">확인</div><div class="fu-t">지난달 보고 내용 (예정·목표·추진)</div><div class="fu-h">${mn}월 숫자로 자동 판정</div><div>결과·현지 답변 메모</div></li>` : '';
  return `<details class="cm-ask fu" ${items.length ? 'open' : ''}><summary>📌 지난달 보고 후속 확인 ${items.length ? `(${items.length - left}/${items.length} 확인)` : ''}</summary>
    <p class="muted sm">${fu ? `${fu.auto ? `지난달 업무노트에 쓴 ${fu.mon}월 코멘트(A~E)` : `「${esc(fu.src)}」${fu.mon ? ` ${fu.mon}월 장표` : ''}`}에서 예정·목표·추진 문장을 뽑았습니다.${fu.auto ? ' 최종 장표 기준으로 바꾸려면 지난달 PPT를 불러오세요.' : ''} 이번 달 숫자로 판정 가능한 것은 ✅/⚠️로 표시합니다.` : '지난달 경영회의 PPT(.pptx)를 불러오면 베트남 장표에서 "~예정 / ~목표 / ~추진 / 협의 중" 문장을 뽑아 이번 달 확인 목록을 만듭니다.'}</p>
    ${items.length ? `<ul class="fu-list">${head}${rows}</ul>` : ''}
    <div class="cm-tools"><button type="button" class="btn sm" id="fuLoad">📂 지난달 PPT 불러오기</button>${items.length ? '<button type="button" class="btn sm" id="fuCopy">📋 미확인 항목 복사(현지 문의용)</button>' : ''}<input type="file" id="fuIn" accept=".pptx" hidden></div></details>`;
}

function renderCmt(k) {
  const P = pptOf(k);
  if (!P) return `<div class="card">⚠️ 이 파일에서 경영회의 장표 표(PPT 시트의 Actual·Budget·Target 표)를 찾지 못했습니다.</div>`;
  const n = P.months.length, mn = P.months[n - 1];
  const head = `<tr><th>(단위: 백만원)</th>${P.months.map((m) => `<th>${m}월</th>`).join('')}<th class="bg">BGT<br><small>${mn}월</small></th><th class="bg">% Ach.</th><th class="tg">Target<br><small>${mn}월</small></th><th class="tg">% Ach.</th></tr>`;
  // 대항목(lv1)만 보이고 세부(lv2)는 ▸ 눌러 펼침 (% 줄은 대항목과 같이 보임)
  const vis = P.rows.filter((r) => !r.h);
  let gi = -1;
  const grp = vis.map((r) => { if (r.lv === 1 && r.kind !== 'pct') gi++; return gi; });
  const hasKids = new Set(vis.filter((r) => r.lv === 2).map((r, x) => grp[vis.indexOf(r)]));
  const body = vis.map((r, x) => {
    const key = r.lv === 1 && r.kind !== 'pct', kid = r.lv === 2, g = grp[x], open = cmOpen.has(g);
    const tg = key && hasKids.has(g) ? `<button type="button" class="cm-tg" data-cmg="${g}" aria-expanded="${open}">${open ? '▾' : '▸'}</button>` : key ? '<span class="cm-tg-sp"></span>' : '';
    return `<tr class="${key ? 'key' : 'sub'}${r.kind === 'pct' ? ' pct' : ''}${kid ? ' cm-kid' : ''}"${kid ? ` data-cmk="${g}"${open ? '' : ' hidden'}` : ''}><th>${tg}${esc(r.label)}</th>${r.a.map((v, j) => `<td class="${j === n - 1 ? 'hl' : ''} ${v < 0 ? 'neg' : ''}">${fP(v, r.kind)}</td>`).join('')}<td>${fP(r.bgt, r.kind)}</td><td class="muted">${fAch(achOf(r, r.bgt), r.kind)}</td><td class="tg">${fP(r.tgt, r.kind)}</td><td class="muted">${fAch(achOf(r, r.tgt), r.kind)}</td></tr>`;
  }).join('');
  const allOpen = hasKids.size && [...hasKids].every((g) => cmOpen.has(g));
  const table = `<div class="cm-tbar"><button type="button" class="btn sm ghost" id="cmAllTg">${allOpen ? '▴ 세부 모두 접기' : '▾ 세부 모두 펼치기'}</button><span class="muted sm">▸ 를 누르면 세부 항목이 펼쳐집니다</span></div><div class="kt-scroll"><table class="rt cm-t"><thead>${head}</thead><tbody>${body}</tbody></table></div><p class="muted sm">출처: '${esc(P.sheet)}' 시트 ${P.row}행 표 (경영회의 장표와 같은 표)</p>${P.auto ? `<p class="sm cm-auto-note">⚠️ 이 파일의 '${esc(P.sheet)}' 시트에는 ${P.auto.map((x) => `${x}월`).join('·')} 열이 아직 없어, '${esc(k.target?.sheet || 'Target PL')}' 시트의 월 실적으로 같은 규칙에 따라 채웠습니다. BGT는 Summary 시트 값(세부 항목 BGT는 비움), Target은 장표 값 그대로입니다. 시트에 ${P.auto[P.auto.length - 1]}월 열을 추가하면 그 값을 씁니다.</p>` : ''}`;
  const notes = k.notes || {};
  const blocks = CM_GROUPS.map((G) => {
    const B = pptBlock(P, G.re);
    if (!B) return '';
    const facts = cmFacts(k, G, B);
    const { accs, lines } = cmGl(k, G);
    const key = `cm${mn}${G.id}`;
    const glHtml = accs.length ? `<details class="cm-gl"><summary>📒 원장 확인 (${k.gl.sheet}, VND)</summary><table class="rt"><tr><th>계정</th><th>${mn - 1}월</th><th>${mn}월</th><th>증감</th><th>≈백만원</th></tr>${accs.map((x) => `<tr><th>${x.a} <small class="muted">${esc(accName(x.a))}</small></th><td>${(x.prev / 1e6).toFixed(1)}M</td><td>${(x.cur / 1e6).toFixed(1)}M</td><td class="${x.cur - x.prev > 0 ? 'neg' : ''}">${x.cur - x.prev > 0 ? '+' : ''}${((x.cur - x.prev) / 1e6).toFixed(1)}M</td><td>${sgn(krwM(x.cur - x.prev), 1)}</td></tr>`).join('')}</table>
      ${lines.length ? `<p class="muted sm">${mn}월 주요 거래 (변동 큰 계정)</p><ul class="cm-lines">${lines.map((l) => `<li><b>${l.a}</b> ${(l.v / 1e6).toFixed(1)}M · ${esc(l.d)}${l.o ? ` <small class="muted">${esc(l.o)}</small>` : ''}</li>`).join('')}</ul>` : ''}</details>` : '';
    return `<div class="cm-block" data-g="${G.id}"><div class="cm-h"><span class="rp-no">${G.id}</span><b>${G.name}</b><button type="button" class="btn ghost sm" data-cm-fill="${G.id}">↺ 자동 문구</button></div>
      <div class="cm-head">${esc(cmHeadline(G, B))}</div>
      ${facts.length ? `<ul class="cm-facts">${facts.map((f) => `<li>${esc(f)}</li>`).join('')}</ul>` : ''}
      ${glHtml}
      <textarea class="cm-text${notes[key] ? '' : ' auto'}" data-id="${k.id}" data-m="${key}" rows="6" placeholder="현지 확인 내용을 넣어 코멘트를 쓰세요 (자동 저장)">${esc(notes[key] || cmAutoText(k, G.id))}</textarea>
      ${notes[key] ? '' : '<p class="muted sm">✍️ 장표 형식 자동 초안입니다. [사유] 자리를 현지 확인 내용으로 바꿔 쓰면 저장됩니다.</p>'}</div>`;
  }).join('');
  const qs = cmQuestions(k);
  const ask = `<details class="cm-ask"><summary>❓ 현지 문의 초안 (${qs.length}건)</summary>${qs.length ? `<ol>${qs.map((x) => `<li>${esc(x)}</li>`).join('')}</ol><button type="button" class="btn sm" id="cmAskCopy">📋 문의 목록 복사</button>` : '<p class="muted">큰 원장 변동이 없습니다.</p>'}</details>`;
  return `<article class="rp cm"><header class="rp-head"><h2>📝 ${k.year}년 ${mn}월 경영회의 베트남 장표 코멘트</h2><p>왼쪽 표는 장표와 같은 숫자, 오른쪽은 A~E 코멘트 초안입니다. 숫자 근거와 원장을 보고 현지 확인 내용을 채운 뒤 복사해 장표에 붙이세요.</p></header>
    <div class="cm-grid"><div class="cm-left">${table}${wasteCard(k, mn)}</div><div class="cm-right">
      <div class="cm-tools">${aiEnabled() ? '<button type="button" class="btn sm primary" id="cmAi">🤖 AI 초안 만들기</button>' : ''}<button type="button" class="btn sm" id="cmCopy">📋 코멘트 전체 복사</button><span class="muted sm" id="cmMsg">${esc(cmStatus)}</span></div>
      ${ask}${blocks}</div></div>${followCard(k, mn)}</article>`;
}
function cmAutoText(k, gid) {
  const P = pptOf(k), G = CM_GROUPS.find((x) => x.id === gid), B = P && pptBlock(P, G.re);
  if (!B) return '';
  return [cmHeadline(G, B), ...cmBullets(k, G, B).map((f) => `  - ${f}`)].join('\n');
}
let cmStatus = '';
const cmOpen = new Set(); // 펼친 대항목 번호 (화면을 다시 그려도 유지)
async function cmAi(k) {
  const msg = $('#cmMsg');
  if (!aiEnabled()) { msg.textContent = '⚙️ 설정에서 Claude API 키를 먼저 넣어주세요.'; return; }
  const P = pptOf(k), mn = P.months[P.months.length - 1];
  const parts = CM_GROUPS.map((G) => {
    const B = pptBlock(P, G.re);
    if (!B) return '';
    const { accs, lines } = cmGl(k, G);
    return `[${G.id} ${G.name}]\n제목줄: ${cmHeadline(G, B).replace(/\n/g, ' ')}\n근거 숫자:\n${cmFacts(k, G, B).map((f) => `- ${f}`).join('\n')}\n원장 계정 증감(VND): ${accs.map((x) => `${x.a}${accName(x.a) ? `(${accName(x.a)})` : ''} ${(x.prev / 1e6).toFixed(1)}M→${(x.cur / 1e6).toFixed(1)}M`).join(', ') || '-'}\n원장 주요 거래: ${lines.map((l) => `${l.a} ${(l.v / 1e6).toFixed(1)}M ${l.d}`).join(' | ') || '-'}\n현지 확인 메모(사용자가 쓴 내용): ${(k.notes || {})[`cm${mn}${G.id}`] || '-'}`;
  }).filter(Boolean).join('\n\n');
  const schema = { type: 'object', additionalProperties: false, required: CM_GROUPS.map((g) => g.id), properties: Object.fromEntries(CM_GROUPS.map((g) => [g.id, { type: 'string' }])) };
  msg.textContent = 'AI가 초안을 쓰는 중…';
  try {
    const out = await aiJson({
      system: '당신은 한국 식품회사 베트남 법인의 월간 경영회의 장표 코멘트를 쓰는 재무 담당자입니다. 각 항목(A~E)마다 한국어 장표 문구를 씁니다. ' +
        '형식: 첫 줄은 주어진 제목줄 그대로(숫자 바꾸지 말 것, 계획 대신 BGT 표기 가능), 그 아래 2~4줄은 "항목명: 내용" 형태의 짧은 개조식(명사형 종결, 예: "~감소", "~예정"). ' +
        '원인은 근거 숫자·원장 거래·현지 확인 메모에 있는 것만 쓰고, 근거가 없으면 "(현지 확인 필요: 무엇을)"이라고 남기세요. 금액은 백만원·원/kg 단위로, 원장 VND는 18.5로 나눠 원화로 환산해 쓰세요. 빈 항목은 빈 문자열.',
      prompt: `${k.year}년 ${mn}월 베트남 법인 실적 장표 코멘트 초안을 써주세요.\n\n${parts}`,
      schema, maxTokens: 4000,
    });
    const notes = { ...(store.find('kpi', k.id)?.notes || {}) };
    for (const G of CM_GROUPS) if (out[G.id]) notes[`cm${mn}${G.id}`] = out[G.id].trim();
    store.upsert('kpi', { id: k.id, notes });
    cmStatus = '✅ AI 초안을 넣었습니다. 현지 확인 후 고쳐 쓰세요.'; ($('#cmMsg') || {}).textContent = cmStatus;
  } catch (e) { cmStatus = '⚠️ ' + (e.message || e); ($('#cmMsg') || {}).textContent = cmStatus; }
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
  // BGT
  const yb = sumR(M.sales.b, 1, n), ya = sumR(S, 1, n);
  if (yb > 0) add(ya >= yb ? 'good' : 'bad', `누계 매출 BGT 달성률 ${Math.round(ya / yb * 100)}% (BGT ${f1(yb)} / 실적 ${f1(ya)}백만원)`);
  if (M.ebit) { const ea = sumR(M.ebit.a, 1, n), eb = sumR(M.ebit.b, 1, n); if (eb) add(ea >= eb ? 'good' : 'bad', `누계 EBIT ${f1(ea)}백만원, BGT(${f1(eb)}) 대비 ${ea - eb > 0 ? '+' : ''}${f1(ea - eb)}`); }
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
  const per = n > 1 ? `1~${n}월` : '1월';
  const i = n - 1, p = n - 2;
  const Q = quarters(n);
  const S = ytd(M.sales);
  const sec = (no, title, sub, body) => `<section class="rp-sec"><h3><span class="rp-no">${no}</span>${title}${sub ? `<small>${sub}</small>` : ''}</h3>${body}</section>`;
  const cell = (v, F, cls = '') => `<td class="${v < 0 ? 'neg ' : ''}${cls}">${F(v)}</td>`;
  const achv = (a, b) => (b > 0 && a >= 0 ? `${Math.round(a / b * 100)}%` : '-'); // 적자·음수 BGT는 달성률 없음

  // 손익 항목 목록 (누계 비율·단가는 합계로 다시 계산)
  const ratioOf = (num) => (w, s, e) => div(sumR(num[w], s, e), sumR(M.sales[w], s, e));
  const priceOf = (it, q) => (w, s, e) => { const t = q ? sumR(q[w], s, e) : 0; return t ? sumR(it[w], s, e) / t * 1000 : 0; };
  const MET = [
    { r: M.sales, name: '매출액', kind: 'amt', cls: 'key' },
    { r: M.qty, name: '매출물량 (톤)', kind: 'ton', cls: 'sub' },
    { r: M.price, name: '매출단가 (원/kg)', kind: 'won', cls: 'sub', agg: M.qty && priceOf(M.sales, M.qty) },
    { r: M.cogs, name: '매출원가', kind: 'cost' },
    { r: M.gp, name: 'GP (매출총이익)', kind: 'amt', cls: 'key' },
    { r: M.gpp, name: 'GP%', kind: 'pct', cls: 'sub', agg: M.gp && ratioOf(M.gp) },
    { r: M.sell, name: '판매비', kind: 'cost' },
    { r: M.cm, name: 'CM (공헌이익)', kind: 'amt', cls: 'key' },
    { r: M.cmp, name: 'CM%', kind: 'pct', cls: 'sub', agg: M.cm && ratioOf(M.cm) },
    { r: M.admin, name: '관리비', kind: 'cost' },
    { r: M.ebit, name: 'EBIT (영업이익)', kind: 'amt', cls: 'key' },
    { r: M.ebitp, name: 'EBIT%', kind: 'pct', cls: 'sub', agg: M.ebit && ratioOf(M.ebit) },
    M.adj && ytd(M.adj) ? { r: M.adj, name: '조정사항', kind: 'amt', cls: 'sub' } : null,
    { r: M.ebitda, name: '조정 EBITDA', kind: 'amt', cls: 'key' },
  ].filter((x) => x && x.r);
  const aggV = (x, w, s, e) => (x.agg ? x.agg(w, s, e) : x.kind === 'pct' || x.kind === 'won' ? 0 : sumR(x.r[w], s, e));

  // Ⅰ. 핵심 요약
  const tile = (r, name, ratioRow) => {
    if (!r) return '';
    const a = ytd(r), b = ytd(r, 'b');
    return `<div class="rp-tile"><div class="t">${name}</div><div class="v ${a < 0 ? 'neg' : ''}">${f1(a)}</div>
      ${ratioRow ? `<div class="s">${ratioRow}</div>` : ''}<div class="s muted">월평균 ${f1(a / n)}</div>
      ${b ? `<div class="s b">BGT ${f1(b)}${achv(a, b) !== '-' ? ` · 달성 <b>${achv(a, b)}</b>` : ''}</div><div class="s">${delta(a, b, 'amt')}</div>` : ''}</div>`;
  };
  const tiles = tile(M.sales, '매출액', M.qty ? `${f1(ytd(M.qty))}톤` : '')
    + tile(M.gp, 'GP', M.gp ? `GP% <b>${pc(div(ytd(M.gp), S))}</b>` : '')
    + tile(M.cm, 'CM', M.cm ? `CM% <b>${pc(div(ytd(M.cm), S))}</b>` : '')
    + tile(M.ebit, 'EBIT', M.ebit ? `EBIT% <b>${pc(div(ytd(M.ebit), S))}</b>` : '')
    + tile(M.ebitda, '조정 EBITDA', '');
  const fs = [...targetFindings(k, n), ...periodFindings(M, n)];
  const grp = (lv, title) => { const l = fs.filter((f) => f.lv === lv); return l.length ? `<div class="rp-find ${lv}"><b>${title}</b><ul>${l.map((f) => `<li>${esc(f.txt.replace(/^⚠️\s*/, ''))}</li>`).join('')}</ul></div>` : ''; };
  const key = `ytd${n}`;
  const s1 = sec('Ⅰ', '핵심 요약', `${per} 누계 · 단위 백만원`, `<div class="rp-tiles">${tiles}</div>
    ${grp('good', '✅ 개선된 점')}${grp('bad', '⚠️ 부진·점검 필요')}${grp('info', 'ℹ️ 참고')}
    <label class="kn rp-memo"><span class="muted sm">✏️ 보고 메모 (자동 저장 · 인쇄에 포함)</span><textarea id="kpiNote" data-id="${k.id}" data-m="${key}" rows="4" placeholder="예: 3Q 이마트 입점으로 매출 확대, 관리비 절감 효과">${esc((k.notes || {})[key] || '')}</textarea></label>`);

  // Ⅱ. Target 대비 (가장 중요: 당월·최근 3개월 vs Target 시나리오)
  const sT = k.target ? sec('Ⅱ', 'Target 대비', `${n}월 실적과 최근 3개월 평균을 월 기준 Target과 비교`, targetTable(k, n)) : '';

  // Ⅲ. 손익 요약 (누계 vs BGT, 당월 vs 전월)
  const s2rows = MET.map((x) => {
    const F = fmtOf(x.kind), a = aggV(x, 'a', 1, n), b = aggV(x, 'b', 1, n);
    const cur = x.r.a[i], prev = p >= 0 ? x.r.a[p] : null;
    return `<tr class="${x.cls || ''}"><th>${x.name}</th>${cell(a, F, 'hl')}<td>${b ? F(b) : '-'}</td><td>${b ? delta(a, b, x.kind) : ''}</td><td>${x.kind === 'pct' || x.kind === 'won' ? '-' : achv(a, b)}</td>
      <td>${x.kind === 'pct' || x.kind === 'won' ? F(a) : F(a / n)}</td>${cell(cur, F)}<td>${prev !== null ? delta(cur, prev, x.kind) : ''}</td></tr>`;
  }).join('');
  const s2 = sec('Ⅲ', '손익 요약', '누계는 BGT와, 당월은 전월과 비교', `<div class="kt-scroll"><table class="rt">
    <thead><tr><th>항목</th><th>${per} 누계</th><th>누계 BGT</th><th>BGT 대비</th><th>달성률</th><th>월평균</th><th>${n}월</th><th>전월 대비</th></tr></thead><tbody>${s2rows}</tbody></table></div>`);

  // Ⅲ. 월별 손익 추이
  const mcols = Array.from({ length: n }, (_, j) => j);
  const s3rows = MET.map((x) => { const F = fmtOf(x.kind); return `<tr class="${x.cls || ''}"><th>${x.name}</th>${mcols.map((j) => cell(x.r.a[j], F)).join('')}${cell(aggV(x, 'a', 1, n), F, 'hl')}</tr>`; }).join('');
  const s3 = sec('Ⅳ', '월별 손익 추이', '단위 백만원 (물량 톤, 단가 원/kg)',
    barChart({ a: M.sales.a, b: M.sales.b, m: n, cnt: n, sel: 0, title: '월별 매출액 (백만원)', vlabels: true })
    + `<div class="kt-scroll"><table class="rt rt-m"><thead><tr><th>항목</th>${mcols.map((j) => `<th>${j + 1}월</th>`).join('')}<th>누계</th></tr></thead><tbody>${s3rows}</tbody></table></div>`);

  // Ⅳ. 분기별 실적
  let s4 = '';
  if (Q.length >= 2) {
    const qm = MET.filter((x) => x.kind !== 'won' && x.r !== M.adj);
    const qv = (x, q) => (x.kind === 'pct' ? x.agg('a', q.s, q.e) : sumR(x.r.a, q.s, q.e) / (q.e - q.s + 1));
    const rows = qm.map((x) => { const F = fmtOf(x.kind), v = Q.map((q) => qv(x, q)); return `<tr class="${x.cls || ''}"><th>${x.name}</th>${v.map((t) => cell(t, F)).join('')}<td>${delta(v[v.length - 1], v[v.length - 2], x.kind)}</td><td>${delta(v[v.length - 1], v[0], x.kind)}</td></tr>`; }).join('');
    const last = Q[Q.length - 1];
    s4 = sec('Ⅴ', '분기별 실적', `금액은 월평균, 비율은 분기 합계 기준${last.part ? ` · ${last.label}는 ${last.s}~${last.e}월만` : ''}`, `<div class="kt-scroll"><table class="rt">
      <thead><tr><th>항목</th>${Q.map((q) => `<th>${q.label} <small>(${q.s}~${q.e}월)</small></th>`).join('')}<th>직전 분기 대비</th><th>${Q[0].label} 대비</th></tr></thead><tbody>${rows}</tbody></table></div>`);
  }

  // Ⅴ. 품목별 실적
  const q0 = Q[0], qL = Q[Q.length - 1];
  const itemRows = M.items.filter((x) => ytd(x) || x.a[i]).sort((a, b) => ytd(b) - ytd(a)).map((x) => {
    const q = M.qtys.find((y) => y.label === x.label);
    const pr = priceOf(x, q), a = ytd(x), t = q ? ytd(q) : 0;
    const p0 = pr('a', q0.s, q0.e), pL = pr('a', qL.s, qL.e);
    return `<tr><th>${esc(x.label)}</th>${cell(a, f1, 'hl')}<td>${pc(div(a, S))}</td><td>${f1(t)}</td><td>${f0(pr('a', 1, n))}</td>${Q.length >= 2 ? `<td>${f0(p0)} → ${f0(pL)}</td><td>${p0 && pL ? delta(pL, p0, 'won') : ''}</td>` : ''}${cell(x.a[i], f1)}<td>${p >= 0 ? delta(x.a[i], x.a[p], 'amt') : ''}</td></tr>`;
  }).join('');
  const itemTotal = `<tr class="key tot"><th>합계</th>${cell(S, f1, 'hl')}<td>100%</td><td>${M.qty ? f1(ytd(M.qty)) : ''}</td><td>${M.qty ? f0(div(S, ytd(M.qty)) * 1000) : ''}</td>${Q.length >= 2 ? '<td></td><td></td>' : ''}${cell(M.sales.a[i], f1)}<td>${p >= 0 ? delta(M.sales.a[i], M.sales.a[p], 'amt') : ''}</td></tr>`;
  const s5 = M.items.length ? sec('Ⅵ', '품목별 실적', '매출 큰 순서', `<div class="kt-scroll"><table class="rt">
    <thead><tr><th>품목</th><th>누계 매출</th><th>비중</th><th>물량 (톤)</th><th>평균 단가</th>${Q.length >= 2 ? `<th>단가 ${q0.label} → ${qL.label}</th><th>단가 변화</th>` : ''}<th>${n}월 매출</th><th>전월 대비</th></tr></thead><tbody>${itemRows}${itemTotal}</tbody></table></div>`
    + stackChart('월별 품목 구성 (백만원)', M.items, n, M.sales.a)) : '';

  // Ⅵ. 거래처별 실적
  const chs = M.chans.filter((x) => ytd(x) || ytd(x, 'b'));
  const span = (x) => { const on = x.a.slice(0, n).map((v) => v >= 0.05); const f = on.indexOf(true), l = on.lastIndexOf(true); return f < 0 ? '-' : f === l ? `${f + 1}월` : `${f + 1}~${l + 1}월`; };
  const chRows = chs.sort((a, b) => ytd(b) - ytd(a)).map((x) => { const a = ytd(x), b = ytd(x, 'b'); return `<tr><th>${esc(x.label)}</th>${cell(a, f1, 'hl')}<td>${pc(div(a, S))}</td><td>${b ? f1(b) : '-'}</td><td>${achv(a, b)}</td><td>${span(x)}</td>${cell(x.a[i], f1)}<td>${p >= 0 && (x.a[i] || x.a[p]) ? delta(x.a[i], x.a[p], 'amt') : ''}</td></tr>`; }).join('');
  const CT = M.chan ? ytd(M.chan) : 0;
  const s6 = chs.length ? sec('Ⅶ', '거래처별 실적', '매출 큰 순서', `<div class="kt-scroll"><table class="rt">
    <thead><tr><th>거래처</th><th>누계 매출</th><th>비중</th><th>누계 BGT</th><th>달성률</th><th>거래 기간</th><th>${n}월 매출</th><th>전월 대비</th></tr></thead><tbody>${chRows}
    <tr class="key tot"><th>합계</th>${cell(CT, f1, 'hl')}<td>${pc(div(CT, S))}</td><td>${M.chan ? f1(ytd(M.chan, 'b')) : ''}</td><td>${M.chan ? achv(CT, ytd(M.chan, 'b')) : ''}</td><td></td>${cell(M.chan ? M.chan.a[i] : 0, f1)}<td></td></tr></tbody></table></div>`
    + (chs.filter((x) => ytd(x)).length > 1 ? stackChart('월별 거래처 구성 (백만원)', M.chans, n, M.chan?.a) : '')) : '';

  // Ⅶ. 수익성 추이
  const s7 = sec('Ⅷ', '수익성 추이', '', `<div class="rp-charts">${ratioChart('GP% · CM% · EBIT%', [M.gpp && { name: 'GP%', v: M.gpp.a, c: '#1f9d62' }, M.cmp && { name: 'CM%', v: M.cmp.a, c: '#2b5ce6' }, M.ebitp && { name: 'EBIT%', v: M.ebitp.a, c: '#d64545' }].filter(Boolean), n)}
    ${M.ebit ? barChart({ a: M.ebit.a, b: M.ebit.b, m: n, cnt: n, sel: 0, title: 'EBIT 월별 (백만원)', vlabels: true }) : ''}</div>`);

  return `<article class="rp"><header class="rp-head"><h2>${k.year}년 ${per} 실적 보고</h2>
    <p>${esc(k.title)}<br><span class="muted">단위: 백만원 · 자료: ${esc(k.file || '')}</span></p></header>
    ${s1}${sT}${s2}${s3}${s4}${s5}${s6}${s7}</article>`;
}

/* ---------- 화면 ---------- */
function list() { return [...(store.get().kpi || [])].sort((a, b) => b.year - a.year || String(a.title).localeCompare(String(b.title))); }

/* ---------- 💰 지원금 수익 (aT 지원사업 등): 'Target … PL' 시트의 '지원' 항목 + 원장의 지원 관련 수익 ---------- */
function subsidyRows(k) {
  const R = k.target?.rows || [];
  const tot = R.filter((r) => /지원/.test(r.label) && r.lv <= 2 && !/반영\s*전|못받|미적용/.test(r.label) && r.a.some((v) => Math.abs(v) >= 0.005));
  return tot.map((t) => ({ tot: t, kids: R.filter((r) => r.lv > t.lv && r.p === t.label && r.a.some((v) => Math.abs(v) >= 0.005)) }));
}
function subsidyCard(k, m, M) {
  const S = subsidyRows(k), i = m - 1, p = m - 2;
  const glL = (k.gl?.lines || []).filter((l) => l.m === m && /^[57]/.test(l.a) && (/^711/.test(l.a) || /hỗ trợ|support|subsid|tài trợ|\baT\b/i.test(l.d)));
  if (!S.length && !glL.length) return '';
  const ab = (v) => Math.abs(v || 0);
  const first = (r) => r.a.findIndex((v) => ab(v) >= 0.005);
  let body = '';
  for (const { tot, kids } of S) {
    const s0 = Math.max(0, first(tot)), cur = ab(tot.a[i]), prv = p >= 0 ? ab(tot.a[p]) : 0;
    const cum = tot.a.slice(0, m).reduce((x, v) => x + ab(v), 0);
    const sales = M.sales?.a[i] || 0, sell = M.sell?.a[i] || 0, eb = M.ebitda?.a[i];
    const cols = Array.from({ length: m - s0 }, (_, j) => s0 + j);
    const rowH = (r, cls = '') => `<tr class="${cls}"><th>${esc(r.label)}</th>${cols.map((c) => `<td class="${c === i ? 'b' : ''}">${ab(r.a[c]) >= 0.005 ? f1(ab(r.a[c])) : '-'}</td>`).join('')}<td class="b">${f1(r.a.slice(0, m).reduce((x, v) => x + ab(v), 0))}</td><td>${cur && r !== tot ? pc(ab(r.a[i]) / cur) : ''}</td></tr>`;
    const stat = (t, v, sub = '') => `<div class="sb-st"><span class="muted sm">${t}</span><b>${v}</b>${sub ? `<small class="muted">${sub}</small>` : ''}</div>`;
    body += `<div class="sb-stats">
        ${stat(`${m}월 지원금`, `${f1(cur)}<small>백만원</small>`, p >= 0 ? `전월 ${f1(prv)} (${cur - prv >= 0 ? '+' : ''}${f1(cur - prv)})` : '')}
        ${stat(`누계 (${s0 + 1}~${m}월)`, `${f1(cum)}<small>백만원</small>`, `${m - s0}개월 · 월평균 ${f1(cum / Math.max(1, m - s0))}`)}
        ${sales ? stat('매출 대비', pc(cur / sales), `${m}월 매출 ${f1(sales)}백만원`) : ''}
        ${sell ? stat('판매비 지원율', pc(cur / (sell + cur)), `지원 전 판매비 ${f1(sell + cur)} → 지원 후 ${f1(sell)}`) : ''}
        ${eb !== undefined ? stat('지원금 없었다면 조정 EBITDA', `${f1(eb - cur)}<small>백만원</small>`, `실제 ${f1(eb)}`) : ''}
      </div>
      <div class="kt-scroll"><table class="kt"><thead><tr><th>(백만원)</th>${cols.map((c) => `<th>${c + 1}월</th>`).join('')}<th>누계</th><th>${m}월 비중</th></tr></thead>
      <tbody>${rowH(tot, 'key')}${kids.map((r) => rowH(r, 'sub')).join('')}</tbody></table></div>`;
  }
  const what = S.some(({ tot }) => /aT/i.test(tot.label))
    ? '<p class="muted sm">ℹ️ <b>aT 지원사업</b> — 한국농수산식품유통공사(aT) 수출 물류 지원. aT 정산서(RD) 기준 운송비 80%, 창고·라벨 90%를 지원받으며, KPI에서는 판매비(운반비·창고·라벨)에서 차감되어 있습니다. 원장에는 회사 부담분만 계상됩니다.</p>' : '';
  const gl = glL.length ? `<p class="muted sm">📒 원장(${esc(k.gl.sheet)})의 ${m}월 지원 관련 수익 거래</p><ul class="cm-lines">${glL.slice(0, 8).map((l) => `<li><b>${l.a}</b> ${(l.v / 1e6).toFixed(1)}M VND (≈${f1(krwM(l.v))}백만원) · ${esc(l.d)}${l.o ? ` <small class="muted">${esc(l.o)}</small>` : ''}</li>`).join('')}</ul>` : '';
  return `<div class="card kt-wrap kpi-sub"><div class="kg-h"><b>💰 ${m}월 지원금 수익</b><small class="muted">출처: '${esc(k.target?.sheet || '')}' 시트 지원 항목</small></div>${body || '<p class="muted">Target PL 시트에 지원금 항목이 없습니다.</p>'}${what}${gl}</div>`;
}

/* ---------- 📒 원장 근거: 판매비·관리비 항목별로 전월↔당월 원장 거래를 비교해 늘고 줄어든 내용을 짧게 ---------- */
// 장표 항목 ↔ 원장 계정 (9월 확인: 64173/18.5=운반비 5.08, 64175=창고 0.24, 64274=사무실 3.64, 64273=여비 2.4, River Park 임차=숙소 0.63)
const GL_ITEM = [
  ['판매비', /인건비/, (a) => /^6411/.test(a)],
  ['판매비', /운반/, (a) => /^64173/.test(a)],
  ['판매비', /창고/, (a) => /^64175/.test(a)],
  ['판매비', /라벨/, (a) => /^64177/.test(a)],
  ['판매비', /^기타/, (a) => /^641(3|8)/.test(a)],
  ['관리비', /인건비/, (a) => /^6421/.test(a)],
  ['관리비', /사무실/, (a) => /^64274/.test(a)],
  ['관리비', /숙소/, (a, d) => /^642/.test(a) && /rent house|apartment|accommodation|management fee c/i.test(d)],
  ['관리비', /여비/, (a) => /^64273/.test(a)],
  ['관리비', /^기타$/, (a, d) => /^642/.test(a) && !/^6421|^64274|^64273/.test(a) && !/rent house|apartment|accommodation|management fee c/i.test(d)],
];
const GL_KO = [[/refrigerated transport/i, '냉장운송비'], [/delivery fee/i, '배송비'], [/accrued for service fee/i, '물류 서비스료'], [/shui|union fee/i, '사회보험·노조비'], [/salary/i, '급여'],
  [/mooncake/i, '월병(고객 선물)'], [/customer sampling/i, '시식용 버섯'], [/sample/i, '샘플'], [/promotional support/i, '판촉 지원금'], [/carton/i, '박스 구매'], [/car rental/i, '차량 렌트'],
  [/office rental/i, '사무실 임차료'], [/rent house|apartment/i, '숙소 임차료'], [/management fee/i, '숙소 관리비'], [/legal/i, '법률 자문료'], [/health check/i, '건강검진'], [/visa/i, '비자'], [/\btrc\b/i, '체류증(TRC)'],
  [/computer rental/i, 'PC 렌트'], [/allocating prepaid/i, '선급비용 상각'], [/accounting|audit|tax/i, '회계·세무 수수료'], [/supplies/i, '소모품']];
const GL_WHO = [[/lotte|\blot$/i, '롯데'], [/cj logistics/i, 'CJ'], [/việt nhật|viet nhat/i, '비엣냣'], [/cobi/i, 'COBI ONE'], [/jang car/i, 'Jang Car'], [/jipyong/i, '지평'], [/bệnh viện|hospital/i, '병원']];
function glReason(k, sec, label, m) {
  const R = GL_ITEM.find(([s, re]) => s === sec && re.test(nzl(label)));
  if (!R || !k.gl?.lines) return '';
  const g = {};
  for (const l of k.gl.lines) {
    if ((l.m !== m && l.m !== m - 1) || !R[2](l.a, l.d)) continue;
    const ko = GL_KO.find(([re]) => re.test(l.d)), who = GL_WHO.find(([re]) => re.test(l.o));
    const name = (ko ? ko[1] : l.d.replace(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?/gi, '').replace(/[\d.,–-]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 28)) + (who ? `(${who[1]})` : '');
    const x = (g[name] ||= { c: 0, p: 0 });
    if (l.m === m) x.c += l.v; else x.p += l.v;
  }
  const top = Object.entries(g).map(([n, x]) => ({ n, d: krwM(x.c - x.p), c: x.c, p: x.p })).filter((x) => Math.abs(x.d) >= 0.1).sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 3);
  return top.map((x) => `${x.n} ${x.d > 0 ? '+' : '-'}${Math.abs(x.d).toFixed(1)}${!x.p ? '(신규)' : !x.c ? '(없음)' : ''}`).join(', ');
}

/* ---------- 매출단가·단위원가 변동 근거 (KPI 파일 숫자로) ---------- */
// 매출단가 변동 = 품목 단가 영향(Σ 단가차 × 당월 물량 / 총물량) + 품목 구성(믹스) 영향(나머지)
function priceReason(M, i, p) {
  const out = [], f0k = (v) => Math.round(Math.abs(v)).toLocaleString('ko-KR');
  const it = M.items.map((x) => ({ n: x.label, q: M.qtys.find((y) => y.label === x.label), pr: M.prices.find((y) => y.label === x.label) })).filter((x) => x.q && x.pr);
  const Q = it.reduce((s, x) => s + (x.q.a[i] || 0), 0), Qp = it.reduce((s, x) => s + (x.q.a[p] || 0), 0);
  if (!Q || !Qp || !M.price) return out;
  const eff = it.map((x) => ({ ...x, e: x.pr.a[i] > 0 && x.pr.a[p] > 0 ? (x.pr.a[i] - x.pr.a[p]) * x.q.a[i] / Q : 0 }));
  const pe = eff.reduce((s, x) => s + x.e, 0), mix = (M.price.a[i] - M.price.a[p]) - pe;
  const sg = (v) => `${v >= 0 ? '+' : '-'}${f0k(v)}원`;
  out.push(`품목 단가 변동 영향 ${sg(pe)}, 품목 구성(믹스) 변화 영향 ${sg(mix)}`);
  for (const x of eff.filter((x) => Math.abs(x.e) >= 10).sort((a, b) => Math.abs(b.e) - Math.abs(a.e)).slice(0, 2)) { const d = x.pr.a[i] - x.pr.a[p]; out.push(`${x.n} 단가 ${f0k(d)}원 ${d >= 0 ? '상승' : '하락'} (${f0k(x.pr.a[p])} → ${f0k(x.pr.a[i])}원/kg)`); }
  const sh = it.map((x) => ({ n: x.n, c: x.q.a[i] / Q, b: x.q.a[p] / Qp, pr: x.pr.a[i] })).filter((x) => Math.abs(x.c - x.b) >= 0.02).sort((a, b) => Math.abs(b.c - b.b) - Math.abs(a.c - a.b)).slice(0, 2);
  if (sh.length && Math.abs(mix) >= 10) out.push(`물량 비중: ${sh.map((x) => `${x.n} ${Math.round(x.b * 100)}% → ${Math.round(x.c * 100)}% (${f0k(x.pr)}원/kg)`).join(', ')}`);
  return out;
}
// 단위원가 하위(상품대·수입제비용·폐기원가) 근거: 'Target … PL' 시트 품목별 금액
function costReason(k, label, m) {
  const T = k.target;
  if (!T) return '';
  const j = m - 1, pj = m - 2, L = nzl(label);
  const tr = (re) => T.rows.find((r) => re.test(nzl(r.label)));
  const kids = (re) => T.rows.filter((r) => r.lv === 3 && re.test(nzl(r.p)));
  const won = (v) => `${v >= 0 ? '+' : '-'}${Math.round(Math.abs(v)).toLocaleString('ko-KR')}원`;
  const mil = (v) => `${v >= 0 ? '+' : '-'}${Math.abs(v).toFixed(1)}`;
  if (L === '상품대') {
    const a = tr(/^kg당원물대/), b = tr(/^kg당수출제비용/);
    return [a && `kg당 원물대 ${won(a.a[j] - a.a[pj])}`, b && `kg당 수출제비용 ${won(b.a[j] - b.a[pj])}`].filter(Boolean).join(', ');
  }
  const amt = L === '수입제비용' ? /^총수입제비용/ : L === '폐기원가' ? /^총폐기금액/ : null;
  if (!amt) return '';
  const tot = tr(amt), q = tr(/^물량/);
  if (!tot) return '';
  const top = kids(amt).map((r) => ({ n: r.label, d: (r.a[j] || 0) - (r.a[pj] || 0) })).filter((x) => Math.abs(x.d) >= 0.1).sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 2);
  const head = `총액 ${f1(tot.a[pj])} → ${f1(tot.a[j])}백만원${q ? `, 물량 ${f1(q.a[pj])} → ${f1(q.a[j])}톤` : ''}`;
  return head + (top.length ? ` / ${top.map((x) => `${x.n} ${mil(x.d)}`).join(', ')}` : '');
}

/* ---------- ✉️ KPI 송부 메일 요약: 사용자가 보내던 메일 형식(전월 대비 1)~8)) ---------- */
function mailText(k, M, m) {
  const i = m - 1, p = m - 2;
  if (p < 0) return '';
  // 메일엔 소수 1자리로 보이는 숫자끼리 뺀 값을 씀 (10.1 - 8.7 = 1.4톤)
  const ab = Math.abs, r1 = (v) => Math.round(v * 10) / 10, n1 = (v) => ab(v).toFixed(1), n0 = (v) => Math.round(ab(v)).toLocaleString('ko-KR');
  const ud = (d) => (d >= 0 ? '증가' : '감소');
  const L = ['안녕하세요', '', '베트남KPI 송부드립니다.', '', `${m - 1}월대비`, ''];
  let no = 0;
  const item = (t) => { L.push(`${++no}) ${t}`); };
  const sub = (t) => L.push(`-> ${t}`);
  const gap = () => L.push('');
  // 1) 매출액 + 거래처별 변동
  const s = M.sales, ds = r1(s.a[i]) - r1(s.a[p]);
  item(`매출액 ${n1(ds)}백만원 ${ud(ds)} (${f1(s.a[i])}백만원)`);
  for (const x of M.chans.filter((x) => ab(x.a[i] - x.a[p]) >= 1).sort((a, b) => ab(b.a[i] - b.a[p]) - ab(a.a[i] - a.a[p]))) {
    const d = r1(x.a[i]) - r1(x.a[p]);
    const q = chanQtyOf(k, x.label), qs = q && ab(q[i] || 0) >= 0.005 ? `, 물량 ${(q[i] || 0).toFixed(1)}톤(전월 ${(q[p] || 0).toFixed(1)})` : '';
    sub(ab(x.a[p]) < 0.05 ? `${x.label} 신규 ${f1(x.a[i])}백만원${qs}` : ab(x.a[i]) < 0.05 ? `${x.label} 매출 없음 (전월 ${f1(x.a[p])}백만원)` : `${x.label} ${n1(d)}백만원 ${ud(d)} (${f1(x.a[i])}백만원${qs})`);
  }
  gap();
  // 2) 물량
  if (M.qty) { const d = r1(M.qty.a[i]) - r1(M.qty.a[p]); item(`매출물량 ${n1(d)}톤 ${ud(d)} (${f1(M.qty.a[i])}톤)`); gap(); }
  // 3) 단가 + 품목별 단가 변동(±5% 이상)
  if (M.price) {
    const d = M.price.a[i] - M.price.a[p];
    item(`kg당 매출단가 ${n0(d)}원 ${ud(d)} (${f0(M.price.a[i])}원/kg)`);
    for (const t of priceReason(M, i, p)) sub(t);
    sub('[원인 확인]');
    gap();
  }
  // 4) 단위원가 (장표 표: 상품대·수입제비용·폐기원가 등 원/kg)
  const P = pptOf(k), pi = P ? P.months.indexOf(m) : -1, pp = P ? P.months.indexOf(m - 1) : -1;
  const blk = (re) => (P && pi >= 0 && pp >= 0 ? pptBlock(P, re) : null);
  const uc = blk(/^단위원가/);
  if (uc) {
    const d = uc.row.a[pi] - uc.row.a[pp];
    item(`kg당 단위원가 ${n0(d)}원 ${ud(d)} (${f0(uc.row.a[pi])}원/kg)`);
    for (const x of uc.kids) { const dd = x.a[pi] - x.a[pp]; if (ab(dd) >= 5) { const why = costReason(k, x.label, m); sub(`kg당 ${x.label} ${n0(dd)}원 ${ud(dd)} ${why ? `(${why})` : '[사유]'}`); } }
    const wl = wasteLine(k, m);
    if (wl) sub(wl);
    gap();
  }
  // 5) GP
  if (M.gp) { item(`GP ${f1(M.gp.a[i])}백만원, GP%는 ${pc(M.gpp?.a[i] || 0)}${M.gpp ? ` (전월 ${pc(M.gpp.a[p])})` : ''}`); gap(); }
  // 6)·7) 판매비·관리비 특이사항: 장표 표의 하위 항목 중 ±0.3백만원 이상
  const costs = (name, re, row) => {
    const b = blk(re), tot = row ? r1(row.a[i]) - r1(row.a[p]) : 0;
    item(`${name} 특이사항 : ${row ? `${f1(row.a[i])}백만원, 전월 대비 ${n1(tot)}백만원 ${ud(tot)}` : ''}`);
    let any = false;
    if (b) for (const x of b.kids.slice().sort((a, c) => ab(c.a[pi] - c.a[pp]) - ab(a.a[pi] - a.a[pp]))) { const dd = x.a[pi] - x.a[pp]; if (ab(dd) >= 0.3) { const why = glReason(k, name, x.label, m); sub(`${x.label} ${n1(dd)}백만원 ${ud(dd)} ${why ? `(원장: ${why})` : '[사유]'}`); any = true; } }
    return any;
  };
  const any6 = costs('판매비', /^판매비/, M.sell);
  for (const { tot, kids } of subsidyRows(k)) {
    const c = ab(tot.a[i]), b = ab(tot.a[p]);
    if (c < 0.05 && b < 0.05) continue;
    const det = kids.filter((x) => ab(x.a[i]) >= 0.05).map((x) => `${x.label} ${f1(ab(x.a[i]))}`).join(', ');
    sub(`${tot.label.replace(/\s*수익$/, '')} 지원금 ${f1(c)}백만원 (전월 ${f1(b)}백만원${det ? ` / ${det}` : ''}) — 판매비에서 차감`);
  }
  if (!any6) sub('[특이사항 확인]');
  gap();
  if (!costs('관리비', /^관리비/, M.admin)) sub('[특이사항 확인]');
  gap();
  // 8) EBITDA
  if (M.ebitda) { const d = r1(M.ebitda.a[i]) - r1(M.ebitda.a[p]); item(`${M.ebitda.label.replace(/\(.*\)/, '').trim()} ${f1(M.ebitda.a[i])}백만원 (전달대비 ${n1(d)}백만원 ${ud(d)})`); gap(); }
  // 9) Target 대비 (장표 숫자)
  if (P && pi === P.months.length - 1) {
    const t = [['매출', /^매출액/], ['GP', /^GP$/i], ['조정 EBITDA', /^조정\s*EBITDA/i]].map(([nm, re]) => { const b = blk(re); if (!b || !b.row.tgt) return ''; const d = b.row.a[pi] - b.row.tgt; return `${nm} ${d >= 0 ? '+' : '-'}${n1(d)}백만원`; }).filter(Boolean);
    if (t.length) { item(`단기 Target 대비 : ${t.join(', ')}`); gap(); }
  }
  L.push('감사합니다');
  return L.join('\n');
}
const chanFixNote = (k, m) => { const f = (k.chanFix || []).filter((x) => x.m === m); return f.length ? `<p class="sm cm-auto-note">ℹ️ ${f.map((x) => `${x.label} ${m}월 매출 ${f1(x.v)}백만원`).join(', ')}은 KPI Summary 거래처 칸이 비어 있어 원장(GL) 거래처 매출로 채웠습니다.</p>` : ''; };
function renderMail(k, M, m) {
  if (m < 2) return '<div class="card">✉️ 메일 초안은 전월과 비교하므로 2월부터 만들 수 있습니다.</div>';
  const key = `mail${m}`, saved = (k.notes || {})[key];
  const auto = mailText(k, M, m);
  // 안내: ① 거래처 합이 매출액보다 모자람(예전 버전으로 불러온 자료) ② 고쳐 쓴(저장된) 내용이 새 숫자와 다름
  const tops = k.rows.filter((r) => r.top && /^매출액/.test(r.label));
  const gap = tops.length > 1 ? (tops[0].a[m - 1] || 0) - k.rows.filter((r) => !r.top && r.grp === tops[1].grp).reduce((s, r) => s + (r.a[m - 1] || 0), 0) : 0;
  const warn = (gap >= 0.5 ? `<p class="sm cm-auto-note">⚠️ 거래처별 매출 합계가 매출액보다 ${f1(gap)}백만원 적습니다(KPI Summary 거래처 칸 비어 있음). 업무노트를 새로고침(Ctrl+Shift+R)한 뒤 KPI 파일을 <b>다시 불러오면</b> 원장(GL) 거래처 매출(이마트=THISO RETAIL)로 채워집니다.</p>` : '')
    + (saved && saved.trim() !== auto.trim() ? '<p class="sm cm-auto-note">✏️ 예전에 고쳐 쓴 내용이 표시되고 있습니다. 새 숫자로 다시 만들려면 「↺ 자동 초안으로 되돌리기」를 누르세요.</p>' : '');
  return `<article class="rp mail"><header class="rp-head"><h2>✉️ ${k.year}년 ${m}월 베트남 KPI 송부 메일 초안</h2>
    <p>전임자가 보내던 메일 형식(전월 대비 1)~8))으로 숫자를 자동으로 채웠습니다. <b>[사유]·[원인 확인]</b> 자리만 현지 확인 내용으로 바꾸고 복사해 메일에 붙이세요. 서명은 메일 프로그램의 서명을 쓰세요.</p>${chanFixNote(k, m)}${warn}</header>
    <div class="card">
      <div class="cm-tools"><button type="button" class="btn sm primary" id="mailCopy">📋 메일 내용 복사</button><button type="button" class="btn sm ghost" id="mailReset">↺ 자동 초안으로 되돌리기</button>
        <span class="muted sm">${saved ? '✏️ 고쳐 쓴 내용 (자동 저장됨)' : '✍️ 자동 초안 — 고쳐 쓰면 자동 저장'}</span></div>
      <textarea class="cm-text mail-text${saved ? '' : ' auto'}" data-id="${k.id}" data-m="${key}" rows="34" spellcheck="false">${esc(saved || auto)}</textarea>
    </div>
    <details class="card"><summary><b>📌 참고: ${m}월 특이사항 (자동)</b></summary><ul class="kf">${[...targetFindings(k, m), ...findings(M, m)].map((f) => `<li class="${f.lv}">${esc(f.txt)}</li>`).join('') || '<li>눈에 띄는 변동이 없습니다.</li>'}</ul></details>
    ${subsidyCard(k, m, M)}</article>`;
}

function render() {
  const box = $('#kpiBody');
  if (!box) return;
  const all = list();
  const sel = $('#kpiSel');
  if (!all.length) {
    sel.hidden = true; $('#kpiMonths').innerHTML = ''; $('#kpiMode').hidden = true; $('#kpiChat').hidden = true;
    box.innerHTML = `<div class="card kpi-empty"><div style="font-size:40px">📈</div><b>KPI 엑셀 파일을 불러와 주세요</b>
      <p class="muted">위의 「📂 KPI 파일 불러오기」를 누르거나 파일을 이 화면에 끌어다 놓으면, Summary 시트의 월별 실적·BGT를 읽어 요약·그래프·특이사항·전월 비교를 보여줍니다.<br>파일은 이 기기 안에서만 읽고, 숫자만 저장되어 탭·PC에서 같이 볼 수 있습니다.</p></div>`;
    return;
  }
  const k = all.find((x) => x.id === ui.id) || all[0];
  if (ui.id !== k.id) { ui.id = k.id; ui.m = k.month; }
  if (!ui.m || ui.m > k.month) ui.m = k.month;
  sel.hidden = all.length < 2;
  $('#kpiMode').hidden = false; $('#kpiChat').hidden = false;
  sel.innerHTML = all.map((x) => `<option value="${x.id}" ${x.id === k.id ? 'selected' : ''}>${x.year}년 · ${esc(x.title)}</option>`).join('');
  const m = ui.m, i = m - 1, p = m - 2;
  $('#kpiMonths').hidden = ui.mode === 'cmt';
  $('#kpiMonths').innerHTML = Array.from({ length: k.month }, (_, j) => `<button type="button" data-m="${j + 1}" class="${j + 1 === m ? 'on' : ''}">${ui.mode === 'ytd' && j ? '1~' : ''}${j + 1}월</button>`).join('');
  const M = model(k);
  if (!M.sales) { box.innerHTML = '<div class="card">⚠️ 이 파일에서 매출액 항목을 찾지 못했습니다.</div>'; return; }
  document.querySelectorAll('#kpiMode [data-mode]').forEach((b) => b.classList.toggle('on', b.dataset.mode === ui.mode));
  // 화면 하나가 오류 나도 멈추지 않고 이유를 보여줌
  const safe = (f, name) => { try { return f(); } catch (e) { console.error(e); return `<div class="card">⚠️ ${name} 화면을 만들다 오류가 났습니다: ${esc(e.message || e)}<br><small class="muted">새로고침(Ctrl+Shift+R) 후에도 같으면 이 문구를 알려주세요.</small></div>`; } };
  const chk = safe(() => checksCard(k), '파일 점검');
  if (ui.mode === 'cmt') { box.innerHTML = chk + safe(() => renderCmt(k), '경영회의 코멘트'); return; }
  if (ui.mode === 'mail') { box.innerHTML = chk + safe(() => renderMail(k, M, m), '메일 초안'); return; }
  if (ui.mode === 'form') { box.innerHTML = safe(() => renderForm(k), '양식 점검'); return; }
  if (ui.mode === 'ytd') {
    box.innerHTML = renderYtd(k, M, m) + '<p class="kpi-foot"><button type="button" class="btn sm" id="kpiPrint">🖨 인쇄·PDF 저장</button> <button type="button" class="btn sm danger" id="kpiDel">🗑 이 KPI 자료 지우기</button></p>';
    return;
  }

  // 요약 카드
  const TM = tmodel(k.target), si = tgtIdx(k.target);
  const card = (row, name, kind = 'amt', sub, trow) => {
    if (!row) return '';
    const c = row.a[i], b = p >= 0 ? row.a[p] : null, bud = row.b[i], tg = trow ? trow.t[si] : null;
    return `<div class="card kpi-card"><div class="kc-t">${name}</div><div class="kc-big ${c < 0 ? 'neg' : ''}">${f1(c)}<small>백만원</small></div>
      ${sub ? `<div class="kc-sub">${sub}</div>` : ''}
      <div class="kc-cmp">${tg !== null ? `<div class="kc-tg"><span>🎯 Target</span> ${f1(tg)} ${delta(c, tg, kind)}</div>` : ''}${b !== null ? `<div><span class="muted">전월</span> ${f1(b)} ${delta(c, b, kind)}</div>` : ''}${bud ? `<div><span class="muted">BGT</span> ${f1(bud)} ${delta(c, bud, kind)}</div>` : ''}</div></div>`;
  };
  const ratio = (r) => (r ? `${r.label.replace(/\(.*\)/, '')} <b>${pc(r.a[i])}</b>${p >= 0 ? ` <small class="muted">(전월 ${pc(r.a[p])})</small>` : ''}` : '');
  const cards = card(M.sales, '매출액', 'amt', M.qty ? `물량 <b>${f1(M.qty.a[i])}톤</b> · 단가 <b>${f0(M.price?.a[i] || 0)}원/kg</b>` : '', TM?.sales)
    + card(M.gp, 'GP (매출총이익)', 'amt', ratio(M.gpp), TM?.gp) + card(M.cm, 'CM (공헌이익)', 'amt', ratio(M.cmp), TM?.cm)
    + card(M.ebit, 'EBIT (영업이익)', 'amt', ratio(M.ebitp), TM?.ebit) + card(M.ebitda, '조정 EBITDA', 'amt', M.adj && M.adj.a[i] ? `조정사항 ${f1(M.adj.a[i])}` : '');
  const tcard = TM ? `<div class="card kt-wrap kpi-target"><div class="kg-h"><b>🎯 ${m}월 Target 대비</b><small class="muted">Target 기준을 눌러 바꿀 수 있습니다</small></div>${targetTable(k, m)}</div>` : '';

  // 특이사항
  const fs = [...targetFindings(k, m), ...findings(M, m)];
  const ic = { bad: '<b class="kd dn">▼</b>', good: '<b class="kd up">▲</b>', info: '<b class="kd" style="color:var(--primary)">●</b>' };
  const note = (k.notes || {})[m] || '';
  const notes = `<div class="card kpi-notes"><div class="kg-h"><b>📌 ${m}월 특이사항</b><small class="muted">전월·BGT와 비교해 자동으로 찾은 내용</small></div>
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
  const table = `<div class="card kt-wrap"><div class="kg-h"><b>📊 ${m}월 전월·BGT 비교</b><small class="muted">단위: 백만원 (물량 톤, 단가 원/kg)</small></div><div class="kt-scroll"><table class="kt">
    <thead><tr><th></th><th>${m}월</th><th>${p >= 0 ? `${m - 1}월` : '전월'}</th><th>전월 대비</th><th>BGT</th><th>BGT 대비</th><th>누계 1~${m}월</th><th>누계 BGT</th><th>누계 BGT 대비</th></tr></thead>
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
  box.innerHTML = `<p class="muted sm kpi-src">📄 ${esc(k.file || '')}${when ? ` · ${when.getMonth() + 1}/${when.getDate()} 불러옴` : ''} · ${esc(k.title)}</p>${chk}${safe(() => closeCard(k), '마감 체크리스트')}${chanFixNote(k, m)}
    <div class="kpi-cards">${cards}</div>${tcard}${subsidyCard(k, m, M)}${notes}<div class="kpi-charts">${charts}</div>${table}${detail}
    <p class="kpi-foot"><button type="button" class="btn sm danger" id="kpiDel">🗑 이 KPI 자료 지우기</button></p>`;
}

/* ---------- 💬 물어보기 (AI 챗봇): 불러온 숫자 전체 + 질문을 Claude에 보냄 (사용자 허락 2026-11-20) ---------- */
const chat = { log: [], busy: false };
function kpiContext() {
  const all = list();
  return all.map((k) => {
    const tops = {}, seen = {};
    for (const r of k.rows) if (r.top) { const base = r.label.replace(/\(.*\)/, '').trim(); seen[base] = (seen[base] || 0) + 1; tops[r.grp] = seen[base] > 1 && /매출액/.test(base) ? '거래처별 매출액' : r.label; }
    const name = (r) => (r.top ? (tops[r.grp] === '거래처별 매출액' ? '거래처별 매출액 합계' : r.label) : `${(tops[r.grp] || '').replace(/^매출액.*/, '품목별 매출액')} > ${r.label}`);
    const n = k.month;
    const fmt = (v) => (Math.abs(v) < 1e-9 ? '0' : String(Math.round(v * 1000) / 1000));
    const lines = k.rows.filter((r) => r.a.some((v) => v) || r.b.some((v) => v)).map((r) => `${name(r)} | 실적 ${r.a.slice(0, n).map(fmt).join(', ')} | BGT ${r.b.map(fmt).join(', ')}`);
    const notes = Object.entries(k.notes || {}).filter(([, v]) => v).map(([m, v]) => `- ${/^ytd/.test(m) ? `1~${m.slice(3)}월 보고` : `${m}월`} 메모: ${v}`);
    const T = k.target;
    const tlines = T ? T.rows.map((r) => `${r.p ? `${r.p} > ` : ''}${r.label} | 실적 ${r.a.slice(0, n).map(fmt).join(', ')} | ${T.scen.map((s, j) => `${s}=${fmt(r.t[j])}`).join(', ')}`) : [];
    const tblock = T ? `\n\n#### Target ('${T.sheet}' 시트, 월 기준 목표치 · 시나리오: ${T.scen.join(' / ')} · 화면 기준: ${T.scen[tgtIdx(T)]})\n형식: 항목 | 실적 1월~${n}월 | 시나리오별 월 Target\n${tlines.join('\n')}` : '';
    return `### ${k.year}년 ${k.title}\n자료: ${k.file || ''} · 결산월(실적 마지막 달): ${n}월\n형식: 항목 | 실적 1월~${n}월 값(쉼표 순서대로) | BGT 1월~12월 값\n${lines.join('\n')}${tblock}${notes.length ? `\n사용자 메모:\n${notes.join('\n')}` : ''}`;
  }).join('\n\n');
}
const CHAT_SYSTEM = (ctx) => `당신은 베트남 법인 월별 실적(KPI)을 분석해 답하는 재무 담당 비서입니다. 아래 <kpi> 자료의 숫자에만 근거해 한국어로 간결하게 답하세요.
- 단위: 금액은 백만원(KRW), 매출물량은 톤, 매출단가는 원/kg, GP%·CM%·EBIT%는 비율(0.35 = 35%).
- 'BGT'는 예산(Budget)입니다. 답에서도 '예산'이라는 말 대신 'BGT'라고 쓰세요.
- 가장 중요한 비교는 Target 대비입니다. Target은 연간 계획이 아니라 '한 달 기준 목표 수준'이므로 당월(또는 최근 월 평균) 실적과 비교하고, 시나리오를 말하지 않으면 화면 기준 시나리오로 답하세요. 비용(원가·판매비·관리비)은 Target보다 많으면 초과(부정적)입니다.
- 누계·분기 비율과 단가는 월 값을 평균하지 말고 합계로 다시 계산하세요(GP% = GP 합계 / 매출 합계, 단가 = 매출 합계 / 물량 합계 × 1000).
- 숫자는 소수 첫째 자리(단가는 원 단위)로, 어느 달·기간인지 함께 쓰세요. 비교는 마크다운 표가 좋으면 표로 쓰세요.
- 자료에 없는 내용(계정별 세부 비용, 점포별, 원인 등)은 추측하지 말고 'Summary 자료에는 없음'이라고 말하되, 숫자로 보이는 흐름은 설명해도 됩니다.
- 보고용 문장을 요청하면 개조식(-로 시작하는 짧은 문장)으로 쓰세요.
<kpi>
${ctx}
</kpi>`;
function mdHtml(md) {
  const out = [], rows = esc(md).split('\n');
  for (let i = 0; i < rows.length; i++) {
    const l = rows[i];
    if (/^\s*\|.*\|\s*$/.test(l)) {
      const tbl = [];
      while (i < rows.length && /^\s*\|.*\|\s*$/.test(rows[i])) tbl.push(rows[i++]);
      i--;
      const cells = (r) => r.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
      const body = tbl.filter((r) => !/^\s*\|[\s:|-]+\|\s*$/.test(r));
      out.push(`<div class="kt-scroll"><table class="rt">${body.map((r, j) => `<tr>${cells(r).map((c) => (j ? `<td>${c}</td>` : `<th>${c}</th>`)).join('')}</tr>`).join('')}</table></div>`);
      continue;
    }
    if (/^#{1,4}\s/.test(l)) { out.push(`<b>${l.replace(/^#+\s/, '')}</b><br>`); continue; }
    if (/^\s*[-*]\s/.test(l)) { out.push(`• ${l.replace(/^\s*[-*]\s/, '')}<br>`); continue; }
    out.push(l + '<br>');
  }
  return out.join('').replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
}
function renderChat() {
  const box = $('#kcLog');
  if (!box) return;
  box.innerHTML = chat.log.map((m, i) => (m.role === 'user' ? `<div class="aq-me">${esc(m.content)}</div>`
    : m.err ? `<div class="aq-bot err">⚠ ${esc(m.err)}</div>`
      : `<div class="aq-bot">${mdHtml(m.content)}<div class="kc-copy"><button type="button" class="btn ghost sm" data-kc-copy="${i}">📋 복사</button></div></div>`)).join('')
    + (chat.busy ? '<div class="aq-bot muted">숫자를 확인하는 중…</div>' : '');
  box.scrollTop = box.scrollHeight;
  $('#kcHint').hidden = !!chat.log.length;
}
async function askKpi(q) {
  q = q.trim();
  if (!q || chat.busy) return;
  if (!aiEnabled()) { chat.log.push({ role: 'user', content: q }, { role: 'assistant', err: '⚙️ 설정에서 Claude API 키를 먼저 넣어주세요.' }); renderChat(); return; }
  chat.log.push({ role: 'user', content: q });
  chat.busy = true; renderChat();
  try {
    const msgs = chat.log.filter((m) => !m.err).slice(-12).map((m) => ({ role: m.role, content: m.content }));
    while (msgs.length && msgs[0].role !== 'user') msgs.shift();
    const ans = await aiChat({ system: CHAT_SYSTEM(kpiContext()), messages: msgs });
    chat.log.push({ role: 'assistant', content: ans });
  } catch (e) {
    chat.log.push({ role: 'assistant', err: e.message || String(e) });
  }
  chat.busy = false; renderChat();
}

export function initKpi({ toast } = {}) {
  if (toast) toastFn = toast;
  const view = $('#view-kpi');
  if (!view) return;
  $('#kcForm').addEventListener('submit', (e) => { e.preventDefault(); const q = $('#kcQ').value; $('#kcQ').value = ''; askKpi(q); });
  $('#kcHint').addEventListener('click', (e) => { const b = e.target.closest('[data-q]'); if (b) askKpi(b.dataset.q); });
  $('#kcReset').addEventListener('click', () => { chat.log = []; renderChat(); });
  $('#kcLog').addEventListener('click', (e) => {
    const b = e.target.closest('[data-kc-copy]');
    if (!b) return;
    navigator.clipboard?.writeText(chat.log[+b.dataset.kcCopy].content).then(() => toastFn('복사했습니다.'), () => toastFn('복사하지 못했습니다.'));
  });
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
  // 📌 지난달 보고 후속 확인
  const fuKey = (k) => { const P = pptOf(k); return P ? `fu${P.months[P.months.length - 1]}` : null; };
  const fuSave = (k, upd) => { const key = fuKey(k); if (!key) return; const notes = { ...(k.notes || {}) }; notes[key] = upd({ ...(notes[key] || autoFollow(k, +key.slice(2)) || { items: [] }) }); store.upsert('kpi', { id: k.id, notes }); };
  view.addEventListener('change', async (e) => {
    const cur = store.find('kpi', ui.id);
    if (!cur) return;
    if (e.target.id === 'wfIn' && e.target.files.length) {
      const f = e.target.files[0];
      e.target.value = '';
      try {
        toastFn('폐기 파일 읽는 중…');
        const w = await readWasteFile(f);
        if (!w || !Object.keys(w.tons).length) { toastFn("이 파일에서 '폐기물량(톤)' 표를 찾지 못했습니다."); return; }
        const c2 = store.find('kpi', cur.id), old = (c2.notes || {}).waste || {}, waste = { ...old };
        // 지난 장표에 적힌 폐기율(r)이 있는 달은 그 율은 그대로 두고 톤만 파일 값으로
        for (const [m, t] of Object.entries(w.tons)) { const o = old[m]; waste[m] = o && typeof o === 'object' && Number.isFinite(o.r) ? { t, r: o.r } : t; }
        store.upsert('kpi', { id: cur.id, notes: { ...(c2.notes || {}), waste, wasteItems: w.items, wasteSrc: f.name } });
        toastFn(`폐기물량 ${Object.keys(w.tons).length}개월을 「${w.sheet}」 시트에서 가져왔습니다.`);
      } catch (err) { toastFn('폐기 파일을 읽지 못했습니다: ' + (err.message || err)); }
      return;
    }
    if (e.target.id === 'fmIn' && e.target.files.length) {
      const f = e.target.files[0];
      lastFile = { name: f.name, buf: await f.arrayBuffer() };
      e.target.value = '';
      formRun(cur, false);
      return;
    }
    if (e.target.id === 'fuIn' && e.target.files.length) {
      const f = e.target.files[0];
      try {
        const r = await readPrevPpt(f);
        if (!r) { toastFn('이 PPT에서 베트남법인 실적 장표를 찾지 못했습니다.'); return; }
        const items = prevItems(r.paras);
        fuSave(cur, () => ({ src: f.name, mon: r.mon, items }));
        const wn = Object.keys(r.waste || {}).length;
        if (wn) { const c2 = store.find('kpi', cur.id); store.upsert('kpi', { id: cur.id, notes: { ...(c2.notes || {}), waste: { ...r.waste, ...((c2.notes || {}).waste || {}) } } }); }
        toastFn(`지난달 장표에서 확인할 항목 ${items.length}개${wn ? `, 폐기물량 ${wn}개월` : ''}을 가져왔습니다.`);
      } catch (err) { toastFn('PPT를 읽지 못했습니다: ' + (err.message || err)); }
      return;
    }
    const fo = e.target.dataset?.fmOk;
    if (fo !== undefined) {
      const key = `formOk${form.month || cur.month}`, ok = { ...((cur.notes || {})[key] || {}) };
      if (e.target.checked) ok[fo] = 1; else delete ok[fo];
      store.upsert('kpi', { id: cur.id, notes: { ...(cur.notes || {}), [key]: ok } });
      return;
    }
    const cl = e.target.dataset?.closeI;
    if (cl !== undefined) {
      const key = `close${cur.month}`, d = { ...((cur.notes || {})[key] || {}) };
      if (e.target.checked) d[cl] = 1; else delete d[cl];
      store.upsert('kpi', { id: cur.id, notes: { ...(cur.notes || {}), [key]: d } });
      return;
    }
    const wm = e.target.dataset?.wasteM;
    if (wm !== undefined) {
      const v = e.target.value.trim().replace(/,/g, ''), waste = { ...((cur.notes || {}).waste || {}) };
      if (v === '' || !Number.isFinite(+v)) delete waste[wm]; else waste[wm] = +v;
      store.upsert('kpi', { id: cur.id, notes: { ...(cur.notes || {}), waste } });
      return;
    }
    const ci = e.target.dataset?.fuI;
    if (ci !== undefined) fuSave(cur, (fu) => ({ ...fu, items: fu.items.map((x, j) => (j === +ci ? { ...x, done: e.target.checked } : x)) }));
  });
  view.addEventListener('toggle', (e) => { if (e.target.classList?.contains('kpi-close')) ui.closeOpen = e.target.open; }, true);
  view.addEventListener('focusout', (e) => {
    const mi = e.target.dataset?.fuM, cur = store.find('kpi', ui.id);
    if (mi === undefined || !cur) return;
    const v = e.target.value.trim();
    fuSave(cur, (fu) => ({ ...fu, items: fu.items.map((x, j) => (j === +mi ? { ...x, memo: v } : x)) }));
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
  const isNote = (el) => el.id === 'kpiNote' || el.classList?.contains('cm-text');
  // 자동 초안(.auto)은 고쳐 쓰기 전까지 저장하지 않음 → 파일을 다시 불러오면 새 숫자로 바뀜
  view.addEventListener('input', (e) => { if (isNote(e.target)) { e.target.classList.remove('auto'); clearTimeout(t); const el = e.target; t = setTimeout(() => saveNote(el), 800); } });
  view.addEventListener('focusout', (e) => { if (isNote(e.target) && !e.target.classList.contains('auto')) { clearTimeout(t); saveNote(e.target); } });
  view.addEventListener('click', (e) => {
    const tb = e.target.closest('[data-tgt]');
    if (tb) { ui.tgt = +tb.dataset.tgt; try { localStorage.setItem('jy-kpitgt', ui.tgt); } catch { /* 저장 불가 */ } render(); return; }
    if (e.target.id === 'kpiPrint') { window.print(); return; }
    const cur = store.find('kpi', ui.id);
    const copy = (txt) => navigator.clipboard?.writeText(txt).then(() => toastFn('복사했습니다.'), () => toastFn('복사하지 못했습니다.'));
    const fill = e.target.closest('[data-cm-fill]');
    if (fill && cur) {
      const ta = fill.closest('.cm-block').querySelector('.cm-text');
      if (ta.value.trim() && !ta.classList.contains('auto') && !confirm('지금 쓴 내용을 자동 문구로 바꿀까요?')) return;
      ta.value = cmAutoText(cur, fill.dataset.cmFill);
      saveNote(ta);
      return;
    }
    if (e.target.id === 'cmAi' && cur) { cmAi(cur); return; }
    const tgb = e.target.closest('[data-cmg]');
    if (tgb) {
      const g = +tgb.dataset.cmg, open = !cmOpen.has(g);
      if (open) cmOpen.add(g); else cmOpen.delete(g);
      tgb.textContent = open ? '▾' : '▸'; tgb.setAttribute('aria-expanded', open);
      view.querySelectorAll(`[data-cmk="${g}"]`).forEach((tr) => { tr.hidden = !open; });
      return;
    }
    if (e.target.id === 'cmAllTg') {
      const gs = [...view.querySelectorAll('[data-cmg]')].map((b) => +b.dataset.cmg), open = !gs.every((g) => cmOpen.has(g));
      gs.forEach((g) => (open ? cmOpen.add(g) : cmOpen.delete(g)));
      view.querySelectorAll('[data-cmg]').forEach((b) => { b.textContent = open ? '▾' : '▸'; b.setAttribute('aria-expanded', open); });
      view.querySelectorAll('[data-cmk]').forEach((tr) => { tr.hidden = !open; });
      e.target.textContent = open ? '▴ 세부 모두 접기' : '▾ 세부 모두 펼치기';
      return;
    }
    if (e.target.id === 'fuLoad') { view.querySelector('#fuIn')?.click(); return; }
    if (e.target.id === 'fuCopy' && cur) {
      const mn = +fuKey(cur).slice(2), fu = (cur.notes || {})[fuKey(cur)] || autoFollow(cur, mn);
      copy(`[${mn}월 베트남 실적 — 지난달 보고 후속 확인 요청]\n` + (fu?.items || []).filter((x) => !x.done).map((x, j) => `${j + 1}. ${x.t}${followHint(cur, x.t, mn) ? `\n   → ${followHint(cur, x.t, mn)}` : ''}\n   → ${mn}월 진행 결과:`).join('\n'));
      return;
    }
    if (e.target.id === 'fmRun' && cur) { formRun(cur, false); return; }
    if (e.target.id === 'fmMake' && cur) { formRun(cur, true); return; }
    if (e.target.id === 'fmPick') { view.querySelector('#fmIn')?.click(); return; }
    if (e.target.id === 'wfLoad') { view.querySelector('#wfIn')?.click(); return; }
    // 의도한 변경 확인 처리(달마다 따로 저장)
    const okAll = e.target.closest('[data-fm-okall]');
    if ((okAll || e.target.id === 'fmOkReset') && cur && form.res) {
      const key = `formOk${form.month || cur.month}`, ok = { ...((cur.notes || {})[key] || {}) };
      if (okAll) { const [sh, kd] = okAll.dataset.fmOkall.split('|'); for (const b of form.res) if (b.sheet === sh) for (const i of b.issues) if (i.kind === kd) ok[`${sh}!${i.addr}`] = 1; }
      store.upsert('kpi', { id: cur.id, notes: { ...(cur.notes || {}), [key]: okAll ? ok : {} } });
      return;
    }
    if (e.target.id === 'mailCopy') { copy(view.querySelector('.mail-text')?.value || ''); return; }
    if (e.target.id === 'mailReset' && cur) {
      const ta = view.querySelector('.mail-text');
      if (!ta.classList.contains('auto') && !confirm('고쳐 쓴 내용을 지우고 자동 초안으로 되돌릴까요?')) return;
      const notes = { ...(cur.notes || {}) }; delete notes[ta.dataset.m];
      store.upsert('kpi', { id: cur.id, notes });
      render();
      return;
    }
    if (e.target.id === 'cmAskCopy' && cur) { copy(cmQuestions(cur).map((x, j) => `${j + 1}. ${x}`).join('\n')); return; }
    if (e.target.id === 'cmCopy' && cur) {
      const txt = [...view.querySelectorAll('.cm-block')].map((b) => { const t = b.querySelector('.cm-text').value.trim() || cmAutoText(cur, b.dataset.g); return `${b.dataset.g}. ${t}`; }).join('\n\n');
      copy(txt);
      return;
    }
    if (e.target.id !== 'kpiDel') return;
    const k = store.find('kpi', ui.id);
    if (k && confirm(`${k.year}년 「${k.title}」 KPI 자료를 지울까요?\n(엑셀 파일은 그대로이고, 사이트에 저장된 숫자와 메모만 지워집니다)`)) { store.remove('kpi', k.id); ui.id = null; render(); }
  });
  // 메모 입력 중에는 다시 그리지 않음 (다른 기기 동기화로 커서가 튀지 않게)
  store.subscribe(() => { const a = document.activeElement; if (!(a && (a.id === 'kpiNote' || a.classList?.contains('cm-text') || a.classList?.contains('fu-memo')))) render(); });
  render();
}