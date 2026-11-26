// 📦 베트남 수불재고 분석 (view-stock)
// 베트남팀 '수불재고_(베트남팀)_MMDD.xlsx'(약 28MB, 56시트)에서 필요한 시트 5개만 JSZip으로 읽어 숫자만 뽑음
//  - daily report(거래처 유형별 실적)(보고용): 기준일, 채널별 당일·MTD·BGT·YTD(매출·물량·단가), 일자별 표
//  - 비교시트: 월별 매출(백만원)·물량·단가(1 daily report 묶음), 직납 로스(kg)(2-2 직납로스 묶음)
//  - SKU별 주간판매실적(kg): 주(일~토)별 품목·SKU 판매 kg (최근 12주만 저장)
//  - II.(회계용)수불부(MTD_창고)ea 단위: 월별 품목 기초·입고·출고·기말(ea·ton)
//  - 3. 거래처 지점별 실적(MTD): 매출액(원) 묶음의 합계 줄 = 채널·점포별 MTD 매출
// 뽑은 숫자는 store 'stock' 항목(기준일별 하나)으로 저장 → 탭·PC 동기화. 원본 파일은 기기 밖으로 보내지 않음
import { store } from './store.js';

const $ = (s, r = document) => r.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const serialDate = (v) => (typeof v === 'number' && v > 40000 && v < 60000 ? new Date(Math.round((v - 25569) * 864e5)) : null);
const ymd = (d) => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
const md = (s) => { const [, m, d] = String(s).split('-'); return `${+m}/${+d}`; };
const num = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
const f0 = (v) => Math.round(num(v)).toLocaleString('ko-KR');
const f1 = (v) => (Math.round(num(v) * 10) / 10).toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const f2 = (v) => (Math.round(num(v) * 100) / 100).toLocaleString('ko-KR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const mil = (won) => f1(num(won) / 1e6); // 원 → 백만원
const pct = (a, b) => (b ? `${Math.round((a / b) * 1000) / 10}%` : '-');
const sgn = (v, f = f1) => (v > 0 ? '+' : v < 0 ? '−' : '±') + f(Math.abs(v));
const rate = (a, b) => (b ? (a / b - 1) * 100 : null);
const ratePct = (a, b) => { const r = rate(a, b); return r === null ? '-' : `${r >= 0 ? '+' : '−'}${Math.abs(Math.round(r))}%`; };

const SHEETS = {
  report: /거래처\s*유형별\s*실적.*보고용/,
  cmp: /^비교시트/,
  sku: /SKU별\s*주간판매실적\s*\(kg\)/i,
  ledger: /^II\..*수불부/,
  store: /거래처\s*지점별\s*실적/,
};

/* ---------------- 읽기 ---------------- */
function grid(cells) {
  const at = (r, c) => cells.get(r + ',' + c)?.v;
  const txt = (r, c) => { const v = at(r, c); return v === undefined || v === null ? '' : String(v).replace(/\s+/g, ' ').trim(); };
  let maxR = 0, maxC = 0;
  for (const x of cells.values()) { if (x.r > maxR) maxR = x.r; if (x.c > maxC) maxC = x.c; }
  const find = (re, cols = null, rows = [1, maxR]) => {
    for (let r = rows[0]; r <= rows[1]; r++) for (let c = 1; c <= maxC; c++) { if (cols && !cols.includes(c)) continue; if (re.test(txt(r, c))) return { r, c }; }
    return null;
  };
  return { at, txt, find, maxR, maxC };
}

function parseReport(g) {
  const out = { kpi: {}, ch: [], days: [] };
  const lab = (re) => { const p = g.find(re, [3]); return p ? num(g.at(p.r, 6)) : 0; };
  out.kpi.today = lab(/^Sales\(당일\)/);
  out.kpi.avg = lab(/당월 평균 일매출/);
  const wd = g.find(/^영업일수$/, [3]);
  out.kpi.days = wd ? num(g.at(wd.r, 6)) : 0;
  const prog = g.find(/영업일수 진도율/, [3]);
  out.kpi.progress = prog ? num(g.at(prog.r, 6)) : 0;
  out.kpi.daysTotal = out.kpi.progress ? Math.round(out.kpi.days / out.kpi.progress) : 0;
  const fx = g.find(/적용환율/, [4]);
  out.kpi.fx = fx ? num(g.at(fx.r, 3)) : 18.5;
  // 기준일: 머리 3행 근처의 날짜값
  for (let r = 1; r <= 6 && !out.date; r++) for (let c = 1; c <= 12; c++) { const d = serialDate(g.at(r, c)); if (d) { out.date = ymd(d); break; } }
  // 채널 표: '당일'·'당월 누계'·'YTD' 머리글 → 열
  const hDay = g.find(/^당일$/), hMtd = g.find(/당월 누계/), hYtd = g.find(/^YTD$/);
  const C = { day: hDay?.c || 7, mtd: hMtd?.c || 8, mbgt: (hMtd?.c || 8) + 1, ytd: hYtd?.c || 10, ybgt: (hYtd?.c || 10) + 1 };
  const sec = [['sales', /^매출액\s*\(원\)/], ['kg', /^물량\s*\(kg\)/], ['price', /^단가\s*\(원\/kg\)/]];
  const rowsOf = {};
  for (const [k, re] of sec) { const p = g.find(re, [5]); if (p) rowsOf[k] = p.r; }
  const vals = (r) => ({ day: num(g.at(r, C.day)), mtd: num(g.at(r, C.mtd)), mbgt: num(g.at(r, C.mbgt)), ytd: num(g.at(r, C.ytd)), ybgt: num(g.at(r, C.ybgt)) });
  if (rowsOf.sales) {
    out.total = {};
    for (const [k] of sec) if (rowsOf[k]) out.total[k] = vals(rowsOf[k]);
    // 채널 행: 매출액 다음 줄부터 물량 머리글 전까지 (E=이름, F=direct/indirect)
    const end = rowsOf.kg || rowsOf.sales + 12;
    for (let r = rowsOf.sales + 1; r < end; r++) {
      const name = g.txt(r, 5); if (!name) continue;
      const type = g.txt(r, 6);
      const i = r - rowsOf.sales;
      const ch = { name, type, sales: vals(r) };
      if (rowsOf.kg) ch.kg = vals(rowsOf.kg + i);
      if (rowsOf.price) ch.price = vals(rowsOf.price + i);
      out.ch.push(ch);
    }
  }
  // 일자별 표: '총 매출액' 머리글 아래 E열 날짜
  const h = g.find(/총 매출액/);
  if (h) {
    const heads = [];
    for (let c = h.c; c <= h.c + 11; c++) heads.push(g.txt(h.r, c));
    out.dayHeads = heads;
    for (let r = h.r + 1; r <= g.maxR; r++) {
      const d = serialDate(g.at(r, h.c - 1));
      if (!d) { if (out.days.length) break; continue; }
      const v = []; for (let c = h.c; c <= h.c + 11; c++) v.push(Math.round(num(g.at(r, c)) * 100) / 100);
      out.days.push({ d: ymd(d), v });
    }
  }
  return out;
}

// 비교시트: 4행 묶음 이름, 5행 월 번호, B열 항목(매출·물량·단가·로스)
function parseCmp(g) {
  const out = { sales: {}, kg: {}, price: {}, loss: {}, lossKg: {} };
  const monRow = (() => { for (let r = 1; r <= 10; r++) { let n = 0; for (let c = 3; c <= 8; c++) if (g.at(r, c) === c - 2) n++; if (n >= 5) return r; } return 5; })();
  const groups = [];
  for (let c = 1; c <= g.maxC; c++) { const t = g.txt(monRow - 1, c); if (t) groups.push({ c, t }); }
  const groupOf = (c) => { let cur = ''; for (const x of groups) if (x.c <= c) cur = x.t; return cur; };
  const rowOf = (re) => { const p = g.find(re, [2]); return p?.r; };
  const rs = { sales: rowOf(/^매출$/), kg: rowOf(/^물량$/), price: rowOf(/^단가$/), loss: rowOf(/^로스$/) };
  for (let c = 1; c <= g.maxC; c++) {
    const m = g.at(monRow, c);
    if (!(Number.isInteger(m) && m >= 1 && m <= 12)) continue;
    const grp = groupOf(c);
    if (/daily report/i.test(grp)) {
      for (const k of ['sales', 'kg', 'price']) if (rs[k] && g.at(rs[k], c) !== undefined) out[k][m] = num(g.at(rs[k], c));
    }
    if (/로스/.test(grp) && rs.loss && g.at(rs.loss, c) !== undefined) {
      out.loss[m] = num(g.at(rs.loss, c));
      if (g.at(rs.loss + 1, c) !== undefined) out.lossKg[m] = num(g.at(rs.loss + 1, c)); // 같은 묶음의 판매물량(비교용)
    }
  }
  return out;
}

// SKU 주간(kg): '시작일' 줄(날짜), 아래 'week' 줄, 그 아래 B=합계, C=품목군, D=SKU
function parseSku(g, keep = 12) {
  const h = g.find(/^시작일$/);
  if (!h) return null;
  const cols = [];
  for (let c = h.c + 1; c <= g.maxC; c++) { const d = serialDate(g.at(h.r, c)); if (d) cols.push({ c, d: ymd(d) }); }
  const rows = [];
  for (let r = h.r + 2; r <= g.maxR; r++) {
    const b = g.txt(r, 2), cat = g.txt(r, 3), sku = g.txt(r, 4);
    const name = sku || cat || b;
    if (!name) continue;
    // B열 제목이 또 나오면 다음 표('일평균 판매중량' 등) 시작 → 첫 표(판매중량 Total)만 읽음
    if (b && rows.some((x) => x.lv === 0)) break;
    const v = cols.map((x) => num(g.at(r, x.c)));
    rows.push({ name, lv: sku ? 2 : cat ? 1 : 0, v, cat: sku ? (cat || [...rows].reverse().find((x) => x.lv === 1)?.name || '') : '' });
  }
  // 판매가 있는 마지막 주까지
  const tot = rows.find((x) => x.lv === 0) || rows[0];
  let last = -1; tot?.v.forEach((v, i) => { if (v > 0) last = i; });
  const from = Math.max(0, last - keep + 1);
  // 시트 점검: 같은 SKU가 두 품목군에 들어 있으면(예: king oyster 300이 팽이·새송이 둘 다) 소계·합계에 두 번 더해짐
  const byName = {};
  rows.filter((x) => x.lv === 2).forEach((x) => (byName[x.name] ||= []).push(x.cat));
  const dups = Object.entries(byName).filter(([, cs]) => new Set(cs).size > 1).map(([name, cats]) => ({ name, cats: [...new Set(cats)] }));
  return { weeks: cols.slice(from, last + 1).map((x) => x.d), rows: rows.map((x) => ({ ...x, v: x.v.slice(from, last + 1).map((n) => Math.round(n * 10) / 10) })), dups };
}

// 수불부: A=월, B=품목, E 기초, F 입고ea, H 입고ton, I aj폐기, M 출고ea, O 출고ton, P 검역, Q 기말ea, S 기말ton
function parseLedger(g) {
  const byM = {};
  for (let r = 1; r <= g.maxR; r++) {
    const m = g.at(r, 1), item = g.txt(r, 2);
    if (!(Number.isInteger(m) && m >= 1 && m <= 12) || !item) continue;
    const row = { item, open: num(g.at(r, 5)), inEa: num(g.at(r, 6)), inTon: num(g.at(r, 8)), aj: num(g.at(r, 9)), outEa: num(g.at(r, 13)), outTon: num(g.at(r, 15)), qc: num(g.at(r, 16)), endEa: num(g.at(r, 17)), endTon: num(g.at(r, 19)), perCt: num(g.at(r, 3)) };
    (byM[m] ||= []).push(row);
  }
  return byM;
}

// 점포별 MTD: '매출액(원)' 아래 머리글 줄 + 합계 줄 (F열 TOTAL부터)
function parseStores(g) {
  const p = g.find(/^매출액\(원\)$/, [3]);
  if (!p) return null;
  const hr = p.r + 1, vr = p.r + 2, out = [];
  let grp = '';
  for (let c = 6; c <= g.maxC; c++) {
    const name = g.txt(hr, c); if (!name) continue;
    const v = Math.round(num(g.at(vr, c)));
    const isGrp = /total|\(특정\)|\(직납\)/i.test(name);
    if (isGrp) grp = name;
    out.push({ name, grp: isGrp ? '' : grp, isGrp, v });
  }
  return out;
}

export async function parseStockFile(buf, fileName) {
  const { openBook, parseSheet } = await import('./kpicheck.js');
  const book = await openBook(buf);
  const pick = (re) => book.sheets.find((s) => re.test(s.name));
  const read = async (re) => { const s = pick(re); if (!s) return null; const xml = await book.zip.file(s.path).async('string'); return grid(parseSheet(xml, book.sst)); };
  const miss = Object.entries(SHEETS).filter(([, re]) => !pick(re)).map(([k]) => k);
  const rg = await read(SHEETS.report);
  if (!rg) throw new Error("'daily report(거래처 유형별 실적)(보고용)' 시트를 찾지 못했습니다 — 수불재고 파일이 맞는지 확인해 주세요.");
  const rep = parseReport(rg);
  const cg = await read(SHEETS.cmp), sg = await read(SHEETS.sku), lg = await read(SHEETS.ledger), stg = await read(SHEETS.store);
  const date = rep.date || ymd(new Date());
  return {
    id: 'st' + date.replace(/-/g, ''),
    date, file: fileName, loadedAt: new Date().toISOString(),
    kpi: rep.kpi, total: rep.total || null, ch: rep.ch, dayHeads: rep.dayHeads || [], days: rep.days.filter((x) => x.d <= date),
    monthly: cg ? parseCmp(cg) : null,
    sku: sg ? parseSku(sg) : null,
    ledger: lg ? parseLedger(lg) : null,
    stores: stg ? parseStores(stg) : null,
    missing: miss,
  };
}

/* ---------------- 분석 ---------------- */
const list = () => [...(store.get().stock || [])].sort((a, b) => b.date.localeCompare(a.date));
const prevOf = (s) => list().find((x) => x.date < s.date);
const monthOf = (s) => +s.date.slice(5, 7);
const dayOf = (s) => +s.date.slice(8, 10);

function ledgerNow(s) {
  const L = s.ledger || {};
  const m = monthOf(s);
  const cur = L[m] || [], prev = L[m - 1] || [];
  const tot = (rows) => rows.find((x) => /^total$/i.test(x.item));
  const items = (rows) => rows.filter((x) => !/^total$/i.test(x.item) && (x.open || x.inEa || x.outEa || x.endEa));
  return { m, cur: items(cur), prev: items(prev), curT: tot(cur), prevT: tot(prev) };
}
// SKU 주간: 다 끝난 마지막 주 (기준일이 그 주 토요일 전이면 그 전 주)
function skuWeekIdx(s) {
  const w = s.sku?.weeks?.length || 0;
  if (!w) return -1;
  const lastStart = new Date(s.sku.weeks[w - 1] + 'T00:00:00Z');
  return (new Date(s.date + 'T00:00:00Z') - lastStart) / 864e5 < 6 ? w - 2 : w - 1;
}
// 재고일수 = 기말 / (당월 출고 ÷ 경과일)
const cover = (end, out, days) => (out > 0 && days ? end / (out / days) : null);

function findings(s) {
  const out = [];
  const k = s.kpi || {}, T = s.total || {};
  const S = T.sales || {}, KG = T.kg || {}, P = T.price || {};
  const add = (lv, t) => out.push({ lv, t });
  const prog = k.progress || 0;
  // 1) 매출 진도
  if (S.mtd) {
    const forecast = prog ? S.mtd / prog : 0;
    const ach = S.mbgt ? S.mtd / S.mbgt : 0;
    if (S.mbgt) add(ach >= prog ? 'good' : 'bad', `당월 매출 MTD ${mil(S.mtd)}백만원 — 월 BGT ${mil(S.mbgt)}백만원 대비 ${pct(S.mtd, S.mbgt)} (영업일 진도 ${Math.round(prog * 100)}%)`);
    const prevM = s.monthly?.sales?.[monthOf(s) - 1];
    if (forecast) add(prevM && forecast * 1e-6 < prevM ? 'bad' : 'info', `이 속도면 월말 매출 약 ${mil(forecast)}백만원 예상${prevM ? ` (전월 ${f1(prevM)}백만원 대비 ${ratePct(forecast / 1e6, prevM)})` : ''}`);
  }
  // 2) 단가
  if (P.mtd && P.ytd) {
    const r = rate(P.mtd, P.ytd);
    if (r !== null && Math.abs(r) >= 3) add(r > 0 ? 'good' : 'bad', `당월 평균 단가 ${f0(P.mtd)}원/kg — 연평균 ${f0(P.ytd)}원 대비 ${ratePct(P.mtd, P.ytd)}`);
  }
  // 3) 채널 — 당월 매출 비중·이번 달 실적 없는 채널
  const chs = (s.ch || []).filter((c) => c.sales);
  const active = chs.filter((c) => c.sales.mtd > 0);
  if (active.length && S.mtd) add('info', `채널 비중: ${active.map((c) => `${c.name}${c.type ? `(${c.type === 'indirect' ? '특정' : '직납'})` : ''} ${pct(c.sales.mtd, S.mtd)}`).join(' · ')}`);
  const stopped = chs.filter((c) => c.sales.mtd === 0 && c.sales.ytd > 0);
  if (stopped.length) add('bad', `이번 달 매출 없는 채널: ${stopped.map((c) => `${c.name}(연누계 ${mil(c.sales.ytd)}백만원)`).join(', ')}`);
  // 4) 최근 7일 vs 그 전 7일
  const ds = (s.days || []).filter((d) => d.v[0] > 0);
  if (ds.length >= 10) {
    const a = ds.slice(-7), b = ds.slice(-14, -7);
    const avg = (arr) => arr.reduce((t, d) => t + d.v[0], 0) / arr.length;
    const r = rate(avg(a), avg(b));
    if (r !== null && Math.abs(r) >= 10) add(r > 0 ? 'good' : 'bad', `최근 7일 일평균 매출 ${mil(avg(a))}백만원 — 그 전 7일(${mil(avg(b))}) 대비 ${ratePct(avg(a), avg(b))}`);
    const best = [...ds].sort((x, y) => y.v[0] - x.v[0])[0];
    add('info', `당월 최고 매출일 ${md(best.d)} ${mil(best.v[0])}백만원`);
  }
  // 5) SKU 주간
  if (s.sku?.weeks?.length >= 3) {
    const i = skuWeekIdx(s);
    const tot = s.sku.rows.find((x) => x.lv === 0);
    if (tot && i >= 1) {
      add(Math.abs(rate(tot.v[i], tot.v[i - 1]) ?? 0) >= 15 ? (tot.v[i] > tot.v[i - 1] ? 'good' : 'bad') : 'info', `주간 판매(${md(s.sku.weeks[i])}~ 주) ${f0(tot.v[i])}kg — 전주 ${f0(tot.v[i - 1])}kg 대비 ${ratePct(tot.v[i], tot.v[i - 1])}`);
      const movers = s.sku.rows.filter((x) => x.lv === 2 && (x.v[i] + x.v[i - 1]) >= 100).map((x) => ({ n: x.name, d: x.v[i] - x.v[i - 1] })).sort((x, y) => Math.abs(y.d) - Math.abs(x.d)).slice(0, 3);
      if (movers.length) add('info', `SKU 주간 변동 큰 순: ${movers.map((x) => `${x.n} ${sgn(x.d, f0)}kg`).join(', ')}`);
      // 시트 점검: 같은 SKU가 두 품목군에 있음 → 그 품목군 소계·전체 합계에 중복으로 더해짐
      for (const d of s.sku.dups || []) {
        const r = s.sku.rows.find((x) => x.lv === 2 && x.name === d.name);
        const extra = r ? r.v[i] * (d.cats.length - 1) : 0;
        add('bad', `SKU 시트 확인 필요: ${d.name}이(가) ${d.cats.join('·')} 두 품목군에 모두 있어 소계·합계에 중복 계산${extra ? ` (지난주 ${f0(extra)}kg 많게 잡힘 → 실제 합계 약 ${f0(tot.v[i] - extra)}kg)` : ''} — 베트남팀에 시트 수정 요청`);
      }
    }
  }
  // 6) 창고 재고
  const L = ledgerNow(s);
  if (L.curT) {
    const cv = cover(L.curT.endTon, L.curT.outTon, dayOf(s));
    add(cv !== null && cv < 7 ? 'bad' : 'info', `창고 기말 재고 ${f2(L.curT.endTon)}톤${cv !== null ? ` — 당월 출고 속도로 약 ${f1(cv)}일분` : ''}${!L.curT.inTon ? ` · ${L.m}월 입고 없음 → 입고 일정 확인 필요` : ` · ${L.m}월 입고 ${f2(L.curT.inTon)}톤`}`);
    const out0 = L.cur.filter((x) => x.endEa <= 0 && x.outEa > 0).map((x) => x.item);
    if (out0.length) add('bad', `재고 소진 품목: ${out0.join(', ')}`);
    const aj = L.cur.filter((x) => x.aj > 0);
    if (aj.length) add('info', `창고 폐기(aj): ${aj.map((x) => `${x.item} ${f0(x.aj)}ea`).join(', ')}`);
  }
  // 7) 직납 로스
  const lm = s.monthly?.loss || {}, m = monthOf(s);
  if (lm[m] !== undefined) {
    const base = s.monthly.lossKg?.[m] || s.monthly.kg?.[m];
    const prevR = lm[m - 1] !== undefined && (s.monthly.lossKg?.[m - 1] || s.monthly.kg?.[m - 1]) ? lm[m - 1] / (s.monthly.lossKg?.[m - 1] || s.monthly.kg?.[m - 1]) : null;
    const r = base ? lm[m] / base : null;
    add(r !== null && prevR !== null && r > prevR ? 'bad' : 'good', `직납 로스 ${m}월 ${f0(lm[m])}kg${r !== null ? ` (판매물량 대비 ${f1(r * 100)}%${prevR !== null ? `, 전월 ${f1(prevR * 100)}%` : ''})` : ''}`);
  }
  // 8) 지난 보고 대비
  const p = prevOf(s);
  if (p && p.date.slice(0, 7) === s.date.slice(0, 7) && p.total?.sales && S.mtd) {
    const dd = Math.round((new Date(s.date) - new Date(p.date)) / 864e5);
    const inc = S.mtd - p.total.sales.mtd;
    add('info', `지난 보고(${md(p.date)}) 이후 ${dd}일간 매출 +${mil(inc)}백만원 (하루 ${mil(inc / Math.max(1, dd))}백만원), 물량 +${f0((KG.mtd || 0) - (p.total.kg?.mtd || 0))}kg`);
  }
  return out;
}

function reportText(s) {
  const k = s.kpi || {}, T = s.total || {};
  const S = T.sales || {}, KG = T.kg || {}, P = T.price || {};
  const m = monthOf(s);
  const L = ledgerNow(s);
  const forecast = k.progress ? S.mtd / k.progress : 0;
  const chs = (s.ch || []).filter((c) => c.sales?.mtd > 0);
  const lines = [];
  lines.push(`[베트남 판매·재고 현황] ${md(s.date)} 기준 (영업일 ${k.days || '-'}/${k.daysTotal || '-'}일, ${Math.round((k.progress || 0) * 100)}%)`);
  lines.push('');
  lines.push(`1) 매출: ${m}월 누계 ${mil(S.mtd)}백만원${S.mbgt ? ` (월 BGT ${mil(S.mbgt)}백만원 대비 ${pct(S.mtd, S.mbgt)})` : ''}, 일평균 ${mil(k.avg)}백만원`);
  if (forecast) lines.push(`   - 월말 예상 ${mil(forecast)}백만원${s.monthly?.sales?.[m - 1] ? ` (전월 ${f1(s.monthly.sales[m - 1])}백만원)` : ''}`);
  if (chs.length) lines.push(`   - ${chs.map((c) => `${c.name}${c.type ? `(${c.type === 'indirect' ? '특정' : '직납'})` : ''} ${mil(c.sales.mtd)}`).join(' / ')}`);
  lines.push(`2) 물량·단가: 누계 ${f1((KG.mtd || 0) / 1000)}톤, 평균 단가 ${f0(P.mtd)}원/kg (연평균 ${f0(P.ytd)}원)`);
  if (s.sku?.weeks?.length >= 2) {
    const f = findings(s).find((x) => x.t.startsWith('주간 판매'));
    if (f) lines.push(`3) ${f.t}`);
  }
  if (L.curT) {
    const cv = cover(L.curT.endTon, L.curT.outTon, dayOf(s));
    lines.push(`4) 창고 재고: 기말 ${f2(L.curT.endTon)}톤${cv !== null ? ` (약 ${f1(cv)}일분)` : ''}, ${m}월 출고 ${f2(L.curT.outTon)}톤 / 입고 ${f2(L.curT.inTon)}톤${!L.curT.inTon ? ' — 입고 일정 확인 필요' : ''}`);
  }
  const lm = s.monthly?.loss?.[m];
  if (lm !== undefined) { const base = s.monthly.lossKg?.[m] || s.monthly.kg?.[m]; lines.push(`5) 직납 로스: ${m}월 ${f0(lm)}kg${base ? ` (판매물량 대비 ${f1((lm / base) * 100)}%)` : ''}`); }
  return lines.join('\n');
}

/* ---------------- 그리기 ---------------- */
// 막대그래프 (HTML) — 화면 폭이 바뀌어도 글자가 늘어나지 않음. 왼쪽 세로축 눈금(단위), 가로 눈금선, 평균선(선택)
function niceStep(raw) {
  const p = 10 ** Math.floor(Math.log10(raw || 1)), r = raw / p;
  return (r <= 1 ? 1 : r <= 2 ? 2 : r <= 2.5 ? 2.5 : r <= 5 ? 5 : 10) * p;
}
function bars(vals, labels, { fmt = f1, hi = -1, unit = '', avg = null, showVals = null, tip = null } = {}) {
  const max = Math.max(...vals, 0) || 1;
  const step = niceStep(max / 4), top = Math.ceil(max / step) * step;
  const ticks = []; for (let t = 0; t <= top + 1e-9; t += step) ticks.push(t);
  const n = vals.length, sv = showVals ?? n <= 16;
  const tf = (t) => (step < 1 ? t.toFixed(step < 0.1 ? 2 : 1) : f0(t));
  return `<div class="st-bc" role="img" aria-label="${esc(unit)} 막대그래프">
    <div class="st-bc-y">${unit ? `<em>${esc(unit)}</em>` : ''}${ticks.map((t) => `<span style="bottom:${(t / top) * 100}%">${tf(t)}</span>`).join('')}</div>
    <div class="st-bc-plot">
      ${ticks.map((t) => `<i class="gl" style="bottom:${(t / top) * 100}%"></i>`).join('')}
      ${avg !== null && avg > 0 ? `<i class="avg" style="bottom:${(avg / top) * 100}%"><b>평균 ${fmt(avg)}</b></i>` : ''}
      <div class="st-bc-bars">${vals.map((v, i) => `<div class="bc${i === hi ? ' hi' : ''}${v ? '' : ' zero'}" title="${esc(tip ? tip(i) : labels[i])}: ${fmt(v)}${unit ? ` ${esc(unit)}` : ''}"><div class="bar" style="height:${(v / top) * 100}%">${sv && v ? `<span>${fmt(v)}</span>` : ''}</div></div>`).join('')}</div>
    </div>
    <div></div>
    <div class="st-bc-x${n > 15 ? ' dense' : ''}">${labels.map((l) => `<span>${esc(l)}</span>`).join('')}</div>
  </div>`;
}

function render(root, s) {
  const k = s.kpi || {}, T = s.total || {};
  const S = T.sales || {}, KG = T.kg || {}, P = T.price || {};
  const m = monthOf(s);
  const L = ledgerNow(s);
  const cv = L.curT ? cover(L.curT.endTon, L.curT.outTon, dayOf(s)) : null;
  const forecast = k.progress ? S.mtd / k.progress : 0;
  const lm = s.monthly?.loss?.[m], lbase = s.monthly?.lossKg?.[m] || s.monthly?.kg?.[m];
  const card = (t, v, sub, cls = '') => `<div class="card kpi-card st-card ${cls}"><div class="kc-t">${t}</div><div class="kc-big">${v}</div><div class="kc-sub muted sm">${sub}</div></div>`;
  const F = findings(s);
  const p = prevOf(s);

  // 채널 표
  const chRows = (s.ch || []).map((c) => `<tr><th>${esc(c.name)} <small class="muted">${esc(c.type === 'indirect' ? '특정' : c.type === 'direct' ? '직납' : c.type)}</small></th>
    <td>${mil(c.sales.day)}</td><td><b>${mil(c.sales.mtd)}</b></td><td>${pct(c.sales.mtd, S.mtd)}</td><td>${f0(c.kg?.mtd)}</td><td>${c.price?.mtd ? f0(c.price.mtd) : '-'}</td><td>${mil(c.sales.ytd)}</td><td>${c.price?.ytd ? f0(c.price.ytd) : '-'}</td></tr>`).join('');
  // 일자별
  const ds = s.days || [];
  const dayVals = ds.map((d) => Math.round(d.v[0] / 1e4) / 100);
  const dayChart = ds.length ? bars(dayVals, ds.map((d) => String(+d.d.slice(8))), { fmt: f1, hi: ds.length - 1, unit: '백만원', avg: k.avg ? k.avg / 1e6 : dayVals.reduce((a, b) => a + b, 0) / dayVals.length, tip: (i) => md(ds[i].d) }) : '';
  const dayTable = ds.length ? `<table class="kt"><thead><tr><th>날짜</th>${(s.dayHeads || []).slice(0, 12).map((h, i) => `<th>${esc(h)}${i < 4 ? '<br><small>(백만원)</small>' : i < 8 ? '<br><small>(kg)</small>' : '<br><small>(원/kg)</small>'}</th>`).join('')}</tr></thead>
    <tbody>${[...ds].reverse().map((d) => `<tr><th>${md(d.d)}</th>${d.v.map((v, i) => `<td>${i < 4 ? mil(v) : i < 8 ? f1(v) : f0(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>` : '';
  // 월별 추이
  const M = s.monthly;
  const mons = M ? Object.keys(M.sales).map(Number).sort((a, b) => a - b) : [];
  const monTable = mons.length ? `<table class="kt"><thead><tr><th></th>${mons.map((x) => `<th>${x}월${x === m ? '<br><small>(진행 중)</small>' : ''}</th>`).join('')}</tr></thead><tbody>
    <tr><th>매출(백만원)</th>${mons.map((x) => `<td>${f1(M.sales[x])}</td>`).join('')}</tr>
    <tr><th>물량(kg)</th>${mons.map((x) => `<td>${f0(M.kg[x])}</td>`).join('')}</tr>
    <tr><th>단가(원/kg)</th>${mons.map((x) => `<td>${f0(M.price[x])}</td>`).join('')}</tr>
    <tr><th>직납 로스(kg)</th>${mons.map((x) => `<td>${M.loss[x] !== undefined ? f0(M.loss[x]) : '-'}</td>`).join('')}</tr>
    <tr><th>로스율</th>${mons.map((x) => { const b = M.lossKg?.[x] || M.kg[x]; return `<td>${M.loss[x] !== undefined && b ? `${f1((M.loss[x] / b) * 100)}%` : '-'}</td>`; }).join('')}</tr>
  </tbody></table>` : '<p class="muted sm">비교시트가 없어 월별 추이를 못 읽었습니다.</p>';
  const monChart = mons.length ? bars(mons.map((x) => M.sales[x]), mons.map((x) => `${x}월`), { fmt: f1, hi: mons.indexOf(m), unit: '백만원' }) : '';
  // SKU 주간
  const W = s.sku;
  const wN = W?.weeks?.length || 0, wShow = Math.min(8, wN), w0 = wN - wShow;
  const wi = W ? skuWeekIdx(s) : -1; // 마지막으로 다 끝난 주
  const skuTable = wN ? `<table class="kt st-sku"><thead><tr><th>품목 · SKU</th>${W.weeks.slice(w0).map((d, j) => `<th>${md(d)}~${w0 + j > wi ? '<br><small>(진행 중)</small>' : ''}</th>`).join('')}<th>전주 대비</th></tr></thead><tbody>
    ${W.rows.filter((r) => r.v.some((v) => v > 0)).map((r) => { const a = r.v[wi], b = r.v[wi - 1]; const dup = r.lv === 2 && (W.dups || []).some((d) => d.name === r.name); return `<tr class="lv${r.lv}${dup ? ' st-dup' : ''}"><th>${esc(r.name)}${dup ? ' <small title="같은 SKU가 두 품목군에 있어 소계에 중복 계산됨 (원본 시트 확인)">⚠ 중복</small>' : ''}</th>${r.v.slice(w0).map((v) => `<td>${v ? f0(v) : '-'}</td>`).join('')}<td class="${a > b ? 'plus' : a < b ? 'minus' : ''}">${b ? ratePct(a, b) : '-'}</td></tr>`; }).join('')}
  </tbody></table><p class="muted sm">주 = 일요일~토요일. 전주 대비 = 다 끝난 마지막 주(${wi >= 0 ? md(W.weeks[wi]) : '-'}~)와 그 전 주</p>` : '<p class="muted sm">SKU 주간 시트를 못 읽었습니다.</p>';
  // 창고 수불
  const ledRow = (x) => { const c = cover(x.endEa, x.outEa, dayOf(s)); return `<tr class="${x.endEa <= 0 && x.outEa > 0 ? 'st-out' : c !== null && c < 7 ? 'st-low' : ''}"><th>${esc(x.item)}</th><td>${f0(x.open)}</td><td>${f0(x.inEa)}</td><td>${f0(x.outEa)}</td><td>${x.aj ? f0(x.aj) : '-'}</td><td><b>${f0(x.endEa)}</b></td><td>${f2(x.endTon)}</td><td>${c === null ? '-' : f1(c)}</td></tr>`; };
  const ledTable = L.cur.length ? `<table class="kt"><thead><tr><th>품목</th><th>기초(ea)</th><th>입고</th><th>출고</th><th>폐기(aj)</th><th>기말</th><th>기말(톤)</th><th>재고일수</th></tr></thead><tbody>
    ${L.curT ? `<tr class="tot"><th>합계</th><td>${f0(L.curT.open)}</td><td>${f0(L.curT.inEa)}</td><td>${f0(L.curT.outEa)}</td><td>${L.curT.aj ? f0(L.curT.aj) : '-'}</td><td><b>${f0(L.curT.endEa)}</b></td><td>${f2(L.curT.endTon)}</td><td>${cv === null ? '-' : f1(cv)}</td></tr>` : ''}
    ${L.cur.map(ledRow).join('')}</tbody></table><p class="muted sm">재고일수 = 기말 ÷ (당월 출고 ÷ 경과일 ${dayOf(s)}일). 빨강 = 재고 소진, 주황 = 7일분 미만${L.prevT ? ` · 전월(${m - 1}월) 출고 ${f2(L.prevT.outTon)}톤, 기말 ${f2(L.prevT.endTon)}톤` : ''}</p>` : '<p class="muted sm">수불부 시트에서 이번 달 표를 못 찾았습니다.</p>';
  // 점포
  const st = s.stores || [];
  const groups = st.filter((x) => x.isGrp && !/^total$/i.test(x.name));
  const allT = st.find((x) => /^total$/i.test(x.name));
  const storeSum = st.length ? `<p class="muted sm">입고일 기준(거래처 지점별 시트) · ${[allT, ...groups].filter((x) => x && x.v > 0).map((x) => `${esc(/^total$/i.test(x.name) ? '전체' : x.name)} <b>${mil(x.v)}</b>`).join(' · ')}</p>` : '';
  const storeHtml = st.length ? storeSum + groups.filter((g) => st.some((x) => x.grp === g.name && x.v > 0)).map((g) => { const kids = st.filter((x) => x.grp === g.name && x.v > 0).sort((a, b) => b.v - a.v); return `<div class="st-grp"><div class="st-gh"><b>${esc(g.name)}</b> <span>${mil(g.v)}백만원</span></div>${kids.length ? `<ul>${kids.map((x) => `<li><span>${esc(x.name)}</span><span class="st-bar"><i style="width:${Math.round((x.v / (kids[0].v || 1)) * 100)}%"></i></span><b>${mil(x.v)}</b></li>`).join('')}</ul>` : '<p class="muted sm">점포별 실적 없음</p>'}</div>`; }).join('') : '<p class="muted sm">점포별 시트를 못 읽었습니다.</p>';

  root.innerHTML = `
    <div class="kpi-cards">
      ${card(`${m}월 매출 (MTD)`, `${mil(S.mtd)}<small>백만원</small>`, S.mbgt ? `월 BGT ${mil(S.mbgt)} 대비 <b>${pct(S.mtd, S.mbgt)}</b> · 진도 ${Math.round((k.progress || 0) * 100)}%` : `일평균 ${mil(k.avg)}백만원`)}
      ${card('월말 예상 매출', forecast ? `${mil(forecast)}<small>백만원</small>` : '-', M?.sales?.[m - 1] ? `전월 ${f1(M.sales[m - 1])}백만원 대비 ${ratePct(forecast / 1e6, M.sales[m - 1])}` : '현재 속도 기준')}
      ${card('판매 물량 (MTD)', `${f1((KG.mtd || 0) / 1000)}<small>톤</small>`, `평균 단가 ${f0(P.mtd)}원/kg (연 ${f0(P.ytd)})`)}
      ${card('창고 재고', L.curT ? `${f2(L.curT.endTon)}<small>톤</small>` : '-', cv !== null ? `약 ${f1(cv)}일분 · ${m}월 입고 ${f2(L.curT.inTon)}톤` : '수불부 없음', cv !== null && cv < 7 ? 'warn' : '')}
      ${card(`${m}월 직납 로스`, lm !== undefined ? `${f0(lm)}<small>kg</small>` : '-', lm !== undefined && lbase ? `판매물량 대비 ${f1((lm / lbase) * 100)}%` : '')}
    </div>
    <div class="card"><h3>🔎 특이사항 <small class="muted">${md(s.date)} 기준${p ? ` · 지난 보고 ${md(p.date)}` : ''}</small></h3><ul class="kf">${F.map((x) => `<li class="${x.lv}">${esc(x.t)}</li>`).join('')}</ul>
      ${s.missing?.length ? `<p class="muted sm">⚠ 못 찾은 시트: ${esc(s.missing.join(', '))}</p>` : ''}</div>
    <div class="card"><h3>📋 보고 문구 <small class="muted">수정해서 복사하세요 (저장은 안 됨)</small></h3><textarea class="st-text" id="stText" rows="11">${esc(reportText(s))}</textarea><div class="st-act"><button type="button" class="btn sm" id="stCopy">📋 복사</button></div></div>
    <div class="card"><h3>🛒 채널별 실적 <small class="muted">매출 백만원 · 물량 kg · 단가 원/kg</small></h3><div class="tbl-wrap"><table class="kt"><thead><tr><th>채널</th><th>당일</th><th>${m}월 누계</th><th>비중</th><th>누계 물량</th><th>누계 단가</th><th>연 누계</th><th>연 단가</th></tr></thead><tbody>
      <tr class="tot"><th>합계</th><td>${mil(S.day)}</td><td><b>${mil(S.mtd)}</b></td><td>100%</td><td>${f0(KG.mtd)}</td><td>${f0(P.mtd)}</td><td>${mil(S.ytd)}</td><td>${f0(P.ytd)}</td></tr>${chRows}</tbody></table></div></div>
    <div class="card"><h3>📅 일자별 매출 <small class="muted">백만원 · 주황 = 기준일 · 점선 = 일평균 · 막대에 마우스를 올리면 금액</small></h3>${dayChart}<details><summary class="muted sm">일자별 표 보기</summary><div class="tbl-wrap">${dayTable}</div></details></div>
    <div class="card"><h3>📈 월별 추이 <small class="muted">비교시트 기준 (이번 달은 진행 중)</small></h3>${monChart}<div class="tbl-wrap">${monTable}</div></div>
    <div class="card"><h3>🍄 SKU별 주간 판매 (kg)</h3><div class="tbl-wrap">${skuTable}</div></div>
    <div class="card"><h3>🏭 창고 수불 (${m}월, ea)</h3><div class="tbl-wrap">${ledTable}</div></div>
    <div class="card"><h3>🏬 점포별 ${m}월 매출 <small class="muted">백만원</small></h3><div class="st-stores">${storeHtml}</div></div>
    <div class="card"><h3>🗒 메모</h3><textarea class="st-text" id="stNote" rows="4" placeholder="이번 주 확인할 것, 현지 답변 등">${esc(s.note || '')}</textarea></div>`;
  $('#stCopy', root).onclick = async () => { try { await navigator.clipboard.writeText($('#stText', root).value); ui.toast?.('복사했습니다.'); } catch { $('#stText', root).select(); document.execCommand('copy'); } };
  $('#stNote', root).addEventListener('change', (e) => store.upsert('stock', { id: s.id, note: e.target.value }));
}

// 홈 '베트남 요약' 카드용: 가장 최근 수불재고 보고의 핵심 숫자 + 빨간 특이사항 2개
export function stockHomeSummary() {
  const s = list()[0];
  if (!s) return null;
  const k = s.kpi || {}, S = s.total?.sales || {};
  const L = ledgerNow(s);
  return {
    date: s.date, mtd: S.mtd || 0, bgt: S.mbgt || 0, progress: k.progress || 0,
    forecast: k.progress ? (S.mtd || 0) / k.progress : 0,
    endTon: L.curT?.endTon ?? null, cover: L.curT ? cover(L.curT.endTon, L.curT.outTon, dayOf(s)) : null,
    alerts: findings(s).filter((x) => x.lv === 'bad').map((x) => x.t).sort((a, b) => /재고|소진|입고/.test(b) - /재고|소진|입고/.test(a)).slice(0, 2),
  };
}

/* ---------------- 화면 ---------------- */
const ui = { id: null, toast: null };
function refresh() {
  const items = list();
  const sel = $('#stSel');
  if (!items.length) {
    sel.hidden = true;
    $('#stBody').innerHTML = '<div class="card empty">📂 수불재고 파일(수불재고_(베트남팀)_MMDD.xlsx)을 불러오면 매출 진도·채널·SKU 주간 판매·창고 재고·로스를 분석합니다.<br><small class="muted">파일은 이 기기에서만 읽고, 뽑아낸 숫자만 저장·동기화합니다.</small></div>';
    return;
  }
  if (!items.some((x) => x.id === ui.id)) ui.id = items[0].id;
  sel.hidden = false;
  sel.innerHTML = items.map((x) => `<option value="${x.id}" ${x.id === ui.id ? 'selected' : ''}>${md(x.date)} 기준 · ${esc(x.file || '')}</option>`).join('');
  const s = items.find((x) => x.id === ui.id);
  const ta = document.activeElement?.id === 'stNote'; // 메모 입력 중엔 다시 그리지 않음
  if (!ta) render($('#stBody'), s);
}

export function initStock({ toast } = {}) {
  ui.toast = toast;
  const inp = $('#stIn');
  $('#stLoad').onclick = () => inp.click();
  $('#stSel').onchange = (e) => { ui.id = e.target.value; refresh(); };
  inp.onchange = async () => {
    const f = inp.files[0];
    inp.value = '';
    if (!f) return;
    const msg = $('#stMsg');
    msg.textContent = `⏳ ${f.name} 읽는 중… (큰 파일이라 10~30초 걸릴 수 있습니다)`;
    try {
      const s = await parseStockFile(await f.arrayBuffer(), f.name);
      const old = store.find('stock', s.id);
      store.upsert('stock', { ...s, note: old?.note || '' });
      ui.id = s.id;
      msg.textContent = `✅ ${f.name} — ${md(s.date)} 기준으로 불러왔습니다.${old ? ' (같은 기준일 자료를 새 파일로 바꿨습니다)' : ''}`;
      refresh();
    } catch (err) {
      msg.textContent = `⚠ 읽지 못했습니다: ${err.message || err}`;
    }
  };
  $('#stDel').onclick = () => {
    const s = store.find('stock', ui.id);
    if (!s || !confirm(`${md(s.date)} 기준 수불재고 분석을 목록에서 지울까요? (원본 파일은 그대로)`)) return;
    store.remove('stock', s.id);
    ui.id = null;
    refresh();
  };
  store.subscribe(() => { if ($('#view-stock')?.classList.contains('active')) refresh(); });
  refresh();
}
export { refresh as refreshStock };
