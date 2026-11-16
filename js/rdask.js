// 💬 RD에 물어보기: 질문 → AI가 '조회 조건'만 만들고(질문·열 이름·계정/사업장 목록만 전송, 금액은 전송 안 함)
// → 이 PC 브라우저 안에서 보관함 RD를 걸러 원래 ERP 양식(같은 열 순서·머리글)으로 표와 엑셀을 만듦
import { loadXLSX } from './excel.js';
import { aiJson, aiEnabled } from './ai.js';
import { findTable, norm, num, Book, Sheet, valCell, blankCell, textCell, numCell, colLet, cellInfo } from './xlsxkit.js';
import { vaultState } from './rdvault.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const won = (n) => (Math.round(n * 100) / 100).toLocaleString('ko-KR');

// RD 종류별: 표를 찾는 머리글, 합계를 낼 숫자 열, 날짜 열, 이름
export const SOURCES = {
  account: { label: '계정별원장', req: ['계정코드', '일자'], nums: ['차변', '대변', '잔액', '환율'], sum: ['차변', '대변'], date: '일자' },
  partner: { label: '거래처원장', req: ['거래처명'], nums: ['기초', '차변', '대변', '잔액'], sum: ['기초', '차변', '대변', '잔액'], date: null },
  receipts: { label: '기간별입고현황', req: ['입고일자', '품목코드'], nums: ['입고수량', '입고단가', '입고금액'], sum: ['입고수량', '입고금액'], date: '입고일자' },
};

/* ---------- RD 파일 → 표 (머리글 + 행) ---------- */
const cache = new Map(); // fileId → {source: {head, rows}}
async function tablesOf(f) {
  if (cache.has(f.id)) return cache.get(f.id);
  const XLSX = await loadXLSX();
  const wb = XLSX.read(f.data, { cellDates: true });
  const out = {};
  for (const sn of wb.SheetNames) {
    const grid = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
    for (const [key, S] of Object.entries(SOURCES)) {
      const t = findTable(grid, S.req);
      if (!t) continue;
      if (key === 'partner' && !t.head.some((h) => h === '계정코드' || h === '계정명')) continue;
      // 이름 없는 열은 빼고, 같은 이름이 또 나오면 '이름(2)'
      const cols = [];
      const seen = {};
      t.head.forEach((h, c) => { if (!h) return; seen[h] = (seen[h] || 0) + 1; cols.push({ c, name: seen[h] > 1 ? `${h}(${seen[h]})` : h }); });
      const di = S.date ? cols.find((x) => x.name === S.date)?.c : null;
      const ni = cols.find((x) => x.name === '거래처명')?.c;
      const rows = [];
      for (let i = t.start; i < grid.length; i++) {
        const r = grid[i] || [];
        if (S.date) { const d = r[di]; if (!(d instanceof Date) || isNaN(d)) continue; } // 전일잔액·합계 줄 제외
        else { const nm = norm(r[ni]); if (!nm || /합계|소계|총계/.test(nm)) continue; }
        rows.push(cols.map((x) => r[x.c] ?? null));
      }
      if (!rows.length) continue;
      const prev = out[key];
      if (prev) prev.rows.push(...rows); else out[key] = { head: cols.map((x) => x.name), rows, layout: { sheet: sn, headerRow: t.headerRow, start: t.start, cols } };
    }
  }
  cache.set(f.id, out);
  return out;
}

// 보관함의 해당 달(또는 지정한 달들) RD에서 종류별로 모두 합친 표
async function gather(source, months) {
  const { files } = vaultState();
  const pick = files.filter((f) => f.types.includes(source) && (!months.length || months.includes(f.month)));
  let head = null;
  const rows = [];
  const used = [];
  let tpl = null; // 엑셀 서식 본보기: 첫 번째 RD 파일의 그 시트
  for (const f of pick) {
    const t = (await tablesOf(f))[source];
    if (!t) continue;
    if (!head) { head = [...t.head]; tpl = { file: f, ...t.layout }; }
    // 파일마다 열 순서가 다를 수 있어 이름으로 맞춤 (없는 열은 뒤에 추가)
    for (const h of t.head) if (!head.includes(h)) head.push(h);
    const idx = head.map((h) => t.head.indexOf(h));
    for (const r of t.rows) rows.push({ cells: idx.map((k) => (k >= 0 ? r[k] : null)), file: f.name, month: f.month });
    used.push(f.name);
  }
  // 법인 열 추가 (사업장 또는 파일 이름으로 판단)
  if (head && rows.length) {
    const si = head.indexOf('사업장');
    head.push('법인');
    for (const r of rows) r.cells.push(corpOf(si >= 0 ? norm(r.cells[si]) : '', r.file));
  }
  return { head: head || [], rows, used, tpl };
}

// 사업장 → 법인 (원래 양식 맨 끝에 '법인' 열을 붙여 법인별 소계·필터에 씀)
export function corpOf(site, file = '') {
  const v = `${site || ''} ${site ? '' : file}`;
  if (/군자/.test(v)) return '군자';
  if (/송이애/.test(v)) return '송이애';
  if (/디에이치/.test(v)) return 'SPC';
  if (/이천팜|대흥농산\s*중부/.test(v)) return '중부/이천팜';
  if (/대흥|닥터애그/.test(v)) return '대흥';
  return norm(site) || '기타';
}

/* ---------- AI에게 보내는 '열 정보' (금액은 보내지 않음) ---------- */
async function schemaFor(months) {
  const parts = [];
  for (const [key, S] of Object.entries(SOURCES)) {
    const g = await gather(key, months);
    if (!g.rows.length) continue;
    const distinct = (name, max = 300) => {
      const i = g.head.indexOf(name);
      if (i < 0) return [];
      return [...new Set(g.rows.map((r) => norm(r.cells[i])).filter(Boolean))].slice(0, max);
    };
    const info = { source: key, label: S.label, columns: g.head, files: g.used, rowCount: g.rows.length };
    if (key === 'account') {
      const ci = g.head.indexOf('계정코드'), ni = g.head.indexOf('계정명');
      info.accounts = [...new Set(g.rows.map((r) => `${norm(r.cells[ci])} ${norm(r.cells[ni])}`))].slice(0, 400);
      info.sites = distinct('사업장', 60);
      info.corps = distinct('법인', 10);
    }
    if (key === 'partner') { info.accounts = distinct('계정명', 100); }
    if (key === 'receipts') { info.plants = distinct('입고공장', 40); info.groups = distinct('품목군', 60); }
    parts.push(info);
  }
  return parts;
}

const SPEC_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['source', 'filters', 'mode', 'subtotalBy', 'summaryBy', 'summarySum', 'sort', 'title', 'explain', 'conditions'],
  properties: {
    source: { type: 'string', enum: ['account', 'partner', 'receipts'] },
    filters: { type: 'array', items: { type: 'object', additionalProperties: false, required: ['col', 'op', 'values'], properties: {
      col: { type: 'string' }, op: { type: 'string', enum: ['in', 'notIn', 'contains', 'notContains', 'startsWith', 'gte', 'lte', 'nonZero'] }, values: { type: 'array', items: { type: 'string' } } } } },
    mode: { type: 'string', enum: ['rows', 'summary'] },
    subtotalBy: { type: 'string' },
    summaryBy: { type: 'array', items: { type: 'string' } },
    summarySum: { type: 'array', items: { type: 'string' } },
    sort: { type: 'array', items: { type: 'object', additionalProperties: false, required: ['col', 'dir'], properties: { col: { type: 'string' }, dir: { type: 'string', enum: ['asc', 'desc'] } } } },
    title: { type: 'string' },
    explain: { type: 'string' },
    conditions: { type: 'array', items: { type: 'string' } },
  },
};

const SYSTEM = `너는 회사 ERP 원장(RD) 조회 도우미다. 사용자의 한국어 요청을 '조회 조건(JSON)'으로 바꾼다. 데이터 값은 받지 않고 열 이름·계정 목록만 받는다.
규칙:
- source: account=계정별원장(계정코드·계정명·사업장·일자·적요·차변·대변·잔액·관리항목…), partner=거래처원장(거래처별 기초·차변·대변·잔액), receipts=기간별입고현황.
- 계정코드가 14로 시작하면 제조, 15로 시작하면 판관이다. "제조 ○○비"는 계정명에 ○○ 포함 + 계정코드 startsWith 14. 계정은 가능한 한 accounts 목록의 정확한 계정코드로 in 조건을 쓴다.
- 법인/회사는 '법인' 열(값: 대흥, 군자, 송이애, 중부/이천팜, SPC)로 in 조건을 쓴다. 특정 공장·지점은 '사업장' contains. "3개 법인"처럼 전체를 말하면 법인 조건을 넣지 않는다.
- 결산전표(적요에 '결산전표' 포함)는 대체 분개라 금액이 이중으로 보일 수 있다. 사용자가 비용·발생액을 보려 하면 적요 notContains '결산전표' 조건을 넣고 conditions에 '결산전표 제외'를 적는다.
- 기간은 일자 열 gte/lte(YYYY-MM-DD)로. 월 범위는 보관함에서 이미 골라 주므로, 사용자가 특정 날짜를 말할 때만 넣는다.
- mode=rows: 원장 행을 원래 양식 그대로(열 그대로) 보여 줌. 여러 법인이면 subtotalBy='법인'(공장·지점별이면 '사업장'). 소계가 필요 없으면 ''. sort는 보통 [법인 asc, 일자 asc].
- mode=summary: "합계", "~별 합계", "요약" 요청이면 summaryBy(묶을 열들)와 summarySum(더할 숫자 열)을 채운다.
- 금액 조건(예: 잔액 1억 이상)은 gte/lte, 0 아닌 행만은 nonZero.
- 열 이름은 반드시 주어진 columns에 있는 이름을 그대로 쓴다.
- title: 엑셀 제목(짧게), explain: 무엇을 찾는지 한 문장, conditions: 화면에 보일 조건 태그(짧은 한국어, 3~6개).
- 이전 조건이 주어지고 사용자가 고쳐 달라고 하면 이전 조건을 바탕으로 바꾼다.`;

/* ---------- 조건 실행 (이 PC 안에서) ---------- */
function runSpec(spec, g) {
  const ix = (name) => g.head.indexOf(name);
  const S = SOURCES[spec.source];
  const val = (r, c) => r.cells[c];
  const asDate = (v) => (v instanceof Date ? `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}` : null);
  let rows = g.rows.filter((r) => spec.filters.every((f) => {
    const c = ix(f.col);
    if (c < 0) return true; // 없는 열 조건은 무시(화면에 알림)
    const v = val(r, c);
    const sv = norm(v instanceof Date ? asDate(v) : v);
    const vals = f.values.map((x) => norm(x));
    switch (f.op) {
      case 'in': return vals.includes(sv);
      case 'notIn': return !vals.includes(sv);
      case 'contains': return vals.some((x) => sv.toLowerCase().includes(x.toLowerCase()));
      case 'notContains': return !vals.some((x) => sv.toLowerCase().includes(x.toLowerCase()));
      case 'startsWith': return vals.some((x) => sv.startsWith(x));
      case 'gte': return v instanceof Date ? asDate(v) >= vals[0] : num(v) >= num(vals[0]);
      case 'lte': return v instanceof Date ? asDate(v) <= vals[0] : num(v) <= num(vals[0]);
      case 'nonZero': return Math.abs(num(v)) >= 0.5;
      default: return true;
    }
  }));
  const ignored = spec.filters.filter((f) => ix(f.col) < 0).map((f) => f.col);
  const sortKeys = (spec.sort || []).filter((s) => ix(s.col) >= 0);
  const cmp = (a, b) => {
    for (const s of sortKeys) {
      const x = val(a, ix(s.col)), y = val(b, ix(s.col));
      const d = typeof x === 'number' || typeof y === 'number' ? num(x) - num(y) : String(x ?? '').localeCompare(String(y ?? ''), 'ko');
      if (d) return s.dir === 'desc' ? -d : d;
    }
    return 0;
  };
  if (spec.mode === 'summary') {
    const by = spec.summaryBy.filter((c) => ix(c) >= 0);
    const sum = (spec.summarySum.length ? spec.summarySum : S.sum).filter((c) => ix(c) >= 0);
    const m = new Map();
    for (const r of rows) {
      const k = by.map((c) => norm(val(r, ix(c)))).join('\u0001');
      if (!m.has(k)) m.set(k, { key: by.map((c) => val(r, ix(c))), sums: sum.map(() => 0), n: 0 });
      const o = m.get(k); o.n++;
      sum.forEach((c, i) => { o.sums[i] += num(val(r, ix(c))); });
    }
    let list = [...m.values()];
    if (sortKeys.length) {
      const sk = sortKeys[0]; const bi = by.indexOf(sk.col), si = sum.indexOf(sk.col);
      list.sort((a, b) => { const d = si >= 0 ? a.sums[si] - b.sums[si] : String(a.key[bi] ?? '').localeCompare(String(b.key[bi] ?? ''), 'ko'); return sk.dir === 'desc' ? -d : d; });
    }
    const head = [...by, ...sum, '건수'];
    const body = list.map((o) => [...o.key, ...o.sums, o.n]);
    const total = [...by.map((_, i) => (i === 0 ? '합계' : '')), ...sum.map((_, i) => list.reduce((t, o) => t + o.sums[i], 0)), list.reduce((t, o) => t + o.n, 0)];
    return { mode: 'summary', head, body, total, numCols: new Set([...sum.map((_, i) => by.length + i), by.length + sum.length]), rowCount: rows.length, ignored };
  }
  rows = [...rows].sort(sortKeys.length ? cmp : () => 0);
  const sumIdx = S.sum.map((c) => ix(c)).filter((c) => c >= 0);
  const numIdx = new Set(S.nums.map((c) => ix(c)).filter((c) => c >= 0));
  const sb = spec.subtotalBy && ix(spec.subtotalBy) >= 0 ? ix(spec.subtotalBy) : -1;
  const groups = [];
  if (sb >= 0) {
    const order = [];
    const m = new Map();
    for (const r of rows) { const k = norm(val(r, sb)); if (!m.has(k)) { m.set(k, []); order.push(k); } m.get(k).push(r); }
    for (const k of order) groups.push({ key: k, rows: m.get(k) });
  } else groups.push({ key: null, rows });
  const tot = (list) => sumIdx.map((c) => list.reduce((t, r) => t + num(val(r, c)), 0));
  return { mode: 'rows', head: g.head, groups, sumIdx, numIdx, total: tot(rows), totalOf: tot, rowCount: rows.length, ignored };
}

/* ---------- 화면 표 / 엑셀 ---------- */
function tableHTML(res, { subtotal = true, max = 300 } = {}) {
  const th = res.head.map((h) => `<th>${esc(h)}</th>`).join('');
  const cell = (v, isNum) => `<td class="${isNum ? 'r' : ''}">${v instanceof Date ? `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}` : isNum && v !== null && v !== '' ? won(num(v)) : esc(v ?? '')}</td>`;
  let n = 0;
  const lines = [];
  if (res.mode === 'summary') {
    for (const r of res.body) { if (n++ >= max) break; lines.push(`<tr>${r.map((v, i) => cell(v, res.numCols.has(i))).join('')}</tr>`); }
    lines.push(`<tr class="aq-tot">${res.total.map((v, i) => cell(v, res.numCols.has(i))).join('')}</tr>`);
  } else {
    const sumNames = res.sumIdx.map((i) => res.head[i]);
    const subRow = (label, sums, cls) => `<tr class="${cls}"><td colspan="${Math.max(1, Math.min(...res.sumIdx, res.head.length))}">${esc(label)} <span class="aq-sumtxt">${sumNames.map((n, k) => `${esc(n)} ${won(sums[k])}`).join(' · ')}</span></td>${res.head.map((_, i) => i).filter((i) => i >= Math.min(...res.sumIdx, res.head.length)).map((i) => { const k = res.sumIdx.indexOf(i); return `<td class="r">${k >= 0 ? won(sums[k]) : ''}</td>`; }).join('')}</tr>`;
    for (const g of res.groups) {
      for (const r of g.rows) { if (n++ < max) lines.push(`<tr>${r.cells.map((v, i) => cell(v, res.numIdx.has(i))).join('')}</tr>`); }
      if (subtotal && g.key !== null && res.groups.length) lines.push(subRow(`${g.key || '(빈칸)'} 소계`, res.totalOf(g.rows), 'aq-sub'));
    }
    lines.push(subRow('합계', res.total, 'aq-tot'));
  }
  return `<div class="aq-xl"><table><tr>${th}</tr>${lines.join('')}</table></div>${n > max ? `<p class="muted sm">화면에는 ${max}행까지만 보입니다. 엑셀에는 ${res.rowCount.toLocaleString()}행이 모두 들어갑니다.</p>` : ''}`;
}

const colOf0 = (tpl, name) => { const x = tpl.cols.find((y) => y.name === name); return x ? x.c + 1 : null; };
// 원본 RD 파일을 본보기로: 머리글 위쪽(제목·머리글 줄)과 열 너비·서식은 그대로, 데이터 줄만 결과로 바꿈
async function toStyledExcel(res, spec, tpl, { subtotal = true, split = false } = {}) {
  const { default: JSZip } = await import('./vendor/jszip.mjs');
  const book = await Book.open(JSZip, tpl.file.data);
  for (const s0 of [...book.sheets]) if (s0.name !== tpl.sheet) book.removeSheet(s0.name);
  const sh = await book.sheet(tpl.sheet);
  sh.expandShared();
  const first = tpl.start + 1; // 1부터 세는 행 번호
  const rowsNo = [...sh.rows.keys()].sort((a, b) => a - b);
  // 데이터 줄 서식: 첫 데이터 줄, 소계 줄 서식: 원본의 합계·월계·누계 줄(없으면 데이터 줄)
  const dCol = SOURCES[spec.source].date ? colOf0(tpl, SOURCES[spec.source].date) : null;
  const dataRn = rowsNo.find((rn) => rn >= first && (dCol ? cellInfo(sh.rows.get(rn).cells.get(dCol))?.v !== undefined : [...sh.rows.get(rn).cells.values()].some((x) => cellInfo(x)?.v !== undefined))) || first;
  const sumRn = rowsNo.find((rn) => rn >= first && [...sh.rows.get(rn).cells.values()].some((x) => /합\s*계|월\s*계|누\s*계|소\s*계/.test(book.text(cellInfo(x)))));
  // 지우기 전에 본보기 줄(데이터·합계)의 서식을 먼저 기억
  const keep = new Map([dataRn, sumRn].filter(Boolean).map((rn) => [rn, { attrs: sh.rows.get(rn)?.attrs || '', cells: new Map(sh.rows.get(rn)?.cells || []) }]));
  const styleAt = (rn, c) => { const x = keep.get(rn)?.cells.get(c); return x ? (x.match(/\ss="(\d+)"/) || [])[1] : undefined; };
  const colOf = new Map(tpl.cols.map((x) => [x.name, x.c + 1])); // 결과 열 이름 → 원본 열 번호
  const outCols = res.head.map((h) => colOf.get(h)).map((c) => c || null); // '법인' 같은 가상 열은 null
  const allCols = new Set([...(keep.get(dataRn)?.cells.keys() || []), ...outCols.filter(Boolean)]);
  const dataAttrs = keep.get(dataRn)?.attrs || '';
  const firstSumCol = Math.min(...res.sumIdx.map((i) => outCols[i]).filter(Boolean));
  const labelCol = Math.min(...[...allCols]);
  const writeRows = (target, groups, withSub) => {
    for (const rn of [...target.rows.keys()]) if (rn >= first) target.rows.delete(rn);
    let rn = first - 1;
    const putRow = (cellsFn, styleRn) => {
      rn++;
      const cells = new Map();
      for (const c of allCols) cells.set(c, cellsFn(c, `${colLet(c)}${rn}`, styleAt(styleRn, c) ?? styleAt(dataRn, c)));
      target.rows.set(rn, { attrs: (styleRn === dataRn ? dataAttrs : keep.get(styleRn)?.attrs || dataAttrs).replace(/\sr="\d+"/, ` r="${rn}"`) || ` r="${rn}"`, cells });
    };
    const sumLine = (label, sums) => putRow((c, ref, st) => {
      if (c === labelCol) return textCell(ref, st, label);
      const i = res.sumIdx.findIndex((k) => outCols[k] === c);
      return i >= 0 ? numCell(ref, st, Math.round(sums[i] * 100) / 100) : blankCell(ref, st);
    }, sumRn || dataRn);
    for (const g of groups) {
      for (const r of g.rows) {
        putRow((c, ref, st) => {
          const i = outCols.indexOf(c);
          const v = i >= 0 ? r.cells[i] : null;
          return v === null || v === undefined || v === '' ? blankCell(ref, st) : valCell(ref, st, v);
        }, dataRn);
      }
      if (withSub && g.key !== null) sumLine(`${g.key || '(빈칸)'} 소계`, res.totalOf(g.rows));
    }
    sumLine('합계', res.totalOf(groups.flatMap((g) => g.rows)));
    // 필터 범위·표시 범위 맞춤
    target.xml = target.xml.replace(/(<autoFilter ref="[A-Z]+\d+:[A-Z]+)(\d+)"/, (m, a) => `${a}${rn}"`)
      .replace(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g, (m, a, r1) => (+r1 >= first ? '' : m))
      .replace(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g, '');
    target.xml = target.xml.replace(/<mergeCells count="\d+">([\s\S]*?)<\/mergeCells>/, (m, body) => { const n = (body.match(/<mergeCell /g) || []).length; return n ? `<mergeCells count="${n}">${body}</mergeCells>` : ''; });
    void firstSumCol;
  };
  if (split && res.groups.length > 1) {
    const baseXml = sh.xml;
    const baseRows = new Map([...sh.rows].filter(([rn]) => rn < first).map(([rn, r]) => [rn, { attrs: r.attrs, cells: new Map(r.cells) }]));
    let k = 0;
    for (const g of res.groups) {
      const name = (g.key || '(빈칸)').replace(/[\\/?*[\]:]/g, ' ').slice(0, 31);
      if (k === 0) { writeRows(sh, [g], false); book.sheets.find((x) => x.name === tpl.sheet).name = name; book.wbXml = book.wbXml.replace(/(<sheet\b[^>]*name=")[^"]+"/, `$1${name.replace(/&/g, '&amp;').replace(/"/g, '&quot;')}"`); book.cache.delete(tpl.sheet); book.cache.set(name, sh); sh.name = name; }
      else {
        const ns = new Sheet(baseXml, name);
        ns.rows = new Map([...baseRows].map(([rn, r]) => [rn, { attrs: r.attrs, cells: new Map(r.cells) }]));
        writeRows(ns, [g], false);
        await book.addSheet(name, ns, {});
      }
      k++;
    }
  } else writeRows(sh, res.groups, subtotal);
  const data = await book.save();
  const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${(spec.title || 'RD_query').replace(/[\\/:*?"<>|]/g, ' ').trim()}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

async function toExcel(res, spec, { subtotal = true, split = false } = {}) {
  const XLSX = await loadXLSX();
  const wb = XLSX.utils.book_new();
  const fmtCells = (ws, aoa, numCols, startRow) => {
    for (let r = startRow; r < aoa.length; r++) for (const c of numCols) {
      const ref = XLSX.utils.encode_cell({ r, c });
      if (ws[ref] && typeof ws[ref].v === 'number') ws[ref].z = '#,##0';
    }
  };
  const widths = (head, aoa) => head.map((h, c) => ({ wch: Math.min(50, Math.max(String(h).length * 2 + 2, ...aoa.slice(2, 200).map((r) => String(r[c] ?? '').length + 2))) }));
  const add = (name, head, rows, numCols) => {
    const aoa = [[spec.title || 'RD 조회'], head, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'yyyy-mm-dd' });
    fmtCells(ws, aoa, numCols, 2);
    ws['!cols'] = widths(head, aoa);
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 1, c: 0 }, e: { r: 1, c: head.length - 1 } }) };
    XLSX.utils.book_append_sheet(wb, ws, name.replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || 'Sheet');
  };
  if (res.mode === 'summary') {
    add('요약', res.head, [...res.body, res.total], [...res.numCols]);
  } else {
    const firstSum = Math.min(...res.sumIdx, res.head.length);
    const subLine = (label, sums) => res.head.map((_, i) => (i === 0 ? label : res.sumIdx.includes(i) ? sums[res.sumIdx.indexOf(i)] : (i < firstSum ? '' : null)));
    const body = (groups) => {
      const out = [];
      for (const g of groups) {
        for (const r of g.rows) out.push(r.cells);
        if (subtotal && g.key !== null) out.push(subLine(`${g.key || '(빈칸)'} 소계`, res.totalOf(g.rows)));
      }
      return out;
    };
    if (split && res.groups.length > 1) {
      for (const g of res.groups) add(g.key || '(빈칸)', res.head, [...g.rows.map((r) => r.cells), subLine('합계', res.totalOf(g.rows))], [...res.numIdx]);
    } else {
      add(SOURCES[spec.source].label, res.head, [...body(res.groups), subLine('합계', res.total)], [...res.numIdx]);
    }
  }
  const data = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${(spec.title || 'RD_query').replace(/[\\/:*?"<>|]/g, ' ').trim()}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

/* ---------- 채팅 ---------- */
const chat = { items: [], last: null };
function monthsInScope() {
  const { files, month } = vaultState();
  return $('#aqScope').value === 'all' ? [...new Set(files.map((f) => f.month))] : [month];
}
function renderLog() {
  $('#aqLog').innerHTML = chat.items.map((it, i) => {
    if (it.me) return `<div class="aq-me">${esc(it.me)}</div>`;
    if (it.err) return `<div class="aq-bot err">⚠ ${esc(it.err)}</div>`;
    if (it.wait) return `<div class="aq-bot muted">조건을 만드는 중…</div>`;
    const r = it.res;
    return `<div class="aq-bot">
      <div>✅ ${esc(it.spec.explain || '')} — <b>${r.rowCount.toLocaleString()}행</b>${r.mode === 'rows' && r.total.length ? ` (${SOURCES[it.spec.source].sum.filter((c) => r.head.includes(c)).map((c, k) => `${c} ${won(r.total[k])}`).join(' · ')})` : ''}</div>
      <div class="aq-cond">${it.spec.local ? '<span class="loc">AI 없이 해석</span>' : ''}<span>${SOURCES[it.spec.source].label} ${it.used.length}개 파일</span>${it.spec.conditions.map((c) => `<span>${esc(c)}</span>`).join('')}${r.ignored.length ? `<span class="bad">무시한 조건: ${esc(r.ignored.join(', '))}</span>` : ''}</div>
      ${r.rowCount ? tableHTML(r, { subtotal: it.subtotal }) : '<div class="empty">조건에 맞는 행이 없습니다. 계정명·법인 이름을 바꿔 다시 물어보세요.</div>'}
      ${r.rowCount ? `<div class="aq-btns"><button class="btn sm primary" data-aq-xl="${i}">⬇ 엑셀로 받기 (${esc(SOURCES[it.spec.source].label)} 양식)</button>
        ${r.mode === 'rows' && r.groups.length > 1 ? `<button class="btn sm" data-aq-split="${i}">${esc(it.spec.subtotalBy)}별 시트로 나눠 받기</button><button class="btn sm" data-aq-sub="${i}">${it.subtotal ? '소계 빼기' : '소계 넣기'}</button>` : ''}</div>` : ''}
    </div>`;
  }).join('');
  $('#aqLog').scrollTop = $('#aqLog').scrollHeight;
}

/* ---------- AI 없이 해석 (자주 쓰는 요청: 계정·제조/판관·법인·잔액·합계) ---------- */
const CORPS = ['대흥', '군자', '송이애', '중부/이천팜', 'SPC'];
export function localSpec(q, schema, prev) {
  const Q = String(q).replace(/\s+/g, '');
  const has = (re) => re.test(Q);
  // 이어서 고치기: "결산전표 포함/빼줘", "소계 빼줘", "군자만" 등은 이전 조건에서 바꿈
  const base = prev && has(/^(결산전표|소계|군자|대흥|송이애|중부|이천팜|SPC|법인별|합계|요약)/) && !has(/비|원장|잔액|입고/) ? JSON.parse(JSON.stringify(prev)) : null;
  const src = base ? base.source : has(/거래처원장|잔액|외상|채권|채무|선급금/) ? 'partner' : has(/입고|수불|품목/) ? 'receipts' : 'account';
  const info = schema.find((x) => x.source === src);
  if (!info) return null;
  const spec = base || { source: src, filters: [], mode: 'rows', subtotalBy: '', summaryBy: [], summarySum: [], sort: [], title: '', explain: '', conditions: [] };
  const conds = base ? [...spec.conditions] : [];
  const setFilter = (col, op, values) => { spec.filters = spec.filters.filter((f) => !(f.col === col && f.op === op)); if (values) spec.filters.push({ col, op, values }); };
  if (!base && src === 'account') {
    // 계정: '계정명' 앞부분(예: 전력비, 수도광열비, 소모품비(포장))이 질문에 있으면, 바로 앞의 '제조/판관'으로 14·15 구분
    const accs = (info.accounts || []).map((a) => { const k = a.indexOf(' '); return { code: a.slice(0, k), name: a.slice(k + 1) }; });
    const bases = [...new Set(accs.map((a) => a.name.split('-')[0].replace(/\s+/g, '')))].sort((a, b) => b.length - a.length);
    const picked = [];
    let rest = Q;
    for (const b of bases) {
      let at = rest.indexOf(b);
      while (at >= 0) {
        const before = rest.slice(Math.max(0, at - 3), at);
        const pre = /제조/.test(before) ? '14' : /판관|판매/.test(before) ? '15' : '';
        const after = rest.slice(at + b.length, at + b.length + 12);
        const subs = accs.filter((a) => a.name.replace(/\s+/g, '').split('-')[0] === b && (!pre || a.code.startsWith(pre)));
        const narrow = subs.filter((a) => { const sub = (a.name.split('-')[1] || '').replace(/\s+/g, ''); return sub && (after.startsWith(sub) || after.startsWith(`-${sub}`) || after.startsWith(`(${sub}`)); });
        picked.push({ base: b, pre, codes: (narrow.length ? narrow : subs).map((a) => a.code) });
        // 세부 계정 이름(예: 생산전력비)까지 지워서 같은 계정을 두 번 잡지 않게
        const subLen = narrow.length ? Math.max(...narrow.map((a) => (a.name.split('-')[1] || '').replace(/\s+/g, '').length)) : 0;
        const cut = b.length + (narrow.length ? Math.min(after.length, subLen + 1) : 0);
        rest = rest.slice(0, at) + ' '.repeat(cut) + rest.slice(at + cut);
        at = rest.indexOf(b);
      }
    }
    if (!picked.length) return null;
    setFilter('계정코드', 'in', [...new Set(picked.flatMap((p) => p.codes))]);
    picked.forEach((p) => conds.push(`계정: ${p.pre === '14' ? '제조 ' : p.pre === '15' ? '판관 ' : ''}${p.base}`));
    spec.title = `${picked.map((p) => `${p.pre === '14' ? '제조' : p.pre === '15' ? '판관' : ''}${p.base}`).join('·')}`;
  }
  if (!base && src === 'partner') {
    const an = has(/매출|채권/) ? '외상매출금' : has(/매입|채무/) ? '외상매입금' : has(/선급금/) ? '선급금' : '';
    if (an) { setFilter('계정명', 'in', [an]); conds.push(`계정: ${an}`); }
    const m = Q.match(/(\d+(?:\.\d+)?)(억|천만|백만|만)?원?이상/);
    if (m) { const unit = { 억: 1e8, 천만: 1e7, 백만: 1e6, 만: 1e4 }[m[2]] || 1; setFilter('잔액', 'gte', [String(+m[1] * unit)]); conds.push(`잔액 ${m[0].replace('이상', '')} 이상`); spec.sort = [{ col: '잔액', dir: 'desc' }]; }
    spec.title = spec.title || `거래처원장 ${an}`.trim();
  }
  if (!base && src === 'receipts') { spec.title = '기간별입고현황'; }
  // 법인
  const corps = CORPS.filter((c) => c.split('/').some((x) => Q.includes(x)));
  if (corps.length && !has(/\d개법인|전체법인|전법인/)) { setFilter('법인', 'in', corps); conds.push(`법인: ${corps.join(', ')}`); }
  else if (!base && has(/\d개법인|전체/)) conds.push('법인: 전체');
  // 결산전표: 원장 금액 조회는 기본 제외
  if (src === 'account') {
    if (has(/결산전표포함|결산포함/)) { setFilter('적요', 'notContains', null); conds.push('결산전표 포함'); }
    else if (!base || has(/결산전표/)) { setFilter('적요', 'notContains', ['결산전표']); if (!conds.includes('결산전표 제외')) conds.push('결산전표 제외'); }
  }
  // 모양: 합계/요약 → 법인·계정별 합계, 아니면 원장 행 + 법인 소계
  if (has(/합계|요약|집계/) && !has(/원장/)) {
    spec.mode = 'summary';
    spec.summaryBy = src === 'account' ? ['법인', '계정명'] : src === 'partner' ? ['법인', '계정명'] : ['법인', '품목군'];
    spec.summarySum = [];
    conds.push('법인·계정별 합계');
  } else if (!base || has(/소계|법인별/)) {
    spec.mode = 'rows';
    spec.subtotalBy = has(/소계빼|소계없/) ? '' : '법인';
    spec.sort = spec.sort.length && src === 'partner' ? spec.sort : [{ col: '법인', dir: 'asc' }, ...(SOURCES[src].date ? [{ col: SOURCES[src].date, dir: 'asc' }] : [])];
    if (spec.subtotalBy) conds.push('법인별 소계');
  }
  spec.conditions = [...new Set(conds)].filter((c) => !(c === '결산전표 제외' && conds.includes('결산전표 포함')));
  spec.explain = `${SOURCES[src].label}에서 ${spec.title || '요청한 데이터'}`;
  spec.local = true;
  return spec;
}

async function ask(q) {
  if (!q.trim()) return;
  const months = monthsInScope();
  chat.items.push({ me: q }, { wait: true });
  renderLog();
  const slot = chat.items.length - 1;
  try {
    const schema = await schemaFor(months);
    if (!schema.length) throw new Error('이 달 보관함에 조회할 RD(계정별원장·거래처원장·기간별입고현황)가 없습니다. RD를 넣거나 범위를 \'보관함 전체\'로 바꿔 보세요.');
    let spec = null;
    if (aiEnabled()) {
      try {
        const prompt = `보관함 RD 정보(값 없음, 열·목록만):\n${JSON.stringify(schema)}\n\n${chat.last ? `이전 조건:\n${JSON.stringify(chat.last)}\n\n` : ''}요청: ${q}`;
        spec = await aiJson({ system: SYSTEM, prompt, schema: SPEC_SCHEMA });
      } catch (e) { spec = null; console.warn(e); }
    }
    if (!spec) spec = localSpec(q, schema, chat.last);
    if (!spec) throw new Error(aiEnabled()
      ? '요청을 해석하지 못했습니다. 계정명(예: 제조 전력비, 판관 소모품비)이나 거래처원장 계정(외상매출금 등)을 넣어 다시 말해 주세요.'
      : '계정명을 찾지 못했습니다. "제조 전력비", "판관 소모품비(포장)"처럼 계정 이름을 넣어 말해 주세요. (설정에 Claude API 키를 넣으면 더 자유롭게 물어볼 수 있습니다)');
    const g = await gather(spec.source, months);
    const res = runSpec(spec, g);
    chat.last = spec;
    chat.items[slot] = { spec, res, used: g.used, tpl: g.tpl, subtotal: true };
  } catch (e) {
    chat.items[slot] = { err: e.message || String(e) };
  }
  renderLog();
}

export function initRdAsk() {
  if (!$('#aqForm')) return;
  $('#aqForm').addEventListener('submit', (e) => { e.preventDefault(); const q = $('#aqQ').value; $('#aqQ').value = ''; ask(q); });
  $('#aqReset').onclick = () => { chat.items = []; chat.last = null; renderLog(); };
  $('#aqLog').addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    const it = chat.items[+(b.dataset.aqXl ?? b.dataset.aqSplit ?? b.dataset.aqSub)];
    if (!it || !it.res) return;
    const styled = it.res.mode === 'rows' && it.tpl;
    const go = (opts) => (styled ? toStyledExcel(it.res, it.spec, it.tpl, opts) : toExcel(it.res, it.spec, opts)).catch((err) => { console.error(err); toExcel(it.res, it.spec, opts); });
    if (b.dataset.aqXl !== undefined) go({ subtotal: it.subtotal });
    else if (b.dataset.aqSplit !== undefined) go({ subtotal: it.subtotal, split: true });
    else if (b.dataset.aqSub !== undefined) { it.subtotal = !it.subtotal; renderLog(); }
  });
}

// 테스트용 (Node에서 조건 실행만 확인)
export const _test = { toStyledExcel, localSpec, runSpec, gather, tablesOf, schemaFor, toExcel, tableHTML, chat, renderLog };
