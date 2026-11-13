// 월마감 보고서 자동 생성 엔진
// 공통: 지난달 보고서 파일 + ERP RD(전체계정 원장)를 넣으면, 같은 양식에 이번 달을 붙인 파일을 만듦
//  · 채권채무 : 거래처원장(외상매입금·외상매출금) → 'YY년MM월매입채무/매출채권' 시트 + 매입비교·매출비교 월 열
//  · 선급금   : 거래처원장(선급금) → 'YY년MM월선급금' 시트 + 선급금증감 월 열
//  · 내부거래 : 계정별원장 → '비용'·'수익' 시트에 내부거래 행 + '내부거래 집계' 월 행
//  · 포장소모품비: 기간별입고현황 → '수불o', 계정별원장(포장 계정) → '수불x' + 회사별 시트 월 열
import { Book, Sheet, norm, num, colNum, colLet, shiftRefs, mapFormulaXml, cellInfo, styleOfXml, valCell, fCell, textCell, numCell, blankCell, esc, unesc, findTable, toSerial } from './xlsxkit.js';

const pad2 = (n) => String(n).padStart(2, '0');
const nextYm = (y, m) => (m === 12 ? [y + 1, 1] : [y, m + 1]);
const prevYm = (y, m) => (m === 1 ? [y - 1, 12] : [y, m - 1]);

/* ================= RD 읽기 ================= */
// 거래처원장: 거래처명·계정(코드/명)·기초·차변·대변·잔액
export function readPartnerLedgers(XLSX, files) {
  const out = [];
  const used = [];
  for (const { name, buf } of files) {
    const wb = XLSX.read(buf, { cellDates: true });
    let got = 0;
    for (const sn of wb.SheetNames) {
      const grid = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
      const t = findTable(grid, ['거래처명']);
      if (!t) continue;
      const H = t.head;
      const ix = (k) => H.findIndex((h) => h === k);
      const cName = ix('거래처명'), cCode = ix('거래처코드'), cAc = ix('계정코드'), cAn = ix('계정명');
      const cB = ix('기초'), cD = ix('차변'), cC = ix('대변'), cE = ix('잔액'), cV = ix('v');
      if (cName < 0 || (cAc < 0 && cAn < 0) || cE < 0) continue;
      for (let i = t.start; i < grid.length; i++) {
        const r = grid[i] || [];
        const nm = r[cName];
        if (nm === null || nm === undefined || norm(nm) === '' || /합계|소계|총계/.test(norm(nm))) continue;
        const ac = cAc >= 0 ? norm(r[cAc]) : '';
        const an = cAn >= 0 ? norm(r[cAn]) : '';
        if (!ac && !an) continue;
        out.push({ v: cV >= 0 ? r[cV] : 0, code: cCode >= 0 ? r[cCode] : null, name: String(nm), acctCode: ac, acctName: an,
          open: num(r[cB]), dr: num(r[cD]), cr: num(r[cC]), bal: num(r[cE]), src: name });
        got++;
      }
    }
    used.push({ name, rows: got });
  }
  return { rows: out, files: used };
}

// 계정별원장(전체계정): 계정코드·계정명·사업장·일자·적요…·차변·대변·잔액·관리항목1~15
export function readAccountLedgers(XLSX, files) {
  const out = [];
  const used = [];
  for (const { name, buf } of files) {
    const wb = XLSX.read(buf, { cellDates: true });
    let got = 0;
    for (const sn of wb.SheetNames) {
      const grid = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
      const t = findTable(grid, ['계정코드', '일자']);
      if (!t) continue;
      const H = t.head;
      for (let i = t.start; i < grid.length; i++) {
        const r = grid[i] || [];
        const o = { src: name, cells: {} };
        H.forEach((h, c) => { if (h) o.cells[h] = o.cells[h] === undefined ? r[c] : o.cells[h]; });
        // 같은 이름 머리글(관리항목1이 두 번 등)은 순서대로 보관
        o.raw = r; o.head = H;
        const d = o.cells['일자'];
        if (!(d instanceof Date) || isNaN(d)) continue; // 전일잔액·합계 줄
        o.date = d;
        o.acctCode = norm(o.cells['계정코드']);
        out.push(o); got++;
      }
    }
    used.push({ name, rows: got });
  }
  return { rows: out, files: used };
}

// 기간별입고현황
export function readReceipts(XLSX, files) {
  const out = [];
  const used = [];
  for (const { name, buf } of files) {
    const wb = XLSX.read(buf, { cellDates: true });
    let got = 0;
    for (const sn of wb.SheetNames) {
      const grid = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
      const t = findTable(grid, ['입고일자', '품목코드']);
      if (!t) continue;
      for (let i = t.start; i < grid.length; i++) {
        const r = grid[i] || [];
        const d = r[t.head.indexOf('입고일자')];
        if (!(d instanceof Date) || isNaN(d)) continue;
        out.push({ src: name, raw: r, head: t.head, date: d }); got++;
      }
    }
    used.push({ name, rows: got });
  }
  return { rows: out, files: used };
}

/* ================= 공통: 월 시트 찾기 ================= */
function monthSheets(book, suffix) {
  const re = new RegExp(`^(\\d{2})년(\\d{1,2})월${suffix}$`);
  return book.sheets.map((s) => {
    const m = s.name.match(re);
    return m ? { name: s.name, y: 2000 + +m[1], m: +m[2] } : null;
  }).filter(Boolean).sort((a, b) => a.y - b.y || a.m - b.m);
}
const monthName = (y, m, suffix) => `${String(y).slice(2)}년${pad2(m)}월${suffix}`;

// 비교표 시트에서 '월 열' 찾기: 데이터 행 수식이 어느 월 시트를 가리키는지
async function monthColumns(book, cmpName, suffix) {
  const sh = await book.sheet(cmpName);
  const re = new RegExp(`'?((\\d{2})년(\\d{1,2})월${suffix})'?!`);
  const cols = new Map();
  for (const [rn, r] of sh.rows) {
    if (rn > 30) continue;
    for (const [c, x] of r.cells) {
      if (cols.has(c) || !x.includes('<f')) continue;
      const f = cellInfo(x).f || '';
      const m = f.match(re);
      if (m && /SUMIFS|VLOOKUP|XLOOKUP|SUM\(/.test(f)) cols.set(c, { sheet: m[1], y: 2000 + +m[2], m: +m[3] });
    }
  }
  return cols;
}

/* ================= ②③ 채권채무 / 선급금 ================= */
export const ARAP_KINDS = {
  ar: { label: '매출채권', suffix: '매출채권', cmp: '매출비교', codes: ['11110401'], names: ['외상매출금'] },
  ap: { label: '매입채무', suffix: '매입채무', cmp: '매입비교', codes: ['12100101'], names: ['외상매입금'] },
  prepay: { label: '선급금', suffix: '선급금', cmp: '선급금증감', codes: ['11110800'], names: ['선급금'] },
};

// 지난 파일에 어떤 보고서가 들어 있는지
export async function detectBook(JSZip, buf) {
  const book = await Book.open(JSZip, buf);
  const n = book.sheets.map((s) => s.name);
  const kinds = [];
  for (const [k, v] of Object.entries(ARAP_KINDS)) if (n.includes(v.cmp) && monthSheets(book, v.suffix).length) kinds.push(k);
  if (n.includes('비용') && n.includes('수익') && n.includes('내부거래 집계')) kinds.push('internal');
  if (n.includes('수불o') && n.includes('수불x')) kinds.push('packing');
  return { kinds, sheets: n };
}

async function addArApMonth(book, kind, rd, target, report) {
  const K = ARAP_KINDS[kind];
  const months = monthSheets(book, K.suffix);
  const last = months[months.length - 1];
  const [Y, M] = target || nextYm(last.y, last.m);
  const newName = monthName(Y, M, K.suffix);
  if (book.has(newName)) throw new Error(`'${newName}' 시트가 이미 있습니다. 다음 달 RD인지 확인하세요.`);
  const tpl = await book.sheet(last.name);
  tpl.expandShared();

  // RD에서 해당 계정 행
  const rows = rd.rows.filter((r) => K.codes.includes(r.acctCode) || K.names.includes(r.acctName));
  if (!rows.length) { report.warnings.push(`${K.label}: RD에서 '${K.names[0]}'(${K.codes[0]}) 계정 행을 찾지 못해 건너뛰었습니다.`); return null; }

  /* --- 비교표 월 열 --- */
  const cmp = await book.sheet(K.cmp);
  const mcols = await monthColumns(book, K.cmp, K.suffix);
  let prevCol = [...mcols.entries()].filter(([, v]) => v.sheet === last.name).map(([c]) => c).sort((a, b) => b - a)[0];
  if (!prevCol) throw new Error(`${K.cmp} 시트에서 ${last.name} 열을 찾지 못했습니다.`);
  // 지난달 시트의 합계 행(마지막 행)
  const tplData = [...tpl.rows.keys()].filter((rn) => rn >= 3 && norm(book.text(tpl.info('C', rn)))).sort((a, b) => a - b);
  const tplLast = tplData[tplData.length - 1];
  const tplTotal = tplLast + 1;

  /* --- 새 월 시트 --- */
  const ns = new Sheet(tpl.xml.replace(/<mergeCells[\s\S]*?<\/mergeCells>/, ''), newName);
  ns.rows = new Map();
  for (const rn of [1, 2]) if (tpl.rows.has(rn)) ns.rows.set(rn, { attrs: tpl.rows.get(rn).attrs, cells: new Map(tpl.rows.get(rn).cells) });
  const st = (c, rn) => tpl.style(c, rn);
  const first = tplData[0] || 3;
  const fCols = [...(tpl.rows.get(first)?.cells.keys() || [])].filter((c) => c > 9 && tpl.info(c, first)?.f !== undefined);
  let newCol = prevCol + 1;
  const newLast = 2 + rows.length, newTotal = newLast + 1;
  const fixCmp = (f) => {
    // 비교표 지난달 열 → 새 열, VLOOKUP 열 번호 보정
    let g = shiftRefs(f, (ref) => (ref.sheet === K.cmp && ref.col === prevCol ? { ...ref, col: newCol } : ref));
    g = g.replace(new RegExp(`VLOOKUP\\(([^,]+),(${K.cmp})!\\$C:\\$([A-Z]+),\\d+,`), (m, a, s, e) => `VLOOKUP(${a},${s}!$C:$${e},${colNum(e) - 2},`);
    return g;
  };
  rows.forEach((r, i) => {
    const rn = 3 + i;
    const vals = [num(r.v), r.code ?? '', r.name, r.acctCode, r.acctName, r.open, r.dr, r.cr, r.bal];
    vals.forEach((v, k) => {
      const c = k + 1;
      const s = st(c, first);
      if (typeof v === 'number' && k >= 5 && v === 0 && tpl.info(c, first)?.v === undefined && k !== 5 && k !== 8) ns.set(c, rn, (ref) => blankCell(ref, s));
      else ns.set(c, rn, (ref) => valCell(ref, s, k === 1 && v !== '' && !isNaN(Number(v)) ? Number(v) : v));
    });
    for (const c of fCols) {
      const f = shiftRefs(tpl.info(c, first).f, (ref) => ((!ref.sheet || ref.sheet === last.name) && ref.row && !ref.rowAbs ? { ...ref, sheet: ref.sheet ? newName : ref.sheet, row: ref.row + (rn - first) } : ref));
      ns.set(c, rn, (ref) => fCell(ref, st(c, first), fixCmp(f)));
    }
  });
  // 합계 행
  const sum = (k) => rows.reduce((t, r) => t + [r.open, r.dr, r.cr, r.bal][k], 0);
  [6, 7, 8, 9].forEach((c, k) => ns.set(c, newTotal, (ref) => numCell(ref, st(c, tplTotal) ?? st(c, first), Math.round(sum(k) * 100) / 100)));
  for (const [c, x] of tpl.rows.get(tplTotal)?.cells || []) {
    if (c <= 9) continue;
    const i = cellInfo(x);
    if (i.f === undefined) continue;
    const f = shiftRefs(i.f, (ref) => {
      if (ref.sheet || !ref.row) return ref;
      if (ref.row === tplTotal) return { ...ref, row: newTotal };
      if (ref.row === tplLast) return { ...ref, row: newLast };
      return ref;
    });
    ns.set(c, newTotal, (ref) => fCell(ref, i.s, fixCmp(f)));
  }
  if (!tpl.rows.get(tplTotal)?.cells.get(10) && fCols.length) {
    // 지난달 합계 행에 확인 수식이 없으면 9월 양식처럼 넣어 둠
  }
  ns.xml = ns.xml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:K${newTotal}"/>`);

  /* --- 비교표: 열 삽입 + 복사 --- */
  const occupied = [...cmp.rows.values()].some((r) => {
    const i = cellInfo(r.cells.get(newCol));
    return i && (i.f !== undefined || i.v !== undefined || i.isText);
  });
  if (occupied) await book.insertColumns(K.cmp, newCol, 1);
  cmp.expandShared();
  const [py, pm] = [last.y, last.m];
  for (const [rn, r] of cmp.rows) {
    const x = r.cells.get(prevCol);
    if (!x) continue;
    const i = cellInfo(x);
    let nx;
    if (i.f !== undefined) {
      let f = shiftRefs(i.f, (ref) => {
        if (ref.sheet === last.name || ref.sheet === K.cmp) {
          if (ref.sheet === last.name) {
            const o = { ...ref, sheet: newName };
            if (ref.row === tplTotal) o.row = newTotal;
            else if (ref.range && ref.part === 2 && ref.row && ref.row < newLast && ref.row >= 20) o.row = Math.max(ref.row, newLast);
            return o;
          }
          return ref;
        }
        if (!ref.sheet && ref.col && !ref.colAbs) return { ...ref, col: ref.col + 1 };
        return ref;
      });
      f = f.split(`'${last.name}'!`).join(`'${newName}'!`).split(`${last.name}!`).join(`'${newName}'!`);
      nx = fCell(`${colLet(newCol)}${rn}`, i.s, f);
    } else {
      const t = book.text(i);
      let v = i.t === 's' || i.isText || i.t === 'str' ? t : (i.v !== undefined ? Number(i.v) : null);
      // 머리글: 월 숫자·'N월'·연도·'YY년'을 새 달로 (지난달 머리글이 틀려 있어도 바로잡음)
      if (typeof v === 'number' && v >= 1 && v <= 12 && Number.isInteger(v)) v = M;
      else if (typeof v === 'number' && v >= 2000 && v <= 2100) v = Y;
      else if (typeof v === 'string') v = v.replace(/^\d{1,2}월/, `${M}월`).replace(/^\d{2}년$/, `${String(Y).slice(2)}년`);
      nx = valCell(`${colLet(newCol)}${rn}`, i.s, v);
    }
    r.cells.set(newCol, nx);
  }
  // 오른쪽 '차액' 열: 지난달-전전달 → 이번달-지난달
  const diffCol = newCol + 1;
  for (const r of cmp.rows.values()) {
    const x = r.cells.get(diffCol);
    if (!x || !x.includes('<f')) continue;
    const f = cellInfo(x).f;
    const refs = []; shiftRefs(f, (ref) => { if (!ref.sheet) refs.push(ref.col); return ref; });
    if (refs.length && refs.every((c) => c === prevCol || c === prevCol - 1)) r.cells.set(diffCol, mapFormulaXml(x, (ref) => (!ref.sheet && ref.col ? { ...ref, col: ref.col + 1 } : ref)));
  }

  /* --- 새 거래처: 비교표 목록 끝에 행 추가 --- */
  const listRows = [...cmp.rows.keys()].filter((rn) => rn >= 3 && cmp.info(prevCol, rn)?.f && norm(book.text(cmp.info(3, rn))));
  const known = new Set(listRows.map((rn) => norm(book.text(cmp.info(3, rn)))));
  const fresh = [...new Set(rows.filter((r) => !known.has(norm(r.name))).map((r) => r.name))];
  if (fresh.length) {
    // 목록은 위에서부터 이어진 덩어리(중간 빈 줄 3줄 이하)의 끝
    const sorted = listRows.sort((a, b) => a - b);
    let Lrow = sorted[0];
    for (const rn of sorted) { if (rn - Lrow > 4) break; Lrow = rn; }
    const at = Lrow + 1;
    await book.insertRows(K.cmp, at, fresh.length);
    // 목록 끝 행을 포함하던 합계 범위(…:열Lrow)를 늘림
    for (const r of cmp.rows.values()) for (const [c, x] of r.cells) {
      if (!x.includes('<f')) continue;
      r.cells.set(c, mapFormulaXml(x, (ref) => (!ref.sheet || ref.sheet === K.cmp) && ref.range && ref.part === 2 && ref.row === Lrow ? { ...ref, row: Lrow + fresh.length } : ref));
    }
    const src = cmp.rows.get(Lrow);
    fresh.forEach((nm, k) => {
      const rn = at + k;
      const cells = new Map();
      for (const [c, x] of src.cells) {
        const i = cellInfo(x);
        if (c === 3) cells.set(c, textCell(`C${rn}`, i.s, nm));
        else if (c === 2) cells.set(c, textCell(`B${rn}`, i.s, '신규'));
        else if (i.f !== undefined) cells.set(c, fCell(`${colLet(c)}${rn}`, i.s, shiftRefs(i.f, (ref) => (ref.row && !ref.rowAbs && (!ref.sheet || ref.sheet === K.cmp) && ref.row === Lrow ? { ...ref, row: rn } : ref))));
        else cells.set(c, blankCell(`${colLet(c)}${rn}`, i.s));
      }
      cmp.rows.set(rn, { attrs: src.attrs.replace(/\sr="\d+"/, ` r="${rn}"`), cells });
    });
    const amt = (nm) => rows.filter((r) => r.name === nm).reduce((t, r) => t + r.bal, 0);
    report.warnings.push(`${K.cmp}: 새 거래처 ${fresh.length}곳을 목록 끝(${at}행~)에 추가했습니다. B열 구분을 '신규'로 넣었으니 바꿔 주세요 — ${fresh.map((n) => `${n}(잔액 ${Math.round(amt(n)).toLocaleString()})`).join(', ')}`);
  }

  /* --- 새 시트 위치: 지난달 시트 옆 (월 순서 방향 유지) --- */
  const names = book.names();
  const li = names.indexOf(last.name);
  const prevOfLast = monthName(...prevYm(last.y, last.m), K.suffix);
  const descending = names[li + 1] && (names[li + 1] === prevOfLast || names[li + 1].replace(/년0?/, '년') === prevOfLast.replace(/년0?/, '년'));
  await book.addSheet(newName, ns, descending ? { before: last.name } : { after: last.name });

  const bal = rows.reduce((t, r) => t + r.bal, 0);
  report.parts.push({ kind, label: K.label, sheet: newName, rows: rows.length, balance: bal, prevSheet: last.name, cmp: K.cmp, col: colLet(newCol), fresh: fresh.length });
  return newName;
}

// 채권채무·선급금 파일 만들기
export async function buildArAp({ XLSX, JSZip, prevBuf, rdFiles, target }) {
  const book = await Book.open(JSZip, prevBuf);
  const rd = readPartnerLedgers(XLSX, rdFiles);
  const report = { warnings: [], info: [], parts: [], rdFiles: rd.files };
  if (!rd.rows.length) throw new Error('RD에서 거래처원장(거래처명·계정·잔액 열)을 찾지 못했습니다.');
  const kinds = Object.keys(ARAP_KINDS).filter((k) => book.has(ARAP_KINDS[k].cmp) && monthSheets(book, ARAP_KINDS[k].suffix).length);
  if (!kinds.length) throw new Error('채권채무(매입비교·매출비교) 또는 선급금(선급금증감) 파일이 아닙니다.');
  let lastSheet = null;
  for (const k of kinds) {
    const s = await addArApMonth(book, k, rd, target, report);
    if (s) lastSheet = s;
  }
  if (!report.parts.length) throw new Error('RD에 해당 계정이 없어 만든 시트가 없습니다.');
  const p0 = report.parts[0];
  const ym = p0.sheet.match(/(\d{2})년(\d{2})월/);
  const data = await book.save({ activeSheet: lastSheet });
  const base = kinds.includes('prepay') ? 'Prepaid' : 'ARAP';
  return { data, report, fileName: `${base}_${ym[1]}${ym[2]}.xlsx` };
}

/* ================= ① 내부거래 집계 ================= */
// 상대 회사 이름(관리항목1) → 집계 구분. 지난 행에서 배운 것을 먼저 쓰고, 없으면 이름으로 판단
const COUNTERPARTY_RULES = [
  [/대흥농산|닥터애그/, '대흥'], [/송이애/, '송이애'], [/이천팜|대흥농산\s*중부/, '중부/이천팜'], [/디에이치/, 'SPC'], [/군자농원/, '군자'],
];
const bizNo = (s) => (String(s ?? '').match(/\(?(\d{3}-\d{2}-\d{5})\)?/) || [])[1] || '';
const partyKey = (s) => bizNo(s) || norm(s).replace(/\(군자\)$/, '');
const memoKey = (s) => norm(s).replace(/\d+월(분)?/g, '').replace(/\(\d+월\)/g, '').replace(/\d{4}-\d{2}/g, '');
function classifyContent(memo) {
  if (/임대료|임차료/.test(memo)) return '물류센터임차료';
  if (/위탁판매|수탁판매/.test(memo) && /판매수수료|위탁판매수수료/.test(memo)) return '위탁판매수수료';
  return '업무위탁수수료';
}

async function addInternalMonth(book, rd, report) {
  const map = await book.sheet('mapping');
  const siteMap = new Map();
  for (const [rn] of map.rows) { const e = norm(book.text(map.info('E', rn))); if (rn > 1 && e) siteMap.set(e, norm(book.text(map.info('F', rn)))); }
  let target = null;
  const results = {};
  for (const sname of ['비용', '수익']) {
    const sh = await book.sheet(sname);
    sh.expandShared();
    const head = new Map(); // 머리글 이름 → 열 (J열부터 RD 열)
    for (const [c, x] of sh.rows.get(1).cells) if (c >= 10) { const h = norm(book.text(cellInfo(x))); if (h && !head.has(h)) head.set(h, c); }
    const dataRows = [...sh.rows.keys()].filter((r) => r >= 2).sort((a, b) => a - b);
    const lastRow = dataRows[dataRows.length - 1];
    const dateOf = (rn) => { const v = book.val(sh, 'M', rn); return typeof v === 'number' ? new Date(Math.round((v - 25569) * 864e5)) : null; };
    // 이력: 계정, 상대회사, 내용
    const partyCol = book.val(sh, 'H', lastRow) !== null && sh.info('H', lastRow)?.f === undefined ? 'H' : 'I';
    const hist = new Map(), partyMap = new Map(), accts = new Set();
    let lastDate = null;
    for (const rn of dataRows) {
      const d = dateOf(rn);
      if (!d) continue;
      if (!lastDate || d > lastDate) lastDate = d;
    }
    const recentFrom = lastDate ? new Date(lastDate.getFullYear(), lastDate.getMonth() - 5, 1) : null;
    for (const rn of dataRows) {
      const d = dateOf(rn);
      const ac = norm(book.val(sh, 'J', rn));
      const x = book.val(sh, 'X', rn);
      const p = norm(book.val(sh, partyCol, rn));
      const D = norm(book.val(sh, 'D', rn));
      if (x && p && p !== '-') partyMap.set(partyKey(x), p);
      if (d && d >= recentFrom && D) accts.add(ac);
      if (D) hist.set(`${ac}|${norm(book.val(sh, 'L', rn))}|${partyKey(x)}|${memoKey(book.val(sh, 'N', rn))}`, D);
    }
    // RD에서 이 시트 계정 + 상대가 내부회사인 행
    const pick = rd.rows.filter((r) => accts.has(r.acctCode));
    const party = (x) => partyMap.get(partyKey(x)) || (COUNTERPARTY_RULES.find(([re]) => re.test(String(x ?? ''))) || [])[1] || null;
    const rows = [];
    for (const r of pick) {
      const memo = norm(r.cells['적요']);
      if (/결산전표|마감손익/.test(memo)) continue;
      const x = r.cells['관리항목1'];
      const p = party(x);
      if (!p) continue;
      rows.push({ r, p, memo });
    }
    if (!rows.length) { report.warnings.push(`${sname}: RD에서 내부거래 행(계정 ${[...accts].join(', ')})을 찾지 못했습니다.`); results[sname] = null; continue; }
    const months = [...new Set(rows.map(({ r }) => `${r.date.getFullYear()}-${r.date.getMonth() + 1}`))];
    const ym = months.map((s) => s.split('-').map(Number)).sort((a, b) => b[0] - a[0] || b[1] - a[1])[0];
    if (!target) target = ym;
    const [Y, M] = target;
    const mine = rows.filter(({ r }) => r.date.getFullYear() === Y && r.date.getMonth() + 1 === M);
    // 이미 같은 달 행이 끝에 있으면 지우고 다시
    let keepTo = lastRow;
    while (keepTo >= 2) { const d = dateOf(keepTo); if (d && d.getFullYear() === Y && d.getMonth() + 1 === M) keepTo--; else break; }
    const removed = lastRow - keepTo;
    for (let rn = keepTo + 1; rn <= lastRow; rn++) sh.rows.delete(rn);
    const tplRow = lastRow; // 지운 행이어도 서식·수식 본보기로 씀(복사해 둔 값 사용)
    const tpl = new Map(sh.rows.get(keepTo).cells);
    const tplRn = keepTo;
    let rn = keepTo;
    const byRule = [];
    for (const { r, p, memo } of mine) {
      rn++;
      const key = `${r.acctCode}|${norm(r.cells['사업장'])}|${partyKey(r.cells['관리항목1'])}|${memoKey(memo)}`;
      let D = hist.get(key);
      if (!D) { D = classifyContent(memo); byRule.push(`${norm(r.cells['사업장'])} · ${memo.slice(0, 40)} → ${D}`); }
      if (!siteMap.has(norm(r.cells['사업장']))) report.warnings.push(`${sname}: 사업장 '${norm(r.cells['사업장'])}'이 mapping 시트 E열에 없습니다 (C열이 #N/A).`);
      for (const [c, x] of tpl) {
        const i = cellInfo(x);
        const ref = `${colLet(c)}${rn}`;
        if (c >= 10) {
          const h = [...head.entries()].find(([, cc]) => cc === c)?.[0];
          let v = h ? r.cells[h] : null;
          if (h === '관리항목1' || /^관리항목\d+$/.test(h || '')) v = r.cells[h];
          if (v instanceof Date) { sh.set(c, rn, numCell(ref, i.s, toSerial(v))); continue; }
          if (['차변', '대변', '잔액', '환율'].includes(h) && v !== null && v !== undefined && v !== '') v = num(v);
          else if (typeof v === 'number' && ['행', '전표번호', '승인번호', '계정코드'].includes(h)) v = String(v);
          sh.set(c, rn, valCell(ref, i.s, v ?? null));
        } else if (c === 4) sh.set(c, rn, textCell(ref, i.s, D));
        else if (colLet(c) === partyCol) sh.set(c, rn, textCell(ref, i.s, p));
        else if (i.f !== undefined) sh.set(c, rn, fCell(ref, i.s, shiftRefs(i.f, (q) => ((!q.sheet || q.sheet === sname) && q.row && !q.rowAbs ? { ...q, row: q.row + (rn - tplRn) } : q))));
        else sh.set(c, rn, blankCell(ref, i.s));
      }
      sh.rows.get(rn).attrs = sh.rows.get(tplRn).attrs.replace(/\sr="\d+"/, ` r="${rn}"`);
    }
    // 필터 범위 늘림
    sh.xml = sh.xml.replace(/(<autoFilter ref="[A-Z]+\d+:[A-Z]+)(\d+)"/, (m, a) => `${a}${rn}"`);
    results[sname] = { rows: mine.length, removed, byRule, total: mine.reduce((t, { r }) => t + num(r.cells[sname === '비용' ? '차변' : '대변']), 0) };
    if (byRule.length) report.warnings.push(`${sname}: 지난 기록에 없는 거래 ${byRule.length}건은 적요로 D열(내용)을 정했습니다 — 확인해 주세요: ${byRule.slice(0, 8).join(' / ')}${byRule.length > 8 ? ' …' : ''}`);
    void tplRow;
  }
  if (!target) throw new Error('RD에서 내부거래 행을 찾지 못했습니다.');
  const [Y, M] = target;

  /* --- 내부거래 집계: 마지막 달 묶음을 복사해 새 달 행 --- */
  const ag = await book.sheet('내부거래 집계');
  ag.expandShared();
  const rowsA = [...ag.rows.keys()].filter((r) => r >= 7 && typeof book.val(ag, 'B', r) === 'number' && typeof book.val(ag, 'C', r) === 'number').sort((a, b) => a - b);
  const last = rowsA[rowsA.length - 1];
  const ly = book.val(ag, 'B', last), lm = book.val(ag, 'C', last);
  let blockStart = last;
  while (blockStart - 1 >= 7 && book.val(ag, 'B', blockStart - 1) === ly && book.val(ag, 'C', blockStart - 1) === lm) blockStart--;
  let aggRows = 0;
  if (ly === Y && lm === M) {
    report.info.push(`내부거래 집계: ${Y}년 ${M}월 행이 이미 있어 그대로 두었습니다 (수식이 비용·수익 시트를 다시 읽음).`);
  } else {
    const size = last - blockStart + 1;
    for (let k = 0; k < size; k++) {
      const src = ag.rows.get(blockStart + k);
      const rn = last + 1 + k;
      const cells = new Map();
      for (const [c, x] of src.cells) {
        const i = cellInfo(x);
        const ref = `${colLet(c)}${rn}`;
        if (c === 2) cells.set(c, numCell(ref, i.s, Y));
        else if (c === 3) cells.set(c, numCell(ref, i.s, M));
        else if (i.f !== undefined) cells.set(c, fCell(ref, i.s, shiftRefs(i.f, (q) => (!q.sheet && q.row && !q.rowAbs ? { ...q, row: q.row + size } : q))));
        else cells.set(c, x.replace(/^<c r="[A-Z]+\d+"/, `<c r="${ref}"`).replace(/<v>[\s\S]*?<\/v>/, (m) => m));
      }
      ag.rows.set(rn, { attrs: src.attrs.replace(/\sr="\d+"/, ` r="${rn}"`), cells });
    }
    aggRows = size;
  }
  report.parts.push({ kind: 'internal', label: '내부거래', y: Y, m: M, cost: results['비용'], income: results['수익'], aggRows });
  return target;
}

export async function buildInternal({ XLSX, JSZip, prevBuf, rdFiles }) {
  const book = await Book.open(JSZip, prevBuf);
  for (const n of ['비용', '수익', '내부거래 집계', 'mapping']) if (!book.has(n)) throw new Error(`내부거래 집계 파일이 아닙니다 ('${n}' 시트 없음).`);
  const rd = readAccountLedgers(XLSX, rdFiles);
  if (!rd.rows.length) throw new Error('RD에서 계정별원장(계정코드·일자 열)을 찾지 못했습니다.');
  const report = { warnings: [], info: [], parts: [], rdFiles: rd.files };
  const [Y, M] = await addInternalMonth(book, rd, report);
  const data = await book.save({ activeSheet: '내부거래 집계' });
  return { data, report, fileName: `Internal_${String(Y).slice(2)}${pad2(M)}.xlsx` };
}

/* ================= ④ 포장소모품비 ================= */
// 원장 시트(수불o/수불x)에 RD 행을 붙임: 머리글 이름으로 열을 맞추고, 오른쪽 계산 열(year/month/…)은 마지막 행 수식을 복사
async function appendRows(book, sname, items, { headRows, dateCol, Y, M }) {
  const sh = await book.sheet(sname);
  sh.expandShared();
  const head = new Map();
  for (const hr of [...headRows].reverse()) for (const [c, x] of sh.rows.get(hr)?.cells || []) { const h = norm(book.text(cellInfo(x))); if (h && !head.has(c)) head.set(c, h); }
  const rowsAll = [...sh.rows.keys()].sort((a, b) => a - b);
  const firstData = Math.max(...headRows) + 1;
  const dateOf = (rn) => { const v = book.val(sh, dateCol, rn); return typeof v === 'number' ? new Date(Math.round((v - 25569) * 864e5)) : null; };
  let last = rowsAll[rowsAll.length - 1];
  while (last > firstData && !dateOf(last)) last--; // 끝의 빈 줄
  let keepTo = last;
  while (keepTo > firstData) { const d = dateOf(keepTo); if (d && d.getFullYear() === Y && d.getMonth() + 1 === M) keepTo--; else break; }
  const tplRn = keepTo;
  const tpl = new Map(sh.rows.get(tplRn).cells);
  for (const rn of rowsAll) if (rn > keepTo) sh.rows.delete(rn);
  let rn = keepTo;
  for (const it of items) {
    rn++;
    const cells = new Map();
    for (const [c, x] of tpl) {
      const i = cellInfo(x);
      const ref = `${colLet(c)}${rn}`;
      if (i.f !== undefined) { cells.set(c, fCell(ref, i.s, shiftRefs(i.f, (q) => ((!q.sheet || q.sheet === sname) && q.row && !q.rowAbs ? { ...q, row: q.row + (rn - tplRn) } : q)))); continue; }
      const h = head.get(c);
      let v = h !== undefined ? it.get(h, c) : null;
      if (i.f !== undefined && head.has(c) && /^(year|month)$/.test(h)) { cells.set(c, fCell(ref, i.s, shiftRefs(i.f, (q) => ((!q.sheet || q.sheet === sname) && q.row && !q.rowAbs ? { ...q, row: q.row + (rn - tplRn) } : q)))); continue; }
      if (v instanceof Date) cells.set(c, numCell(ref, i.s, toSerial(v)));
      else cells.set(c, valCell(ref, i.s, v ?? null));
    }
    // 본보기 행에 없는 열의 값도 넣기
    for (const [c, h] of head) {
      if (cells.has(c)) continue;
      const v = it.get(h, c);
      if (v === null || v === undefined || v === '') continue;
      cells.set(c, v instanceof Date ? numCell(`${colLet(c)}${rn}`, undefined, toSerial(v)) : valCell(`${colLet(c)}${rn}`, undefined, v));
    }
    sh.rows.set(rn, { attrs: sh.rows.get(tplRn)?.attrs?.replace(/\sr="\d+"/, ` r="${rn}"`) || ` r="${rn}"`, cells });
  }
  sh.xml = sh.xml.replace(/(<autoFilter ref="[A-Z]+\d+:[A-Z]+)(\d+)"/, (m, a, e) => (+e > 4 ? `${a}${rn}"` : m));
  return { added: items.length, removed: last - keepTo };
}

// 회사별 시트: 월마다 3칸(수불o·수불x판관·수불x제조) 묶음 + (있으면) 왼쪽 월별 요약 열
async function addPackingMonthCols(book, sname, others, Y, M, report) {
  const sh = await book.sheet(sname);
  sh.expandShared();
  const r3 = sh.rows.get(3)?.cells || new Map();
  const monthOf = (c) => { const v = book.val(sh, c, 3); return typeof v === 'number' ? v : null; };
  const cols3 = [...r3.keys()].filter((c) => monthOf(c) !== null).sort((a, b) => a - b);
  if (!cols3.length) return null;
  const lm = Math.max(...cols3.map(monthOf));
  if (lm === M) { report.info.push(`${sname}: ${M}월 열이 이미 있습니다.`); return null; }
  let gStart = cols3.find((c) => monthOf(c) === lm);
  const gSize = cols3.filter((c) => monthOf(c) === lm).length;
  const firstGroup = cols3[0];
  // 요약 열: 묶음 왼쪽, 7행에 월 번호(1,=D7+1,…)
  const sumCols = [...(sh.rows.get(7)?.cells.keys() || [])].filter((c) => c < firstGroup && (book.val(sh, c, 7) !== null && typeof book.val(sh, c, 7) === 'number')).sort((a, b) => a - b);
  let sumLast = null;
  if (sumCols.length) {
    sumLast = sumCols[sumCols.length - 1];
    await book.insertColumns(sname, sumLast + 1, 1);
    gStart += 1;
  }
  const detailStart = firstGroup + (sumLast ? 1 : 0);
  const newStart = gStart + gSize;
  // 새 묶음 자리가 차 있으면 밀기
  const occupied = [...sh.rows.values()].some((r) => [...Array(gSize).keys()].some((k) => { const i = cellInfo(r.cells.get(newStart + k)); return i && (i.f !== undefined || i.v !== undefined || i.isText); }));
  if (occupied) await book.insertColumns(sname, newStart, gSize);
  const fam = new Set([sname, ...others]);
  for (const [rn, r] of sh.rows) {
    for (let k = 0; k < gSize; k++) {
      const x = r.cells.get(gStart + k);
      if (!x) continue;
      const i = cellInfo(x);
      const ref = `${colLet(newStart + k)}${rn}`;
      if (i.f !== undefined) {
        const f = shiftRefs(i.f, (q) => ((!q.sheet || fam.has(q.sheet)) && q.col && !q.colAbs ? { ...q, col: q.col + gSize } : q));
        r.cells.set(newStart + k, fCell(ref, i.s, f));
      } else {
        let v = i.t === 's' || i.isText || i.t === 'str' ? book.text(i) : (i.v !== undefined ? Number(i.v) : null);
        if (rn === 2 && typeof v === 'number') v = Y;
        else if ((rn === 3 || rn === 6) && typeof v === 'number') v = M;
        r.cells.set(newStart + k, valCell(ref, i.s, v));
      }
    }
  }
  // 6행 병합(월 이름) 추가
  sh.xml = sh.xml.replace(/<mergeCells count="(\d+)">/, (m, n) => `<mergeCells count="${+n + 1}">`).replace('</mergeCells>', `<mergeCell ref="${colLet(newStart)}6:${colLet(newStart + gSize - 1)}6"/></mergeCells>`);
  const detailEnd = newStart + gSize - 1;
  if (sumLast) {
    const nc = sumLast + 1;
    for (const [rn, r] of sh.rows) {
      const x = r.cells.get(sumLast);
      if (!x) continue;
      const i = cellInfo(x);
      const ref = `${colLet(nc)}${rn}`;
      if (i.f !== undefined) r.cells.set(nc, fCell(ref, i.s, shiftRefs(i.f, (q) => (!q.sheet && q.col && !q.colAbs ? { ...q, col: q.col + 1 } : q))));
      else r.cells.set(nc, valCell(ref, i.s, i.t === 's' || i.isText || i.t === 'str' ? book.text(i) : (i.v !== undefined ? Number(i.v) : null)));
    }
    // 요약 수식 범위를 전체 월 묶음으로
    for (const r of sh.rows.values()) for (let c = sumCols[0]; c <= nc; c++) {
      const x = r.cells.get(c);
      if (!x || !x.includes('<f')) continue;
      r.cells.set(c, mapFormulaXml(x, (q) => (!q.sheet && q.range && q.part === 2 && q.col && q.col >= detailStart ? { ...q, col: detailEnd } : q)));
    }
  }
  return { sheet: sname, col: colLet(newStart), summaryCol: sumLast ? colLet(sumLast + 1) : null };
}

export async function buildPacking({ XLSX, JSZip, prevBuf, rdFiles }) {
  const book = await Book.open(JSZip, prevBuf);
  for (const n of ['수불o', '수불x', 'Mapping']) if (!book.has(n)) throw new Error(`포장소모품비 파일이 아닙니다 ('${n}' 시트 없음).`);
  const report = { warnings: [], info: [], parts: [] };
  const rec = readReceipts(XLSX, rdFiles);
  const led = readAccountLedgers(XLSX, rdFiles);
  report.rdFiles = rdFiles.map((f) => ({ name: f.name, rows: (rec.files.find((x) => x.name === f.name)?.rows || 0) + (led.files.find((x) => x.name === f.name)?.rows || 0) }));
  // 포장 계정: Mapping 시트 P열
  const mp = await book.sheet('Mapping');
  const codes = new Set();
  for (const rn of mp.rows.keys()) { const v = book.val(mp, 'P', rn); if (v !== null && /^\d{6,}$/.test(String(v))) codes.add(String(v)); }
  const ledRows = led.rows.filter((r) => codes.has(r.acctCode));
  const dates = [...rec.rows.map((r) => r.date), ...ledRows.map((r) => r.date)];
  if (!dates.length) throw new Error('RD에서 기간별입고현황 또는 포장 계정(계정별원장) 행을 찾지 못했습니다.');
  const maxD = new Date(Math.max(...dates));
  const Y = maxD.getFullYear(), M = maxD.getMonth() + 1;
  const inMonth = (d) => d.getFullYear() === Y && d.getMonth() + 1 === M;
  // 수불o
  const recM = rec.rows.filter((r) => inMonth(r.date));
  if (recM.length) {
    const res = await appendRows(book, '수불o', recM.map((r) => ({ get: (h) => { const k = r.head.indexOf(h); return k >= 0 ? r.raw[k] : null; } })), { headRows: [3, 4], dateCol: 'A', Y, M });
    report.parts.push({ kind: 'packing', label: '수불o (기간별입고현황)', rows: res.added, removed: res.removed, total: recM.reduce((t, r) => t + num(r.raw[r.head.indexOf('입고금액')]), 0) });
  } else report.warnings.push(`수불o: RD에 ${M}월 기간별입고현황이 없어 건너뛰었습니다.`);
  // 수불x
  const ledM = ledRows.filter((r) => inMonth(r.date));
  if (ledM.length) {
    const res = await appendRows(book, '수불x', ledM.map((r) => ({ get: (h) => {
      const v = r.cells[h];
      if (typeof v === 'number' && ['행', '전표번호', '승인번호', '계정코드'].includes(h)) return String(v);
      return v ?? null;
    } })), { headRows: [3, 4], dateCol: 'D', Y, M });
    report.parts.push({ kind: 'packing', label: '수불x (포장 계정 원장)', rows: res.added, removed: res.removed, total: ledM.reduce((t, r) => t + num(r.cells['차변']), 0) });
  } else report.warnings.push(`수불x: RD 계정별원장에 포장 계정(${[...codes].join(', ')}) ${M}월 행이 없어 건너뛰었습니다.`);
  // 회사별 시트
  const comp = [];
  for (const s of book.sheets) {
    if (['수불o', '수불x', 'Mapping', '포장소모품비 예산정리'].includes(s.name)) continue;
    const sh = await book.sheet(s.name);
    if (sh.rows.get(7) && [...(sh.rows.get(7).cells.values())].some((x) => /수불o/.test(book.text(cellInfo(x))))) comp.push(s.name);
  }
  // 다른 시트를 참조하는 합계 시트는 나중에(참조 대상 열이 먼저 생기도록)
  for (const n of comp) {
    const r = await addPackingMonthCols(book, n, comp.filter((x) => x !== n), Y, M, report);
    if (r) report.parts.push({ kind: 'packing', label: `${n} 시트 ${M}월 열`, rows: null, col: r.col, summaryCol: r.summaryCol });
  }
  report.warnings.push(`'포장소모품비 예산정리' 시트의 ${M}월 금액(백만원)은 손으로 넣는 칸이라 비워 두었습니다.`);
  const data = await book.save({ activeSheet: comp[0] });
  return { data, report, fileName: `Packing_${String(Y).slice(2)}${pad2(M)}.xlsx` };
}
