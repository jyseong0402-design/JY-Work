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
      const cB = ix('기초'), cD = ix('차변'), cC = ix('대변'), cE = ix('잔액'), cV = ix('v'), cSite = ix('사업장');
      if (cName < 0 || (cAc < 0 && cAn < 0) || cE < 0) continue;
      for (let i = t.start; i < grid.length; i++) {
        const r = grid[i] || [];
        const nm = r[cName];
        if (nm === null || nm === undefined || norm(nm) === '' || /합계|소계|총계/.test(norm(nm))) continue;
        const ac = cAc >= 0 ? norm(r[cAc]) : '';
        const an = cAn >= 0 ? norm(r[cAn]) : '';
        if (!ac && !an) continue;
        out.push({ v: cV >= 0 ? r[cV] : 0, code: cCode >= 0 ? r[cCode] : null, name: String(nm), acctCode: ac, acctName: an, site: cSite >= 0 ? norm(r[cSite]) : '',
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
  if (n.includes('매출채권') && n.includes('매입채무') && n.includes('맵핑')) kinds.push('ar', 'ap', 'arapList');
  if (n.includes('닥터애그(CAPEX Actual)')) kinds.push('capex');
  if (n.includes('수불o') && n.includes('수불x')) kinds.push('packing');
  if (n.some((x) => /raw\s*materials/i.test(x))) kinds.push('rawmat');
  if (n.some((x) => /전력비/.test(x)) && n.some((x) => /가스|등유|벙커/.test(x))) kinds.push('energy');
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

  /* --- 범위 확인: 이번 RD '기초' = 지난달 시트 '잔액' 이어야 같은 법인 범위로 받은 것 --- */
  {
    const scope = norm(book.text(cmp.info('A', 1))).replace(/^ERP\s*>\s*/, '');
    if (scope) report.info.push(`${K.label}: 지난달 파일의 거래처원장 조회 범위 — ${scope}`);
    const prevBal = new Map();
    let prevSum = 0;
    for (const rn of tplData) {
      const nm = norm(book.text(tpl.info('C', rn)));
      const v = num(book.val(tpl, 'I', rn));
      prevBal.set(nm, (prevBal.get(nm) || 0) + v);
      prevSum += v;
    }
    const openBal = new Map();
    let openSum = 0;
    for (const r of rows) { const nm = norm(r.name); openBal.set(nm, (openBal.get(nm) || 0) + r.open); openSum += r.open; }
    const diffs = [];
    for (const nm of new Set([...prevBal.keys(), ...openBal.keys()])) {
      const d = (openBal.get(nm) || 0) - (prevBal.get(nm) || 0);
      if (Math.abs(d) >= 1) diffs.push({ nm, prev: prevBal.get(nm) || 0, open: openBal.get(nm) || 0, d });
    }
    diffs.sort((a, b) => Math.abs(b.d) - Math.abs(a.d));
    const sites = [...new Set(rows.map((r) => r.site).filter(Boolean))];
    const check = { prevSum, openSum, diff: openSum - prevSum, vendors: diffs.length, sites };
    report.checks = (report.checks || []).concat({ label: K.label, ...check });
    const won = (n) => Math.round(n).toLocaleString();
    if (Math.abs(check.diff) >= 1 || diffs.length) {
      report.warnings.unshift(`⚠ ${K.label} 범위 확인: 이번 RD 기초 합계 ${won(openSum)}원 ≠ 지난달(${last.name}) 잔액 합계 ${won(prevSum)}원 (차이 ${won(openSum - prevSum)}원, 거래처 ${diffs.length}곳). ERP 거래처원장을 지난달과 같은 법인 범위로 받았는지 확인하세요${scope ? ` (지난달: ${scope})` : ''}.`
        + (diffs.length ? ` 차이 큰 거래처: ${diffs.slice(0, 5).map((x) => `${x.nm} 기초 ${won(x.open)} / 지난달 ${won(x.prev)}`).join(', ')}${diffs.length > 5 ? ' …' : ''}` : '')
        + (sites.length ? ` · RD에 들어 있는 사업장: ${sites.join(', ')}` : ''));
    } else {
      report.info.push(`${K.label} 범위 확인 ✓ 이번 RD 기초 합계가 지난달 잔액 합계(${won(prevSum)}원)와 같습니다${sites.length ? ` · 사업장: ${sites.join(', ')}` : ''}.`);
    }
  }

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
  if (book.has('매출채권') && book.has('매입채무') && book.has('맵핑') && !book.has('매출비교')) return buildArApList({ XLSX, book, rdFiles });
  // RD가 여러 법인 것이면(보관함 등) 지난달 잔액과 기초가 가장 잘 맞는 파일만 사용
  const pickInfo = [];
  if (rdFiles.length > 1) {
    const prevBal = new Map();
    for (const K of Object.values(ARAP_KINDS)) {
      const ms = monthSheets(book, K.suffix);
      if (!ms.length) continue;
      const tpl = await book.sheet(ms[ms.length - 1].name);
      for (const rn of tpl.rows.keys()) { const nm = norm(book.text(tpl.info('C', rn))); if (rn >= 3 && nm) prevBal.set(`${K.suffix}|${nm}`, (prevBal.get(`${K.suffix}|${nm}`) || 0) + num(book.val(tpl, 'I', rn))); }
    }
    let best = null;
    for (const f of rdFiles) {
      const rows = readPartnerLedgers(XLSX, [f]).rows;
      let hit = 0, tot = 0;
      for (const r of rows) {
        const K = Object.values(ARAP_KINDS).find((k) => k.codes.includes(r.acctCode) || k.names.includes(r.acctName));
        if (!K) continue;
        const pv = prevBal.get(`${K.suffix}|${norm(r.name)}`);
        if (pv === undefined && !r.open) continue;
        tot++; if (pv !== undefined && Math.abs(pv - r.open) < 1) hit++;
      }
      if (tot && (!best || hit / tot > best.score)) best = { f, score: hit / tot };
    }
    if (best) { pickInfo.push(`RD ${rdFiles.length}개 중 지난달 잔액과 기초가 맞는 '${best.f.name}'를 썼습니다.`); rdFiles = [best.f]; }
  }
  const rd = readPartnerLedgers(XLSX, rdFiles);
  const report = { warnings: [], info: [...pickInfo], parts: [], rdFiles: rd.files };
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
  let recM = rec.rows.filter((r) => inMonth(r.date));
  // 수불o에 지금까지 들어간 (구분·입고공장) 조합만 — 전체 RD(제품·상품·원재료 섞임)를 넣어도 안전
  {
    const so = await book.sheet('수불o');
    const hc = {};
    for (const [c, x] of so.rows.get(3)?.cells || []) hc[norm(book.text(cellInfo(x)))] = c;
    if (hc['구분'] && hc['입고공장']) {
      const seen = new Set();
      for (const rn of so.rows.keys()) if (rn > 4) { const k = `${norm(book.val(so, hc['구분'], rn))}|${norm(book.val(so, hc['입고공장'], rn))}`; if (k !== '|') seen.add(k); }
      if (seen.size) {
        const key = (r) => `${norm(r.raw[r.head.indexOf('구분')])}|${norm(r.raw[r.head.indexOf('입고공장')])}`;
        const skip = recM.filter((r) => !seen.has(key(r)));
        if (skip.length) {
          const g = new Map(); skip.forEach((r) => g.set(key(r), (g.get(key(r)) || 0) + 1));
          report.info.push(`수불o: 지금까지 쓰지 않던 구분·공장 ${skip.length}행은 뺐습니다 (${[...g].sort((a, z) => z[1] - a[1]).slice(0, 5).map(([k, n]) => `${k.replace('|', ' · ')} ${n}`).join(', ')}${g.size > 5 ? ` 외 ${g.size - 5}종` : ''})`);
          recM = recM.filter((r) => seen.has(key(r)));
        }
        // 같은 구분·공장이 여러 RD 파일에 있으면(전용 RD + 전체 RD) 첫 파일 것만 — 이중 집계 방지
        const owner = new Map();
        for (const r of recM) if (!owner.has(key(r))) owner.set(key(r), r.src);
        const dup = recM.filter((r) => owner.get(key(r)) !== r.src);
        if (dup.length) { report.info.push(`수불o: 같은 구분·공장이 다른 RD에도 있어 ${dup.length}행은 한 번만 셌습니다.`); recM = recM.filter((r) => owner.get(key(r)) === r.src); }
      }
    }
  }
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

/* ================= 공통: RD 시트를 A1 기준 표로 ================= */
function gridOf(XLSX, ws) {
  const g = [];
  if (!ws || !ws['!ref']) return g;
  const R = XLSX.utils.decode_range(ws['!ref']);
  for (let r = R.s.r; r <= R.e.r; r++) for (let c = R.s.c; c <= R.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;
    (g[r + 1] ||= [])[c + 1] = cell.v;
  }
  return g;
}
const gv = (g, r, c) => (g[r] ? g[r][c] ?? null : null);
const isNum = (v) => typeof v === 'number' && isFinite(v);
const same = (a, b) => Math.abs(a - b) <= Math.max(0.5, Math.abs(a) * 1e-9);

/* ================= ⑦ 전력비·수도광열비 (에너지 사용현황 → 계획대비 계정증감) ================= */
// 시트 모양: 블록(사용량/단가/요금)마다 '년'·'월' 머리글 행 + 월별 행, 오른쪽에 계량기 열, 그 뒤 Total 열
// get(r,c) 로 값을 읽는 어떤 표든(RD=SheetJS, 보고서=Book) 같은 방식으로 분석
function energyLayout(get, maxRow, maxCol) {
  let yc = 0, mc = 0, hr = 0;
  for (let r = 1; r <= Math.min(maxRow, 40) && !hr; r++) for (let c = 1; c < maxCol; c++) {
    if (norm(get(r, c)) === '년' && norm(get(r, c + 1)) === '월') { yc = c; mc = c + 1; hr = r; break; }
  }
  if (!hr) return null;
  // 계량기 열: 월 열 다음 ~ Total/합계 열 앞
  let end = maxCol;
  for (let c = mc + 1; c <= maxCol; c++) {
    let tot = false;
    for (let r = 1; r <= hr + 1; r++) if (/^total$|^합계$/i.test(norm(get(r, c)))) tot = true;
    if (tot) { end = c - 1; break; }
  }
  const label = (c) => { for (let r = hr + 1; r >= 1; r--) { const v = norm(get(r, c)); if (v && !isNum(get(r, c))) return v; } return ''; };
  const top = (c) => { for (let r = 1; r <= hr; r++) { const v = norm(get(r, c)); if (v && !isNum(get(r, c))) return v; } return ''; };
  const meters = [];
  for (let c = mc + 1; c <= end; c++) meters.push({ c, label: label(c), top: top(c) });
  const blocks = [];
  let cur = null, lastY = null;
  for (let r = 1; r <= maxRow; r++) {
    const y = get(r, yc), m = get(r, mc);
    if (norm(y) === '년' && norm(m) === '월') { cur = { hr: r, title: norm(get(r - 1, yc)), rows: new Map() }; blocks.push(cur); lastY = null; continue; }
    if (!cur) continue;
    let Y = isNum(y) ? y : (y === null && isNum(m) ? lastY : null);
    if (isNum(Y) && isNum(m) && m >= 1 && m <= 12 && Y > 2000) { lastY = Y; cur.rows.set(Y * 100 + m, r); }
  }
  return { yc, mc, hr, meters, blocks };
}

export function readEnergyRd(XLSX, files) {
  const out = []; // {file, sheet, kind:'power'|'gas', L, get}
  for (const { name, buf } of files) {
    const wb = XLSX.read(buf, { cellDates: false });
    for (const sn of wb.SheetNames) {
      const kind = /전력비/.test(sn) ? 'power' : /가스|등유|벙커/.test(sn) ? 'gas' : null;
      if (!kind) continue;
      const g = gridOf(XLSX, wb.Sheets[sn]);
      let maxCol = 0; for (const r of g) if (r && r.length - 1 > maxCol) maxCol = r.length - 1;
      const get = (r, c) => gv(g, r, c);
      const L = energyLayout(get, g.length - 1, maxCol);
      if (L && L.blocks.length) out.push({ file: name, sheet: sn, kind, L, get });
    }
  }
  return out;
}

export async function buildEnergy({ XLSX, JSZip, prevBuf, rdFiles }) {
  const book = await Book.open(JSZip, prevBuf);
  const report = { warnings: [], info: [], parts: [] };
  const rds = readEnergyRd(XLSX, rdFiles);
  report.rdFiles = rdFiles.map((f) => ({ name: f.name, rows: rds.filter((x) => x.file === f.name).length }));
  if (!rds.length) throw new Error('RD에서 전력비·가스(등유·벙커C유) 사용현황 표를 찾지 못했습니다.');
  // 이번 달: RD 첫 블록(사용량)에 값이 있는 가장 최근 달
  let ym = 0;
  for (const s of rds) {
    const b = s.L.blocks[0];
    for (const [k, r] of b.rows) if (k > ym && s.L.meters.some((mt) => { const v = s.get(r, mt.c); return isNum(v) && v !== 0; })) ym = k;
  }
  if (!ym) throw new Error('RD 사용량 표에 값이 있는 달이 없습니다.');
  const Y = Math.floor(ym / 100), M = ym % 100;
  let first = null;
  for (const s of book.sheets) {
    const kind = /전력비/.test(s.name) ? 'power' : /가스|등유|벙커/.test(s.name) ? 'gas' : null;
    if (!kind) continue;
    const src = rds.filter((x) => x.kind === kind);
    if (!src.length) { report.warnings.push(`'${s.name}': 맞는 RD 시트가 없어 건너뛰었습니다.`); continue; }
    const sh = await book.sheet(s.name);
    sh.expandShared();
    const get = (r, c) => book.val(sh, c, r);
    let L = energyLayout(get, sh.maxRow, sh.maxCol);
    if (!L) { report.warnings.push(`'${s.name}': 년·월 머리글을 찾지 못했습니다.`); continue; }
    // 이번 달 행이 없으면 각 블록 맨 위(최근 달 순서)에 한 줄 끼움 — 아래 블록부터
    const need = L.blocks.filter((b) => !b.rows.has(ym));
    // 1) 모든 블록에 빈 줄을 먼저 끼우고(아래 블록부터) 2) 최종 위치에서 옆 줄 수식을 복사 — 블록끼리 참조(단가=요금/사용량)가 맞게
    const ins = [];
    for (const b of [...need].reverse()) {
      const keys = [...b.rows.keys()];
      if (!keys.length) continue;
      const firstRow = Math.min(...b.rows.values());
      const desc = b.rows.get(Math.max(...keys)) === firstRow;
      const at = desc ? firstRow : Math.max(...b.rows.values()) + 1;
      await book.insertRows(s.name, at, 1);
      ins.forEach((x) => { if (x.at >= at) x.at++; });
      ins.push({ at, desc });
    }
    for (const { at, desc } of ins) {
      const d = desc ? -1 : 1;
      const tRow = sh.rows.get(at - d);
      const cells = new Map();
      for (const [c, x] of tRow?.cells || []) {
        const i = cellInfo(x);
        const ref = `${colLet(c)}${at}`;
        if (c === L.yc) cells.set(c, numCell(ref, i.s, Y));
        else if (c === L.mc) cells.set(c, numCell(ref, i.s, M));
        else if (i.f !== undefined) cells.set(c, fCell(ref, i.s, shiftRefs(i.f, (q) => ((!q.sheet || q.sheet === s.name) && q.row && !q.rowAbs ? { ...q, row: q.row + d } : q))));
        else cells.set(c, blankCell(ref, i.s));
      }
      sh.rows.set(at, { attrs: (tRow?.attrs || '').replace(/\sr="\d+"/, ` r="${at}"`) || ` r="${at}"`, cells });
    }
    if (need.length) { report.info.push(`'${s.name}': ${Y}년 ${M}월 행을 블록마다 새로 넣었습니다(위 행 수식 복사).`); L = energyLayout(get, sh.maxRow, sh.maxCol); }
    let wrote = 0, kept = 0;
    const unmatched = [];
    L.blocks.forEach((b, bi) => {
      const tr = b.rows.get(ym);
      if (!tr) return;
      const hist = [...b.rows.entries()].filter(([k]) => k < ym).sort((a, z) => z[0] - a[0]).slice(0, 12);
      const occ = new Map();
      for (const mt of L.meters) {
        const key = `${mt.label}#${(occ.get(mt.label) || 0) + 1}`; occ.set(mt.label, (occ.get(mt.label) || 0) + 1);
        const ti = sh.info(mt.c, tr);
        if (ti && ti.f !== undefined) { kept++; continue; } // 수식 칸(단가 등)은 그대로
        // 같은 블록 RD 열 중 지난 달 값이 가장 잘 맞는 열
        let best = null;
        for (const s2 of src) {
          const rb = s2.L.blocks[bi];
          if (!rb) continue;
          const occ2 = new Map();
          for (const m2 of s2.L.meters) {
            const key2 = `${m2.label}#${(occ2.get(m2.label) || 0) + 1}`; occ2.set(m2.label, (occ2.get(m2.label) || 0) + 1);
            let eq = 0, bad = 0;
            for (const [k, r] of hist) {
              const a = get(r, mt.c), r2 = rb.rows.get(k);
              const v = r2 ? s2.get(r2, m2.c) : null;
              if (isNum(a) && isNum(v) && a !== 0) { if (same(a, v)) eq++; else bad++; }
            }
            const tv = rb.rows.get(ym) ? s2.get(rb.rows.get(ym), m2.c) : null;
            let score = eq >= 2 && bad <= eq / 4 ? eq * 10 - bad * 30 : -1;
            if (score < 0 && key2 === key && mt.label) score = 1; // 기록이 없으면 머리글 이름(같은 이름은 순서)으로
            if (score < 0) continue;
            if (isNum(tv)) score += 0.5;
            if (!best || score > best.score) best = { score, v: tv };
          }
        }
        if (!best) { if (isNum(get(hist[0]?.[1], mt.c))) unmatched.push(mt.label || colLet(mt.c)); continue; }
        const v = isNum(best.v) ? best.v : null;
        if (v === null && !ti) continue;
        sh.set(mt.c, tr, (ref) => valCell(ref, sh.style(mt.c, tr), v));
        if (v !== null) wrote++;
      }
    });
    report.parts.push({ kind: 'energy', label: `${s.name}`, rows: wrote, note: `${Y}년 ${M}월 · 값 ${wrote}칸${kept ? ` · 수식 ${kept}칸 유지` : ''}` });
    if (unmatched.length) report.warnings.push(`'${s.name}': RD에서 짝을 못 찾은 열 — ${[...new Set(unmatched)].join(', ')} (손으로 확인)`);
    first ||= s.name;
  }
  const rest = book.sheets.filter((s) => !/전력비|가스|등유|벙커/.test(s.name)).map((s) => s.name);
  if (rest.length) report.info.push(`${rest.join(', ')} 시트는 그대로 두었습니다.`);
  const data = await book.save({ activeSheet: first });
  return { data, report, fileName: `계획대비계정증감_전력비·수도광열비_${String(Y).slice(2)}${pad2(M)}.xlsx` };
}

/* ================= ⑧ Raw Material (기간별입고현황 원재료 → 법인별 원재료 구매 물량·금액) ================= */
// 구역(D열 제목) → RD 입고공장
const RM_SECTIONS = [
  { re: /^팽이/, plants: [/디에이치아그로/, /디에이치팜/, /대흥농산\s*북부/] },
  { re: /느타리.*대흥/, plants: [/대흥농산\s*남부/] },
  { re: /느타리.*송이애/, plants: [/^송이애\s*\(이천\)/] },
  { re: /^새송이/, plants: [/이천팜/] },
];
const RM_ALIAS = [['콘콥', '콘컵'], ['케이포박', '케이폭', '케이폭박'], ['면실박', '면실피']];
const rmKey = (s) => norm(s).replace(/\s/g, '').replace(/\(kg\)$/i, '');
const rmSame = (a, b) => { a = rmKey(a); b = rmKey(b); return a === b || RM_ALIAS.some((g) => g.includes(a) && g.includes(b)); };

export async function buildRawMaterial({ XLSX, JSZip, prevBuf, rdFiles }) {
  const book = await Book.open(JSZip, prevBuf);
  const sname = book.sheets.find((s) => /raw\s*materials/i.test(s.name))?.name;
  if (!sname) throw new Error("Raw Material 파일이 아닙니다 ('raw materials' 시트 없음).");
  const report = { warnings: [], info: [], parts: [] };
  const rec = readReceipts(XLSX, rdFiles);
  report.rdFiles = rec.files;
  const rows = rec.rows.map((r) => {
    const g = (h) => { const k = r.head.indexOf(h); return k >= 0 ? r.raw[k] : null; };
    return { date: r.date, plant: norm(g('입고공장')), kind: norm(g('구분')), item: norm(g('품목명')), unit: norm(g('단위')), qty: num(g('입고수량')), amt: num(g('입고금액')), how: norm(g('입고구분')) };
  }).filter((r) => r.kind === '원재료');
  if (!rows.length) throw new Error('RD(기간별입고현황)에서 구분이 \'원재료\'인 행을 찾지 못했습니다.');
  const maxD = new Date(Math.max(...rows.map((r) => r.date)));
  const Y = maxD.getFullYear(), M = maxD.getMonth() + 1;
  const mrows = rows.filter((r) => r.date.getFullYear() === Y && r.date.getMonth() + 1 === M);
  const sh = await book.sheet(sname);
  sh.expandShared();
  const V = (c, r) => book.val(sh, c, r);
  // 이번 달 열: 1행 Act, 2행 물량/매입액, 6행 년, 7행 월
  let qc = 0, ac = 0;
  for (let c = 1; c <= sh.maxCol; c++) {
    if (norm(V(c, 1)) !== 'Act' || V(c, 6) !== Y || V(c, 7) !== M) continue;
    if (norm(V(c, 2)) === '물량') qc = c; else if (norm(V(c, 2)) === '매입액') ac = c;
  }
  if (!qc || !ac) throw new Error(`${Y}년 ${M}월 물량·매입액 열이 없습니다 (새해 열은 먼저 엑셀에서 만들어 주세요).`);
  // 구역·품목 행
  const items = [];
  let sec = null;
  for (let r = 8; r <= sh.maxRow; r++) {
    const d = norm(V(4, r)), e = norm(V(5, r));
    if (d && !e) { sec = RM_SECTIONS.find((x) => x.re.test(d)) ? { title: d, cfg: RM_SECTIONS.find((x) => x.re.test(d)) } : (/표고/.test(d) ? { title: d, pyo: r } : null); if (sec?.pyo) items.push({ r, pyo: true }); continue; }
    if (sec?.cfg && e && !/total|others/i.test(e)) items.push({ r, sec, label: e, mult: (String(V(3, r) || '').match(/(\d+(?:\.\d+)?)\s*kg/i) || [])[1] });
  }
  const used = new Set();
  let wrote = 0;
  const put = (c, r, v, f) => {
    const i = sh.info(c, r);
    if (f) { sh.setF(c, r, f); return; }
    if (v === null && (!i || i.f !== undefined)) return;
    sh.set(c, r, (ref) => valCell(ref, i?.s, v));
  };
  for (const it of items) {
    if (it.pyo) {
      // 표고: 미배양 배지 = 이천 생산입고(kg) + 원삼 구매(EA). 금액 중 생산분 원가는 RD에 없음
      const pr = mrows.filter((x) => /표고배지/.test(x.item) && /미배양/.test(x.item));
      pr.forEach((x) => used.add(x));
      const q = pr.reduce((t, x) => t + x.qty, 0), buy = pr.filter((x) => x.amt).reduce((t, x) => t + x.amt, 0);
      if (q) { put(qc, it.r, q); wrote++; }
      report.warnings.push(`표고(${it.r}행): 물량 ${Math.round(q).toLocaleString()}(미배양 배지 생산입고+구매)은 넣었고, 금액은 RD에 생산분 원가가 없어 비워 둠 — 구매분 ${Math.round(buy).toLocaleString()}원 + 이천 생산분 원가를 넣어 주세요.`);
      continue;
    }
    const hit = mrows.filter((x) => it.sec.cfg.plants.some((p) => p.test(x.plant)) && rmSame(x.item, it.label));
    hit.forEach((x) => used.add(x));
    const q = hit.reduce((t, x) => t + x.qty, 0), a = hit.reduce((t, x) => t + x.amt, 0);
    const pos = hit.some((x) => x.unit === '포');
    if (!hit.length) { put(qc, it.r, null); put(ac, it.r, null); continue; }
    if (pos && it.mult) put(qc, it.r, null, `${Math.round(q * 1000) / 1000}*${it.mult}`);
    else put(qc, it.r, Math.round(q));
    put(ac, it.r, a ? Math.round(a * 100) / 100 : null);
    if (pos && !it.mult) report.warnings.push(`${it.sec.title} ${it.label}: RD 단위가 '포'인데 kg 환산값(C열)이 없어 포 수량 그대로 넣었습니다.`);
    wrote++;
  }
  const left = mrows.filter((x) => !used.has(x) && !/표고배지.*배양완료/.test(x.item));
  const agg = new Map();
  for (const x of left) { const k = `${x.plant} · ${x.item}(${x.unit})`; const o = agg.get(k) || { q: 0, a: 0 }; o.q += x.qty; o.a += x.amt; agg.set(k, o); }
  if (agg.size) report.info.push(`파일에 넣지 않은 원재료 입고: ${[...agg].map(([k, o]) => `${k} ${Math.round(o.q).toLocaleString()} / ${Math.round(o.a).toLocaleString()}원`).join(' · ')}`);
  report.parts.push({ kind: 'rawmat', label: `${sname} ${Y}년 ${M}월`, rows: wrote, col: `${colLet(qc)}·${colLet(ac)}`, total: mrows.filter((x) => used.has(x)).reduce((t, x) => t + x.amt, 0) });
  const data = await book.save({ activeSheet: sname });
  return { data, report, fileName: `${M}월 Raw Material_${String(Y).slice(2)}${pad2(M)}.xlsx` };
}

/* ================= ⑬ 채권·채무 원장형 (송이애·군자 등) ================= */
// '매출채권'·'매입채무' 시트 = 달마다 거래처원장 행을 아래로 붙임 (A열 'YYYY.M.'), J·K열은 수식(맵핑 구분·증감)
const LIST_KINDS = [
  { sheet: '매출채권', label: '매출채권(외상매출금)', codes: ['11110401', '1110401'], names: ['외상매출금'], mapCol: 'A' },
  { sheet: '매입채무', label: '매입채무(외상매입금)', codes: ['12100101', '2100101'], names: ['외상매입금'], mapCol: 'H' },
];
async function buildArApList({ XLSX, book, rdFiles }) {
  const report = { warnings: [], info: [], parts: [] };
  const corp = (norm(book.val(await book.sheet('매출채권'), 'A', 2)).match(/거래처원장\s*-\s*(\S+)/) || [])[1] || '';
  // RD 파일별로 읽어 두고, 지난달 잔액과 기초가 가장 잘 맞는 파일(=같은 법인)을 고름
  const perFile = rdFiles.map((f) => ({ f, rd: readPartnerLedgers(XLSX, [f]) })).filter((x) => x.rd.rows.length);
  if (!perFile.length) throw new Error('RD에서 거래처원장(거래처명·계정·잔액 열)을 찾지 못했습니다.');
  report.rdFiles = perFile.map((x) => ({ name: x.f.name, rows: x.rd.rows.length }));
  const prevMonth = async (k) => {
    const sh = await book.sheet(k.sheet);
    const rows = [...sh.rows.keys()].filter((r) => r >= 5 && norm(book.val(sh, 'C', r))).sort((a, b) => a - b);
    if (!rows.length) return null;
    const lastV = norm(book.val(sh, 'A', rows[rows.length - 1]));
    const mine = rows.filter((r) => norm(book.val(sh, 'A', r)) === lastV);
    const bal = new Map();
    for (const r of mine) bal.set(norm(book.val(sh, 'C', r)), (bal.get(norm(book.val(sh, 'C', r))) || 0) + num(book.val(sh, 'I', r)));
    return { sh, lastV, first: mine[0], last: rows[rows.length - 1], bal };
  };
  const prevs = [];
  for (const k of LIST_KINDS) prevs.push(await prevMonth(k));
  const isK = (k, r) => k.codes.includes(r.acctCode) || k.names.includes(r.acctName);
  let best = null;
  for (const x of perFile) {
    let hit = 0, tot = 0;
    LIST_KINDS.forEach((k, i) => {
      const pv = prevs[i]; if (!pv) return;
      for (const r of x.rd.rows.filter((r) => isK(k, r))) { if (!r.open && !pv.bal.get(norm(r.name))) continue; tot++; if (Math.abs((pv.bal.get(norm(r.name)) || 0) - r.open) < 1) hit++; }
    });
    const score = tot ? hit / tot : 0;
    if (!best || score > best.score) best = { ...x, score, hit, tot };
  }
  if (perFile.length > 1) report.info.push(`RD ${perFile.length}개 중 지난달 잔액과 기초가 맞는 '${best.f.name}'를 썼습니다.`);
  if (best.tot && best.score < 0.9) report.warnings.unshift(`범위 확인: RD 기초가 지난달 잔액과 ${best.hit}/${best.tot}곳만 맞습니다 — ${corp || '이 법인'} 전체로 조회한 거래처원장인지 확인해 주세요.`);
  else if (best.tot) report.info.push(`범위 확인 ✓ RD 기초 = 지난달 잔액 (${best.hit}/${best.tot}곳 일치)`);
  let Y = 0, M = 0;
  for (let i = 0; i < LIST_KINDS.length; i++) {
    const k = LIST_KINDS[i], pv = prevs[i];
    if (!pv) { report.warnings.push(`'${k.sheet}' 시트에서 지난달 행을 찾지 못했습니다.`); continue; }
    const items = best.rd.rows.filter((r) => isK(k, r));
    if (!items.length) { report.warnings.push(`RD에 ${k.label} 행이 없어 '${k.sheet}'는 그대로 두었습니다.`); continue; }
    const m = pv.lastV.match(/^(\d{4})\.(\d{1,2})\.?$/);
    if (!m) throw new Error(`'${k.sheet}' A열의 달 표시('${pv.lastV}')를 읽지 못했습니다.`);
    [Y, M] = nextYm(+m[1], +m[2]);
    const label = `${Y}.${M}.`;
    const sh = pv.sh;
    sh.expandShared();
    // 같은 달을 다시 만들면 그 달 행을 지우고 새로
    let base = pv.last, removed = 0;
    if (pv.lastV === label) { for (let r = pv.first; r <= pv.last; r++) { sh.rows.delete(r); removed++; } base = pv.first - 1; }
    const tplRn = pv.lastV === label ? base : pv.last;
    const tpl = new Map(sh.rows.get(tplRn)?.cells || []);
    const tplAttrs = (sh.rows.get(tplRn)?.attrs || '').replace(/\s(thickBot|thickTop)="1"/g, '');
    let rn = base;
    const mapSh = await book.sheet('맵핑');
    const mapped = new Set();
    for (const r of mapSh.rows.keys()) { const v = norm(book.val(mapSh, k.mapCol, r)); if (v) mapped.add(v); }
    const fresh = [];
    for (const it of items) {
      rn++;
      const vals = { A: label, B: it.code === null || it.code === undefined ? null : (/^\d+$/.test(String(it.code)) ? Number(it.code) : String(it.code)), C: it.name,
        D: /^\d+$/.test(it.acctCode) ? Number(it.acctCode) : it.acctCode, E: it.acctName, F: it.open, G: it.dr || null, H: it.cr || null, I: it.bal };
      const cells = new Map();
      // 코드 열은 지난 행이 글자면 글자, 숫자면 숫자로
      for (const c of ['B', 'D']) { const ti = cellInfo(tpl.get(colNum(c))); if (vals[c] !== null && ti && (ti.t === 's' || ti.t === 'inlineStr' || ti.t === 'str')) vals[c] = String(vals[c]); }
      for (const c of ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']) cells.set(colNum(c), valCell(`${c}${rn}`, styleOfXml(tpl.get(colNum(c))), vals[c]));
      for (const c of ['J', 'K']) {
        const i = cellInfo(tpl.get(colNum(c)));
        if (i?.f !== undefined) cells.set(colNum(c), fCell(`${c}${rn}`, i.s, shiftRefs(i.f, (q) => ((!q.sheet || q.sheet === k.sheet) && q.row && !q.rowAbs ? { ...q, row: q.row + (rn - tplRn) } : q))));
      }
      sh.rows.set(rn, { attrs: tplAttrs.replace(/\sr="\d+"/, ` r="${rn}"`) || ` r="${rn}"`, cells });
      if (!mapped.has(norm(it.name)) && (it.open || it.dr || it.cr || it.bal)) fresh.push(it.name);
    }
    report.parts.push({ kind: 'arapList', label: k.label, sheet: k.sheet, rows: items.length, removed, balance: items.reduce((t, r) => t + r.bal, 0), month: label });
    if (fresh.length) report.warnings.push(`'${k.sheet}': 맵핑 시트에 없는 거래처 ${fresh.length}곳 — 구분(J열)이 비어 있으니 맵핑 ${k.mapCol}열에 추가해 주세요: ${[...new Set(fresh)].slice(0, 8).join(', ')}${fresh.length > 8 ? ' 외' : ''}`);
  }
  if (!report.parts.length) throw new Error('RD에 외상매출금·외상매입금 행이 없어 만든 내용이 없습니다.');
  report.info.push('오른쪽 요약표(구분별 합계)는 손대지 않았습니다.');
  const data = await book.save({ activeSheet: '매출채권' });
  return { data, report, fileName: `(${corp || 'AR_AP'}) ${String(Y).slice(2)}년${pad2(M)}월 거래처별 외상매입매출_${String(Y).slice(2)}${pad2(M)}.xlsx` };
}

/* ================= ⑮ CAPEX 재분류 검증 (대흥 통합 Capex Actual → 닥터애그(CAPEX Actual)) ================= */
// 대흥 파일 'CAPEX 닥터애그…' 시트: C 구분1 · D 구분2 · F 투자내역 · 월 열(1~12월, 원) → 검증파일 맨 아래에 그 달 행(A 연도·B 월·E·F·I·그 달 열, 천원)
export function readCapexSource(XLSX, files) {
  for (const { name, buf } of files) {
    const wb = XLSX.read(buf, { cellDates: true });
    const sn = wb.SheetNames.find((x) => /^CAPEX\s*닥터애그/.test(x));
    if (!sn) continue;
    const g = gridOf(XLSX, wb.Sheets[sn]);
    let hr = 0, cInv = 0, cG1 = 0, cG2 = 0, m1 = 0;
    for (let r = 1; r <= 10 && !hr; r++) {
      const row = g[r] || [];
      const ix = row.findIndex((v) => norm(v) === '투자내역');
      if (ix > 0) { hr = r; cInv = ix; cG1 = row.findIndex((v) => norm(v) === '구분1'); cG2 = row.findIndex((v) => norm(v) === '구분2'); m1 = ix + 1; }
    }
    if (!hr) continue;
    return { name, sheet: sn, get: (r, c) => gv(g, r, c), maxRow: g.length - 1, hr, cInv, cG1, cG2, m1 };
  }
  return null;
}

export async function buildCapex({ XLSX, JSZip, prevBuf, rdFiles }) {
  const book = await Book.open(JSZip, prevBuf);
  const SN = '닥터애그(CAPEX Actual)';
  if (!book.has(SN)) throw new Error(`CAPEX 재분류 검증 파일이 아닙니다 ('${SN}' 시트 없음).`);
  const src = readCapexSource(XLSX, rdFiles);
  if (!src) throw new Error("대흥 '통합 Capex Actual' 파일('CAPEX 닥터애그' 시트)을 RD 칸에 넣어 주세요.");
  const report = { warnings: [], info: [], parts: [], rdFiles: [{ name: src.name, rows: src.maxRow }] };
  const sh = await book.sheet(SN);
  sh.expandShared();
  const V = (c, r) => book.val(sh, c, r);
  // 지난 마지막 달
  const rows = [...sh.rows.keys()].filter((r) => r >= 9 && typeof V(1, r) === 'number' && typeof V(2, r) === 'number').sort((a, b) => a - b);
  if (!rows.length) throw new Error(`'${SN}'에서 연도·월(A·B열)이 있는 행을 찾지 못했습니다.`);
  const lastRn = rows[rows.length - 1];
  const [ly, lm] = [V(1, lastRn), V(2, lastRn)];
  // 이번 달: 원본 파일에서 지난달 다음 달 열에 금액이 있으면 그 달 (같은 달을 다시 만들면 그 달 교체)
  const amt = (M, r) => num(src.get(r, src.m1 + M - 1));
  const srcRows = (M) => { const o = []; for (let r = src.hr + 1; r <= src.maxRow; r++) if (norm(src.get(r, src.cInv)) && amt(M, r)) o.push(r); return o; };
  let [Y, M] = nextYm(ly, lm);
  if (Y !== ly) throw new Error(`${Y}년 1월은 검증파일에 새해 열(${Y} Actual)을 먼저 만든 뒤 진행해 주세요.`);
  if (!srcRows(M).length && srcRows(lm).length) { M = lm; report.info.push(`원본에 ${lm + 1}월 금액이 아직 없어 ${lm}월을 다시 만들었습니다.`); }
  const picked = srcRows(M);
  if (!picked.length) throw new Error(`원본 파일에 ${M}월 금액이 있는 행이 없습니다.`);
  // 검증파일의 '{Y} Actual' 블록에서 M월 열
  let start = 0;
  for (let c = 1; c <= sh.maxCol && !start; c++) if (norm(V(c, 3)) === `${Y} Actual`) start = c;
  if (!start) throw new Error(`'${SN}' 3행에서 '${Y} Actual' 열을 찾지 못했습니다.`);
  const mc = start + M - 1;
  // 같은 달 행이 이미 있으면 지우고 다시
  const same = rows.filter((r) => V(1, r) === Y && V(2, r) === M);
  let base = lastRn;
  if (same.length) { for (const r of same) sh.rows.delete(r); base = same[0] - 1; report.info.push(`기존 ${M}월 ${same.length}행을 새로 바꿨습니다.`); }
  const tplRn = same.length ? base : lastRn;
  const tpl = sh.rows.get(tplRn);
  const st = (c) => styleOfXml(tpl?.cells.get(c)) ?? styleOfXml(tpl?.cells.get(9));
  // 금액 칸 서식 = 본보기 행의 금액 칸(그 행의 달 열) 서식
  const tplM = V(1, tplRn) === Y ? start + V(2, tplRn) - 1 : mc;
  const amtStyle = styleOfXml(tpl?.cells.get(tplM)) ?? st(mc);
  const dVal = V(4, tplRn);
  // 재분류 배우기: 지난달들의 원본 구분 → 검증파일에서 고친 구분 (투자내역·달·금액이 같은 행끼리)
  const txt = (v) => String(v ?? '').trim();
  const learn = { g1: new Map(), g2: new Map() };
  const vRows = rows.filter((r) => V(1, r) === Y && V(2, r) < M);
  for (let m = 1; m < M; m++) {
    const done = new Map();
    for (const r of vRows) if (V(2, r) === m) { const k = `${norm(V(9, r))}|${Math.round(num(V(start + m - 1, r)))}`; if (!done.has(k)) done.set(k, r); }
    for (const r of srcRows(m)) {
      const vr = done.get(`${norm(src.get(r, src.cInv))}|${Math.round(amt(m, r) / 1000)}`);
      if (!vr) continue;
      for (const [k, sc, vc] of [['g1', src.cG1, 5], ['g2', src.cG2, 6]]) {
        const a = norm(src.get(r, sc)), b = norm(V(vc, vr));
        if (!a || !b) continue;
        const cnt = learn[k].get(a) || new Map(); cnt.set(b, (cnt.get(b) || 0) + 1); learn[k].set(a, cnt);
      }
    }
  }
  const cats = { g1: new Set(), g2: new Set() };
  for (const r of rows) { const a = norm(V(5, r)), b = norm(V(6, r)); if (a) cats.g1.add(a); if (b) cats.g2.add(b); }
  const changed = [], unknown = new Set();
  const classify = (k, v, inv) => {
    if (!v) return null;
    const c = learn[k].get(v);
    let out = c ? [...c].sort((a, z) => z[1] - a[1])[0][0] : null;
    if (!out && cats[k].has(v)) out = v;
    if (!out) out = [...cats[k]].filter((x) => v.startsWith(x)).sort((a, z) => z.length - a.length)[0] || null;
    if (!out) { unknown.add(v); return v; }
    if (out !== v) changed.push(`${inv}: ${v} → ${out}`);
    return out;
  };
  let rn = base, total = 0;
  for (const r of picked) {
    rn++;
    const cells = new Map();
    const put = (c, v) => cells.set(c, valCell(`${colLet(c)}${rn}`, st(c), v));
    put(1, Y); put(2, M);
    if (dVal !== null && dVal !== '') put(4, dVal);
    const inv = txt(src.get(r, src.cInv));
    put(5, classify('g1', norm(src.get(r, src.cG1)), inv)); put(6, classify('g2', norm(src.get(r, src.cG2)), inv)); put(9, inv);
    const a = Math.round(amt(M, r)) / 1000;
    cells.set(mc, valCell(`${colLet(mc)}${rn}`, amtStyle, a)); total += a;
    sh.rows.set(rn, { attrs: (tpl?.attrs || '').replace(/\sr="\d+"/, ` r="${rn}"`) || ` r="${rn}"`, cells });
  }
  sh.xml = sh.xml.replace(/(<autoFilter ref="[A-Z]+\d+:[A-Z]+)(\d+)"/, (m0, a, e) => (+e >= base ? `${a}${rn}"` : m0));
  report.parts.push({ kind: 'capex', label: `${SN} ${Y}년 ${M}월`, rows: picked.length, note: `${colLet(mc)}열 · 합계 ${Math.round(total).toLocaleString()}천원` });
  report.warnings.push('송이애(CAPEX Actual)·군자(CAPEX Actual)는 각 법인 KPI 파일을 보고 손으로 넣어 주세요.');
  if (changed.length) report.info.push(`재분류 ${changed.length}건 (지난 달들의 검증 기록대로): ${[...new Set(changed.map((x) => x.split(': ')[1]))].join(', ')}`);
  if (unknown.size) report.warnings.push(`검증파일에 없던 분류라 그대로 둔 것: ${[...unknown].join(', ')} — 맞는 분류로 E·F열을 고쳐 주세요.`);
  const data = await book.save({ activeSheet: SN });
  return { data, report, fileName: `${Y}년${pad2(M)}월_Dr.AG_Capex_재분류_검증_${String(Y).slice(2)}${pad2(M)}.xlsx` };
}
