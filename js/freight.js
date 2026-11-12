// 운반비 검증 엔진
// 입력: ① 운반비 집계파일(타부서, '현황' 시트) ② ERP 계정별원장(전체계정 또는 운반비 계정, 여러 개) ③ 지난 검증파일
// 출력: 지난 검증파일에 현황N월·운반비N월 시트를 추가(같은 달이 있으면 다시 만듦)하고 계정원장에 N월 행을 붙인 엑셀
//
// 원칙
// - 기존 시트·서식·수식은 건드리지 않고 XML을 직접 고쳐서 글꼴·서식이 그대로 유지되게 함
// - 원장: 운반보관비 3개 계정, 차변만 사용(결산전표 등 대변 행 제외), 매출차감 거래처(오아시스) 차감액은 0
// - 운반비N월 시트는 지난달 시트를 복사해서 만들고, 납품처 금액(H열)만 집계파일 값으로 바꿈
// - 차이가 나는 지급처는 U열(사유)에 '제안)'으로 후보를 적어 둠 → 사용자가 확인·수정
//
// 브라우저와 Node 모두에서 쓰도록 XLSX(SheetJS)·JSZip은 밖에서 넘겨받음

export const FREIGHT_CODES = ['1510424001', '1510424002', '1510424003']; // 운반보관비-선행/후행/기타
export const SALES_DEDUCT = ['오아시스']; // 매출차감 거래처: 차변 음수 행은 0 처리
// 집계파일 현황 2행(회사) → 검증파일 운반비 시트 C열(사업장) 이름
export const COMPANY_ALIAS = { 닥터애그: '대흥', '닥터애그_중부': '중부이천팜' };

const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const colNum = (letters) => letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colLet = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const toSerial = (d) => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 864e5;
const isDate = (d) => d instanceof Date && !isNaN(d);
const num = (v) => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
};

/* ---------------- 1. 집계파일 '현황' 읽기 ---------------- */
export function readAggregate(XLSX, buf) {
  const wb = XLSX.read(buf, { cellDates: true, cellStyles: false });
  const name = wb.SheetNames.find((n) => n === '현황') || wb.SheetNames.find((n) => n.startsWith('현황'));
  if (!name) throw new Error("집계파일에 '현황' 시트가 없습니다.");
  const ws = wb.Sheets[name];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  const hr = grid.findIndex((row) => row && norm(row[1]) === '날짜');
  if (hr < 0) throw new Error("현황 시트에서 '날짜' 머리글을 찾지 못했습니다.");
  const maxC = Math.max(...grid.slice(hr, hr + 4).map((r) => (r ? r.length : 0)));
  const cols = [];
  let company = '', place = '';
  for (let c = 3; c < maxC; c++) {
    const co = norm(grid[hr]?.[c]);
    if (co) { company = co; place = ''; }
    const p = norm(grid[hr + 1]?.[c]);
    const carrier = norm(grid[hr + 2]?.[c]);
    const channel = norm(grid[hr + 3]?.[c]);
    if (!p && !carrier) continue; // 구분용 빈 열
    if (p) place = p;
    cols.push({ c, company, sajang: COMPANY_ALIAS[company] || company, place, placeRaw: p, carrier, channel });
  }
  const rowOf = (label) => grid.findIndex((r, i) => i > hr && r && norm(r[1]) === label);
  const totalRow = rowOf('Total');
  if (totalRow < 0) throw new Error("현황 시트에서 'Total' 행을 찾지 못했습니다.");
  const firstDate = grid.slice(hr + 1).map((r) => r?.[1]).find(isDate);
  if (!firstDate) throw new Error('현황 시트에서 날짜를 찾지 못했습니다.');
  for (const col of cols) col.total = num(grid[totalRow]?.[col.c]);
  return { sheetName: name, ws, grid, hr, totalRow, cols, year: firstDate.getFullYear(), month: firstDate.getMonth() + 1, merges: ws['!merges'] || [] };
}

/* ---------------- 2. ERP 계정별원장 읽기 (여러 파일) ---------------- */
export function readLedgers(XLSX, bufs, { year, month }) {
  const rows = [];
  const perFile = [];
  for (const { name, buf } of bufs) {
    const wb = XLSX.read(buf, { cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const g = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const h = g.findIndex((r) => r && r.some((x) => norm(x) === '계정코드'));
    if (h < 0) { perFile.push({ name, error: "'계정코드' 머리글을 찾지 못함" }); continue; }
    const H = g[h].map(norm);
    const ix = (k) => H.indexOf(k);
    const krw = ix('KRW') >= 0 ? ix('KRW') : ix('차변');
    const idx = {
      code: ix('계정코드'), acc: ix('계정명'), site: ix('사업장'), date: ix('일자'), memo: ix('적요'), memo2: ix('적요2'),
      appr: ix('승인번호'), dr: krw, cr: krw + 1, bal: krw + 2, fx: ix('환율'), slip: ix('전표번호'), line: ix('행'),
      dept: ix('귀속부서'), m1: ix('관리항목1'), m2: ix('관리항목2'),
    };
    let taken = 0, credit = 0, deduct = 0, other = 0;
    const sites = new Set();
    for (let i = h + 1; i < g.length; i++) {
      const r = g[i];
      if (!r) continue;
      const d = r[idx.date];
      if (!isDate(d)) continue; // 전일잔액·월계·누계·2행 머리글
      const code = norm(r[idx.code]).replace(/\.0$/, '');
      if (!FREIGHT_CODES.includes(code) && !norm(r[idx.acc]).startsWith('운반보관비')) continue;
      if (d.getFullYear() !== year || d.getMonth() + 1 !== month) { other++; continue; }
      let dr = num(r[idx.dr]);
      if (!dr) { credit++; continue; } // 차변만 사용 (결산전표 등 제외)
      const text = `${r[idx.memo] ?? ''} ${r[idx.m1] ?? ''}`;
      let note = '';
      if (dr < 0 && SALES_DEDUCT.some((k) => text.includes(k))) { dr = 0; deduct++; note = '매출차감 → 0'; }
      const str = (k) => (idx[k] >= 0 && r[idx[k]] !== null && r[idx[k]] !== undefined ? String(r[idx[k]]).replace(/\.0$/, '') : '');
      rows.push({
        code, acc: norm(r[idx.acc]), site: norm(r[idx.site]), date: d, memo: str('memo'), memo2: str('memo2'), appr: str('appr'),
        dr, cr: num(r[idx.cr]), bal: r[idx.bal] === null || r[idx.bal] === '' ? null : num(r[idx.bal]),
        fx: r[idx.fx] === null || r[idx.fx] === '' ? null : num(r[idx.fx]), slip: str('slip'), line: str('line'), dept: str('dept'),
        m1: str('m1'), m2: str('m2'), note, file: name,
      });
      sites.add(norm(r[idx.site]));
      taken++;
    }
    perFile.push({ name, taken, credit, deduct, other, sites: [...sites] });
  }
  return { rows, perFile };
}

/* ---------------- 3. xlsx(zip) 내부 다루기 ---------------- */
async function openBook(JSZip, buf) {
  const zip = await JSZip.loadAsync(buf);
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const relXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rels = {};
  for (const m of relXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = m[0].match(/Id="([^"]+)"/)[1];
    const target = m[0].match(/Target="([^"]+)"/)[1];
    rels[id] = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
  }
  const sheets = [...wbXml.matchAll(/<sheet\b[^>]*\/>/g)].map((m) => ({
    tag: m[0],
    name: unesc(m[0].match(/name="([^"]+)"/)[1]),
    sheetId: +m[0].match(/sheetId="(\d+)"/)[1],
    rid: m[0].match(/r:id="([^"]+)"/)[1],
  }));
  sheets.forEach((s) => { s.path = rels[s.rid]; });
  return { zip, wbXml, relXml, sheets };
}

// 셀 XML 파싱 (값 읽기용): { 'A1': {attrs, inner, v, f, t, s} }
function parseCells(xml) {
  const cells = {};
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const attrs = m[2];
    const inner = m[3] || '';
    cells[m[1]] = {
      attrs, inner,
      t: (attrs.match(/\st="([^"]+)"/) || [])[1],
      s: (attrs.match(/\ss="(\d+)"/) || [])[1],
      v: (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1],
      f: (inner.match(/<f\b[^>]*>([\s\S]*?)<\/f>/) || [])[1],
      fTag: (inner.match(/<f\b[^>]*\/?>/) || [])[0],
      isText: (inner.match(/<is>[\s\S]*?<\/is>/) || [])[0],
    };
  }
  return cells;
}

async function sharedStrings(zip) {
  const f = zip.file('xl/sharedStrings.xml');
  if (!f) return [];
  const xml = await f.async('string');
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => unesc([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join('')));
}
const cellText = (c, sst) => {
  if (!c) return '';
  if (c.t === 's') return sst[+c.v] ?? '';
  if (c.isText) return unesc([...c.isText.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join(''));
  return c.v !== undefined ? unesc(c.v) : '';
};

const textCell = (ref, s, text) => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${esc(text)}</t></is></c>`;
const numCell = (ref, s, v) => `<c r="${ref}"${s ? ` s="${s}"` : ''}><v>${v}</v></c>`;
const fCell = (ref, s, f) => `<c r="${ref}"${s ? ` s="${s}"` : ''}><f>${esc(f)}</f></c>`;
const blankCell = (ref, s) => `<c r="${ref}"${s ? ` s="${s}"` : ''}/>`;

// 한 셀 바꾸기(없으면 행에 추가). rows: Map(rowNum → {attrs, cells: Map(col → xml)})
function parseRows(sheetXml) {
  const sd = sheetXml.match(/<sheetData>([\s\S]*?)<\/sheetData>|<sheetData\/>/);
  const rows = new Map();
  if (sd && sd[1]) {
    for (const m of sd[1].matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const rn = +m[1].match(/\sr="(\d+)"/)[1];
      const cells = new Map();
      for (const c of (m[2] || '').matchAll(/<c r="([A-Z]+)(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) cells.set(colNum(c[1]), c[0]);
      rows.set(rn, { attrs: m[1].replace(/\sspans="[^"]*"/, ''), cells });
    }
  }
  return rows;
}
function buildSheetData(rows) {
  const out = [];
  for (const rn of [...rows.keys()].sort((a, b) => a - b)) {
    const r = rows.get(rn);
    const cells = [...r.cells.keys()].sort((a, b) => a - b).map((k) => r.cells.get(k)).join('');
    out.push(cells ? `<row${r.attrs}>${cells}</row>` : `<row${r.attrs}/>`);
  }
  return `<sheetData>${out.join('')}</sheetData>`;
}
function setCell(rows, ref, xml) {
  const [, col, rn] = ref.match(/([A-Z]+)(\d+)/);
  if (!rows.has(+rn)) rows.set(+rn, { attrs: ` r="${rn}"`, cells: new Map() });
  rows.get(+rn).cells.set(colNum(col), xml);
}
const styleOf = (rows, ref) => {
  const [, col, rn] = ref.match(/([A-Z]+)(\d+)/);
  const x = rows.get(+rn)?.cells.get(colNum(col));
  return x ? (x.match(/\ss="(\d+)"/) || [])[1] : undefined;
};

/* ---------------- 4. 검증파일 만들기 ---------------- */
export async function buildVerification({ XLSX, JSZip, aggBuf, ledgerBufs, prevBuf, extraMapping = [] }) {
  const report = { warnings: [], info: [] };
  const agg = readAggregate(XLSX, aggBuf);
  const { year: Y, month: M } = agg;
  const led = readLedgers(XLSX, ledgerBufs, { year: Y, month: M });
  report.month = M; report.year = Y; report.ledgerFiles = led.perFile;
  if (!led.rows.length) throw new Error(`원장에서 ${Y}년 ${M}월 운반보관비 차변 행을 찾지 못했습니다. 기간·파일을 확인해주세요.`);

  const book = await openBook(JSZip, prevBuf);
  const { zip } = book;
  const sst = await sharedStrings(zip);
  const byName = (n) => book.sheets.find((s) => s.name === n);
  const freightNames = book.sheets.map((s) => s.name.match(/^운반비(\d+)월$/)).filter(Boolean).map((m) => +m[1]);
  const prevM = freightNames.filter((n) => n !== M).sort((a, b) => ((M - a + 12) % 12) - ((M - b + 12) % 12))[0];
  if (!prevM) throw new Error("지난 검증파일에서 '운반비N월' 시트를 찾지 못했습니다.");
  const tplFreight = byName(`운반비${prevM}월`);
  const tplStatus = byName(`현황${prevM}월`) || book.sheets.find((s) => /^현황\d+월$/.test(s.name));
  const ledgerSheet = byName('계정원장');
  const mapSheet = byName('mapping');
  if (!ledgerSheet || !mapSheet) throw new Error("지난 검증파일에 '계정원장' 또는 'mapping' 시트가 없습니다.");
  report.template = tplFreight.name;
  const replacing = !!byName(`운반비${M}월`);
  if (replacing) report.info.push(`이미 있는 ${M}월 시트를 새로 만들었습니다.`);

  /* --- mapping 읽기 --- */
  const mapXml = await zip.file(mapSheet.path).async('string');
  const mc = parseCells(mapXml);
  const mapCol = (kc, vc) => {
    const out = new Map();
    for (const ref of Object.keys(mc)) {
      const [, col, rn] = ref.match(/([A-Z]+)(\d+)/);
      if (col !== kc || +rn < 2) continue;
      const k = norm(cellText(mc[ref], sst));
      if (k && !out.has(k)) out.set(k, norm(cellText(mc[`${vc}${rn}`], sst)));
    }
    return out;
  };
  const siteToCo = mapCol('C', 'D'), accToLead = mapCol('F', 'G'), vendorToPayee = mapCol('K', 'L'), payeeToCh = mapCol('O', 'P');
  // 사용자가 화면에서 입력한 새 거래처 → 지급처 매핑을 mapping 시트 K:L 아래에 추가
  const addMap = extraMapping.filter((x) => norm(x.vendor) && norm(x.payee) && !vendorToPayee.has(norm(x.vendor)));
  if (addMap.length) {
    const mr = parseRows(mapXml);
    const lastK = Math.max(...Object.keys(mc).filter((ref) => /^K\d+$/.test(ref) && norm(cellText(mc[ref], sst))).map((ref) => +ref.slice(1)));
    const st = (c) => styleOf(mr, `${c}${lastK}`);
    let n = lastK;
    for (const x of addMap) {
      n++;
      if (x.co) setCell(mr, `I${n}`, textCell(`I${n}`, st('I'), x.co));
      if (x.lead) setCell(mr, `J${n}`, textCell(`J${n}`, st('J'), x.lead));
      setCell(mr, `K${n}`, textCell(`K${n}`, st('K'), norm(x.vendor)));
      setCell(mr, `L${n}`, textCell(`L${n}`, st('L'), norm(x.payee)));
      vendorToPayee.set(norm(x.vendor), norm(x.payee));
    }
    const newMapXml = mapXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, buildSheetData(mr))
      .replace(/(<autoFilter ref="[A-Z]+\d+:[A-Z]+)(\d+)"/, (m, a, b) => `${a}${Math.max(+b, n)}"`);
    zip.file(mapSheet.path, newMapXml);
    report.info.push(`mapping 시트에 새 거래처 ${addMap.length}곳을 추가했습니다.`);
    book.mapAddedTo = n;
  }

  /* --- 계정원장: M월 행 지우고 새로 붙이기 --- */
  const ledXml = await zip.file(ledgerSheet.path).async('string');
  const ledRows = parseRows(ledXml);
  const ledCells = parseCells(ledXml);
  const serialMonth = (s) => { const d = new Date(Date.UTC(1899, 11, 30) + s * 864e5); return [d.getUTCFullYear(), d.getUTCMonth() + 1]; };
  let removed = 0, lastKeep = 2;
  const existing = []; // 남는 기존 행(대사 계산용)
  for (const rn of [...ledRows.keys()].sort((a, b) => a - b)) {
    if (rn <= 2) continue;
    const dc = ledCells[`D${rn}`];
    if (dc && dc.v && !isNaN(+dc.v)) {
      const [yy, mm] = serialMonth(+dc.v);
      if (yy === Y && mm === M) { ledRows.delete(rn); removed++; continue; }
      existing.push({
        rn, site: norm(cellText(ledCells[`C${rn}`], sst)), acc: norm(cellText(ledCells[`B${rn}`], sst)), y: yy, m: mm,
        dr: num(ledCells[`H${rn}`]?.v), m1: norm(cellText(ledCells[`O${rn}`], sst)), memo: cellText(ledCells[`E${rn}`], sst),
      });
    }
    lastKeep = Math.max(lastKeep, rn);
  }
  if (removed) report.info.push(`계정원장의 기존 ${M}월 행 ${removed}건을 새 원장으로 바꿨습니다.`);
  // 마지막 데이터 행을 서식·수식 본보기로
  const tplRow = ledRows.get(lastKeep);
  const tplCell = (col) => tplRow?.cells.get(colNum(col)) || '';
  const tplStyle = (col) => (tplCell(col).match(/\ss="(\d+)"/) || [])[1];
  const tplFormula = (col) => {
    let f = (tplCell(col).match(/<f>([\s\S]*?)<\/f>/) || [])[1];
    if (!f) { // 공유 수식이면 기본 수식으로
      f = { S: 'YEAR($D{r})', T: 'MONTH($D{r})' }[col];
      return f || null;
    }
    f = unesc(f);
    return f.replace(/(\$?[A-Z]{1,3})(\$?)(\d+)(?!\d)/g, (all, c, d, n) => (d || +n !== lastKeep ? all : `${c}{r}`));
  };
  const defaultF = {
    Q: 'VLOOKUP($C{r},mapping!$C$1:$D$7,2,0)', R: 'VLOOKUP($B{r},mapping!$F$1:$G$5,2,0)', S: 'YEAR($D{r})', T: 'MONTH($D{r})',
    U: 'VLOOKUP($O{r},mapping!$K:$L,2,0)', V: 'IF(R{r}="선행",R{r},VLOOKUP(U{r},mapping!$O$1:$P$17,2,0))',
  };
  let r = lastKeep;
  const newLedger = [];
  const unmapped = new Map();
  for (const x of led.rows) {
    r++;
    const st = tplStyle;
    const cells = [
      textCell(`A${r}`, st('A'), x.code), textCell(`B${r}`, st('B'), x.acc), textCell(`C${r}`, st('C'), x.site),
      numCell(`D${r}`, st('D'), toSerial(x.date)), textCell(`E${r}`, st('E'), x.memo), textCell(`F${r}`, st('F'), x.memo2),
      textCell(`G${r}`, st('G'), x.appr), numCell(`H${r}`, st('H'), x.dr), numCell(`I${r}`, st('I'), x.cr),
      x.bal === null ? blankCell(`J${r}`, st('J')) : numCell(`J${r}`, st('J'), x.bal),
      x.fx === null ? blankCell(`K${r}`, st('K')) : numCell(`K${r}`, st('K'), x.fx),
      textCell(`L${r}`, st('L'), x.slip), textCell(`M${r}`, st('M'), x.line), textCell(`N${r}`, st('N'), x.dept),
      x.m1 ? textCell(`O${r}`, st('O'), x.m1) : blankCell(`O${r}`, st('O')), x.m2 ? textCell(`P${r}`, st('P'), x.m2) : blankCell(`P${r}`, st('P')),
      ...['Q', 'R', 'S', 'T', 'U', 'V'].map((c) => fCell(`${c}${r}`, st(c), (tplFormula(c) || defaultF[c]).replaceAll('{r}', r))),
    ];
    ledRows.set(r, { attrs: ` r="${r}"`, cells: new Map(cells.map((c) => [colNum(c.match(/r="([A-Z]+)/)[1]), c])) });
    const co = siteToCo.get(x.site) || '';
    const lead = accToLead.get(x.acc) || '';
    const payee = vendorToPayee.get(norm(x.m1));
    if (x.m1 && payee === undefined) { const u = unmapped.get(x.m1) || { amount: 0, co, lead, memo: x.memo }; u.amount += x.dr; unmapped.set(x.m1, u); }
    newLedger.push({ ...x, rn: r, co, lead, payee: payee ?? '', y: Y, m: M });
  }
  if (!siteToCo.size) report.warnings.push('mapping 시트에서 사업장→회사 표를 읽지 못했습니다.');
  report.unmapped = [...unmapped].map(([vendor, u]) => ({ vendor, ...u }));
  // 지운 행에 공유수식 '원본'이 있었으면, 남은 셀들이 원본을 잃지 않도록 일반 수식으로 바꿈
  const masters = new Set();
  for (const [, row] of ledRows) for (const [, x] of row.cells) { const m = x.match(/<f t="shared"[^>]*\sref="[^"]+"[^>]*si="(\d+)"|<f t="shared"[^>]*si="(\d+)"[^>]*\sref="[^"]+"/); if (m) masters.add(m[1] || m[2]); }
  for (const [rn, row] of ledRows) {
    for (const [c, x] of row.cells) {
      const m = x.match(/<f t="shared"[^>]*si="(\d+)"[^>]*\/>/);
      if (!m || masters.has(m[1])) continue;
      const f = defaultF[colLet(c)];
      row.cells.set(c, f ? x.replace(/<f t="shared"[^>]*\/>/, `<f>${esc(f.replaceAll('{r}', rn))}</f>`) : x.replace(/<f t="shared"[^>]*\/>/, ''));
    }
  }
  let newLedXml = ledXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, buildSheetData(ledRows))
    .replace(/<dimension ref="([A-Z]+\d+):[A-Z]+\d+"\/>/, `<dimension ref="$1:V${r}"/>`)
    .replace(/(<autoFilter ref="[A-Z]+\d+:[A-Z]+)\d+"/, `$1${r}"`);
  zip.file(ledgerSheet.path, newLedXml);
  let wbXml = book.wbXml.replace(/(<definedName[^>]*>계정원장!\$A\$\d+:\$V\$)\d+</, `$1${r}<`);
  if (book.mapAddedTo) wbXml = wbXml.replace(/(<definedName[^>]*>mapping!\$I\$1:\$L\$)(\d+)</, (m, a, b) => `${a}${Math.max(+b, book.mapAddedTo)}<`);

  // 원장 회사 확인: 지난달 원장에 있던 회사가 이번 원장에 없으면 경고
  const coOf = (site) => siteToCo.get(site) || site;
  const prevCos = new Set(existing.filter((e) => e.y === Y && e.m === prevM).map((e) => coOf(e.site)));
  const nowCos = new Set(newLedger.map((e) => e.co || e.site));
  const missingCos = [...prevCos].filter((c) => !nowCos.has(c));
  if (missingCos.length) report.warnings.push(`지난달 원장에는 있던 회사가 이번 원장에 없습니다: ${missingCos.join(', ')} — 해당 회사 원장 파일을 함께 넣었는지 확인하세요.`);

  /* --- 운반비M월: 지난달 시트 복사 --- */
  const tplXml = await zip.file(tplFreight.path).async('string');
  const tc = parseCells(tplXml);
  const T = (ref) => cellText(tc[ref], sst);
  const fr = parseRows(tplXml);
  // 납품처 목록 (C~J, 4행~)
  const lastListRow = Math.max(...[...fr.keys()].filter((n) => n >= 4 && n < 100 && ['C', 'E', 'F'].some((c) => norm(T(`${c}${n}`)))));
  const listRows = [];
  let curCo = '', curPlace = '';
  for (let n = 4; n <= lastListRow + 3; n++) {
    const co = norm(T(`C${n}`)), place = norm(T(`E${n}`)), carrier = norm(T(`F${n}`));
    if (!co && !place && !carrier) { listRows.push({ n, blank: true, co: curCo }); continue; }
    if (co && co !== curCo) { curCo = co; curPlace = ''; }
    if (place) curPlace = place;
    listRows.push({ n, co: curCo, place: curPlace, carrier, payee: norm(T(`I${n}`)), channel: norm(T(`G${n}`)), hasFormula: !!tc[`H${n}`]?.f });
  }
  const carrierToPayee = new Map(listRows.filter((x) => x.carrier && x.payee).map((x) => [x.carrier, x.payee]));
  const key = (co, place, carrier) => `${norm(co)}|${norm(place)}|${norm(carrier)}`;
  const aggMap = new Map(agg.cols.map((c) => [key(c.sajang, c.place, c.carrier), c]));
  const used = new Set();
  const changes = [];
  for (const x of listRows) {
    if (x.blank || !x.carrier || x.carrier === '합계' || x.hasFormula) continue;
    const k = key(x.co, x.place, x.carrier);
    const a = aggMap.get(k);
    const ref = `H${x.n}`;
    const s = styleOf(fr, ref);
    if (a) {
      used.add(k);
      setCell(fr, ref, numCell(ref, s, a.total));
      x.amount = a.total;
    } else {
      setCell(fr, ref, numCell(ref, s, 0));
      x.amount = 0;
      x.missing = true;
    }
  }
  // 집계파일에만 있는 납품처 처리
  //  1) 같은 회사·같은 납품처인데 운송사만 바뀐 경우 → 그 행의 운송사를 바꿈 (지급처는 다른 행 기준이 있으면 따르고, 없으면 유지)
  //  2) 같은 회사·같은 운송사인데 납품처 이름만 바뀐 경우 → 그 행의 납품처 이름을 바꿈
  //  3) 그 밖에는 같은 회사 구역 뒤 빈 행에 새로 추가
  const missingRows = () => listRows.filter((x) => !x.blank && x.missing && !x.reused);
  const fresh = agg.cols.filter((a) => !used.has(key(a.sajang, a.place, a.carrier)) && a.carrier !== '합계');
  const unplaced = [];
  for (const a of fresh) {
    const m = missingRows().find((x) => x.co === a.sajang && x.place === a.place);
    if (!m) { unplaced.push(a); continue; }
    m.reused = true;
    const payee = carrierToPayee.get(a.carrier) || m.payee || a.carrier;
    setCell(fr, `F${m.n}`, textCell(`F${m.n}`, styleOf(fr, `F${m.n}`), a.carrier));
    setCell(fr, `I${m.n}`, textCell(`I${m.n}`, styleOf(fr, `I${m.n}`), payee));
    setCell(fr, `H${m.n}`, numCell(`H${m.n}`, styleOf(fr, `H${m.n}`), a.total));
    changes.push({ type: 'change', text: `${a.sajang} · ${a.place}: 운송사 ${m.carrier} → ${a.carrier} (${m.n}행, 지급처 '${payee}')` });
    Object.assign(m, { carrier: a.carrier, payee, amount: a.total });
  }
  const unplaced2 = [];
  for (const a of unplaced) {
    const m = missingRows().find((x) => x.co === a.sajang && x.carrier === a.carrier && norm(T(`E${x.n}`)));
    if (!m) { unplaced2.push(a); continue; }
    m.reused = true;
    setCell(fr, `E${m.n}`, textCell(`E${m.n}`, styleOf(fr, `E${m.n}`), a.placeRaw || a.place));
    setCell(fr, `H${m.n}`, numCell(`H${m.n}`, styleOf(fr, `H${m.n}`), a.total));
    changes.push({ type: 'change', text: `${a.sajang}: 납품처 ${m.place} → ${a.place} (${m.n}행, 운송사 ${a.carrier})` });
    Object.assign(m, { place: a.place, amount: a.total });
  }
  for (const a of unplaced2) {
    const coRows = listRows.filter((x) => x.co === a.sajang && !x.blank);
    const lastOfCo = coRows.length ? Math.max(...coRows.map((x) => x.n)) : lastListRow;
    const slot = listRows.filter((x) => x.blank && !x.taken).sort((p, q) => p.n - q.n).find((x) => x.n > lastOfCo) || null;
    if (!slot) { changes.push({ type: 'new-fail', text: `${a.sajang} · ${a.place} · ${a.carrier} ${Math.round(a.total).toLocaleString()}원: 넣을 빈 행이 없어 추가하지 못함` }); continue; }
    slot.taken = true;
    const n = slot.n, sr = lastOfCo;
    const st = (c) => styleOf(fr, `${c}${sr}`);
    const payee = carrierToPayee.get(a.carrier) || a.carrier;
    setCell(fr, `C${n}`, textCell(`C${n}`, st('C'), a.sajang));
    setCell(fr, `E${n}`, textCell(`E${n}`, st('E'), a.placeRaw || a.place));
    setCell(fr, `F${n}`, textCell(`F${n}`, st('F'), a.carrier));
    setCell(fr, `G${n}`, textCell(`G${n}`, st('G'), a.channel));
    setCell(fr, `H${n}`, numCell(`H${n}`, st('H'), a.total));
    setCell(fr, `I${n}`, textCell(`I${n}`, st('I'), payee));
    setCell(fr, `J${n}`, fCell(`J${n}`, st('J'), `G${n}`));
    listRows.push({ n, co: a.sajang, place: a.place, carrier: a.carrier, payee, amount: a.total });
    changes.push({ type: 'new', text: `${a.sajang} · ${a.place} · ${a.carrier} ${Math.round(a.total).toLocaleString()}원 → ${n}행에 새로 추가 (지급처 '${payee}' 확인 필요)` });
  }
  for (const m of listRows.filter((x) => x.missing && !x.reused)) {
    changes.push({ type: 'missing', text: `${m.co} · ${m.place} · ${m.carrier}: 이번 집계파일에 없어 0원 처리 (${m.n}행)` });
  }
  report.listChanges = changes;
  // 월·연도 셀
  const setNum = (ref, v) => setCell(fr, ref, numCell(ref, styleOf(fr, ref), v));
  setNum('B1', M);
  if (num(tc.X7?.v) === prevM && !tc.X7?.f) setNum('X7', M);
  for (let c = colNum('X'); c <= colNum('AE'); c++) {
    const ref = `${colLet(c)}6`;
    if (tc[ref] && !tc[ref].f && num(tc[ref].v) >= 2000 && num(tc[ref].v) !== Y) setNum(ref, Y);
  }

  /* --- 대사 계산 (엑셀 수식과 같은 방식) --- */
  const allLedger = [
    ...existing.map((e) => ({ ...e, co: coOf(e.site), lead: accToLead.get(e.acc) || '', payee: vendorToPayee.get(e.m1) ?? '' })),
    ...newLedger,
  ];
  const monthLed = allLedger.filter((e) => e.y === Y && e.m === M);
  const opsBy = (co, payee) => listRows.filter((x) => x.co === co && x.payee === payee).reduce((s, x) => s + (x.amount || 0), 0);
  // 지난달 사유: P(거래처) → U
  const prevNotes = new Map();
  for (const ref of Object.keys(tc)) {
    const [, col, rn] = ref.match(/([A-Z]+)(\d+)/);
    if (col === 'P') { const p = norm(T(ref)); const u = norm(T(`U${rn}`)); if (p && u && !u.startsWith('제안)')) prevNotes.set(p, u); }
  }
  // 스필(UNIQUE/FILTER) 구역: P열 배열수식 시작 셀
  const spills = Object.entries(tc).filter(([ref, c]) => ref.startsWith('P') && /t="array"/.test(c.fTag || ''))
    .map(([ref, c]) => ({ row: +ref.slice(1), ref: (c.fTag.match(/ref="([^"]+)"/) || [])[1] }));
  const recon = [];
  for (const sp of spills) {
    const co = norm(T(`N${sp.row}`)), lead = norm(T(`O${sp.row}`));
    const vendors = [];
    for (const e of monthLed) if (e.co === co && e.lead === lead && !vendors.includes(e.m1)) vendors.push(e.m1);
    // 넣을 수 있는 행 수: 같은 회사 구역에서, 지난달 스필 범위 아래로 P·Q가 비어 있는 행까지
    const spillEnd = sp.ref ? +sp.ref.split(':')[1].slice(1) : sp.row;
    let cap = 0;
    for (let n = sp.row; n < sp.row + 60; n++) {
      if (n > sp.row && norm(T(`N${n}`)) !== co) break;
      if (n > spillEnd && (norm(T(`P${n}`)) || norm(T(`Q${n}`)))) break;
      cap++;
    }
    if (vendors.length > cap) report.warnings.push(`${co} 지급처가 ${vendors.length}곳이라 운반비${M}월 P${sp.row} 아래 칸(${cap}칸)을 넘습니다 — 엑셀에서 #SPILL! 이 나면 행을 추가하세요.`);
    vendors.forEach((v, i) => {
      const payee = vendorToPayee.get(v) ?? '(매핑 없음)';
      const S = monthLed.filter((e) => e.co === co && e.lead === lead && e.m1 === v && e.payee === payee).reduce((s, e) => s + e.dr, 0);
      const R = opsBy(co, payee);
      const lines = monthLed.filter((e) => e.co === co && e.m1 === v).sort((a, b) => Math.abs(b.dr) - Math.abs(a.dr));
      recon.push({ co, lead, row: sp.row + i, vendor: v, payee, ops: R, ledger: S, diff: R - S, prevNote: prevNotes.get(v) || '', lines: lines.slice(0, 4).map((e) => ({ amount: e.dr, memo: e.memo })) });
    });
  }
  // 회사별 합계
  const cos = [...new Set(listRows.filter((x) => !x.blank && x.co).map((x) => x.co))];
  report.companies = cos.map((co) => {
    const ops = listRows.filter((x) => x.co === co).reduce((s, x) => s + (x.amount || 0), 0);
    const ledger = monthLed.filter((e) => e.co === co && e.lead === '선행').reduce((s, e) => s + e.dr, 0);
    return { co, ops, ledger, diff: ops - ledger };
  });
  report.recon = recon;

  /* --- U열 사유: 지난달 사유 지우고, 차이 있는 지급처에 제안 --- */
  for (const ref of Object.keys(tc)) {
    const [, col, rn] = ref.match(/([A-Z]+)(\d+)/);
    if (col === 'U' && +rn >= 4 && norm(T(`N${rn}`)) && !tc[ref].f) setCell(fr, ref, blankCell(ref, tc[ref].s));
  }
  for (const x of recon) {
    if (Math.abs(x.diff) < 1) continue;
    const ref = `U${x.row}`;
    const parts = [];
    if (x.prevNote) parts.push(`지난달: ${x.prevNote}`);
    if (x.lines.length) parts.push(`원장: ${x.lines.slice(0, 2).map((l) => `${Math.round(l.amount).toLocaleString()} ${norm(l.memo).slice(0, 40)}`).join(' / ')}`);
    x.suggestion = `제안) ${parts.join(' | ') || '사유 확인 필요'}`;
    setCell(fr, ref, textCell(ref, styleOf(fr, ref) || tc[ref]?.s, x.suggestion));
  }

  /* --- 수식 정리: 시트 이름 바꾸기, 계산값 비우기(열 때 다시 계산), 스필 결과 지우기 --- */
  const oldName = tplFreight.name, newName = `운반비${M}월`;
  const spillRanges = spills.map((s) => s.ref).filter(Boolean).map((ref) => { const [a, b] = ref.split(':'); return [+a.slice(1), +b.slice(1)]; });
  for (const [rn, row] of fr) {
    for (const [c, xml] of row.cells) {
      let x = xml;
      if (/<f[\s>]/.test(x)) {
        x = x.replaceAll(`${oldName}!`, `${newName}!`).replaceAll(`'${oldName}'!`, `'${newName}'!`);
        x = x.replace(/<v>[\s\S]*?<\/v>/, '').replace(/\st="(str|n|b|e)"/, '');
      } else if (c === colNum('P') && spillRanges.some(([a, b]) => rn > a && rn <= b)) {
        x = blankCell(`P${rn}`, (x.match(/\ss="(\d+)"/) || [])[1]);
      }
      row.cells.set(c, x);
    }
  }
  let newFreightXml = tplXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, buildSheetData(fr))
    .replace(/<sheetView\b([^>]*?)(\/?)>/, (m, a, sl) => `<sheetView showGridLines="0"${a.replace(/\s(showGridLines|tabSelected)="[^"]*"/g, '')}${sl}>`);

  /* --- 현황M월: 지난달 현황 시트의 서식에 이번 집계 값 --- */
  const stXml = await zip.file(tplStatus.path).async('string');
  const sr = parseRows(stXml);
  const out = new Map();
  const range = XLSX.utils.decode_range(agg.ws['!ref']);
  for (let R0 = range.s.r; R0 <= range.e.r; R0++) {
    const rn = R0 + 1;
    const tplR = sr.get(rn);
    const row = { attrs: tplR ? tplR.attrs : ` r="${rn}"`, cells: new Map() };
    for (let C0 = range.s.c; C0 <= range.e.c; C0++) {
      const ref = `${colLet(C0 + 1)}${rn}`;
      const cell = agg.ws[ref];
      let s = tplR?.cells.get(C0 + 1)?.match(/\ss="(\d+)"/)?.[1];
      if (!s && tplR) { for (let k = C0; k >= 1 && !s; k--) s = tplR.cells.get(k)?.match(/\ss="(\d+)"/)?.[1]; }
      if (!cell || cell.v === undefined || cell.v === null || cell.v === '') { if (s) row.cells.set(C0 + 1, blankCell(ref, s)); continue; }
      let xml;
      if (cell.v instanceof Date) xml = numCell(ref, s, toSerial(cell.v));
      else if (typeof cell.v === 'number') xml = numCell(ref, s, cell.v);
      else if (typeof cell.v === 'boolean') xml = `<c r="${ref}"${s ? ` s="${s}"` : ''} t="b"><v>${cell.v ? 1 : 0}</v></c>`;
      else xml = textCell(ref, s, cell.v);
      row.cells.set(C0 + 1, xml);
    }
    if (row.cells.size || tplR) out.set(rn, row);
  }
  const mergeXml = agg.merges.length ? `<mergeCells count="${agg.merges.length}">${agg.merges.map((m) => `<mergeCell ref="${XLSX.utils.encode_range(m)}"/>`).join('')}</mergeCells>` : '';
  let newStatusXml = stXml.replace(/<sheetData>[\s\S]*?<\/sheetData>|<sheetData\/>/, buildSheetData(out))
    .replace(/<mergeCells[\s\S]*?<\/mergeCells>/, '')
    .replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="${agg.ws['!ref']}"/>`)
    .replace(/<sheetView\b([^>]*?)(\/?)>/, (m, a, sl) => `<sheetView showGridLines="0"${a.replace(/\s(showGridLines|tabSelected)="[^"]*"/g, '')}${sl}>`)
    .replace(/<(drawing|legacyDrawing|tableParts)\b[^>]*\/>/g, '')
    .replace(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g, '');
  if (mergeXml) newStatusXml = newStatusXml.replace('</sheetData>', `</sheetData>${mergeXml}`);

  /* --- 시트 추가/교체 --- */
  const ctPath = '[Content_Types].xml';
  let ct = await zip.file(ctPath).async('string');
  let relXml = book.relXml;
  const nextRid = () => `rId${Math.max(...[...relXml.matchAll(/Id="rId(\d+)"/g)].map((m) => +m[1])) + 1}`;
  const nextSheetFile = () => { let i = 1; while (zip.file(`xl/worksheets/sheet${i}.xml`)) i++; return i; };
  let nextSheetId = Math.max(...book.sheets.map((s) => s.sheetId)) + 1;
  // 시트에 딸린 메모·도형 복사
  async function copySheetRels(srcPath, dstPath, sheetXml) {
    const relPath = srcPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
    const rf = zip.file(relPath);
    if (!rf) return sheetXml.replace(/<(drawing|legacyDrawing|tableParts)\b[^>]*\/>/g, '');
    let rx = await rf.async('string');
    for (const m of [...rx.matchAll(/<Relationship\b[^>]*\/>/g)]) {
      const target = m[0].match(/Target="([^"]+)"/)[1];
      const type = m[0].match(/Type="[^"]*\/([^"/]+)"/)[1];
      if (!['comments', 'vmlDrawing', 'drawing'].includes(type)) { rx = rx.replace(m[0], ''); continue; }
      const abs = 'xl/' + target.replace('../', '');
      const ext = abs.split('.').pop();
      const base = abs.replace(/\d*\.\w+$/, '');
      let i = 1; while (zip.file(`${base}${i}.${ext}`)) i++;
      const nAbs = `${base}${i}.${ext}`;
      zip.file(nAbs, await zip.file(abs).async('uint8array'));
      const subRel = abs.replace(/([^/]+)$/, '_rels/$1.rels');
      if (zip.file(subRel)) zip.file(nAbs.replace(/([^/]+)$/, '_rels/$1.rels'), await zip.file(subRel).async('uint8array'));
      if (ext === 'xml') {
        const ctm = ct.match(new RegExp(`<Override PartName="/${abs.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}" ContentType="([^"]+)"/>`));
        if (ctm) ct = ct.replace('</Types>', `<Override PartName="/${nAbs}" ContentType="${ctm[1]}"/></Types>`);
      }
      rx = rx.replace(m[0], m[0].replace(target, '../' + nAbs.replace('xl/', '')));
    }
    zip.file(dstPath.replace('worksheets/', 'worksheets/_rels/') + '.rels', rx);
    return sheetXml;
  }
  async function putSheet(name, xml, srcPathForRels, afterName, beforeName) {
    const ex = byName(name);
    if (ex) {
      const relPath = ex.path.replace('worksheets/', 'worksheets/_rels/') + '.rels';
      if (zip.file(relPath)) zip.remove(relPath);
      xml = srcPathForRels ? await copySheetRels(srcPathForRels, ex.path, xml) : xml.replace(/<(drawing|legacyDrawing|tableParts)\b[^>]*\/>/g, '');
      zip.file(ex.path, xml);
      return;
    }
    const fileNo = nextSheetFile();
    const path = `xl/worksheets/sheet${fileNo}.xml`;
    xml = srcPathForRels ? await copySheetRels(srcPathForRels, path, xml) : xml.replace(/<(drawing|legacyDrawing|tableParts)\b[^>]*\/>/g, '');
    zip.file(path, xml);
    const rid = nextRid();
    relXml = relXml.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${fileNo}.xml"/></Relationships>`);
    ct = ct.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
    const tag = `<sheet name="${esc(name)}" sheetId="${nextSheetId++}" r:id="${rid}"/>`;
    const list = [...wbXml.matchAll(/<sheet\b[^>]*\/>/g)].map((m) => m[0]);
    const nameOfTag = (t) => unesc(t.match(/name="([^"]+)"/)[1]);
    let pos = list.length;
    const ai = afterName ? list.findIndex((t) => nameOfTag(t) === afterName) : -1;
    const bi = beforeName ? list.findIndex((t) => nameOfTag(t) === beforeName) : -1;
    if (ai >= 0) pos = ai + 1; else if (bi >= 0) pos = bi;
    list.splice(pos, 0, tag);
    wbXml = wbXml.replace(/<sheets>[\s\S]*?<\/sheets>/, `<sheets>${list.join('')}</sheets>`);
    // 시트 위치가 바뀌면 시트 전용 이름(필터 범위)의 번호도 밀어줌
    wbXml = wbXml.replace(/localSheetId="(\d+)"/g, (m, i) => `localSheetId="${+i >= pos ? +i + 1 : i}"`);
    book.sheets.push({ name, path });
  }
  await putSheet(newName, newFreightXml, tplFreight.path, tplFreight.name, null);
  await putSheet(`현황${M}월`, newStatusXml, null, null, tplStatus.name);
  // 새 운반비 시트를 열린 탭으로
  const names = [...wbXml.matchAll(/<sheet\b[^>]*name="([^"]+)"/g)].map((m) => unesc(m[1]));
  wbXml = wbXml.replace(/activeTab="\d+"/, `activeTab="${names.indexOf(newName)}"`).replace(/firstSheet="\d+"/, 'firstSheet="0"');
  // 열 때 전체 재계산, 계산 순서 파일 제거
  wbXml = wbXml.replace(/<calcPr\b([^>]*?)\/>/, (m, a) => `<calcPr${a.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`);
  if (!/<calcPr/.test(wbXml)) wbXml = wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  zip.remove('xl/calcChain.xml');
  ct = ct.replace(/<Override[^>]*calcChain[^>]*\/>/, '');
  relXml = relXml.replace(/<Relationship[^>]*calcChain[^>]*\/>/, '');
  // 다른 시트의 '선택됨' 표시 해제
  for (const s of book.sheets) {
    if (!s.path || s.name === newName) continue;
    const f = zip.file(s.path);
    if (!f) continue;
    const x = await f.async('string');
    if (/tabSelected="1"/.test(x)) zip.file(s.path, x.replace(/\stabSelected="1"/, ''));
  }
  zip.file('xl/workbook.xml', wbXml);
  zip.file('xl/_rels/workbook.xml.rels', relXml);
  zip.file(ctPath, ct);

  report.ledgerCount = newLedger.length;
  report.ledgerTotal = newLedger.reduce((s, e) => s + e.dr, 0);
  report.opsTotal = listRows.reduce((s, x) => s + (x.amount || 0), 0);
  const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { data, report, fileName: `운반비_검증_${String(Y).slice(2)}${String(M).padStart(2, '0')}.xlsx` };
}
