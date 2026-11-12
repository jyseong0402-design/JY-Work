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
// 운반비 검증에서 0으로 두는 납품처(후행물류비 등). 현황 시트에는 집계 값 그대로, 운반비N월 H열만 0
export const ZERO_PLACES = ['엑소푸레쉬'];
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
  return { wb, sheetName: name, ws, grid, hr, totalRow, cols, year: firstDate.getFullYear(), month: firstDate.getMonth() + 1, merges: ws['!merges'] || [] };
}

/* ---------------- 1-2. 집계파일 '납품처' 시트 (채널 구분의 기준) ---------------- */
// 채널(새 구분) → 운반비 시트 J열(요약표 분류)
const KNOWN_CH = ['경매', 'Hyper/SSM', '온라인', 'F/S', 'B2B', '기타'];
export const channelGroup = (ch) => (ch === '간선_외부매입' ? '매입운반비' : String(ch).startsWith('간선') ? '간선' : ch);

export function readDeliverySheet(XLSX, wb, month) {
  const ws = wb.Sheets['납품처'];
  if (!ws) return null;
  const g = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = g.findIndex((r) => r && r.some((x) => norm(x) === '채널'));
  if (hr < 0) return null;
  const cB = g[hr].findIndex((x) => norm(x) === '채널');
  const org = XLSX.utils.decode_range(ws['!ref'] || 'A1').s;
  const monthCol = (m) => {
    const lab = `${String(m).padStart(2, '0')}월`;
    const c = (g[hr] || []).findIndex((x) => norm(x) === lab || norm(x) === `${m}월`);
    if (c < 0) return null;
    const sub = (g[hr + 1] || []).map(norm);
    let amt = c, plt = null;
    for (let k = c; k < c + 6; k++) { if (sub[k] === '운반비') { amt = k; break; } }
    for (let k = amt + 1; k < amt + 4; k++) { if (sub[k] === 'PLT') { plt = k; break; } }
    return { amt, plt };
  };
  const cur = monthCol(month), prev = monthCol(month === 1 ? 12 : month - 1);
  const rows = [];
  for (let i = hr + 2; i < g.length; i++) {
    const r = g[i];
    if (!r || !r.slice(cB, cB + 3).some((x) => norm(x))) continue;
    const ch = norm(r[cB]);
    const ref = XLSX.utils.encode_cell({ r: i + org.r, c: (cur ? cur.amt : 0) + org.c });
    const f = cur && ws[ref] && ws[ref].f ? String(ws[ref].f) : '';
    rows.push({
      i, ch, client: norm(r[cB + 1]), place: norm(r[cB + 2]), clientRaw: r[cB + 1] ?? '', placeRaw: r[cB + 2] ?? '',
      kind: /요약$/.test(ch) ? 'sum' : ch === '총합계' ? 'total' : /^[A-Z]+\d+(\+[A-Z]+\d+)+$/.test(f.replace(/\$/g, '')) ? 'group' : 'row',
      cur: cur ? num(r[cur.amt]) : 0, curPlt: cur && cur.plt !== null ? num(r[cur.plt]) : 0,
      prev: prev ? num(r[prev.amt]) : 0, prevPlt: prev && prev.plt !== null ? num(r[prev.plt]) : 0,
    });
  }
  return { rows, hasCur: !!cur, hasPrev: !!prev };
}

// 현황 열(회사·납품처·운송사, Total) → 납품처 시트의 채널. 금액이 같은 행(또는 같은 거래처 행들의 합)을 찾고, 이름이 비슷한 쪽을 고름
function tokens(s) { return new Set(String(s).replace(/[()\[\]_,\-\->\/·]/g, ' ').split(/\s+/).map((t) => t.trim()).filter((t) => t.length >= 2)); }
export function matchChannels(cols, dv) {
  const out = new Map();
  if (!dv) return out;
  const det = dv.rows.filter((r) => r.kind === 'row' || r.kind === 'group');
  const groups = new Map();
  for (const r of dv.rows.filter((x) => x.kind === 'row')) {
    const k = `${r.ch}|${r.client}`;
    if (!groups.has(k)) groups.set(k, { ch: r.ch, client: r.client, place: '', cur: 0, rows: [] });
    const gg = groups.get(k); gg.cur += r.cur; gg.rows.push(r);
  }
  const rest = [];
  const kOf = (col) => `${norm(col.sajang)}|${norm(col.place)}|${norm(col.carrier)}`;
  for (const col of cols) {
    if (!col.total || col.carrier === '합계') continue;
    const T = col.total;
    let cands = det.filter((r) => Math.abs(r.cur - T) <= 1).map((r) => ({ ch: r.ch, text: `${r.client} ${r.place}`, row: r }));
    if (!cands.length) cands = [...groups.values()].filter((gq) => gq.rows.length > 1 && Math.abs(gq.cur - T) <= 1).map((gq) => ({ ch: gq.ch, text: gq.client, group: gq }));
    if (!cands.length) { rest.push(col); continue; }
    const mine = tokens(`${col.place} ${col.carrier} ${col.company}`);
    const score = (c) => [...tokens(c.text)].filter((t) => [...mine].some((m) => m.includes(t) || t.includes(m))).length;
    cands.sort((a, b) => score(b) - score(a));
    const best = cands[0];
    const sure = cands.length === 1 || score(best) > score(cands[1]) || cands.every((c) => c.ch === best.ch);
    out.set(kOf(col), { ch: best.ch, sure, row: best.row || null, group: best.group || null, aggCh: norm(col.channel) });
  }
  // 2차: 같은 회사의 남은 열 2~3개를 합친 금액이 한 행과 같으면 (예: GS 네트웍스 + GS 신미래 = GS리테일 피킹비)
  const used = new Set([...out.values()].map((m) => m.row && m.row.i));
  const free = det.filter((r) => r.kind === 'row' && r.cur && !used.has(r.i));
  const left = new Set(rest);
  for (const r of free) {
    const pool = rest.filter((c) => left.has(c));
    let hit = null;
    for (let x = 0; x < pool.length && !hit; x++) for (let y = x + 1; y < pool.length && !hit; y++) {
      if (pool[x].company !== pool[y].company) continue;
      if (Math.abs(pool[x].total + pool[y].total - r.cur) <= 1) hit = [pool[x], pool[y]];
      for (let z = y + 1; z < pool.length && !hit; z++) if (pool[z].company === pool[x].company && Math.abs(pool[x].total + pool[y].total + pool[z].total - r.cur) <= 1) hit = [pool[x], pool[y], pool[z]];
    }
    if (hit) for (const c of hit) { left.delete(c); out.set(kOf(c), { ch: r.ch, sure: true, row: null, group: { rows: [r] }, combined: true, aggCh: norm(c.channel) }); }
  }
  // 3차: 금액이 안 맞으면 이름으로 (운송사 = 납품처 칸, 지역명 겹침). 확실하지 않음 표시
  for (const c of rest.filter((q) => left.has(q))) {
    const mine = tokens(`${c.place} ${c.carrier}`);
    let best = null, bs = 0;
    for (const r of det) {
      const sc = [...tokens(`${r.client} ${r.place}`)].filter((t) => [...mine].some((m) => m.includes(t) || t.includes(m))).length + (norm(r.place) === norm(c.carrier) ? 1 : 0);
      if (sc > bs) { bs = sc; best = r; }
    }
    if (best && bs >= 2) out.set(kOf(c), { ch: best.ch, sure: false, row: null, byName: `${best.client} ${best.place}`, aggCh: norm(c.channel) });
  }
  return out;
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
export async function buildVerification({ XLSX, JSZip, aggBuf, ledgerBufs, prevBuf, extraMapping = [], overrides = {} }) {
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
    listRows.push({ n, co: curCo, place: curPlace, carrier, payee: norm(T(`I${n}`)), channel: norm(T(`G${n}`)), hasFormula: !!tc[`H${n}`]?.f, prev: tc[`H${n}`]?.f ? null : num(tc[`H${n}`]?.v) });
  }
  const key = (co, place, carrier) => `${norm(co)}|${norm(place)}|${norm(carrier)}`;
  // 납품처 금액: 사용자가 고친 값 > 0 처리 규칙 > 집계 Total
  const keyOfAgg = (a) => key(a.sajang, a.place, a.carrier);
  const amt = (a) => {
    const k = keyOfAgg(a);
    if (overrides[k] !== undefined && overrides[k] !== null && overrides[k] !== '') return num(overrides[k]);
    if (ZERO_PLACES.some((z) => a.place.includes(z) || a.carrier.includes(z))) return 0;
    return a.total;
  };
  const carrierToPayee = new Map(listRows.filter((x) => x.carrier && x.payee).map((x) => [x.carrier, x.payee]));
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
      x.amount = amt(a); x.agg = a.total; x.key = keyOfAgg(a);
      setCell(fr, ref, numCell(ref, s, x.amount));
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
    setCell(fr, `H${m.n}`, numCell(`H${m.n}`, styleOf(fr, `H${m.n}`), amt(a)));
    changes.push({ type: 'change', text: `${a.sajang} · ${a.place}: 운송사 ${m.carrier} → ${a.carrier} (${m.n}행, 지급처 '${payee}')` });
    Object.assign(m, { carrier: a.carrier, payee, amount: amt(a), agg: a.total, key: keyOfAgg(a) });
  }
  const unplaced2 = [];
  for (const a of unplaced) {
    const m = missingRows().find((x) => x.co === a.sajang && x.carrier === a.carrier && norm(T(`E${x.n}`)));
    if (!m) { unplaced2.push(a); continue; }
    m.reused = true;
    setCell(fr, `E${m.n}`, textCell(`E${m.n}`, styleOf(fr, `E${m.n}`), a.placeRaw || a.place));
    setCell(fr, `H${m.n}`, numCell(`H${m.n}`, styleOf(fr, `H${m.n}`), amt(a)));
    changes.push({ type: 'change', text: `${a.sajang}: 납품처 ${m.place} → ${a.place} (${m.n}행, 운송사 ${a.carrier})` });
    Object.assign(m, { place: a.place, amount: amt(a), agg: a.total, key: keyOfAgg(a) });
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
    setCell(fr, `H${n}`, numCell(`H${n}`, st('H'), amt(a)));
    setCell(fr, `I${n}`, textCell(`I${n}`, st('I'), payee));
    setCell(fr, `J${n}`, fCell(`J${n}`, st('J'), `G${n}`));
    listRows.push({ n, co: a.sajang, place: a.place, carrier: a.carrier, payee, amount: amt(a), agg: a.total, key: keyOfAgg(a), prev: null });
    changes.push({ type: 'new', text: `${a.sajang} · ${a.place} · ${a.carrier} ${Math.round(a.total).toLocaleString()}원 → ${n}행에 새로 추가 (지급처 '${payee}' 확인 필요)` });
  }
  for (const m of listRows.filter((x) => x.missing && !x.reused)) {
    changes.push({ type: 'missing', text: `${m.co} · ${m.place} · ${m.carrier}: 이번 집계파일에 없어 0원 처리 (${m.n}행)` });
  }
  report.listChanges = changes;
  report.places = listRows.filter((x) => !x.blank && x.carrier && x.carrier !== '합계' && !x.hasFormula).map((x) => ({
    row: x.n, co: x.co, place: x.place, carrier: x.carrier, key: x.key || key(x.co, x.place, x.carrier),
    amount: x.amount || 0, agg: x.agg ?? null, prev: x.prev ?? null,
    zeroed: x.agg !== undefined && x.agg !== null && (x.amount || 0) === 0 && x.agg !== 0 && overrides[x.key] === undefined,
    overridden: x.key !== undefined && overrides[x.key] !== undefined && overrides[x.key] !== '',
  }));
  for (const p of report.places) {
    if (p.overridden) report.info.push(`수기 수정: ${p.co} · ${p.place} · ${p.carrier} ${Math.round(p.agg ?? 0).toLocaleString()} → ${Math.round(p.amount).toLocaleString()}원 (${p.row}행)`);
    if (p.zeroed) report.info.push(`0 처리(규칙): ${p.co} · ${p.place} · ${p.carrier} 집계 ${Math.round(p.agg).toLocaleString()}원 → 0원 (${p.row}행)`);
  }
  // 월·연도 셀
  const setNum = (ref, v) => setCell(fr, ref, numCell(ref, styleOf(fr, ref), v));
  setNum('B1', M);
  if (num(tc.X7?.v) === prevM && !tc.X7?.f) setNum('X7', M);
  for (let c = colNum('X'); c <= colNum('AE'); c++) {
    const ref = `${colLet(c)}6`;
    if (tc[ref] && !tc[ref].f && num(tc[ref].v) >= 2000 && num(tc[ref].v) !== Y) setNum(ref, Y);
  }

  for (let c = colNum('Y'); c <= colNum('AE'); c++) {
    const ref = `${colLet(c)}7`;
    if (tc[ref] && !tc[ref].f && num(tc[ref].v) === prevM) setNum(ref, M);
  }

  /* --- 오른쪽 월별 표: 채널별 후행물류비(1월~당월 열)에 당월 열 추가, 채널별 선행물류비는 당월 열로 옮김 --- */
  {
    const mCol = (m) => colLet(colNum('Y') + m - 1);
    const pc = mCol(prevM), cc = mCol(M);
    const sortedRows = [...fr.keys()].sort((a, b) => a - b);
    const titleRow = (txt) => sortedRows.find((n) => norm(T(`W${n}`)) === txt);
    const leadTitle = titleRow('채널별 선행물류비');
    const lagTitle = titleRow('채널별 후행물류비');
    const lastRow = sortedRows[sortedRows.length - 1];
    const tok = (f, from, to) => f.replace(new RegExp(`(^|[^A-Z$!])(\\$?)${from}(?=\\$?\\d)`, 'g'), (m0, p1, p2) => p1 + p2 + to);
    // 후행: 머리글(1..지난달) 다음 칸에 당월 번호, 각 행은 지난달 열 수식을 복사(월 번호 칸만 당월로)
    if (lagTitle && M !== 1) {
      const hdr = sortedRows.find((n) => n > lagTitle && num(tc[`Y${n}`]?.v) === 1 && num(tc[`${pc}${n}`]?.v) === prevM);
      if (hdr && !tc[`${cc}${hdr}`]?.v) {
        setCell(fr, `${cc}${hdr}`, numCell(`${cc}${hdr}`, styleOf(fr, `${pc}${hdr}`), M));
        const adj = [];
        for (let n = hdr + 1; n < (leadTitle || lastRow + 1); n++) {
          const c0 = tc[`${pc}${n}`];
          if (!c0 || !c0.f) continue;
          let f = unesc(c0.f).split(`${pc}$${hdr}`).join(`${cc}$${hdr}`);
          const mAdj = f.match(/\)\/10\^6([+-][\d.]+)$/);
          if (mAdj) { f = f.slice(0, -mAdj[1].length); adj.push(`${T(`W${n}`)} ${T(`X${n}`)}(지난달 ${mAdj[1]})`); }
          setCell(fr, `${cc}${n}`, fCell(`${cc}${n}`, styleOf(fr, `${pc}${n}`), f));
        }
        report.info.push(`채널별 후행물류비 표에 ${M}월 열(${cc})을 추가했습니다.`);
        if (adj.length) report.warnings.push(`채널별 후행물류비 ${M}월 열: 지난달 수식 끝에 손으로 더한 값이 있던 곳은 빼고 넣었습니다 — ${adj.join(', ')}. 이번 달 가산값이 있으면 ${cc}열에 직접 더해 주세요.`);
      }
    }
    // 선행: 지난달 열의 수식들을 당월 열로 옮기고 지난달 열은 비움
    if (leadTitle) {
      let moved = 0;
      for (let n = leadTitle + 1; n <= lastRow; n++) {
        const row = fr.get(n); if (!row) continue;
        const xml = row.cells.get(colNum(pc));
        if (!xml || !/<f[\s>\/]/.test(xml)) continue;
        const curStyle = styleOf(fr, `${cc}${n}`);
        const nx = xml.replace(`r="${pc}${n}"`, `r="${cc}${n}"`)
          .replace(/(<f\b[^>]*?ref=")([^"]+)(")/, (m0, a1, r1, a2) => a1 + r1.replace(new RegExp(`\\b${pc}(?=\\d)`, 'g'), cc) + a2)
          .replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/, (m0, a1, body, a2) => a1 + tok(body, pc, cc) + a2);
        row.cells.set(colNum(cc), nx);
        row.cells.set(colNum(pc), blankCell(`${pc}${n}`, curStyle));
        moved++;
      }
      if (moved) report.info.push(`채널별 선행물류비 표를 ${prevM}월 열(${pc})에서 ${M}월 열(${cc})로 옮겼습니다 (${moved}칸).`);
    }
  }

  /* --- 채널 구분(G열: 집계파일 납품처 시트 기준, J열: 요약표 분류) --- */
  const dv = readDeliverySheet(XLSX, agg.wb, M);
  const chMap = matchChannels(agg.cols, dv);
  report.channelChanges = [];
  if (!dv) report.warnings.push("집계파일에 '납품처' 시트가 없어 채널 구분은 지난달 그대로 두었습니다.");
  else {
    let unmatched = [];
    for (const x of listRows) {
      if (x.blank || !x.carrier || x.carrier === '합계' || x.hasFormula && !x.key) continue;
      const k = x.key || key(x.co, x.place, x.carrier);
      const m = chMap.get(k);
      const oldG = norm(T(`G${x.n}`)), oldJ = norm(T(`J${x.n}`)) || oldG;
      if (!m) { if ((x.amount || 0) !== 0) unmatched.push(`${x.co} · ${x.place} · ${x.carrier}`); continue; }
      const newG = m.ch, newJ = channelGroup(m.ch);
      x.channel = newG; x.group = newJ;
      if (newG !== oldG) setCell(fr, `G${x.n}`, textCell(`G${x.n}`, styleOf(fr, `G${x.n}`) || styleOf(fr, `G${x.n - 1}`), newG));
      x.jSet = newJ;
      const aggDiff = m.aggCh && KNOWN_CH.includes(m.aggCh) && channelGroup(m.aggCh) !== newJ ? m.aggCh : '';
      if (newG !== oldG || newJ !== oldJ) report.channelChanges.push({ row: x.n, co: x.co, place: x.place, carrier: x.carrier, from: oldG + (oldJ !== oldG ? ` / ${oldJ}` : ''), to: newG + (newJ !== newG ? ` / ${newJ}` : ''), sure: m.sure, aggCh: aggDiff });
    }
    // J열 표기 통일: 분류가 G와 같으면 =G수식(원래 양식), 다르면(간선_*, 간선_외부매입) 분류 글자
    for (const x of listRows) {
      if (x.blank || !x.carrier || x.carrier === '합계') continue;
      const g = x.channel || norm(T(`G${x.n}`));
      if (!g) continue;
      const j = x.jSet || channelGroup(g);
      const jr = `J${x.n}`, js = styleOf(fr, jr) || styleOf(fr, `J${x.n - 1}`);
      setCell(fr, jr, j === g ? fCell(jr, js, `G${x.n}`) : textCell(jr, js, j));
    }
    const byName = [...chMap.entries()].filter(([, m]) => m.byName);
    if (byName.length) report.warnings.push(`납품처 시트 금액과 맞지 않아 이름으로 채널을 정한 곳(확인 필요): ${byName.map(([k, m]) => `${k.split('|').slice(1).join(' · ')} → ${m.ch} (${m.byName})`).join(', ')}`);
    const tot = dv.rows.find((r) => r.kind === 'total');
    const aggSum = agg.cols.filter((c) => c.carrier !== '합계' && !ZERO_PLACES.some((z) => norm(c.place).includes(z))).reduce((t, c) => t + (c.total || 0), 0);
    if (tot && Math.abs(tot.cur - aggSum) > 1) report.warnings.push(`집계파일 안에서 '납품처' 시트 총합계(${Math.round(tot.cur).toLocaleString()})와 '현황' 합계(${Math.round(aggSum).toLocaleString()})가 ${Math.round(aggSum - tot.cur).toLocaleString()}원 다릅니다. 집계 담당 부서 확인이 필요할 수 있습니다.`);
    if (unmatched.length) report.warnings.push(`집계파일 납품처 시트에서 채널을 찾지 못해 지난달 구분을 그대로 둔 곳: ${unmatched.join(', ')}`);
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
      let R = opsBy(co, payee);
      // 운영 금액(R)을 특정 셀로 직접 지정한 수식(예: R10 = H25)이 있으면 엑셀과 같게 계산하고, 거래처가 바뀌었으면 알림
      const rf = tc[`R${sp.row + i}`]?.f;
      const direct = rf && unesc(rf).match(/^\$?H\$?(\d+)$/);
      if (direct) {
        const target = listRows.find((x) => x.n === +direct[1]);
        R = target?.amount || 0;
        const prevVendor = norm(T(`P${sp.row + i}`));
        if (prevVendor !== v) report.warnings.push(`운반비${M}월 R${sp.row + i}은 H${direct[1]}(${target ? `${target.place}/${target.carrier}` : '빈 행'})를 직접 가리키는 수식인데, 이번 달 ${sp.row + i}행 거래처가 '${prevVendor}' → '${v}'로 바뀌었습니다. 엑셀에서 R${sp.row + i} 수식을 확인하세요.`);
      }
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

  /* --- '납품처' 시트: 전월 vs 당월 비교 (집계파일 납품처 시트 기준, 검증에서 고친 금액 반영) --- */
  const dlv = byName('납품처');
  if (dlv && dv && dv.hasCur) {
    const dXml = await zip.file(dlv.path).async('string');
    const dr = parseRows(dXml);
    const dc = parseCells(dXml);
    const DT = (ref) => norm(cellText(dc[ref], sst));
    // 본보기 서식: 채널별 상세행, 묶음행, 요약행, 총합계행
    const styleRow = (rn) => { const r0 = dr.get(rn); const st = {}; if (r0) for (const [c, x] of r0.cells) st[c] = (x.match(/\ss="(\d+)"/) || [])[1]; return st; };
    const tplRows = [...dr.keys()].filter((n) => n >= 5).sort((a, b) => a - b);
    const roleRow = {};
    for (const n of tplRows) {
      const b = DT(`B${n}`);
      if (!b) continue;
      const isGroup = /^=?SUM\(|^SUM\(/.test(unesc(dc[`E${n}`]?.f || '')) && !/요약|합계/.test(b);
      const k = /요약$/.test(b) ? `sum:${b}` : b === '총합계' ? 'total' : isGroup ? `group:${b}` : `row:${b}`;
      if (!roleRow[k]) roleRow[k] = n;
      if (/요약$/.test(b) && !roleRow.sumAny) roleRow.sumAny = n;
      if (!/요약|합계/.test(b) && !isGroup && !roleRow.rowAny) roleRow.rowAny = n;
      if (isGroup && !roleRow.groupAny) roleRow.groupAny = n;
    }
    const baseCh = (ch) => (String(ch).startsWith('간선') ? '간선' : ch);
    const pickStyle = (r) => {
      const n = r.kind === 'total' ? roleRow.total
        : r.kind === 'sum' ? (roleRow[`sum:${r.ch}`] || roleRow[`sum:${baseCh(r.ch.replace(/ 요약$/, ''))} 요약`] || roleRow.sumAny)
        : r.kind === 'group' ? (roleRow[`group:${r.ch}`] || roleRow.groupAny || roleRow[`row:${r.ch}`])
        : (roleRow[`row:${r.ch}`] || roleRow[`row:${baseCh(r.ch)}`] || roleRow.rowAny);
      return styleRow(n || 6);
    };
    // 검증에서 고친 금액(수기 수정·0 처리)을 당월에 반영: 한 행에 맞춰진 경우만
    const delta = new Map();
    for (const x of listRows) {
      if (!x.key || x.agg === undefined || x.agg === null) continue;
      const d = (x.amount || 0) - x.agg;
      if (Math.abs(d) < 1) continue;
      const m = chMap.get(x.key);
      const tgt = m && (m.row || (m.combined && m.group.rows[0]));
      if (tgt) delta.set(tgt.i, (delta.get(tgt.i) || 0) + d);
      else if (m) report.warnings.push(`'납품처' 시트: ${x.place}/${x.carrier}의 수정 금액(${Math.round(d).toLocaleString()})을 넣을 행을 특정하지 못해 집계 값 그대로 두었습니다.`);
    }
    const chDelta = new Map(); let totDelta = 0;
    for (const r of dv.rows) if (delta.has(r.i)) { chDelta.set(baseCh(r.ch), (chDelta.get(baseCh(r.ch)) || 0) + delta.get(r.i)); totDelta += delta.get(r.i); }
    const out = new Map();
    for (const n of [...dr.keys()].filter((k) => k < 5)) out.set(n, dr.get(n));
    const pm = M === 1 ? 12 : M - 1;
    const setHdr = (ref, text) => setCell(out, ref, textCell(ref, styleOf(dr, ref), text));
    setHdr('E3', `${pm}월`); setHdr('H3', `${M}월`);
    let n = 4;
    const cols = 'ABCDEFGHIJKLM'.split('');
    for (const r of dv.rows) {
      n++;
      const st = pickStyle(r);
      const cells = new Map();
      const put = (col, xml) => cells.set(colNum(col), xml);
      const S = (col) => st[colNum(col)];
      if (S('A')) put('A', blankCell(`A${n}`, S('A')));
      put('B', r.ch ? textCell(`B${n}`, S('B'), r.ch) : blankCell(`B${n}`, S('B')));
      put('C', r.clientRaw !== '' ? textCell(`C${n}`, S('C'), r.clientRaw) : blankCell(`C${n}`, S('C')));
      put('D', r.placeRaw !== '' ? textCell(`D${n}`, S('D'), r.placeRaw) : blankCell(`D${n}`, S('D')));
      const d = r.kind === 'total' ? totDelta : r.kind === 'sum' ? (chDelta.get(baseCh(r.ch.replace(/ 요약$/, ''))) || 0) : (delta.get(r.i) || 0);
      const v = (x) => (x ? numCell : (ref, s2) => blankCell(ref, s2));
      put('E', r.prev ? numCell(`E${n}`, S('E'), r.prev) : blankCell(`E${n}`, S('E')));
      put('F', r.prevPlt ? numCell(`F${n}`, S('F'), r.prevPlt) : blankCell(`F${n}`, S('F')));
      put('G', fCell(`G${n}`, S('G'), `IFERROR(E${n}/F${n},0)`));
      put('H', (r.cur + d) ? numCell(`H${n}`, S('H'), r.cur + d) : blankCell(`H${n}`, S('H')));
      put('I', r.curPlt ? numCell(`I${n}`, S('I'), r.curPlt) : blankCell(`I${n}`, S('I')));
      put('J', fCell(`J${n}`, S('J'), `IFERROR(H${n}/I${n},0)`));
      put('K', fCell(`K${n}`, S('K'), `H${n}-E${n}`));
      put('L', fCell(`L${n}`, S('L'), `I${n}-F${n}`));
      put('M', fCell(`M${n}`, S('M'), `J${n}-G${n}`));
      void v; void cols;
      const tplAttrs = dr.get(roleRow.rowAny)?.attrs || '';
      const ht = (tplAttrs.match(/\s(ht="[^"]+"|customHeight="1")/g) || []).join('');
      out.set(n, { attrs: ` r="${n}"${ht}`, cells });
    }
    let newD = dXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, buildSheetData(out))
      .replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="B3:M${n}"/>`)
      .replace(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g, '')
      .replace(/<sheetView\b([^>]*?)(\/?)>/, (m0, a, sl) => `<sheetView showGridLines="0"${a.replace(/\s(showGridLines|tabSelected)="[^"]*"/g, '')}${sl}>`);
    // 머리글(3~4행) 병합만 남김
    newD = newD.replace(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/, (m0, body) => {
      const keep = [...body.matchAll(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g)].filter((mm) => +mm[4] <= 4).map((mm) => mm[0]);
      return keep.length ? `<mergeCells count="${keep.length}">${keep.join('')}</mergeCells>` : '';
    });
    zip.file(dlv.path, newD);
    report.deliverySheet = { rows: dv.rows.length, prev: pm, cur: M, adjusted: delta.size };
    report.info.push(`'납품처' 시트를 ${pm}월 vs ${M}월 비교로 새로 만들었습니다 (${dv.rows.length}행${delta.size ? `, 검증에서 고친 금액 ${delta.size}곳 반영` : ''}).`);
  } else if (dlv && !dv) {
    report.warnings.push("집계파일 '납품처' 시트를 읽지 못해 검증파일 '납품처' 시트는 그대로 두었습니다.");
  }

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
  return { data, report, fileName: `Freight_Check_${String(Y).slice(2)}${String(M).padStart(2, "0")}.xlsx` };
}
