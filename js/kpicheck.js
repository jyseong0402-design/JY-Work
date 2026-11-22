// 🧮 베트남 KPI 양식 점검 + 다음 달 양식 만들기
//  - 엑셀 파일(xlsx)을 JSZip으로 열어 시트 XML을 직접 읽음 (수식·서식 번호 그대로)
//  - 시트마다 '월 머리글'(1~12, N월, Jan…, 날짜)이 이어진 묶음을 찾아 당월 열을 전월 열과 셀 단위로 비교
//  - 다음 달 양식: 당월 열의 수식·서식을 다음 달 열로 '오른쪽 채우기'처럼 복사(상대 참조만 한 칸 이동)

// 원본 데이터를 붙여넣는 시트(점검·채우기 제외)
export const RAW_SHEET = /^(GL|TB|PL$|Detail of|매출원가|폐기원가|수불부|회차별|샘플|Mapping|단가표|Actual 단가)/i;

const unxml = (s) => String(s ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const enxml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export const colNum = (L) => { let n = 0; for (const ch of L) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
export const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1 - m) / 26; } return s; };

/* ---------- 수식 참조 이동 ---------- */
// 문자열("…")과 시트 이름('…'!)은 건드리지 않고, A1·$A$1·A:A·1:1 참조의 상대 부분만 옮김
const REF = /(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\d(A-Za-z_])|(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?![\d(A-Za-z_])|(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\d(A-Za-z_])/g;
function mapRefs(f, fn) {
  let out = '', i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') { const j = f.indexOf('"', i + 1); const e = j < 0 ? f.length : j + 1; out += f.slice(i, e); i = e; continue; }
    if (ch === "'") { let j = i + 1; while (j < f.length && !(f[j] === "'" && f[j + 1] !== "'")) j += f[j] === "'" ? 2 : 1; out += f.slice(i, j + 1); i = j + 1; continue; }
    // 참조가 아닌 글자 덩어리(함수 이름·정의된 이름) 판단을 위해 앞 글자 확인
    let j = i; while (j < f.length && f[j] !== '"' && f[j] !== "'") j++;
    out += f.slice(i, j).replace(REF, (m, ...g) => {
      const off = g[g.length - 2], src = g[g.length - 1], prev = src[off - 1 + 0] ?? '';
      // 앞이 글자·숫자·_·. 이면 이름의 일부(예: LOG10, Sheet1.A) → 그대로
      const before = src[off - 1];
      if (before && /[A-Za-z0-9_.]/.test(before)) return m;
      return fn(g, m, prev);
    });
    i = j;
  }
  return out;
}
export function shiftFormula(f, dr, dc) {
  return mapRefs(f, (g, m) => {
    if (g[1] !== undefined) { // A1
      const [ca, cl, ra, rw] = [g[0], g[1], g[2], g[3]];
      const c = ca ? colNum(cl) : colNum(cl) + dc, r = ra ? +rw : +rw + dr;
      if (c < 1 || r < 1) return '#REF!';
      return `${ca}${colName(c)}${ra}${r}`;
    }
    if (g[5] !== undefined) { // A:A
      const c1 = g[4] ? colNum(g[5]) : colNum(g[5]) + dc, c2 = g[6] ? colNum(g[7]) : colNum(g[7]) + dc;
      return `${g[4]}${colName(c1)}:${g[6]}${colName(c2)}`;
    }
    const r1 = g[8] ? +g[9] : +g[9] + dr, r2 = g[10] ? +g[11] : +g[11] + dr; // 1:1
    return `${g[8]}${r1}:${g[10]}${r2}`;
  });
}
// 위치와 상관없는 비교용 형태(R1C1 비슷하게): 상대 참조는 칸 차이로
export function relForm(f, r, c) {
  return mapRefs(f, (g) => {
    if (g[1] !== undefined) return `${g[2] ? 'R' + g[3] : 'R[' + (+g[3] - r) + ']'}${g[0] ? 'C' + colNum(g[1]) : 'C[' + (colNum(g[1]) - c) + ']'}`;
    if (g[5] !== undefined) return `${g[4] ? 'C' + colNum(g[5]) : 'C[' + (colNum(g[5]) - c) + ']'}:${g[6] ? 'C' + colNum(g[7]) : 'C[' + (colNum(g[7]) - c) + ']'}`;
    return `${g[8] ? 'R' + g[9] : 'R[' + (+g[9] - r) + ']'}:${g[10] ? 'R' + g[11] : 'R[' + (+g[11] - r) + ']'}`;
  });
}

/* ---------- 통합문서 열기 ---------- */
export async function openBook(buf) {
  const { default: JSZip } = await import('./vendor/jszip.mjs');
  const zip = await JSZip.loadAsync(buf);
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rmap = Object.fromEntries([...rels.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
  const sheets = [...wb.matchAll(/<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"/g)].map((m) => ({ name: unxml(m[1]), path: 'xl/' + rmap[m[2]].replace(/^\/?xl\//, '') }));
  let sst = [];
  const sf = zip.file('xl/sharedStrings.xml');
  if (sf) sst = [...(await sf.async('string')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => unxml([...m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((x) => x[1]).join('')));
  return { zip, sheets, sst, wbXml: wb };
}

/* ---------- 시트 읽기 ---------- */
// cells: Map key "r,c" → {r,c,s,t,v,f(전개된 수식),shared,raw}
export function parseSheet(xml, sst) {
  const cells = new Map(), masters = {};
  const rowRe = /<row [^>]*?r="(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g;
  let m;
  while ((m = rowRe.exec(xml))) {
    if (!m[2]) continue;
    const cRe = /<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let c;
    while ((c = cRe.exec(m[2]))) {
      const col = colNum(c[1]), row = +c[2], attr = c[3], inner = c[4] || '';
      const s = +((attr.match(/\bs="(\d+)"/) || [])[1] || 0), t = (attr.match(/\bt="(\w+)"/) || [])[1] || '';
      let v = (inner.match(/<v>([^<]*)<\/v>/) || [])[1];
      if (v !== undefined) v = t === 's' ? sst[+v] : t === 'str' || t === 'e' ? unxml(v) : +v;
      else if (t === 'inlineStr') v = unxml([...inner.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map((x) => x[1]).join(''));
      const fm = inner.match(/<f([^>]*)>([^<]*)<\/f>|<f([^>]*)\/>/);
      let f = null, shared = null;
      if (fm) {
        const fa = fm[1] ?? fm[3] ?? '', ftxt = fm[2] !== undefined ? unxml(fm[2]) : '';
        const si = (fa.match(/si="(\d+)"/) || [])[1];
        if (/t="shared"/.test(fa) && si !== undefined) {
          shared = si;
          if (ftxt) { masters[si] = { f: ftxt, r: row, c: col }; f = ftxt; } else f = { si };
        } else if (!/t="array"|t="dataTable"/.test(fa) || ftxt) f = ftxt;
      }
      cells.set(row + ',' + col, { r: row, c: col, s, t, v, f, shared });
    }
  }
  for (const x of cells.values()) if (x.f && typeof x.f === 'object') { const ms = masters[x.f.si]; x.f = ms ? shiftFormula(ms.f, x.r - ms.r, x.c - ms.c) : null; }
  return cells;
}

/* ---------- 월 머리글 묶음 찾기 ---------- */
const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
function monthOf(x) {
  if (!x) return 0;
  const v = x.v;
  if (typeof v === 'number') {
    if (v >= 1 && v <= 12 && Number.isInteger(v)) return v;
    if (v > 40000 && v < 60000) return new Date(Math.round((v - 25569) * 864e5)).getUTCMonth() + 1;
    return 0;
  }
  const s = String(v ?? '').trim();
  let mm = /^(\d{1,2})\s*월$/.exec(s); if (mm && +mm[1] >= 1 && +mm[1] <= 12) return +mm[1];
  mm = /^([A-Za-z]{3})/.exec(s); if (mm && MON[mm[1].toLowerCase()] && s.length <= 4) return MON[mm[1].toLowerCase()];
  return 0;
}
// 같은 행에서 왼쪽→오른쪽으로 월이 1씩 늘어나는 열 3개 이상 = 묶음
export function monthBlocks(cells) {
  const byRow = new Map();
  for (const x of cells.values()) { if (!byRow.has(x.r)) byRow.set(x.r, []); byRow.get(x.r).push(x); }
  const blocks = [];
  for (const [r, list] of byRow) {
    const ms = list.map((x) => ({ c: x.c, m: monthOf(x), str: typeof x.v === 'string' })).filter((x) => x.m).sort((a, b) => a.c - b.c);
    // 숫자 1~12만 있는 행은 머리글이 아닐 수 있어 위쪽(20행 이내) 또는 문자 월이 섞인 행만
    if (!ms.length || (r > 20 && !ms.some((x) => x.str) && !ms.some((x) => list.find((y) => y.c === x.c)?.v > 40000))) continue;
    let run = [ms[0]];
    const flush = () => { if (run.length >= 3) blocks.push({ row: r, cols: Object.fromEntries(run.map((x) => [x.m, x.c])), first: run[0].c, last: run[run.length - 1].c }); };
    for (let i = 1; i < ms.length; i++) {
      const p = run[run.length - 1];
      if (ms[i].c === p.c + 1 && ms[i].m === p.m + 1) run.push(ms[i]); else { flush(); run = [ms[i]]; }
    }
    flush();
  }
  // 같은 열 범위의 묶음이 여러 행에 있으면(연도 행·월 행 등) 아래쪽 머리글 기준으로 데이터 범위를 나눔
  blocks.sort((a, b) => a.row - b.row);
  const maxR = Math.max(0, ...[...cells.values()].map((x) => x.r));
  for (const b of blocks) {
    const next = blocks.find((o) => o.row > b.row + 1 && o.first <= b.last && o.last >= b.first);
    b.to = next ? next.row - 1 : maxR;
  }
  // 바로 아래 행에 같은 묶음(예: 4행 연도/5행 월)이 붙어 있으면 위 것은 버림
  return blocks.filter((b) => !blocks.some((o) => o.row === b.row + 1 && o.first === b.first && o.last === b.last));
}

/* ---------- 점검 ---------- */
// month: 당월. 결과 [{sheet, block, issues:[{kind, addr, prev, cur}] , stats}]
export function checkSheet(name, cells, month) {
  const out = [];
  for (const b of monthBlocks(cells)) {
    const cm = b.cols[month], cp = b.cols[month - 1];
    const head = `${colName(b.first)}${b.row}:${colName(b.last)}${b.row}`;
    if (!cp) continue;
    if (!cm) { out.push({ sheet: name, head, missing: true, issues: [], stats: { rows: 0 } }); continue; }
    const issues = [];
    let rows = 0, same = 0;
    for (let r = b.row + 1; r <= b.to; r++) {
      const P = cells.get(r + ',' + cp), C = cells.get(r + ',' + cm);
      if (!P && !C) continue;
      rows++;
      const addr = `${colName(cm)}${r}`;
      const pf = P?.f, cf = C?.f;
      if (pf && !cf) issues.push({ kind: C && C.v !== undefined && C.v !== '' ? 'value' : 'blank', addr, prev: `=${pf}`, cur: C?.v ?? '' });
      else if (pf && cf) {
        if (relForm(pf, r, cp) !== relForm(cf, r, cm)) issues.push({ kind: 'formula', addr, prev: `=${pf}`, cur: `=${cf}`, want: `=${shiftFormula(pf, 0, cm - cp)}` });
        else same++;
      } else if (!pf && cf && P && P.v !== undefined && P.v !== '') issues.push({ kind: 'newformula', addr, prev: P.v, cur: `=${cf}` });
      if (P && C && P.s !== C.s && (pf || cf || P.v !== undefined)) issues.push({ kind: 'style', addr, prev: P.s, cur: C.s });
    }
    out.push({ sheet: name, head, issues, stats: { rows, same } });
  }
  return out;
}

export async function checkBook(book, month, onProgress = () => {}) {
  const res = [];
  const list = book.sheets.filter((s) => !RAW_SHEET.test(s.name.trim()));
  for (let i = 0; i < list.length; i++) {
    const s = list[i];
    onProgress(`${i + 1}/${list.length} ${s.name}`);
    const f = book.zip.file(s.path);
    if (!f || f._data?.uncompressedSize > 30e6) continue; // 너무 큰 시트는 건너뜀
    const cells = parseSheet(await f.async('string'), book.sst);
    res.push(...checkSheet(s.name, cells, month));
  }
  return res;
}

/* ---------- 열 끼워 넣기 (표가 당월에서 끝날 때 다음 달 열을 만들기 위해) ---------- */
// 수식 안의 참조를 시트 이름과 함께 찾아서 fn(sheet|null, ref문자열) 으로 바꿈 (문자열 "…" 안은 그대로)
const QREF = /((?:'(?:[^']|'')+'|[A-Za-z0-9_.ㄱ-힣]+)!)?(\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})(?![\d(A-Za-z_])/g;
function mapQRefs(f, fn) {
  let out = '', i = 0;
  while (i < f.length) {
    if (f[i] === '"') { const j = f.indexOf('"', i + 1); const e = j < 0 ? f.length : j + 1; out += f.slice(i, e); i = e; continue; }
    let j = i; while (j < f.length && f[j] !== '"') j++;
    const seg = f.slice(i, j);
    out += seg.replace(QREF, (m, pre, ref, off) => {
      const before = seg[off - 1];
      if (!pre && before && /[A-Za-z0-9_.]/.test(before)) return m;
      const sh = pre ? pre.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'") : null;
      return (pre || '') + fn(sh, ref);
    });
    i = j;
  }
  return out;
}
// 한 참조(A1, A1:B2, A:B)의 열 번호를 fc(col)로 바꿈
const mapRefCols = (ref, fc) => ref.replace(/(\$?)([A-Z]{1,3})(\$?\d+)?/g, (m, a, L, rest) => `${a}${colName(fc(colNum(L)))}${rest || ''}`);
const shiftRange = (rg, at) => rg.split(/\s+/).map((x) => mapRefCols(x, (c) => (c >= at ? c + 1 : c))).join(' ');

// 시트 at열에 빈 열을 끼우고, 모든 시트의 수식·정의된 이름에서 그 시트 at열 이후 참조를 한 칸 밀기
export async function insertColumn(book, sheetName, at) {
  const own = book.sheets.find((s) => s.name === sheetName);
  const adj = (f, self) => mapQRefs(f, (sh, ref) => ((sh ? sh === sheetName : self) ? mapRefCols(ref, (c) => (c >= at ? c + 1 : c)) : ref));
  for (const s of book.sheets) {
    const file = book.zip.file(s.path);
    if (!file) continue;
    const self = s === own;
    const size = file._data?.uncompressedSize || 0;
    if (!self && size > 30e6) continue; // 원본 데이터 큰 시트는 수식이 거의 없음
    let xml = await file.async('string');
    if (!self && !xml.includes(enxml(sheetName.replace(/'/g, "''")) + (/[^A-Za-z0-9_.ㄱ-힣]/.test(sheetName) ? "'!" : '!'))) continue;
    // 공유 수식은 셀마다 풀어서 씀(밀고 나면 기준 위치가 달라지므로)
    if (/t="shared"/.test(xml)) {
      const cells = parseSheet(xml, book.sst);
      // 빈 셀(<c …/>)을 먼저 걸러야 다음 셀까지 삼키지 않음
      xml = xml.replace(/<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (m, L, R, a, inner) => {
        if (inner === undefined || !/<f[^>]*t="shared"/.test(inner)) return m;
        const x = cells.get(R + ',' + colNum(L));
        return x?.f ? `<c r="${L}${R}"${a}>${inner.replace(/<f[^>]*\/>|<f[^>]*>[^<]*<\/f>/, `<f>${enxml(x.f)}</f>`)}</c>` : m;
      });
    }
    xml = xml.replace(/<f>([^<]*)<\/f>/g, (m, f) => `<f>${enxml(adj(unxml(f), self))}</f>`).replace(/<f ([^>]*)>([^<]+)<\/f>/g, (m, a, f) => `<f ${a}>${enxml(adj(unxml(f), self))}</f>`);
    if (self) {
      xml = xml.replace(/<c r="([A-Z]+)(\d+)"/g, (m, L, R) => { const c = colNum(L); return c >= at ? `<c r="${colName(c + 1)}${R}"` : m; });
      xml = xml.replace(/(<(?:mergeCell|hyperlink|autoFilter|dimension|f) [^>]*?ref=")([^"]+)"/g, (m, p, rg) => `${p}${shiftRange(rg, at)}"`)
        .replace(/(topLeftCell=")([^"]+)"/g, (m, p, rg) => `${p}${shiftRange(rg, at)}"`)
        .replace(/(sqref=")([^"]+)"/g, (m, p, rg) => `${p}${shiftRange(rg, at)}"`)
        .replace(/(<selection [^>]*?activeCell=")([^"]+)"/g, (m, p, rg) => `${p}${shiftRange(rg, at)}"`)
        // 열 너비 정의: 마지막 열(XFD=16384)을 넘지 않게
        .replace(/<col ([^>]*?)min="(\d+)"([^>]*?)max="(\d+)"/g, (m, a, mn, b, mx) => { mn = +mn; mx = +mx; if (mn >= at) { mn++; mx++; } else if (mx >= at) mx++; return `<col ${a}min="${Math.min(mn, 16384)}"${b}max="${Math.min(mx, 16384)}"`; })
        .replace(/<row ([^>]*?)spans="(\d+):(\d+)"/g, (m, a, s1, s2) => `<row ${a}spans="${s1}:${+s2 >= at ? +s2 + 1 : s2}"`);
    }
    book.zip.file(s.path, xml);
  }
  // 정의된 이름(인쇄 영역 등)
  book.wbXml = book.wbXml.replace(/(<definedName [^>]*>)([^<]+)(<\/definedName>)/g, (m, a, f, b) => a + enxml(adj(unxml(f), false)) + b);
  book.zip.file('xl/workbook.xml', book.wbXml);
}

/* ---------- 다음 달 양식 만들기 ---------- */
// 당월 열(month)의 수식·서식을 다음 달 열로 복사. 다음 달 열에 이미 수식이 있으면 그대로 둠.
function setCellInRow(rowXml, col, row, cellXml) {
  // rowXml = '<row ...>...</row>'. 열 순서에 맞게 끼우거나 바꿈
  const ref = colName(col) + row;
  const re = new RegExp(`<c r="${ref}"(?:[^>]*?\\/>|[^>]*?>[\\s\\S]*?<\\/c>)`);
  if (re.test(rowXml)) return rowXml.replace(re, cellXml);
  const open = rowXml.match(/^<row [^>]*?>/)[0];
  if (/\/>$/.test(open)) return open.replace(/\/>$/, '>') + cellXml + '</row>';
  const body = rowXml.slice(open.length, -'</row>'.length);
  const cs = [...body.matchAll(/<c r="([A-Z]+)\d+"(?:[^>]*?\/>|[^>]*?>[\s\S]*?<\/c>)/g)];
  const after = cs.find((x) => colNum(x[1]) > col);
  const pos = after ? after.index : body.length;
  return open + body.slice(0, pos) + cellXml + body.slice(pos) + '</row>';
}
// 머리글 칸 다음 달 값 만들기
const ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function nextHeaderXml(x, r, c, month) {
  const s = x.s ? ` s="${x.s}"` : '', ref = `${colName(c)}${r}`;
  if (x.f) {
    // 수식 머리글(예: =EDATE(AW234,1))도 계산값을 같이 넣어야 다음 단계에서 월로 알아봄
    const inner = nextHeaderXml({ ...x, f: null }, r, c, month).match(/^<c [^>]*?(?: t="(\w+)")?>([\s\S]*)<\/c>$/);
    const v = inner && /^<v>/.test(inner[2]) ? inner[2] : inner ? `<v>${enxml(unxml((inner[2].match(/<t>([^<]*)<\/t>/) || [])[1] || ''))}</v>` : '';
    const t = inner && inner[1] === 'inlineStr' ? ' t="str"' : '';
    return `<c r="${ref}"${s}${t}><f>${enxml(shiftFormula(x.f, 0, 1))}</f>${v}</c>`;
  }
  const v = x.v;
  if (typeof v === 'number') {
    if (v > 40000) { const d = new Date(Math.round((v - 25569) * 864e5)); const n = Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); return `<c r="${ref}"${s}><v>${Math.round(n / 864e5 + 25569)}</v></c>`; }
    return `<c r="${ref}"${s}><v>${month + 1}</v></c>`;
  }
  const t = String(v).replace(/^(\d{1,2})(\s*월)$/, (m, a, b) => `${+a + 1}${b}`).replace(/^([A-Za-z]{3})$/, (m, a) => ABBR[ABBR.findIndex((z) => z.toLowerCase() === a.toLowerCase()) + 1] || a);
  return `<c r="${ref}"${s} t="inlineStr"><is><t>${enxml(t)}</t></is></c>`;
}
async function writeCells(book, s, list) {
  let xml = await book.zip.file(s.path).async('string');
  const byRow = new Map();
  for (const x of list) { if (!byRow.has(x.r)) byRow.set(x.r, []); byRow.get(x.r).push(x); }
  xml = xml.replace(/<row [^>]*?r="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, (rowXml, rn) => {
    const e = byRow.get(+rn);
    if (!e) return rowXml;
    for (const x of e) rowXml = setCellInRow(rowXml, x.c, +rn, x.xml);
    return rowXml;
  });
  book.zip.file(s.path, xml);
}
// 끼워 넣은 표의 '당월' 열(BGT 당월·Target·% Ach.·누계 등) 수식: 당월(구) 열을 가리키던 참조 → 새 달 열로
async function movePointers(book, ins, blockMap) {
  const s = book.sheets.find((x) => x.name === ins.sheet);
  const cells = parseSheet(await book.zip.file(s.path).async('string'), book.sst);
  const edits = [];
  // 대상 열: 이 표의 새 달 열 오른쪽 ~ 같은 머리글 줄에서 오른쪽 다음 월 표 시작 전 (다른 월 표의 데이터는 건드리지 않음)
  const own = (blockMap.get(ins.sheet) || []);
  const nextFirst = Math.min(...own.filter((o) => o.row === ins.row && o.first > ins.at).map((o) => o.first), Infinity);
  const inAnyBlock = (r, c) => own.some((o) => r > o.row && r <= o.to && c >= o.first && c <= o.last);
  for (const x of cells.values()) {
    if (!x.f || x.r <= ins.row || x.r > ins.to || x.c <= ins.at || x.c >= nextFirst || inAnyBlock(x.r, x.c)) continue;
    let changed = false;
    const nf = mapQRefs(x.f, (sh, ref) => {
      const parts = ref.split(':');
      if (!sh) {
        // 같은 시트: 구 당월 열(at-1) → at, 1월~구 당월 범위(누계)는 끝을 at으로
        // 1월부터 시작하는 범위(누계·평균)는 끝을 새 달 열까지
        if (parts.length === 2) { const a = colNum(parts[0].replace(/[$\d]/g, '')), b = colNum(parts[1].replace(/[$\d]/g, '')); if (a === ins.first && b >= ins.first && b < ins.at) { changed = true; return `${parts[0]}:${mapRefCols(parts[1], () => ins.at)}`; } return ref; }
        const c = colNum(ref.replace(/[$\d]/g, ''));
        if (c === ins.at - 1) { changed = true; return mapRefCols(ref, () => ins.at); }
        return ref;
      }
      // 다른 시트의 월 표에서 당월 열을 가리키면 다음 달 열로
      if (parts.length !== 1) return ref;
      const c = colNum(ref.replace(/[$\d]/g, '')), r = +ref.replace(/[^\d]/g, '');
      const b = (blockMap.get(sh) || []).find((o) => r > o.row && r <= o.to && o.cols[ins.month] === c && o.cols[ins.month + 1]);
      if (b) { changed = true; return mapRefCols(ref, () => b.cols[ins.month + 1]); }
      return ref;
    });
    if (changed) edits.push({ r: x.r, c: x.c, xml: `<c r="${colName(x.c)}${x.r}"${x.s ? ` s="${x.s}"` : ''}><f>${enxml(nf)}</f></c>` });
  }
  if (edits.length) await writeCells(book, s, edits);
  return edits.length;
}
async function blockMapOf(book) {
  const map = new Map();
  for (const s of book.sheets) {
    if (RAW_SHEET.test(s.name.trim())) continue;
    const f = book.zip.file(s.path);
    if (!f || f._data?.uncompressedSize > 30e6) continue;
    map.set(s.name, monthBlocks(parseSheet(await f.async('string'), book.sst)));
  }
  return map;
}

export async function makeNextMonth(book, month, onProgress = () => {}) {
  const next = month + 1, report = { filled: [], manual: [], noCol: [], months: [], inserted: [], pointers: [] };
  const list = book.sheets.filter((s) => !RAW_SHEET.test(s.name.trim()));
  // ① 당월에서 끝나는 표에는 다음 달 열을 끼워 넣음 (오른쪽 표부터)
  const ins = [];
  if (next <= 12) {
    for (const s of list) {
      const f = book.zip.file(s.path);
      if (!f || f._data?.uncompressedSize > 30e6) continue;
      const cells = parseSheet(await f.async('string'), book.sst);
      const bl = monthBlocks(cells).filter((b) => b.cols[month] && !b.cols[next] && b.last === b.cols[month]).sort((a, b) => b.first - a.first || a.row - b.row);
      const done = new Set();
      for (const b of bl) {
        onProgress(`열 끼워 넣는 중: ${s.name}`);
        const at = b.last + 1;
        // 사용자 방식(10월 파일로 확인): 월 바로 오른쪽에 새 열을 끼워 넣음(빈 예비 열이 있어도 그대로 둠)
        // 같은 열 위치의 표가 여러 개(위아래)면 열은 한 번만 끼움
        const fresh = !done.has(b.last);
        if (fresh) {
          await insertColumn(book, s.name, at);
          for (const o of ins) if (o.sheet === s.name && o.at >= at) { o.at++; o.first++; } // 같은 시트 오른쪽 표는 한 칸 밀림
          done.add(b.last);
        }
        // 머리글(그 위 행 서식 포함)
        const cur = parseSheet(await book.zip.file(s.path).async('string'), book.sst);
        const hx = cur.get(b.row + ',' + (at - 1));
        const w = [];
        if (hx) w.push({ r: b.row, c: at, xml: nextHeaderXml(hx, b.row, at, month) });
        for (const rr of [b.row - 1, b.row - 2]) { const u = cur.get(rr + ',' + (at - 1)); if (u && u.s && !cur.get(rr + ',' + at)) w.push({ r: rr, c: at, xml: `<c r="${colName(at)}${rr}" s="${u.s}"/>` }); }
        // 같은 머리글 줄 오른쪽의 '당월' 표시(BGT Sep·Target Sep 등)도 다음 달로
        const inBlock = monthBlocks(cur).filter((o) => o.row === b.row);
        for (const x of cur.values()) if (x.r === b.row && x.c > at && !x.f && monthOf(x) === month && !inBlock.some((o) => x.c >= o.first && x.c <= o.last)) w.push({ r: x.r, c: x.c, xml: nextHeaderXml(x, x.r, x.c, month) });
        if (w.length) await writeCells(book, s, w);
        ins.push({ sheet: s.name, at, row: b.row, to: b.to, first: b.first, month });
        if (fresh) report.inserted.push(`${s.name} ${colName(at)}열 (${next}월)`);
      }
    }
  }
  for (let i = 0; i < list.length; i++) {
    const s = list[i];
    onProgress(`${i + 1}/${list.length} ${s.name}`);
    const file = book.zip.file(s.path);
    if (!file || file._data?.uncompressedSize > 30e6) continue;
    let xml = await file.async('string');
    const cells = parseSheet(xml, book.sst);
    const edits = new Map(); // row → [{col, cellXml}]
    let nFill = 0;
    for (const b of monthBlocks(cells)) {
      const cm = b.cols[month];
      if (!cm) continue;
      const cn = b.cols[next];
      if (!cn) { if (next <= 12) report.noCol.push(`${s.name} ${colName(b.first)}${b.row}~${colName(b.last)}${b.row} (${Math.max(...Object.keys(b.cols).map(Number))}월까지)`); continue; }
      for (let r = b.row + 1; r <= b.to; r++) {
        const C = cells.get(r + ',' + cm), N = cells.get(r + ',' + cn);
        if (!C) continue;
        if (C.f) {
          if (N?.f) continue; // 이미 수식 있음
          const nf = shiftFormula(C.f, 0, cn - cm);
          if (!edits.has(r)) edits.set(r, []);
          edits.get(r).push({ col: cn, xml: `<c r="${colName(cn)}${r}"${C.s ? ` s="${C.s}"` : ''}><f>${enxml(nf)}</f></c>` });
          nFill++;
        } else if (C.v !== undefined && C.v !== '' && typeof C.v === 'number' && !(N && N.v !== undefined)) {
          // 당월이 손으로 넣은 숫자인 칸 → 다음 달도 직접 입력 필요
          report.manual.push(`${s.name}!${colName(cn)}${r}`);
          if (C.s && (!N || N.s !== C.s)) { if (!edits.has(r)) edits.set(r, []); edits.get(r).push({ col: cn, xml: `<c r="${colName(cn)}${r}" s="${C.s}"/>` }); }
        }
      }
    }
    // 결산월 칸(B3 등): 값이 당월 숫자인 B3 → 다음 달
    const b3 = cells.get('3,2');
    if (b3 && !b3.f && b3.v === month) { if (!edits.has(3)) edits.set(3, []); edits.get(3).push({ col: 2, xml: `<c r="B3"${b3.s ? ` s="${b3.s}"` : ''}><v>${next}</v></c>` }); report.months.push(`${s.name}!B3 ${month} → ${next}`); }
    if (!edits.size) continue;
    xml = xml.replace(/<row [^>]*?r="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, (rowXml, rn) => {
      const e = edits.get(+rn);
      if (!e) return rowXml;
      edits.delete(+rn);
      for (const x of e.sort((a, b) => a.col - b.col)) rowXml = setCellInRow(rowXml, x.col, +rn, x.xml);
      return rowXml;
    });
    book.zip.file(s.path, xml);
    if (nFill) report.filled.push(`${s.name}: ${nFill}칸`);
  }
  // ③ 끼워 넣은 표의 당월 열(BGT·Target·% Ach.·누계) 참조를 다음 달로
  if (ins.length) {
    onProgress('당월 참조 옮기는 중…');
    const bm = await blockMapOf(book);
    for (const x of ins) { const n = await movePointers(book, x, bm); if (n) report.pointers.push(`${x.sheet}: ${n}칸`); }
  }
  // 엑셀이 열 때 전부 다시 계산하도록, 계산 순서 파일은 지움(엑셀이 새로 만듦)
  let wb = book.wbXml;
  wb = /<calcPr[^>]*\/>/.test(wb) ? wb.replace(/<calcPr([^>]*)\/>/, (m, a) => `<calcPr${a.replace(/\s*fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`) : wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  book.zip.file('xl/workbook.xml', wb);
  if (book.zip.file('xl/calcChain.xml')) {
    book.zip.remove('xl/calcChain.xml');
    const ct = await book.zip.file('[Content_Types].xml').async('string');
    book.zip.file('[Content_Types].xml', ct.replace(/<Override [^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, ''));
    const rp = 'xl/_rels/workbook.xml.rels';
    const rl = await book.zip.file(rp).async('string');
    book.zip.file(rp, rl.replace(/<Relationship [^>]*Target="[^"]*calcChain\.xml"[^>]*\/>/, ''));
  }
  onProgress('파일 만드는 중…');
  const blob = await book.zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, report };
}
