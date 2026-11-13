// 엑셀(xlsx) XML 직접 편집 도구 — 서식·글꼴을 그대로 두고 필요한 셀·시트만 고침
// 브라우저와 Node 모두에서 사용 (JSZip은 밖에서 넘겨받음)

export const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
export const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
export const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
export const colNum = (letters) => letters.replace(/\$/g, '').split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export const colLet = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
export const num = (v) => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
};
export const toSerial = (d) => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 864e5;
export const fromSerial = (n) => new Date(Math.round((n - 25569) * 864e5));
const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/* ---------- 셀 XML ---------- */
export const textCell = (ref, s, text) => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${esc(text)}</t></is></c>`;
export const numCell = (ref, s, v) => `<c r="${ref}"${s ? ` s="${s}"` : ''}><v>${v}</v></c>`;
export const fCell = (ref, s, f) => `<c r="${ref}"${s ? ` s="${s}"` : ''}><f>${esc(f)}</f></c>`;
export const blankCell = (ref, s) => `<c r="${ref}"${s ? ` s="${s}"` : ''}/>`;
export const valCell = (ref, s, v) => (v === null || v === undefined || v === '' ? blankCell(ref, s)
  : typeof v === 'number' ? numCell(ref, s, v) : v instanceof Date ? numCell(ref, s, toSerial(v)) : textCell(ref, s, String(v)));

export function cellInfo(xml) {
  if (!xml) return null;
  const m = xml.match(/^<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)$/);
  if (!m) return null;
  const attrs = m[1], inner = m[2] || '';
  const fm = inner.match(/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/);
  return {
    s: (attrs.match(/\ss="(\d+)"/) || [])[1],
    t: (attrs.match(/\st="([^"]+)"/) || [])[1],
    v: (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1],
    f: fm ? unesc(fm[2] || '') : undefined,
    fAttrs: fm ? fm[1] : '',
    isText: (inner.match(/<is>[\s\S]*?<\/is>/) || [])[0],
  };
}
export const styleOfXml = (xml) => (xml ? (xml.match(/\ss="(\d+)"/) || [])[1] : undefined);

/* ---------- 시트 (행 단위) ---------- */
export class Sheet {
  constructor(xml, name) {
    this.name = name;
    this.xml = xml;
    this.rows = new Map();
    const sd = xml.match(/<sheetData>([\s\S]*?)<\/sheetData>|<sheetData\/>/);
    if (sd && sd[1]) {
      for (const m of sd[1].matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
        const rn = +m[1].match(/\sr="(\d+)"/)[1];
        const cells = new Map();
        for (const c of (m[2] || '').matchAll(/<c r="([A-Z]+)(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) cells.set(colNum(c[1]), c[0]);
        this.rows.set(rn, { attrs: m[1].replace(/\sspans="[^"]*"/, ''), cells });
      }
    }
  }
  get maxRow() { return this.rows.size ? Math.max(...this.rows.keys()) : 0; }
  get maxCol() { let m = 0; for (const r of this.rows.values()) for (const c of r.cells.keys()) if (c > m) m = c; return m; }
  xmlAt(col, row) { return this.rows.get(row)?.cells.get(typeof col === 'string' ? colNum(col) : col); }
  info(col, row) { return cellInfo(this.xmlAt(col, row)); }
  style(col, row) { return styleOfXml(this.xmlAt(col, row)); }
  set(col, row, xmlFn) {
    const c = typeof col === 'string' ? colNum(col) : col;
    if (!this.rows.has(row)) this.rows.set(row, { attrs: ` r="${row}"`, cells: new Map() });
    const ref = `${colLet(c)}${row}`;
    this.rows.get(row).cells.set(c, typeof xmlFn === 'function' ? xmlFn(ref) : xmlFn);
  }
  del(col, row) { this.rows.get(row)?.cells.delete(typeof col === 'string' ? colNum(col) : col); }
  setVal(col, row, v, s) { this.set(col, row, (ref) => valCell(ref, s ?? this.style(col, row), v)); }
  setF(col, row, f, s) { this.set(col, row, (ref) => fCell(ref, s ?? this.style(col, row), f)); }
  build() {
    const out = [];
    for (const rn of [...this.rows.keys()].sort((a, b) => a - b)) {
      const r = this.rows.get(rn);
      const cells = [...r.cells.keys()].sort((a, b) => a - b).map((k) => r.cells.get(k)).join('');
      out.push(cells ? `<row${r.attrs}>${cells}</row>` : `<row${r.attrs}/>`);
    }
    let xml = this.xml.replace(/<sheetData>[\s\S]*?<\/sheetData>|<sheetData\/>/, `<sheetData>${out.join('')}</sheetData>`);
    const mr = this.maxRow, mc = this.maxCol;
    if (mr && mc) xml = xml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${colLet(mc)}${mr}"/>`);
    return xml;
  }
  // 공유 수식을 각 셀의 일반 수식으로 풀기 (셀 복사·이동 전에 필요)
  expandShared() {
    const masters = new Map();
    for (const [rn, r] of this.rows) for (const [c, x] of r.cells) {
      const m = x.match(/<f\b[^>]*t="shared"[^>]*ref="[^"]*"[^>]*si="(\d+)"[^>]*>([\s\S]*?)<\/f>|<f\b[^>]*t="shared"[^>]*si="(\d+)"[^>]*ref="[^"]*"[^>]*>([\s\S]*?)<\/f>/);
      if (m) masters.set(m[1] ?? m[3], { r: rn, c, f: unesc(m[2] ?? m[4]) });
    }
    if (!masters.size) return;
    for (const [rn, r] of this.rows) for (const [c, x] of r.cells) {
      const m = x.match(/<f\b([^>]*?)t="shared"([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/);
      if (!m) continue;
      const si = (m[0].match(/si="(\d+)"/) || [])[1];
      const ms = masters.get(si);
      if (!ms) continue;
      const f = rn === ms.r && c === ms.c ? ms.f : shiftRefs(ms.f, (ref) => (ref.sheet ? ref : {
        ...ref, col: ref.col && !ref.colAbs ? ref.col + (c - ms.c) : ref.col, row: ref.row && !ref.rowAbs ? ref.row + (rn - ms.r) : ref.row,
      }));
      r.cells.set(c, x.replace(m[0], `<f>${esc(f)}</f>`));
    }
  }
}

/* ---------- 수식 참조 이동 ---------- */
// fn(ref) → ref. ref = { sheet, col, row, colAbs, rowAbs } (열 전체 참조는 row=null, 행 전체는 col=null)
const SHEET = "(?:'(?:[^']|'')+'|[A-Za-z0-9_\\.\\u3131-\\uD7A3]+)!";
const REF_RE = new RegExp(`(${SHEET})?(?:(\\$?)([A-Z]{1,3})(\\$?)(\\d+)(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d+))?|(\\$?)([A-Z]{1,3}):(\\$?)([A-Z]{1,3})(?![\\d(])|(\\$?)(\\d+):(\\$?)(\\d+))`, 'g');
export function shiftRefs(formula, fn) {
  // 문자열 상수 "..."는 건드리지 않음
  return String(formula).split(/("(?:[^"]|"")*")/).map((part, i) => {
    if (i % 2 === 1) return part;
    return part.replace(REF_RE, (m, sh, a1, c1, a2, r1, a3, c2, a4, r2, b1, cc1, b2, cc2, d1, rr1, d2, rr2, off, str) => {
      const prev = str[off - 1];
      if (prev && /[A-Za-z0-9_.]/.test(prev) && !sh) return m; // 함수 이름 등의 일부
      const after = str[off + m.length];
      if (after === '(') return m; // 함수 호출(LOG10( 등)
      const sheet = sh ? sh.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'") : null;
      const mk = (o) => `${o.colAbs ? '$' : ''}${o.col ? colLet(o.col) : ''}${o.rowAbs ? '$' : ''}${o.row ?? ''}`;
      if (c1) {
        const p = fn({ sheet, col: colNum(c1), row: +r1, colAbs: !!a1, rowAbs: !!a2, part: 1, range: !!c2 });
        if (!c2) return (sh || '') + mk(p);
        const q = fn({ sheet, col: colNum(c2), row: +r2, colAbs: !!a3, rowAbs: !!a4, part: 2, range: true });
        return (sh || '') + mk(p) + ':' + mk(q);
      }
      if (cc1) {
        const p = fn({ sheet, col: colNum(cc1), row: null, colAbs: !!b1, rowAbs: false, part: 1, range: true });
        const q = fn({ sheet, col: colNum(cc2), row: null, colAbs: !!b2, rowAbs: false, part: 2, range: true });
        return (sh || '') + `${b1}${colLet(p.col)}:${b2}${colLet(q.col)}`;
      }
      const p = fn({ sheet, col: null, row: +rr1, colAbs: false, rowAbs: !!d1, part: 1, range: true });
      const q = fn({ sheet, col: null, row: +rr2, colAbs: false, rowAbs: !!d2, part: 2, range: true });
      return (sh || '') + `${d1}${p.row}:${d2}${q.row}`;
    });
  }).join('');
}
// 셀 XML 안의 수식에 shiftRefs 적용
export function mapFormulaXml(xml, fn) {
  return xml.replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/, (m, a, body, b) => a + esc(shiftRefs(unesc(body), fn)) + b);
}
const sqrefShift = (sq, fn) => sq.split(' ').map((r) => shiftRefs(r, fn)).join(' ');

/* ---------- 통합문서 ---------- */
export class Book {
  static async open(JSZip, buf) {
    const b = new Book();
    b.zip = await JSZip.loadAsync(buf);
    b.wbXml = await b.zip.file('xl/workbook.xml').async('string');
    b.relXml = await b.zip.file('xl/_rels/workbook.xml.rels').async('string');
    b.ct = await b.zip.file('[Content_Types].xml').async('string');
    const rels = {};
    for (const m of b.relXml.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = m[0].match(/Id="([^"]+)"/)[1];
      const target = m[0].match(/Target="([^"]+)"/)[1];
      rels[id] = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
    }
    b.sheets = [...b.wbXml.matchAll(/<sheet\b[^>]*\/>/g)].map((m) => ({
      name: unesc(m[0].match(/name="([^"]+)"/)[1]),
      sheetId: +m[0].match(/sheetId="(\d+)"/)[1],
      path: rels[m[0].match(/r:id="([^"]+)"/)[1]],
      hidden: /state="(hidden|veryHidden)"/.test(m[0]),
    }));
    const sf = b.zip.file('xl/sharedStrings.xml');
    b.sst = sf ? [...(await sf.async('string')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => unesc([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join(''))) : [];
    b.cache = new Map();
    return b;
  }
  names() { return [...this.wbXml.matchAll(/<sheet\b[^>]*name="([^"]+)"/g)].map((m) => unesc(m[1])); }
  has(name) { return this.sheets.some((s) => s.name === name); }
  async sheet(name) {
    if (this.cache.has(name)) return this.cache.get(name);
    const s = this.sheets.find((x) => x.name === name);
    if (!s) return null;
    const sh = new Sheet(await this.zip.file(s.path).async('string'), name);
    this.cache.set(name, sh);
    return sh;
  }
  text(info) {
    if (!info) return '';
    if (info.t === 's') return this.sst[+info.v] ?? '';
    if (info.isText) return unesc([...info.isText.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join(''));
    return info.v !== undefined ? unesc(info.v) : '';
  }
  // 값 읽기: 숫자면 숫자, 아니면 글자 (수식은 마지막 계산값)
  val(sh, col, row) {
    const i = sh.info(col, row);
    if (!i) return null;
    if (i.t === 's' || i.t === 'inlineStr' || i.t === 'str') return this.text(i);
    if (i.v === undefined) return null;
    const n = Number(i.v);
    return isNaN(n) ? unesc(i.v) : n;
  }
  // 모든 시트 수식(+정의된 이름)에서 targetSheet 참조를 fn으로 이동. 고친 시트는 캐시에 올라감
  async shiftRefsEverywhere(targetSheet, fn, { skipSelf = false } = {}) {
    const re = new RegExp(reEsc(targetSheet).replace(/'/g, "''?") + "'?!");
    for (const s of this.sheets) {
      if (skipSelf && s.name === targetSheet) continue;
      const cached = this.cache.get(s.name);
      const raw = cached ? null : await this.zip.file(s.path).async('string');
      if (!cached && !re.test(raw)) continue;
      const sh = cached || await this.sheet(s.name);
      for (const r of sh.rows.values()) for (const [c, x] of r.cells) {
        if (!x.includes('<f') || !re.test(unesc(x))) continue;
        r.cells.set(c, mapFormulaXml(x, (ref) => (ref.sheet === targetSheet ? fn(ref) : ref)));
      }
    }
    this.wbXml = this.wbXml.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (m, a, body, b) =>
      a + esc(shiftRefs(unesc(body), (ref) => (ref.sheet === targetSheet ? fn(ref) : ref))) + b);
  }
  // name 시트의 at 열부터 n칸 오른쪽으로 밀기 (엑셀 '열 삽입'과 같음)
  async insertColumns(name, at, n = 1) {
    const sh = await this.sheet(name);
    sh.expandShared();
    const mv = (ref) => {
      if (!ref.col) return ref;
      if (ref.range && ref.part === 1 && ref.col >= at) return { ...ref, col: ref.col + n };
      if (ref.col >= at) return { ...ref, col: ref.col + n };
      return ref;
    };
    for (const r of sh.rows.values()) {
      const moved = new Map();
      for (const [c, x] of r.cells) {
        let nx = x.includes('<f') ? mapFormulaXml(x, (ref) => (ref.sheet && ref.sheet !== name ? ref : mv(ref))) : x;
        const nc = c >= at ? c + n : c;
        if (nc !== c) nx = nx.replace(/^<c r="[A-Z]+(\d+)"/, `<c r="${colLet(nc)}$1"`);
        nx = nx.replace(/(<f\b[^>]*\sref=")([^"]+)"/, (m, a, r0) => `${a}${shiftRefs(r0, mv)}"`);
        moved.set(nc, nx);
      }
      r.cells = moved;
    }
    await this.shiftRefsEverywhere(name, mv, { skipSelf: true });
    let xml = sh.xml;
    xml = xml.replace(/<cols>([\s\S]*?)<\/cols>/, (m, body) => '<cols>' + body.replace(/<col\b([^>]*?)min="(\d+)"([^>]*?)max="(\d+)"/g, (mm, a, mi, b, ma) => {
      let lo = +mi, hi = +ma;
      if (lo >= at) { lo += n; hi += n; } else if (hi >= at) hi += n;
      return `<col${a}min="${lo}"${b}max="${hi}"`;
    }) + '</cols>');
    xml = xml.replace(/<mergeCell ref="([^"]+)"/g, (m, r) => `<mergeCell ref="${shiftRefs(r, mv)}"`);
    xml = xml.replace(/sqref="([^"]+)"/g, (m, r) => `sqref="${sqrefShift(r, mv)}"`);
    xml = xml.replace(/(<autoFilter ref=")([^"]+)"/, (m, a, r) => `${a}${shiftRefs(r, mv)}"`);
    sh.xml = xml;
  }
  // name 시트의 at 행부터 n줄 아래로 밀기 (엑셀 '행 삽입')
  async insertRows(name, at, n = 1) {
    const sh = await this.sheet(name);
    sh.expandShared();
    const mv = (ref) => {
      if (!ref.row) return ref;
      if (ref.row >= at) return { ...ref, row: ref.row + n };
      return ref;
    };
    const moved = new Map();
    for (const [rn, r] of sh.rows) {
      const nr = rn >= at ? rn + n : rn;
      const cells = new Map();
      for (const [c, x] of r.cells) {
        let nx = x.includes('<f') ? mapFormulaXml(x, (ref) => (ref.sheet && ref.sheet !== name ? ref : mv(ref))) : x;
        if (nr !== rn) nx = nx.replace(/^<c r="([A-Z]+)\d+"/, `<c r="$1${nr}"`);
        nx = nx.replace(/(<f\b[^>]*\sref=")([^"]+)"/, (m, a, r0) => `${a}${shiftRefs(r0, mv)}"`);
        cells.set(c, nx);
      }
      moved.set(nr, { attrs: r.attrs.replace(/\sr="\d+"/, ` r="${nr}"`), cells });
    }
    sh.rows = moved;
    await this.shiftRefsEverywhere(name, mv, { skipSelf: true });
    let xml = sh.xml;
    xml = xml.replace(/<mergeCell ref="([^"]+)"/g, (m, r) => `<mergeCell ref="${shiftRefs(r, mv)}"`);
    xml = xml.replace(/sqref="([^"]+)"/g, (m, r) => `sqref="${sqrefShift(r, mv)}"`);
    xml = xml.replace(/(<autoFilter ref=")([^"]+)"/, (m, a, r) => `${a}${shiftRefs(r, mv)}"`);
    sh.xml = xml;
  }
  async copySheetRels(srcPath, dstPath, sheetXml) {
    const relPath = srcPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
    const rf = this.zip.file(relPath);
    if (!rf) return sheetXml.replace(/<(drawing|legacyDrawing|tableParts)\b[^>]*\/>/g, '');
    let rx = await rf.async('string');
    for (const m of [...rx.matchAll(/<Relationship\b[^>]*\/>/g)]) {
      const target = m[0].match(/Target="([^"]+)"/)[1];
      const type = m[0].match(/Type="[^"]*\/([^"/]+)"/)[1];
      if (!['comments', 'vmlDrawing', 'drawing'].includes(type) || /TargetMode="External"/.test(m[0])) {
        if (!/TargetMode="External"/.test(m[0])) rx = rx.replace(m[0], '');
        continue;
      }
      const abs = 'xl/' + target.replace('../', '');
      if (!this.zip.file(abs)) { rx = rx.replace(m[0], ''); continue; }
      const ext = abs.split('.').pop();
      const base = abs.replace(/\d*\.\w+$/, '');
      let i = 1; while (this.zip.file(`${base}${i}.${ext}`)) i++;
      const nAbs = `${base}${i}.${ext}`;
      this.zip.file(nAbs, await this.zip.file(abs).async('uint8array'));
      const subRel = abs.replace(/([^/]+)$/, '_rels/$1.rels');
      if (this.zip.file(subRel)) this.zip.file(nAbs.replace(/([^/]+)$/, '_rels/$1.rels'), await this.zip.file(subRel).async('uint8array'));
      if (ext === 'xml') {
        const ctm = this.ct.match(new RegExp(`<Override PartName="/${reEsc(abs)}" ContentType="([^"]+)"/>`));
        if (ctm) this.ct = this.ct.replace('</Types>', `<Override PartName="/${nAbs}" ContentType="${ctm[1]}"/></Types>`);
      }
      rx = rx.replace(m[0], m[0].replace(target, '../' + nAbs.replace('xl/', '')));
    }
    this.zip.file(dstPath.replace('worksheets/', 'worksheets/_rels/') + '.rels', rx);
    return sheetXml;
  }
  // 새 시트 추가: sheet(Sheet 객체), relsFrom(메모·도형을 복사할 원본 시트 이름), after/before 위치
  async addSheet(name, sheet, { relsFrom, after, before } = {}) {
    if (this.has(name)) throw new Error(`이미 '${name}' 시트가 있습니다`);
    let i = 1; while (this.zip.file(`xl/worksheets/sheet${i}.xml`)) i++;
    const path = `xl/worksheets/sheet${i}.xml`;
    const src = relsFrom && this.sheets.find((s) => s.name === relsFrom);
    let xml = sheet.build();
    xml = xml.replace(/\stabSelected="1"/, '');
    xml = src ? await this.copySheetRels(src.path, path, xml) : xml.replace(/<(drawing|legacyDrawing|tableParts)\b[^>]*\/>/g, '');
    sheet.xml = xml; sheet.name = name;
    this.zip.file(path, xml);
    const rid = `rId${Math.max(...[...this.relXml.matchAll(/Id="rId(\d+)"/g)].map((m) => +m[1])) + 1}`;
    this.relXml = this.relXml.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/></Relationships>`);
    this.ct = this.ct.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
    const sheetId = Math.max(...this.sheets.map((s) => s.sheetId)) + 1;
    const tag = `<sheet name="${esc(name)}" sheetId="${sheetId}" r:id="${rid}"/>`;
    const list = [...this.wbXml.matchAll(/<sheet\b[^>]*\/>/g)].map((m) => m[0]);
    const nm = (t) => unesc(t.match(/name="([^"]+)"/)[1]);
    let pos = list.length;
    const ai = after ? list.findIndex((t) => nm(t) === after) : -1;
    const bi = before ? list.findIndex((t) => nm(t) === before) : -1;
    if (ai >= 0) pos = ai + 1; else if (bi >= 0) pos = bi;
    list.splice(pos, 0, tag);
    this.wbXml = this.wbXml.replace(/<sheets>[\s\S]*?<\/sheets>/, `<sheets>${list.join('')}</sheets>`);
    this.wbXml = this.wbXml.replace(/localSheetId="(\d+)"/g, (m, k) => `localSheetId="${+k >= pos ? +k + 1 : k}"`);
    this.sheets.push({ name, path, sheetId });
    this.cache.set(name, sheet);
    this.added = (this.added || []).concat(path);
    return pos;
  }
  // 저장: 고친 시트 반영, 열 때 전체 재계산, 계산값(<v>)은 남겨 둠(엑셀이 다시 계산)
  async save({ activeSheet } = {}) {
    for (const [name, sh] of this.cache) {
      const s = this.sheets.find((x) => x.name === name);
      if (!s) continue;
      let xml = sh.build();
      if (activeSheet) xml = xml.replace(/\stabSelected="1"/, '');
      if (activeSheet === name) xml = xml.replace(/<sheetView\b/, '<sheetView tabSelected="1"');
      this.zip.file(s.path, xml);
    }
    if (activeSheet) {
      for (const s of this.sheets) {
        if (this.cache.has(s.name)) continue;
        const x = await this.zip.file(s.path).async('string');
        if (/tabSelected="1"/.test(x)) this.zip.file(s.path, x.replace(/\stabSelected="1"/, ''));
      }
      const idx = this.names().indexOf(activeSheet);
      if (idx >= 0) {
        if (/activeTab="\d+"/.test(this.wbXml)) this.wbXml = this.wbXml.replace(/activeTab="\d+"/, `activeTab="${idx}"`);
        else this.wbXml = this.wbXml.replace(/<workbookView\b/, `<workbookView activeTab="${idx}"`);
        this.wbXml = this.wbXml.replace(/firstSheet="\d+"/, 'firstSheet="0"');
      }
    }
    this.wbXml = this.wbXml.replace(/<calcPr\b([^>]*?)\/>/, (m, a) => `<calcPr${a.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`);
    if (!/<calcPr/.test(this.wbXml)) this.wbXml = this.wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    this.zip.remove('xl/calcChain.xml');
    this.ct = this.ct.replace(/<Override[^>]*calcChain[^>]*\/>/, '');
    this.relXml = this.relXml.replace(/<Relationship[^>]*calcChain[^>]*\/>/, '');
    this.zip.file('xl/workbook.xml', this.wbXml);
    this.zip.file('xl/_rels/workbook.xml.rels', this.relXml);
    this.zip.file('[Content_Types].xml', this.ct);
    return this.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
}

/* ---------- RD(ERP 엑셀) 읽기: 머리글 행을 찾아 표로 ---------- */
// rows: sheet_to_json(header:1) 결과. want: 찾을 머리글 이름 목록. 머리글이 두 줄(KRW 아래 기초/차변…)이어도 합쳐 읽음
export function findTable(grid, required) {
  for (let i = 0; i < Math.min(grid.length, 40); i++) {
    const r = (grid[i] || []).map(norm);
    if (required.every((k) => r.some((x) => x === k))) {
      const sub = (grid[i + 1] || []).map(norm);
      const head = r.map((h, c) => h || sub[c] || '');
      // 아래 줄에 기초/차변/대변/잔액 같은 하위 머리글이 있으면 그것을 우선
      sub.forEach((s, c) => { if (s && ['기초', '차변', '대변', '잔액', 'year', 'month'].includes(s)) head[c] = s; });
      const twoLine = sub.some((s) => ['기초', '차변', '대변', '잔액'].includes(s));
      return { headerRow: i, head, start: i + (twoLine ? 2 : 1) };
    }
  }
  return null;
}
