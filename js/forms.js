// 📄 월마감 양식: 업무별 완성본 파일(양식)을 이 기기(IndexedDB 'jywork-forms')에 보관 + 파일명 규칙 처리
// 파일명 규칙 표기: {YYYY} {YY} {MM} {M} {YYMM} {END}(말일) {Mon}(Jan…) — {MM+1}처럼 달을 더하거나 뺄 수 있음

const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const TOKEN = /\{(YYYY|YYMM|YY|MM|M|END|Mon)([+-]\d+)?\}/g;
const p2 = (n) => String(n).padStart(2, '0');

// 규칙 → 그 달의 실제 이름 (ym = 'YYYY-MM')
export function fillName(pattern, ym) {
  if (!pattern) return '';
  const [y0, m0] = ym.split('-').map(Number);
  return String(pattern).replace(TOKEN, (_, k, off) => {
    const d = new Date(y0, m0 - 1 + (+off || 0), 1);
    const Y = d.getFullYear(), M = d.getMonth() + 1;
    switch (k) {
      case 'YYYY': return String(Y);
      case 'YY': return String(Y).slice(2);
      case 'YYMM': return String(Y).slice(2) + p2(M);
      case 'MM': return p2(M);
      case 'M': return String(M);
      case 'END': return p2(new Date(Y, M, 0).getDate());
      case 'Mon': return MON[M - 1];
      default: return _;
    }
  });
}

// 규칙 → 어느 달 파일이든 알아보는 정규식 (버전 _v1a 등은 아무 버전, 뒤에 붙은 말은 extra로)
const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const TOKEN_RE = { YYYY: '\\d{4}', YYMM: '\\d{4}', YY: '\\d{2}', MM: '\\d{1,2}', M: '\\d{1,2}', END: '\\d{1,2}', Mon: '[A-Za-z]{3}' };
function literalRe(s) {
  return s.split(/([_-]v\d+[a-z]*)/i).map((part, i) => (i % 2
    ? '(?:[_-]v\\d+[a-z0-9.]*)?'
    : escRe(part).replace(/\s+/g, '\\s*').replace(/년/g, '년\\s*'))).join('');
}
// 반환: { re, toks:[{k,off}] } — 달 표기는 캡처해서 어느 달 파일인지 알아냄 (마지막 앞 캡처 = 뒤에 붙은 말)
export function patternRe(pattern) {
  if (!pattern || !/\S/.test(pattern)) return null;
  let body = '', last = 0;
  const toks = [];
  for (const m of String(pattern).matchAll(TOKEN)) {
    body += literalRe(pattern.slice(last, m.index)) + `(${TOKEN_RE[m[1]]})`;
    toks.push({ k: m[1], off: +m[2] || 0 });
    last = m.index + m[0].length;
  }
  body += literalRe(pattern.slice(last));
  return { re: new RegExp(`^\\s*${body}(.*?)\\.(xlsx|xlsm|xlsb|xls|pptx|ppt|pdf|docx|csv)$`, 'i'), toks };
}
// 파일 이름 속 달 → 비교용 숫자 (YYYYMM, 연도 없으면 달만)
function periodOf(m, toks) {
  let y = 0, mo = 0;
  toks.forEach(({ k, off }, i) => {
    const v = m[i + 1];
    if (k === 'YYYY') y = +v;
    else if (k === 'YY' && !y) y = 2000 + +v;
    else if (k === 'YYMM') { y = 2000 + +v.slice(0, 2); mo = +v.slice(2) - off; }
    else if ((k === 'MM' || k === 'M') && !mo) mo = +v - off;
  });
  return y * 100 + mo;
}

// 파일 여러 개 → 업무별로 가장 알맞은 파일 (target 달 파일 → 다른 말이 덜 붙은 것 → 최근 달 → 최근 수정본)
// target = 'YYYY-MM' (주면 그 달 파일을 먼저 고름)
export function matchFiles(tasks, files, target) {
  const res = tasks.map((t) => ({ t, p: patternRe(t.file) })).filter((x) => x.p);
  const byTask = new Map();
  const unmatched = [];
  const tg = target ? +target.replace('-', '') : 0;
  const hit = (per) => (tg && (per === tg || (per < 100 && per === tg % 100)) ? 1 : 0);
  const better = (a, b) => hit(a.per) > hit(b.per) || (hit(a.per) === hit(b.per) && (a.extra < b.extra || (a.extra === b.extra && (a.per > b.per || (a.per === b.per && a.f.lastModified > b.f.lastModified)))));
  for (const f of files) {
    let best = null;
    for (const { t, p } of res) {
      const m = p.re.exec(f.name);
      if (!m) continue;
      const extra = (m[p.toks.length + 1] || '').trim().length;
      if (!best || extra < best.extra || (extra === best.extra && t.file.length > best.t.file.length)) best = { t, extra, per: periodOf(m, p.toks), f };
    }
    if (!best) { unmatched.push(f); continue; }
    const cur = byTask.get(best.t.id);
    if (!cur || better(best, cur)) byTask.set(best.t.id, best);
  }
  return { matched: [...byTask.values()].map((x) => ({ ...x, hit: !!hit(x.per) })), unmatched };
}

/* ---------- 📁 업무 폴더 연결 (PC 크롬·엣지: File System Access) ----------
   업무 폴더(예: Desktop\업무)를 한 번 고르면, 업무마다 적힌 저장 위치(folder)를 열어 파일명 규칙에 맞는 완성본을 찾음 */
export const folderOk = () => typeof window !== 'undefined' && 'showDirectoryPicker' in window;
const fdb = () => new Promise((res, rej) => {
  const r = indexedDB.open('jywork-folders', 1);
  r.onupgradeneeded = () => { if (!r.result.objectStoreNames.contains('h')) r.result.createObjectStore('h'); };
  r.onsuccess = () => res(r.result);
  r.onerror = () => rej(r.error);
});
export async function getWorkRoot() {
  try {
    const d = await fdb();
    return await new Promise((res) => { const q = d.transaction('h').objectStore('h').get('work'); q.onsuccess = () => res(q.result || null); q.onerror = () => res(null); });
  } catch { return null; }
}
async function putWorkRoot(h) {
  const d = await fdb();
  return new Promise((res, rej) => { const t = d.transaction('h', 'readwrite'); t.objectStore('h').put(h, 'work'); t.oncomplete = res; t.onerror = () => rej(t.error); });
}
export async function pickWorkRoot() {
  const h = await window.showDirectoryPicker({ id: 'jywork-work', mode: 'read' });
  await putWorkRoot(h);
  return h;
}
// 읽기 권한 (ask = 버튼 누른 직후라 허용 창을 띄울 수 있을 때)
export async function workPerm(h, ask) {
  if (!h) return false;
  let p = await h.queryPermission({ mode: 'read' });
  if (p !== 'granted' && ask) p = await h.requestPermission({ mode: 'read' });
  return p === 'granted';
}
const norm = (s) => s.normalize('NFC').replace(/\s+/g, ' ').trim().toLowerCase();
async function childDir(dir, name) {
  try { return await dir.getDirectoryHandle(name); } catch { /* 이름이 조금 다를 수 있음 */ }
  for await (const [n, h] of dir.entries()) if (h.kind === 'directory' && norm(n) === norm(name)) return h;
  return null;
}
// 'A\B\C' → 폴더 핸들 (업무 폴더보다 위를 골랐으면 그 안의 '업무' 폴더부터 찾아 봄)
export async function openPath(root, path) {
  const parts = String(path).split(/[\\/]+/).filter(Boolean);
  const walk = async (start) => { let d = start; for (const p of parts) { d = await childDir(d, p); if (!d) return null; } return d; };
  return (await walk(root)) || (await (async () => { const w = await childDir(root, '업무'); return w ? walk(w) : null; })());
}
export async function listFiles(dir) {
  const out = [];
  for await (const [name, h] of dir.entries()) {
    if (h.kind !== 'file' || name.startsWith('~$')) continue;
    out.push({ name, h });
  }
  return out;
}

/* ---------- IndexedDB: files(파일 내용) / meta(이름·크기·날짜만 — 목록 표시용) ---------- */
let dbp = null;
function db() {
  if (!dbp) {
    dbp = new Promise((res, rej) => {
      const r = indexedDB.open('jywork-forms', 1);
      r.onupgradeneeded = () => {
        if (!r.result.objectStoreNames.contains('files')) r.result.createObjectStore('files');
        if (!r.result.objectStoreNames.contains('meta')) r.result.createObjectStore('meta');
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => { dbp = null; rej(r.error); };
    });
  }
  return dbp;
}
async function run(stores, mode, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(stores, mode);
    const out = fn(t);
    t.oncomplete = () => res(out && 'result' in out ? out.result : out);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error || new Error('저장하지 못했습니다 (저장 공간 부족일 수 있음)'));
  });
}

// 모든 업무의 양식 정보 { taskId: {name,size,at,lastModified} }
export async function formMetaAll() {
  const d = await db();
  return new Promise((res) => {
    const out = {};
    const q = d.transaction('meta').objectStore('meta').openCursor();
    q.onsuccess = () => { const c = q.result; if (c) { out[c.key] = c.value; c.continue(); } else res(out); };
    q.onerror = () => res(out);
  });
}
export async function formPut(taskId, file) {
  const data = new Uint8Array(await file.arrayBuffer());
  const meta = { name: file.name, size: file.size, type: file.type || '', at: Date.now(), lastModified: file.lastModified || Date.now() };
  await run(['files', 'meta'], 'readwrite', (t) => { t.objectStore('files').put({ ...meta, data }, taskId); t.objectStore('meta').put(meta, taskId); });
  return meta;
}
export async function formGetFile(taskId) {
  const v = await run(['files'], 'readonly', (t) => t.objectStore('files').get(taskId));
  if (!v) return null;
  return new File([v.data], v.name, { type: v.type || '', lastModified: v.lastModified || v.at });
}
export const formDel = (taskId) => run(['files', 'meta'], 'readwrite', (t) => { t.objectStore('files').delete(taskId); t.objectStore('meta').delete(taskId); });
