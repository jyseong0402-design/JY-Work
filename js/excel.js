// 엑셀 내보내기/가져오기 (SheetJS)
import { uid } from './store.js';

const XLSX_URL = new URL('./vendor/xlsx.full.min.js', import.meta.url).href; // SheetJS 0.18.5 (Apache-2.0)
let loading = null;

function loadXLSX() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (!loading) {
    loading = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = XLSX_URL;
      s.onload = () => resolve(window.XLSX);
      s.onerror = () => { loading = null; reject(new Error('엑셀 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해주세요.')); };
      document.head.appendChild(s);
    });
  }
  return loading;
}

const list = (a) => (Array.isArray(a) ? a.join('\n') : a || '');
const split = (s) => String(s || '').split(/\n/).map((x) => x.trim()).filter(Boolean);

const COLS = {
  events: [
    ['ID', 'id'], ['날짜', 'date'], ['시작', 'start'], ['종료', 'end'], ['제목', 'title'], ['분류', 'category'],
    ['장소', 'place'], ['참석자', 'attendees'], ['메모', 'note'], ['완료', 'done'],
  ],
  meetings: [
    ['ID', 'id'], ['날짜', 'date'], ['노트', 'series'], ['제목', 'title'], ['참석자', 'attendees'], ['태그', 'tags'], ['요약', 'summary'],
    ['결정사항', 'decisions'], ['할일', 'actions'], ['키워드', 'keywords'], ['요약방식', 'summarizedBy'], ['필기(쪽)', 'inkPages'], ['원문', 'content'],
  ],
  tasks: [
    ['ID', 'id'], ['할일', 'text'], ['담당', 'owner'], ['마감', 'due'], ['완료', 'done'], ['출처회의', 'source'], ['회의ID', 'meetingId'],
  ],
};
const SHEET = { events: '일정', meetings: '회의메모', tasks: '할일' };

function toRow(kind, item) {
  const row = {};
  for (const [label, key] of COLS[kind]) {
    let v = item[key];
    if (key === 'done') v = v ? 'O' : '';
    else if (key === 'actions') v = (v || []).map((a) => `${a.text}${a.owner ? ' / ' + a.owner : ''}${a.due ? ' / ' + a.due : ''}`).join('\n');
    else if (Array.isArray(v)) v = list(v);
    row[label] = v ?? '';
  }
  return row;
}

function fromRow(kind, row) {
  const item = {};
  for (const [label, key] of COLS[kind]) {
    let v = row[label];
    if (v === undefined) continue;
    if (key === 'inkPages') v = Number(v) || 0;
    else if (key === 'done') v = v === true || String(v).trim().toUpperCase() === 'O' || v === 'TRUE';
    else if (key === 'decisions' || key === 'keywords') v = split(v);
    else if (key === 'actions') v = split(v).map((line) => {
      const [text, owner = '', due = ''] = line.split(' / ');
      return { text, owner, due };
    });
    else if (v instanceof Date) v = v.toISOString().slice(0, 10);
    else v = String(v);
    item[key] = v;
  }
  if (!item.id) item.id = uid();
  return item;
}

function autoWidth(ws, rows) {
  const keys = rows.length ? Object.keys(rows[0]) : [];
  ws['!cols'] = keys.map((k) => ({
    wch: Math.min(60, Math.max(k.length * 2 + 2, ...rows.map((r) => String(r[k] ?? '').split('\n')[0].length * 1.6))),
  }));
}

export async function exportExcel(data, filename = '업무기록.xlsx', kinds = ['events', 'meetings', 'tasks']) {
  const XLSX = await loadXLSX();
  const wb = XLSX.utils.book_new();
  for (const kind of kinds) {
    const rows = (data[kind] || []).slice().sort((a, b) => String(a.date || a.due || '').localeCompare(String(b.date || b.due || '')))
      .map((x) => toRow(kind, x));
    const ws = rows.length
      ? XLSX.utils.json_to_sheet(rows)
      : XLSX.utils.aoa_to_sheet([COLS[kind].map(([l]) => l)]);
    autoWidth(ws, rows);
    XLSX.utils.book_append_sheet(wb, ws, SHEET[kind]);
  }
  XLSX.writeFile(wb, filename);
}

export async function importExcel(file) {
  const XLSX = await loadXLSX();
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { cellDates: true });
  const out = { events: [], meetings: [], tasks: [] };
  for (const kind of Object.keys(SHEET)) {
    const ws = wb.Sheets[SHEET[kind]];
    if (!ws) continue;
    out[kind] = XLSX.utils.sheet_to_json(ws, { defval: '' }).map((r) => fromRow(kind, r));
  }
  return out;
}
