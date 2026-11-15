// ✅ 월마감 현황: 내가 만들 파일(업무) · 필요파일 · 파일 생산자 목록 + 월별 '받음'·'완료' 체크
// 업무 1개 = store 'closing' 항목 1개 { id, order, name, file, timing, auto, needs:[{id,name,producer}], marks:{'YYYY-MM':{done, recv:{needId:date}}} }
import { store, settings, uid, todayStr } from './store.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const AUTO = { '': '연결 없음', freight: '🚚 운반비 검증', arap: '💳 채권·채무', prepay: '💰 선급금', internal: '🔁 내부거래 집계', packing: '📦 포장소모품비' };
const ym = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
const md = (s) => (s ? `${+s.slice(5, 7)}/${+s.slice(8, 10)}` : '');

const ui = { month: ym(new Date()), editing: null, draft: null, openAuto: null };

const list = () => [...(store.get().closing || [])].sort((a, b) => (a.order ?? 0) - (b.order ?? 0) || String(a.createdAt).localeCompare(String(b.createdAt)));
const markOf = (t) => (t.marks && t.marks[ui.month]) || { done: '', recv: {} };
function setMark(t, fn) {
  const marks = { ...(t.marks || {}) };
  const m = { done: '', recv: {}, ...(marks[ui.month] || {}) };
  m.recv = { ...(m.recv || {}) };
  fn(m);
  marks[ui.month] = m;
  store.upsert('closing', { id: t.id, marks });
}

function monthLabel(s) { const [y, m] = s.split('-'); return `${y}년 ${+m}월`; }
function shiftMonth(n) { const [y, m] = ui.month.split('-').map(Number); ui.month = ym(new Date(y, m - 1 + n, 1)); render(); }

function taskView(t, i) {
  const mk = markOf(t);
  const needs = t.needs || [];
  const got = needs.filter((n) => mk.recv?.[n.id]).length;
  return `<div class="bd-task ${mk.done ? 'is-done' : ''}" data-id="${t.id}" draggable="true">
    <div class="bd-row">
      <span class="bd-no">${i + 1}</span>
      <div class="bd-main"><span class="bd-drag" title="끌어서 순서 바꾸기">⠿</span><b>${esc(t.name)}</b>${t.file ? `<div class="bd-file">${esc(t.file)}</div>` : ''}</div>
      <span class="bd-timing">${esc(t.timing || '')}</span>
      <span class="bd-need-sum">${needs.length ? `필요파일 <b>${got}/${needs.length}</b>` : ''}</span>
      <span class="bd-acts">
        ${t.auto ? `<button class="btn sm primary" data-auto="${t.auto}" title="백업자료 생성으로 이동">📊 바로 만들기</button>` : ''}
        <label class="bd-done ${mk.done ? 'on' : ''}"><input type="checkbox" data-done="${t.id}" ${mk.done ? 'checked' : ''}/> ${mk.done ? `완료 ${md(mk.done)}` : '완료'}</label>
        <button class="bd-ib" data-edit="${t.id}" title="편집">✎</button>
        <button class="bd-ib" data-del="${t.id}" title="삭제">🗑</button>
      </span>
    </div>
    ${needs.length ? `<div class="bd-needs">${needs.map((n) => `<label class="bd-need ${mk.recv?.[n.id] ? 'on' : ''}">
        <input type="checkbox" data-recv="${t.id}|${n.id}" ${mk.recv?.[n.id] ? 'checked' : ''}/>
        <span class="bd-nname">${esc(n.name)}</span><span class="bd-prod">${esc(n.producer || '')}</span>
        <span class="bd-rdate">${mk.recv?.[n.id] ? `받음 ${md(mk.recv[n.id])}` : '<span class="muted">미수취</span>'}</span></label>`).join('')}</div>` : ''}
  </div>`;
}

function editView(d, i) {
  return `<div class="bd-task bd-editing" data-id="${d.id || ''}">
    <div class="bd-row"><span class="bd-no">${i + 1}</span><div class="bd-main"><b>${d.id ? '업무 편집' : '새 업무'}</b></div><span></span><span></span>
      <span class="bd-acts">${d.id ? '<button class="bd-ib" data-tmove="-1" title="위로">▲</button><button class="bd-ib" data-tmove="1" title="아래로">▼</button>' : ''}<button class="btn sm primary" data-save>저장</button><button class="btn sm" data-cancel>취소</button></span></div>
    <div class="bd-form">
      <label><span>업무명</span><input data-f="name" value="${esc(d.name)}" placeholder="예: 운반비_집계"/></label>
      <label><span>작성 파일명</span><input data-f="file" value="${esc(d.file)}" placeholder="예: 운반비_집계_2600_v1a_검증"/></label>
      <label><span>업무 시기</span><input data-f="timing" value="${esc(d.timing)}" placeholder="예: 전사 매출마감 이후"/></label>
      <label><span>바로 만들기 연결</span><select data-f="auto">${Object.entries(AUTO).map(([k, v]) => `<option value="${k}" ${d.auto === k ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
    </div>
    <div class="bd-nform">
      <div class="bd-nhead"><span>필요파일</span><span>파일 생산자 (담당자·부서)</span><span></span></div>
      ${d.needs.map((n, k) => `<div class="bd-nrow" data-k="${k}"><input data-n="name" value="${esc(n.name)}" placeholder="파일명"/><input data-n="producer" value="${esc(n.producer)}" placeholder="누가 주는지"/><button class="bd-ib" data-nup="${k}" title="위로">▲</button><button class="bd-ib" data-ndel="${k}" title="삭제">🗑</button></div>`).join('')}
      <button class="bd-add" data-nadd>＋ 필요파일 추가</button>
    </div>
  </div>`;
}

function readDraft() {
  const box = $('#bdList .bd-editing');
  if (!box || !ui.draft) return;
  box.querySelectorAll('[data-f]').forEach((el) => { ui.draft[el.dataset.f] = el.value; });
  box.querySelectorAll('.bd-nrow').forEach((row) => {
    const n = ui.draft.needs[+row.dataset.k];
    if (!n) return;
    row.querySelectorAll('[data-n]').forEach((el) => { n[el.dataset.n] = el.value; });
  });
}

export function renderBoard() {
  if (!$('#view-board')) return;
  const items = list();
  $('#bdMonth').textContent = monthLabel(ui.month);
  let done = 0, needAll = 0, needGot = 0;
  for (const t of items) {
    const mk = markOf(t);
    if (mk.done) done++;
    for (const n of t.needs || []) { needAll++; if (mk.recv?.[n.id]) needGot++; }
  }
  $('#bdStats').innerHTML = `
    <div class="bd-stat"><small>전체 업무</small><b>${items.length}</b></div>
    <div class="bd-stat"><small>완료</small><b class="ok">${done}</b></div>
    <div class="bd-stat"><small>남은 업무</small><b>${items.length - done}</b></div>
    <div class="bd-stat"><small>못 받은 필요파일</small><b class="${needAll - needGot ? 'wn' : ''}">${needAll - needGot}</b></div>
    <div class="bd-prog"><i style="width:${items.length ? Math.round((done / items.length) * 100) : 0}%"></i></div>`;
  const rows = items.map((t, i) => (ui.editing === t.id ? editView(ui.draft, i) : taskView(t, i)));
  if (ui.editing === 'new') rows.push(editView(ui.draft, items.length));
  $('#bdList').innerHTML = rows.join('') || `<div class="empty">아직 업무가 없습니다. <b>＋ 업무 추가</b>로 만들거나, <b>구글시트에서 가져오기</b>로 한 번에 채우세요.</div>`;
  const url = settings.get().closingSheet || '';
  $('#bdSheetUrl').value = url;
}
const render = renderBoard;

function startEdit(id) {
  readDraft();
  if (id === 'new') ui.draft = { name: '', file: '', timing: '', auto: '', needs: [{ id: uid(), name: '', producer: '' }] };
  else { const t = store.find('closing', id); ui.draft = JSON.parse(JSON.stringify({ needs: [], auto: '', ...t })); }
  ui.editing = id;
  render();
  $('#bdList .bd-editing input')?.focus();
}
function save() {
  readDraft();
  const d = ui.draft;
  if (!d.name.trim()) { alert('업무명을 입력해 주세요.'); return; }
  const needs = d.needs.filter((n) => n.name.trim() || n.producer.trim()).map((n) => ({ id: n.id || uid(), name: n.name.trim(), producer: n.producer.trim() }));
  const base = { name: d.name.trim(), file: d.file.trim(), timing: d.timing.trim(), auto: d.auto || '', needs };
  if (ui.editing === 'new') store.upsert('closing', { ...base, order: (list().at(-1)?.order ?? 0) + 1, marks: {} });
  else store.upsert('closing', { id: ui.editing, ...base });
  ui.editing = null; ui.draft = null;
  render();
}

/* ---------- 구글시트에서 가져오기 (공유: 링크가 있는 모든 사용자) ---------- */
function sheetIds(url) {
  const id = (url.match(/\/d\/([a-zA-Z0-9_-]+)/) || [])[1];
  const gid = (url.match(/[#&?]gid=(\d+)/) || [])[1] || '0';
  return id ? { id, gid } : null;
}
function fetchSheet(url) {
  const ids = sheetIds(url);
  if (!ids) return Promise.reject(new Error('구글 스프레드시트 주소가 아닙니다.'));
  return new Promise((res, rej) => {
    const cb = `__bdSheet${Date.now()}`;
    const s = document.createElement('script');
    const done = () => { delete window[cb]; s.remove(); };
    window[cb] = (r) => { done(); if (r.status === 'error') rej(new Error('시트를 읽지 못했습니다. 공유가 "링크가 있는 모든 사용자"인지 확인하세요.')); else res(r.table); };
    s.onerror = () => { done(); rej(new Error('시트를 읽지 못했습니다. 공유 설정과 인터넷 연결을 확인하세요.')); };
    s.src = `https://docs.google.com/spreadsheets/d/${ids.id}/gviz/tq?gid=${ids.gid}&headers=0&tqx=out:json;responseHandler:${cb}`;
    document.head.appendChild(s);
    setTimeout(() => { if (window[cb]) { done(); rej(new Error('시간 초과: 시트 공유 설정을 확인하세요.')); } }, 15000);
  });
}
// 시트 양식: B=업무, C=필요파일(ㄴ…), D=파일명, E=업무시기/파일생산자
export function parseSheetRows(rows) {
  const tasks = [];
  let cur = null;
  const txt = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();
  for (const r of rows) {
    const c = (i) => txt(r[i]);
    const b = c(1), cc = c(2), d = c(3), e = c(4);
    if (b === '업무' || cc === '필요파일' || /파일$/.test(cc) && d === '파일명') continue;
    if (b) {
      cur = { name: b, file: d, timing: e, needs: [] };
      tasks.push(cur);
      continue;
    }
    if (!cur) continue;
    const label = cc.replace(/^ㄴ\s*/, '');
    if (!label && !d) continue;
    cur.needs.push({ name: label ? (d && d !== label ? `${label} (${d})` : label) : d, producer: e });
  }
  return tasks;
}
const guessAuto = (name) => (/운반비/.test(name) ? 'freight' : /내부거래/.test(name) ? 'internal' : /포장소모품비/.test(name) ? 'packing' : /^AR\/AP$/.test(name) ? 'arap' : '');

async function importSheet() {
  const url = $('#bdSheetUrl').value.trim();
  if (!url) { alert('구글 스프레드시트 주소를 붙여 넣어 주세요.'); return; }
  settings.set({ closingSheet: url });
  const btn = $('#bdImport');
  btn.disabled = true; btn.textContent = '읽는 중…';
  try {
    const table = await fetchSheet(url);
    const rows = table.rows.map((row) => (row.c || []).map((x) => (x ? (x.f ?? x.v) : '')));
    const tasks = parseSheetRows(rows);
    if (!tasks.length) throw new Error('업무 줄을 찾지 못했습니다. 시트 양식(B열 업무, C열 필요파일)을 확인하세요.');
    const have = list();
    let added = 0, updated = 0;
    let order = (have.at(-1)?.order ?? 0);
    for (const t of tasks) {
      const ex = have.find((x) => x.name === t.name);
      if (ex) {
        const names = new Set((ex.needs || []).map((n) => n.name));
        const more = t.needs.filter((n) => !names.has(n.name)).map((n) => ({ id: uid(), ...n }));
        if (more.length || (!ex.file && t.file) || (!ex.timing && t.timing)) {
          store.upsert('closing', { id: ex.id, file: ex.file || t.file, timing: ex.timing || t.timing, needs: [...(ex.needs || []), ...more] });
          updated++;
        }
      } else {
        store.upsert('closing', { ...t, needs: t.needs.map((n) => ({ id: uid(), ...n })), auto: guessAuto(t.name), order: ++order, marks: {} });
        added++;
      }
    }
    $('#bdImportMsg').textContent = `가져오기 완료: 새 업무 ${added}개${updated ? `, 기존 업무 ${updated}개에 빠진 내용 추가` : ''} (이미 있는 내용은 그대로 둠)`;
  } catch (err) {
    $('#bdImportMsg').textContent = '⚠ ' + (err.message || err);
  } finally {
    btn.disabled = false; btn.textContent = '가져오기';
  }
}

/* ---------- 순서 바꾸기 (끌어 놓기) ---------- */
function reorder(dragId, targetId, after) {
  const items = list().filter((t) => t.id !== dragId);
  const drag = store.find('closing', dragId);
  let at = items.findIndex((t) => t.id === targetId);
  if (at < 0) return;
  if (after) at++;
  items.splice(at, 0, drag);
  items.forEach((t, i) => { if (t.order !== i + 1) store.upsert('closing', { id: t.id, order: i + 1 }); });
}

export function initBoard({ openAuto } = {}) {
  if (!$('#view-board')) return;
  ui.openAuto = openAuto;
  $('#bdPrev').onclick = () => shiftMonth(-1);
  $('#bdNext').onclick = () => shiftMonth(1);
  $('#bdAdd').onclick = () => startEdit('new');
  $('#bdImport').onclick = importSheet;
  const box = $('#bdList');
  box.addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.done) {
      const item = store.find('closing', t.dataset.done);
      setMark(item, (m) => { m.done = t.checked ? todayStr() : ''; });
    } else if (t.dataset.recv) {
      const [tid, nid] = t.dataset.recv.split('|');
      setMark(store.find('closing', tid), (m) => { if (t.checked) m.recv[nid] = todayStr(); else delete m.recv[nid]; });
    }
  });
  box.addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.edit) startEdit(b.dataset.edit);
    else if (b.dataset.del) {
      const t = store.find('closing', b.dataset.del);
      if (confirm(`'${t.name}' 업무를 삭제할까요? 필요파일 목록과 체크 기록도 함께 지워집니다.`)) store.remove('closing', t.id);
    } else if (b.hasAttribute('data-save')) save();
    else if (b.hasAttribute('data-cancel')) { ui.editing = null; ui.draft = null; render(); }
    else if (b.hasAttribute('data-nadd')) { readDraft(); ui.draft.needs.push({ id: uid(), name: '', producer: '' }); render(); $$last('#bdList .bd-nrow input')?.focus(); }
    else if (b.dataset.ndel !== undefined) { readDraft(); ui.draft.needs.splice(+b.dataset.ndel, 1); render(); }
    else if (b.dataset.nup !== undefined) { readDraft(); const k = +b.dataset.nup; if (k > 0) { const n = ui.draft.needs; [n[k - 1], n[k]] = [n[k], n[k - 1]]; } render(); }
    else if (b.dataset.tmove) {
      readDraft();
      const items = list(); const i = items.findIndex((t) => t.id === ui.editing); const j = i + +b.dataset.tmove;
      if (i >= 0 && j >= 0 && j < items.length) { reorder(ui.editing, items[j].id, +b.dataset.tmove > 0); render(); }
    } else if (b.dataset.auto && ui.openAuto) ui.openAuto(b.dataset.auto);
  });
  let dragId = null;
  box.addEventListener('dragstart', (e) => {
    const t = e.target.closest('.bd-task');
    if (!t || ui.editing) { e.preventDefault(); return; }
    dragId = t.dataset.id; t.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
  });
  box.addEventListener('dragover', (e) => { if (dragId) e.preventDefault(); });
  box.addEventListener('drop', (e) => {
    e.preventDefault();
    const t = e.target.closest('.bd-task');
    if (dragId && t && t.dataset.id !== dragId) {
      const r = t.getBoundingClientRect();
      reorder(dragId, t.dataset.id, e.clientY > r.top + r.height / 2);
    }
    dragId = null;
  });
  box.addEventListener('dragend', () => { dragId = null; box.querySelectorAll('.dragging').forEach((x) => x.classList.remove('dragging')); });
  store.subscribe(() => { if (!ui.editing) render(); });
  render();
}
const $$last = (s) => { const a = document.querySelectorAll(s); return a[a.length - 1]; };
