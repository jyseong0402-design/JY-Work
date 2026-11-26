// ✅ 월마감 현황: 내가 만들 파일(업무) · 필요파일 · 파일 생산자 목록 + 월별 '받음'·'완료' 체크
// 업무 1개 = store 'closing' 항목 1개 { id, order, name, file(파일명 규칙), folder(저장 위치), note, timing, auto,
//   needs:[{id,name,dept,person,rd}], marks:{'YYYY-MM':{done, recv:{needId:date}}} }
// 📄 양식(완성본 파일)은 이 기기 IndexedDB에만 보관 (forms.js)
import { store, uid, todayStr } from './store.js';
import { CLOSING_SEED, SEED_OLD, NEED_OLD } from './closing-seed.js';
import { fillName, matchFiles, formMetaAll, formPut, formGetFile, formDel, folderOk, getWorkRoot, pickWorkRoot, workPerm, openPath, listFiles } from './forms.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const AUTO = { '': '연결 없음', freight: '🚚 운반비 검증', arap: '💳 채권·채무', prepay: '💰 선급금', internal: '🔁 내부거래 집계', packing: '📦 포장소모품비', energy: '⚡ 전력비·수도광열비', rawmat: '🌾 Raw Material(원재료)', capex: '🏗 CAPEX 재분류 검증' };
const ym = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
const md = (s) => (s ? `${+s.slice(5, 7)}/${+s.slice(8, 10)}` : '');
const mdt = (t) => { const d = new Date(t); return `${d.getFullYear() % 100}.${d.getMonth() + 1}.${d.getDate()}`; };
const kb = (n) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)}MB` : `${Math.max(1, Math.round(n / 1024)).toLocaleString()}KB`);
const isXlsx = (n) => /\.xls[xm]$/i.test(n || '');

// 예전 '파일 생산자' 한 칸 → 부서/담당자 (사람 이름은 담당자, 조직 이름은 부서)
const ORG = /팀|기획|법인|ERP|부서|본부|센터|회계|영업부|생산부|재무|운영/;
export function splitProducer(p) {
  const v = String(p || '').trim();
  if (!v) return { dept: '', person: '' };
  return ORG.test(v) && !/님/.test(v) ? { dept: v, person: '' } : { dept: '', person: v };
}
const needOf = (n) => (n.dept !== undefined || n.person !== undefined ? { ...n, dept: n.dept || '', person: n.person || '' } : { ...n, ...splitProducer(n.producer) });

const ui = { month: ym(new Date()), editing: null, draft: null, openAuto: null, setPrev: null, forms: {}, formFor: null };

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
// 보고 있는 달 = 일하는 달, 파일 이름·폴더는 그 전달(결산월) 기준 (11월에 하는 일 = 10월 마감 자료)
const closeOf = (s) => { const [y, m] = s.split('-').map(Number); return ym(new Date(y, m - 2, 1)); };
const fill = (s) => fillName(s, closeOf(ui.month));

function formChip(t) {
  const f = ui.forms[t.id];
  return f
    ? `<button class="btn sm" data-fdl="${t.id}" title="양식 내려받기 · ${esc(f.name)} (${kb(f.size)}, ${mdt(f.at)} 등록)">📄 양식</button>`
    : `<button class="btn sm ghost" data-fup="${t.id}" title="완성본 파일을 이 업무의 양식으로 등록">📎 양식 등록</button>`;
}

function taskView(t, i) {
  const mk = markOf(t);
  const needs = t.needs || [];
  const got = needs.filter((n) => mk.recv?.[n.id]).length;
  const rdN = needs.filter((n) => n.rd).length;
  return `<div class="bd-task ${mk.done ? 'is-done' : ''}" data-id="${t.id}" draggable="true">
    <div class="bd-row">
      <span class="bd-no">${i + 1}</span>
      <div class="bd-main"><span class="bd-drag" title="끌어서 순서 바꾸기">⠿</span><b>${esc(t.name)}</b>
        ${t.file ? `<div class="bd-file" title="파일명 규칙: ${esc(t.file)}">${esc(fill(t.file))}</div>` : ''}
        ${t.folder ? `<div class="bd-folder" title="저장 위치 (업무 폴더 기준)">📁 ${esc(fill(t.folder))}</div>` : ''}</div>
      <span class="bd-timing">${esc(t.timing || '')}</span>
      <span class="bd-need-sum">${needs.length ? `필요파일 <b>${got}/${needs.length}</b>${rdN ? `<br><span class="muted">RD ${rdN}</span>` : ''}` : ''}</span>
      <span class="bd-acts">
        ${formChip(t)}
        ${t.auto ? `<button class="btn sm primary" data-make="${t.id}" title="이 업무의 파일 만들기">📊 파일 생성</button>` : ''}
        <label class="bd-done ${mk.done ? 'on' : ''}"><input type="checkbox" data-done="${t.id}" ${mk.done ? 'checked' : ''}/> ${mk.done ? `완료 ${md(mk.done)}` : '완료'}</label>
        <button class="bd-ib" data-edit="${t.id}" title="편집">✎</button>
        <button class="bd-ib" data-del="${t.id}" title="삭제">🗑</button>
      </span>
    </div>
    ${t.note ? `<div class="bd-note">${esc(fill(t.note))}</div>` : ''}
    ${needs.length ? `<div class="bd-needs"><div class="bd-need bd-need-h"><span></span><span>필요파일</span><span>부서</span><span>담당자</span><span></span></div>${needs.map(needOf).map((n) => `<label class="bd-need ${mk.recv?.[n.id] ? 'on' : ''}">
        <input type="checkbox" data-recv="${t.id}|${n.id}" ${mk.recv?.[n.id] ? 'checked' : ''}/>
        <span class="bd-nname">${n.rd ? '<span class="tag rd">RD</span> ' : ''}${esc(fill(n.name))}</span><span class="bd-dept">${n.dept ? `<span class="tag">${esc(n.dept)}</span>` : '<span class="muted">—</span>'}</span><span class="bd-prod">${esc(n.person) || '<span class="muted">—</span>'}</span>
        <span class="bd-rdate">${mk.recv?.[n.id] ? `${n.rd ? '준비' : '받음'} ${md(mk.recv[n.id])}${mk.auto?.[n.id] && mk.auto[n.id] !== 'off' ? ' <small class="bd-auto" title="RD 보관함에 이 RD가 있어 자동 체크됨">자동</small>' : ''}` : `<span class="muted">${n.rd ? '미준비' : '미수취'}</span>`}</span></label>`).join('')}</div>` : ''}
  </div>`;
}

function editView(d, i) {
  const f = d.id && ui.forms[d.id];
  return `<div class="bd-task bd-editing" data-id="${d.id || ''}">
    <div class="bd-row"><span class="bd-no">${i + 1}</span><div class="bd-main"><b>${d.id ? '업무 편집' : '새 업무'}</b></div><span></span><span></span>
      <span class="bd-acts">${d.id ? '<button class="bd-ib" data-tmove="-1" title="위로">▲</button><button class="bd-ib" data-tmove="1" title="아래로">▼</button>' : ''}<button class="btn sm primary" data-save>저장</button><button class="btn sm" data-cancel>취소</button></span></div>
    <div class="bd-form">
      <label><span>업무명</span><input data-f="name" value="${esc(d.name)}" placeholder="예: 운반비_집계"/></label>
      <label><span>작성 파일명 규칙 <i class="muted">{YYYY} {YY} {MM} {M} {YYMM} {END} → 그 달로 바뀜</i></span><input data-f="file" value="${esc(d.file)}" placeholder="예: 운반비_집계_{YYMM}_v1a_검증"/></label>
      <label><span>업무 시기</span><input data-f="timing" value="${esc(d.timing)}" placeholder="예: 전사 매출마감 이후"/></label>
      <label><span>바로 만들기 연결</span><select data-f="auto">${Object.entries(AUTO).map(([k, v]) => `<option value="${k}" ${d.auto === k ? 'selected' : ''}>${v}</option>`).join('')}</select></label>
      <label class="bd-wide"><span>저장 위치 (업무 폴더 기준)</span><input data-f="folder" value="${esc(d.folder)}" placeholder="예: 4.경영회의\\경영회의 부속자료\\{YYYY}년\\{M}월"/></label>
      <label class="bd-wide"><span>작업 메모 (수기 부분, 함께 저장할 파일 등)</span><textarea data-f="note" rows="2">${esc(d.note)}</textarea></label>
      ${d.id ? `<div class="bd-wide bd-formfile"><span>📄 양식 파일 (이 기기에 보관)</span> ${f ? `<b>${esc(f.name)}</b> <span class="muted">${kb(f.size)} · ${mdt(f.at)} 등록</span> <button class="btn sm" data-fdl="${d.id}">내려받기</button><button class="btn sm" data-fup="${d.id}">바꾸기</button><button class="btn sm" data-fdel="${d.id}">빼기</button>` : `<span class="muted">없음</span> <button class="btn sm" data-fup="${d.id}">📎 등록</button>`}</div>` : ''}
    </div>
    <div class="bd-nform">
      <div class="bd-nhead"><span>필요파일</span><span>부서</span><span>담당자</span><span title="ERP 등에서 직접 뽑는 로우데이터">RD</span><span></span><span></span></div>
      ${d.needs.map((n, k) => `<div class="bd-nrow" data-k="${k}"><input data-n="name" value="${esc(n.name)}" placeholder="파일명"/><input data-n="dept" value="${esc(n.dept)}" placeholder="부서"/><input data-n="person" value="${esc(n.person)}" placeholder="담당자"/><label class="bd-rdchk" title="ERP 등에서 직접 뽑는 로우데이터"><input type="checkbox" data-n="rd" ${n.rd ? 'checked' : ''}/><span>RD</span></label><button class="bd-ib" data-nup="${k}" title="위로">▲</button><button class="bd-ib" data-ndel="${k}" title="삭제">🗑</button></div>`).join('')}
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
    row.querySelectorAll('[data-n]').forEach((el) => { n[el.dataset.n] = el.type === 'checkbox' ? el.checked : el.value; });
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
  const nForms = items.filter((t) => ui.forms[t.id]).length;
  $('#bdStats').innerHTML = `
    <div class="bd-stat"><small>전체 업무</small><b>${items.length}</b></div>
    <div class="bd-stat"><small>완료</small><b class="ok">${done}</b></div>
    <div class="bd-stat"><small>남은 업무</small><b>${items.length - done}</b></div>
    <div class="bd-stat"><small>못 받은 필요파일</small><b class="${needAll - needGot ? 'wn' : ''}">${needAll - needGot}</b></div>
    <div class="bd-prog"><i style="width:${items.length ? Math.round((done / items.length) * 100) : 0}%"></i></div>
    <div class="bd-formstat muted sm">📁 ${monthLabel(closeOf(ui.month))} 마감 자료 · 📄 양식 등록 ${nForms}/${items.length} (이 기기) · <span class="tag rd">RD</span> = ERP 등에서 직접 뽑는 로우데이터</div>`;
  const rows = items.map((t, i) => (ui.editing === t.id ? editView(ui.draft, i) : taskView(t, i)));
  if (ui.editing === 'new') rows.push(editView(ui.draft, items.length));
  $('#bdList').innerHTML = rows.join('') || `<div class="empty">아직 업무가 없습니다. <b>＋ 업무 추가</b>로 만들거나, 아래 <b>📥 목록 파일 불러오기</b>로 넣으세요.</div>`;
}
const render = renderBoard;

function startEdit(id) {
  readDraft();
  if (id === 'new') ui.draft = { name: '', file: '', folder: '', note: '', timing: '', auto: '', needs: [{ id: uid(), name: '', dept: '', person: '', rd: false }] };
  else { const t = store.find('closing', id); ui.draft = JSON.parse(JSON.stringify({ needs: [], auto: '', folder: '', note: '', ...t })); ui.draft.needs = ui.draft.needs.map(needOf); }
  ui.editing = id;
  render();
  $('#bdList .bd-editing input')?.focus();
}
function save() {
  readDraft();
  const d = ui.draft;
  if (!d.name.trim()) { alert('업무명을 입력해 주세요.'); return; }
  const needs = d.needs.filter((n) => n.name.trim() || n.dept.trim() || n.person.trim()).map((n) => ({ id: n.id || uid(), name: n.name.trim(), dept: n.dept.trim(), person: n.person.trim(), ...(n.rd ? { rd: true } : {}) }));
  const base = { name: d.name.trim(), file: d.file.trim(), folder: (d.folder || '').trim(), note: (d.note || '').trim(), timing: d.timing.trim(), auto: d.auto || '', needs };
  if (ui.editing === 'new') store.upsert('closing', { ...base, order: (list().at(-1)?.order ?? 0) + 1, marks: {} });
  else store.upsert('closing', { id: ui.editing, ...base });
  ui.editing = null; ui.draft = null;
  render();
}

/* ---------- 목록 파일(JSON) 불러오기·저장: 이름 등은 공개 코드에 넣지 않고 내 계정에만 ---------- */
const FILE_TAG = 'jywork-closing-list';
function downloadBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
function exportList() {
  const tasks = list().map((t) => ({ name: t.name, file: t.file || '', folder: t.folder || '', note: t.note || '', timing: t.timing || '', auto: t.auto || '', needs: (t.needs || []).map(needOf).map((n) => ({ name: n.name, dept: n.dept, person: n.person, ...(n.rd ? { rd: true } : {}) })) }));
  downloadBlob(new Blob([JSON.stringify({ app: FILE_TAG, version: 2, tasks }, null, 1)], { type: 'application/json' }), `closing-list-${todayStr()}.json`);
}
async function importList(file) {
  let j;
  try { j = JSON.parse(await file.text()); } catch { throw new Error('목록 파일을 읽지 못했습니다 (JSON 형식이 아님).'); }
  if (!j || j.app !== FILE_TAG || !Array.isArray(j.tasks)) throw new Error('월마감 목록 파일이 아닙니다.');
  const have = list();
  let added = 0, updated = 0;
  let order = have.at(-1)?.order ?? 0;
  const mkNeed = (n) => ({ id: uid(), name: n.name || '', dept: n.dept, person: n.person, ...(n.rd ? { rd: true } : {}) });
  for (const t of j.tasks) {
    if (!t || !t.name) continue;
    const needs = (t.needs || []).filter((n) => n && (n.name || n.producer || n.dept || n.person)).map(needOf);
    const ex = have.find((x) => x.name === t.name);
    if (ex) {
      const cur = (ex.needs || []).map(needOf);
      let ch = false;
      for (const n of needs) {
        const m = cur.find((x) => x.name === n.name);
        if (!m) { cur.push(mkNeed(n)); ch = true; }
        else {
          if (!m.dept && n.dept) { m.dept = n.dept; ch = true; }
          if (!m.person && n.person) { m.person = n.person; ch = true; }
          if (!m.rd && n.rd) { m.rd = true; ch = true; }
        }
      }
      const patch = { id: ex.id };
      if (ch) patch.needs = cur;
      for (const k of ['file', 'folder', 'note', 'timing', 'auto']) if (!ex[k] && t[k]) patch[k] = t[k];
      if (Object.keys(patch).length > 1) { store.upsert('closing', patch); updated++; }
    } else {
      store.upsert('closing', { name: t.name, file: t.file || '', folder: t.folder || '', note: t.note || '', timing: t.timing || '', auto: t.auto || '', needs: needs.map(mkNeed), order: ++order, marks: {} });
      added++;
    }
  }
  return { added, updated };
}

/* ---------- 기본 목록 갱신 (이미 목록이 있는 기기에 한 번): 사용자가 손대지 않은 값만 새 값으로 ---------- */
function upgradeSeed() {
  const seedIds = CLOSING_SEED.map((s) => s.id);
  for (const s of CLOSING_SEED) {
    const t = store.find('closing', s.id);
    if (!t) continue;
    const old = SEED_OLD[s.id] || {};
    const patch = {};
    if (s.file && (!t.file || t.file === old.file)) patch.file = s.file;
    if (old.name && t.name === old.name) patch.name = s.name;
    if (s.folder && !t.folder) patch.folder = s.folder;
    if (s.note && !t.note) patch.note = s.note;
    if (s.auto && !t.auto) patch.auto = s.auto;
    const cur = (t.needs || []).map(needOf);
    let ch = false;
    for (const n of s.needs) {
      const m = cur.find((x) => x.id === n.id);
      if (m) {
        if (NEED_OLD[n.id] !== undefined && m.name === NEED_OLD[n.id] && m.name !== n.name) {
          m.name = n.name;
          if (n.id === 'seed08n01') { m.dept = ''; m.person = ''; }  // 예전엔 담당자 3명이 한 칸에 있었음
          if (!m.dept && n.dept) m.dept = n.dept;
          if (!m.person && n.person) m.person = n.person;
          ch = true;
        }
        if (n.rd && !m.rd && m.name === n.name) { m.rd = true; ch = true; }
      } else if (NEED_OLD[n.id] === undefined && !cur.some((x) => x.name === n.name)) {
        cur.push({ ...n }); ch = true;  // 새로 정리한 필요파일 (예전 목록에서 사용자가 지운 것은 되살리지 않음)
      }
    }
    if (ch) {
      const rank = (x) => { const k = s.needs.findIndex((n) => n.id === x.id); return k < 0 ? 999 : k; };
      patch.needs = cur.map((x, k) => [x, k]).sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1]).map(([x]) => { const { producer, ...rest } = x; return rest; });
    }
    if (Object.keys(patch).length) store.upsert('closing', { id: t.id, ...patch });
  }
  // 새로 추가된 업무 (지운 업무는 되살리지 않음 — store.seed가 삭제 기록 확인)
  const have = new Set((store.get().closing || []).map((t) => t.id));
  let order = list().at(-1)?.order ?? 0;
  const fresh = CLOSING_SEED.filter((s) => !have.has(s.id) && +s.id.slice(4) >= 26).map((s) => ({ ...s, order: ++order, marks: {} }));
  if (fresh.length) store.seed('closing', fresh);
  return seedIds;
}

/* ---------- 📄 양식 등록 ---------- */
async function loadForms() {
  try { ui.forms = await formMetaAll(); } catch { ui.forms = {}; }
  if (!ui.editing) render();
}
async function putForm(taskId, file) {
  try {
    ui.forms[taskId] = await formPut(taskId, file);
  } catch (e) { alert(`양식을 저장하지 못했습니다: ${e.message || e}`); return false; }
  return true;
}
async function downloadForm(taskId) {
  const f = await formGetFile(taskId).catch(() => null);
  if (!f) { alert('이 기기에 저장된 양식이 없습니다.'); return; }
  downloadBlob(f, f.name);
}
async function bulkForms(files) {
  const msg = $('#bdLoadMsg');
  msg.textContent = `파일 ${files.length}개 확인 중…`;
  const { matched, unmatched } = matchFiles(list().filter((t) => t.file), files);
  const ok = [];
  for (const { t, f } of matched) if (await putForm(t.id, f)) ok.push(`${t.name} ← ${f.name}`);
  render();
  msg.innerHTML = `📄 양식 ${ok.length}개 등록${ok.length ? `<details><summary>등록한 것 보기</summary><ul class="bd-msglist">${ok.map((x) => `<li>${esc(x)}</li>`).join('')}</ul></details>` : ''}${unmatched.length ? `<details><summary>업무와 맞지 않은 파일 ${unmatched.length}개 (로우데이터·받은 파일 등)</summary><ul class="bd-msglist">${unmatched.map((f) => `<li>${esc(f.name)}</li>`).join('')}</ul></details>` : ''}`;
}

/* ---------- 📁 폴더에서 자동 등록: 업무마다 저장 위치(결산월로 채움)를 열어 규칙에 맞는 완성본을 양식으로 ---------- */
async function folderLabel() {
  const b = $('#bdAutoReg');
  if (!b) return;
  if (!folderOk()) { b.title = 'PC 크롬·엣지에서만 됩니다 (탭은 📂 버튼으로 직접 고르기)'; return; }
  const h = await getWorkRoot();
  $('#bdRootName').innerHTML = h ? `연결: <b>${esc(h.name)}</b> <button class="linkbtn" id="bdRootChange">바꾸기</button>` : '';
  const c = $('#bdRootChange');
  if (c) c.onclick = () => autoFromFolder(true);
}
async function autoFromFolder(change) {
  const msg = $('#bdLoadMsg');
  if (!folderOk()) { msg.textContent = '⚠ 폴더 자동 등록은 PC의 크롬·엣지에서만 됩니다. 탭에서는 📂 버튼으로 파일을 직접 골라 주세요.'; return; }
  let root = change ? null : await getWorkRoot();
  try {
    if (!root) {
      if (!change && !confirm("처음 한 번 '업무' 폴더를 골라 주세요.\n(바탕화면\\업무 — 업무마다 적힌 저장 위치를 이 폴더 기준으로 찾습니다. 읽기만 하고 파일을 바꾸지 않습니다)")) return;
      root = await pickWorkRoot();
      await folderLabel();
    }
    if (!(await workPerm(root, true))) { msg.textContent = '⚠ 폴더 읽기를 허용해야 찾을 수 있습니다.'; return; }
  } catch (e) { if (e?.name !== 'AbortError') msg.textContent = `⚠ 폴더를 열지 못했습니다: ${e.message || e}`; return; }

  const close = closeOf(ui.month);
  const tasks = list().filter((t) => t.file && t.folder);
  const groups = new Map();
  for (const t of tasks) { const p = fillName(t.folder, close); if (!groups.has(p)) groups.set(p, []); groups.get(p).push(t); }
  const ok = [], same = [], miss = [], noDir = [];
  let k = 0;
  for (const [path, ts] of groups) {
    msg.textContent = `📁 폴더 확인 중… (${++k}/${groups.size}) ${path}`;
    const dir = await openPath(root, path).catch(() => null);
    if (!dir) { noDir.push(`${path} (${ts.map((t) => t.name).join(', ')})`); continue; }
    const files = [];
    for (const { h } of await listFiles(dir)) { try { files.push(await h.getFile()); } catch { /* 잠긴 파일 */ } }
    const { matched } = matchFiles(ts, files, close);
    for (const t of ts) {
      const m = matched.find((x) => x.t.id === t.id);
      if (!m) { miss.push(`${t.name} — ${fillName(t.file, close)}`); continue; }
      const cur = ui.forms[t.id];
      if (cur && cur.name === m.f.name && cur.lastModified === m.f.lastModified) { same.push(t.name); continue; }
      if (await putForm(t.id, m.f)) ok.push(`${t.name} ← ${m.f.name}${m.hit ? '' : ' (이번 결산월 파일이 없어 가장 최근 파일)'}`);
    }
  }
  render();
  const ul = (a) => `<ul class="bd-msglist">${a.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>`;
  msg.innerHTML = `📁 ${monthLabel(close)} 마감 폴더에서 양식 <b>${ok.length}</b>개 등록${same.length ? ` · 이미 최신 ${same.length}개` : ''}`
    + (ok.length ? `<details><summary>등록한 것</summary>${ul(ok)}</details>` : '')
    + (miss.length ? `<details><summary>폴더에 파일이 없는 업무 ${miss.length}개</summary>${ul(miss)}</details>` : '')
    + (noDir.length ? `<details><summary>찾지 못한 폴더 ${noDir.length}개</summary>${ul(noDir)}</details>` : '');
}

/* ---------- 📊 파일 생성 팝업: 백업자료 생성 화면을 그대로 띄움 ---------- */
let mkTask = null;
async function openMake(taskId) {
  const t = store.find('closing', taskId);
  if (!t || !t.auto || !ui.openAuto) return;
  mkTask = t.id;
  const dlg = $('#dlgMake');
  const body = $('#mkBody');
  const view = $('#view-auto');
  // 백업자료 생성 화면의 선택칸·첨부칸·결과를 팝업으로 옮김 (닫으면 제자리로)
  [...view.querySelectorAll(':scope > .auto-pick, :scope > .auto-panel')].forEach((el) => body.appendChild(el));
  ui.openAuto(t.auto);
  const idx = list().findIndex((x) => x.id === t.id);
  $('#mkTitle').textContent = `📊 파일 생성 · ${AUTO[t.auto] || ''}`;
  $('#mkSub').textContent = `${idx + 1}. ${t.name} — ${monthLabel(ui.month)}`;
  $('#mkDone').hidden = true;
  if (!dlg.open) dlg.showModal();
  // 등록한 양식(지난달 완성본)이 있으면 ① 지난달 파일 칸에 바로 넣음
  if (ui.forms[t.id] && isXlsx(ui.forms[t.id].name) && ui.setPrev) {
    const f = await formGetFile(t.id).catch(() => null);
    if (f && mkTask === t.id) ui.setPrev(t.auto, f);
  }
}
function closeMake() {
  const dlg = $('#dlgMake');
  const view = $('#view-auto');
  [...$('#mkBody').children].forEach((el) => view.appendChild(el));
  mkTask = null;
  if (dlg.open) dlg.close();
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

export function initBoard({ openAuto, setPrev } = {}) {
  if (!$('#view-board')) return;
  ui.openAuto = openAuto;
  ui.setPrev = setPrev;
  // 처음 쓰는 기기: 기본 목록을 넣어 둠 (한 번만)
  try {
    if (!localStorage.getItem('jywork-closing-seeded')) {
      if (!(store.get().closing || []).length) {
        store.seed('closing', CLOSING_SEED.map((t) => ({ ...t, marks: {} })));
        localStorage.setItem('jywork-closing-auto2', '1');
        localStorage.setItem('jywork-closing-v3', '1');
      }
      localStorage.setItem('jywork-closing-seeded', '1');
    }
    // 새로 자동화된 자료: 이미 목록이 있는 기기에서도 '바로 만들기' 연결 (연결이 비어 있을 때만, 한 번)
    if (!localStorage.getItem('jywork-closing-auto2')) {
      for (const [id, k] of [['seed08', 'energy'], ['seed09', 'rawmat']]) {
        const t = store.find('closing', id);
        if (t && !t.auto) store.upsert('closing', { ...t, auto: k });
      }
      localStorage.setItem('jywork-closing-auto2', '1');
    }
    // 2026-11-19: 파일명 규칙·저장 위치·로우데이터 정리 + 새 업무 (한 번)
    if (!localStorage.getItem('jywork-closing-v3')) {
      upgradeSeed();
      localStorage.setItem('jywork-closing-v3', '1');
    }
    // 1.19.1: Sales estimate 완성본 위치를 부속자료 폴더로 (손대지 않았을 때만)
    if (!localStorage.getItem('jywork-closing-v4')) {
      const t = store.find('closing', 'seed16');
      const s = CLOSING_SEED.find((x) => x.id === 'seed16');
      if (t && t.folder === '6.전략기획\\영업추정') store.upsert('closing', { id: t.id, folder: s.folder, note: s.note });
      localStorage.setItem('jywork-closing-v4', '1');
    }
  } catch { /* 저장 불가 */ }
  $('#bdPrev').onclick = () => shiftMonth(-1);
  $('#bdNext').onclick = () => shiftMonth(1);
  $('#bdAdd').onclick = () => startEdit('new');
  $('#bdSave').onclick = exportList;
  $('#bdLoad').onclick = () => $('#bdLoadIn').click();
  $('#bdLoadIn').onchange = async (e) => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    try {
      const r = await importList(f);
      $('#bdLoadMsg').textContent = `불러오기 완료: 새 업무 ${r.added}개${r.updated ? `, 기존 업무 ${r.updated}개에 빠진 내용 추가` : ''} (이미 있는 내용은 그대로 둠)`;
    } catch (err) { $('#bdLoadMsg').textContent = '⚠ ' + (err.message || err); }
  };
  // 📄 양식: 한 업무 / 여러 파일 한꺼번에
  $('#bdFormIn').onchange = async (e) => {
    const f = e.target.files[0]; e.target.value = '';
    const id = ui.formFor; ui.formFor = null;
    if (!f || !id) return;
    readDraft();
    if (await putForm(id, f)) render();
  };
  $('#bdForms').onclick = () => $('#bdFormsIn').click();
  $('#bdAutoReg').onclick = () => autoFromFolder(false);
  folderLabel();
  $('#bdFormsIn').onchange = async (e) => {
    const fs = [...e.target.files]; e.target.value = '';
    if (fs.length) await bulkForms(fs);
  };
  $('#mkClose').onclick = closeMake;
  $('#dlgMake').addEventListener('close', () => { if ($('#mkBody').children.length) closeMake(); });
  $('#dlgMake').addEventListener('cancel', (e) => { e.preventDefault(); closeMake(); });
  window.addEventListener('backup:made', (e) => {
    if (!mkTask || !$('#dlgMake').open) return;
    $('#mkDoneMsg').innerHTML = `✅ <b>${esc(e.detail?.fileName || '')}</b> 완성 — 내려받은 뒤 완료로 표시하세요`;
    $('#mkDone').hidden = false;
  });
  $('#mkMarkDone').onclick = () => {
    const t = mkTask && store.find('closing', mkTask);
    if (t) setMark(t, (m) => { m.done = todayStr(); });
    $('#mkMarkDone').textContent = '✔ 완료로 표시했습니다';
    setTimeout(() => { $('#mkMarkDone').textContent = '✔ 이 업무 완료로 표시'; closeMake(); }, 700);
  };
  const box = $('#bdList');
  box.addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.done) {
      const item = store.find('closing', t.dataset.done);
      setMark(item, (m) => { m.done = t.checked ? todayStr() : ''; });
    } else if (t.dataset.recv) {
      const [tid, nid] = t.dataset.recv.split('|');
      setMark(store.find('closing', tid), (m) => {
        if (t.checked) m.recv[nid] = todayStr(); else delete m.recv[nid];
        // RD 보관함이 자동 체크한 줄을 직접 끄면 다시 자동으로 켜지 않게 표시, 직접 켜면 자동 표시 없앰
        if (m.auto?.[nid]) { m.auto = { ...m.auto }; if (t.checked) delete m.auto[nid]; else m.auto[nid] = 'off'; }
      });
    }
  });
  box.addEventListener('click', async (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.edit) startEdit(b.dataset.edit);
    else if (b.dataset.del) {
      const t = store.find('closing', b.dataset.del);
      if (confirm(`'${t.name}' 업무를 삭제할까요? 필요파일 목록과 체크 기록도 함께 지워집니다.`)) {
        store.remove('closing', t.id);
        if (ui.forms[t.id]) { delete ui.forms[t.id]; formDel(t.id).catch(() => {}); }
      }
    } else if (b.dataset.fup) { ui.formFor = b.dataset.fup; $('#bdFormIn').click(); }
    else if (b.dataset.fdl) downloadForm(b.dataset.fdl);
    else if (b.dataset.fdel) {
      if (!confirm('이 업무에 등록한 양식 파일을 이 기기에서 뺄까요?')) return;
      readDraft();
      await formDel(b.dataset.fdel).catch(() => {});
      delete ui.forms[b.dataset.fdel];
      render();
    } else if (b.hasAttribute('data-save')) save();
    else if (b.hasAttribute('data-cancel')) { ui.editing = null; ui.draft = null; render(); }
    else if (b.hasAttribute('data-nadd')) { readDraft(); ui.draft.needs.push({ id: uid(), name: '', dept: '', person: '', rd: false }); render(); $$last('#bdList .bd-nrow input')?.focus(); }
    else if (b.dataset.ndel !== undefined) { readDraft(); ui.draft.needs.splice(+b.dataset.ndel, 1); render(); }
    else if (b.dataset.nup !== undefined) { readDraft(); const k = +b.dataset.nup; if (k > 0) { const n = ui.draft.needs; [n[k - 1], n[k]] = [n[k], n[k - 1]]; } render(); }
    else if (b.dataset.tmove) {
      readDraft();
      const items = list(); const i = items.findIndex((t) => t.id === ui.editing); const j = i + +b.dataset.tmove;
      if (i >= 0 && j >= 0 && j < items.length) { reorder(ui.editing, items[j].id, +b.dataset.tmove > 0); render(); }
    } else if (b.dataset.make) openMake(b.dataset.make);
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
  loadForms();
}
const $$last = (s) => { const a = document.querySelectorAll(s); return a[a.length - 1]; };

// 🏠 홈 요약: 이번 달 진행률 + 바로 작업 가능한 업무 / 파일 기다리는 업무
export function homeClosingHTML() {
  const month = ym(new Date());
  const items = list();
  if (!items.length) return '<div class="empty">월마감 업무가 없습니다.</div>';
  let done = 0, miss = 0;
  const ready = [], wait = [];
  for (const t of items) {
    const mk = (t.marks && t.marks[month]) || { done: '', recv: {} };
    const needs = t.needs || [];
    const left = needs.filter((n) => !mk.recv?.[n.id]);
    miss += left.length;
    if (mk.done) { done++; continue; }
    if (needs.length && !left.length) ready.push(t);
    else if (left.length) wait.push({ t, left });
  }
  const pct = Math.round((done / items.length) * 100);
  const row = (cls, label, t, sub) => `<div class="hc-row" data-go-view="board"><span class="st ${cls}">${label}</span><span class="hc-name">${esc(t.name)}</span><span class="muted sm">${esc(sub)}</span></div>`;
  return `<div class="bd-prog"><i style="width:${pct}%"></i></div>
    <p class="muted sm" style="margin:6px 0 8px">${monthLabel(month)} · ${items.length}개 중 <b>${done}</b>개 완료 · 못 받은 필요파일 ${miss}개</p>
    ${ready.slice(0, 3).map((t) => row('ok', '작업 가능', t, '필요파일 모두 받음')).join('')}
    ${wait.slice(0, Math.max(0, 4 - Math.min(ready.length, 3))).map(({ t, left }) => row('wn', '대기', t, `${fillName(left[0].name, closeOf(month))}${left.length > 1 ? ` 외 ${left.length - 1}` : ''} 미수취`)).join('')}`;
}
