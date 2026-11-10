import { store, settings, todayStr } from './store.js';
import { localSummarize } from './summarize.js';
import { aiSummarize, aiAnswer, aiEnabled } from './ai.js';
import { exportExcel, importExcel } from './excel.js';

/* ---------- 공통 도우미 ---------- */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const parseD = (s) => new Date(s + 'T00:00:00');
const fmt = (s) => {
  if (!s) return '';
  const d = parseD(s);
  return `${d.getMonth() + 1}/${d.getDate()} (${DOW[d.getDay()]})`;
};
const addDays = (s, n) => { const d = parseD(s); d.setDate(d.getDate() + n); return todayStr(d); };
const weekRange = (offset = 0) => {
  const t = new Date();
  const mon = new Date(t); mon.setDate(t.getDate() - ((t.getDay() + 6) % 7) + offset * 7);
  const from = todayStr(mon);
  return [from, addDays(from, 6)];
};
const inRange = (d, from, to) => !!d && (!from || d >= from) && (!to || d <= to);
const splitList = (s) => String(s || '').split(/[,，]/).map((x) => x.trim()).filter(Boolean);
const lines = (s) => String(s || '').split('\n').map((x) => x.trim()).filter(Boolean);

function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => t.classList.remove('show'), 2200);
}

function highlight(text, q) {
  const safe = esc(text);
  if (!q) return safe;
  const words = q.split(/\s+/).filter(Boolean).map((w) => esc(w).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  if (!words.length) return safe;
  return safe.replace(new RegExp(`(${words.join('|')})`, 'gi'), '<mark>$1</mark>');
}

/* ---------- 화면 전환 ---------- */
function show(view) {
  $$('.nav button').forEach((b) => b.classList.toggle('active', b.dataset.view === view));
  $$('.view').forEach((v) => v.classList.toggle('active', v.id === 'view-' + view));
  location.hash = view;
  window.scrollTo(0, 0);
}
$$('.nav button').forEach((b) => b.addEventListener('click', () => show(b.dataset.view)));

/* ---------- 항목 렌더링 ---------- */
function eventItem(e, q) {
  const time = e.start ? `${e.start}${e.end ? '~' + e.end : ''}` : '종일';
  return `<div class="item ${e.done ? 'done' : ''}" data-open="event" data-id="${e.id}">
    <div class="when">${fmt(e.date)}<br>${esc(time)}</div>
    <div class="body"><div class="title">${highlight(e.title, q)}</div>
      <div class="sub">${[e.category && `<span class="tag gray">${esc(e.category)}</span>`, e.place && '📍' + highlight(e.place, q), e.attendees && '👥 ' + highlight(e.attendees, q)].filter(Boolean).join(' ')}</div>
      ${e.note ? `<div class="sub">${highlight(e.note, q)}</div>` : ''}
    </div></div>`;
}

function meetingItem(m, q) {
  const tags = splitList(m.tags).map((t) => `<span class="tag">${esc(t)}</span>`).join('');
  const sum = m.summary || (m.content || '').slice(0, 140);
  return `<div class="item" data-open="meetingView" data-id="${m.id}">
    <div class="when">${fmt(m.date)}</div>
    <div class="body"><div class="title">${highlight(m.title, q)}</div>
      <div class="sub">${highlight(sum, q)}</div>
      <div>${tags}${m.decisions?.length ? `<span class="tag green">결정 ${m.decisions.length}</span>` : ''}${m.actions?.length ? `<span class="tag red">할일 ${m.actions.length}</span>` : ''}${m.summary ? '' : '<span class="tag gray">요약 전</span>'}</div>
    </div></div>`;
}

function taskItem(t, q) {
  const today = todayStr();
  const over = !t.done && t.due && t.due < today;
  return `<div class="item ${t.done ? 'done' : ''}" data-open="task" data-id="${t.id}">
    <input type="checkbox" data-toggle="${t.id}" ${t.done ? 'checked' : ''} aria-label="완료" />
    <div class="body"><div class="title">${highlight(t.text, q)}</div>
      <div class="sub">${t.owner ? '👤 ' + highlight(t.owner, q) + ' ' : ''}${t.due ? `<span class="tag ${over ? 'red' : 'gray'}">${over ? '기한 지남 · ' : '~'}${fmt(t.due)}</span>` : ''}${t.source ? ` <span class="muted">· ${esc(t.source)}</span>` : ''}</div>
    </div></div>`;
}

const listHTML = (arr, fn, q, emptyMsg = '항목이 없습니다.') => (arr.length ? arr.map((x) => fn(x, q)).join('') : `<div class="empty">${emptyMsg}</div>`);
const byDate = (a, b) => `${a.date}${a.start || ''}`.localeCompare(`${b.date}${b.start || ''}`);
const byDue = (a, b) => (a.due || '9999').localeCompare(b.due || '9999');

/* ---------- 홈 ---------- */
function renderDashboard() {
  const { events, meetings, tasks } = store.get();
  const today = todayStr();
  const d = new Date();
  $('#todayLabel').textContent = `${d.getFullYear()}년 ${d.getMonth() + 1}월 ${d.getDate()}일 ${DOW[d.getDay()]}요일`;
  const [wf, wt] = weekRange();
  const todayEv = events.filter((e) => e.date === today);
  const weekEv = events.filter((e) => inRange(e.date, today, wt) || (e.date === today)).sort(byDate);
  const open = tasks.filter((t) => !t.done);
  const overdue = open.filter((t) => t.due && t.due < today);
  $('#stats').innerHTML = `
    <div class="stat" data-go="calendar"><b>${todayEv.length}</b><span>오늘 일정</span></div>
    <div class="stat" data-preset="week"><b>${events.filter((e) => inRange(e.date, wf, wt)).length}</b><span>이번 주 일정</span></div>
    <div class="stat" data-go="tasks"><b>${open.length}</b><span>미완료 할일</span></div>
    <div class="stat ${overdue.length ? 'warn' : ''}" data-preset="overdue"><b>${overdue.length}</b><span>기한 지난 할일</span></div>`;
  $('#dashEvents').innerHTML = listHTML(weekEv.slice(0, 8), eventItem, '', '이번 주 남은 일정이 없습니다.');
  $('#dashTasks').innerHTML = listHTML(open.sort(byDue).slice(0, 8), taskItem, '', '미완료 할일이 없습니다 🎉');
  $('#dashMeetings').innerHTML = listHTML([...meetings].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 5), meetingItem, '', '아직 회의메모가 없습니다. “+ 회의메모”로 시작해보세요.');
}

/* ---------- 일정(달력) ---------- */
let calMonth = new Date(); calMonth.setDate(1);
let selDate = todayStr();

function renderCalendar() {
  const { events, meetings } = store.get();
  const y = calMonth.getFullYear(), m = calMonth.getMonth();
  $('#calTitle').textContent = `${y}년 ${m + 1}월`;
  const start = new Date(y, m, 1 - new Date(y, m, 1).getDay());
  const today = todayStr();
  let html = DOW.map((d) => `<div class="dow">${d}</div>`).join('');
  const cells = Math.ceil((new Date(y, m, 1).getDay() + new Date(y, m + 1, 0).getDate()) / 7) * 7;
  for (let i = 0; i < cells; i++) {
    const day = new Date(start); day.setDate(start.getDate() + i);
    const ds = todayStr(day);
    const evs = events.filter((e) => e.date === ds).sort(byDate);
    const ms = meetings.filter((x) => x.date === ds && !evs.some((e) => e.id === x.eventId));
    const items = [...evs.map((e) => `<div class="ev">${esc(e.start ? e.start + ' ' : '')}${esc(e.title)}</div>`), ...ms.map((x) => `<div class="ev m">📝 ${esc(x.title)}</div>`)];
    html += `<div class="d ${day.getMonth() !== m ? 'other' : ''} ${ds === today ? 'today' : ''} ${ds === selDate ? 'sel' : ''} ${day.getDay() === 0 ? 'sun' : ''}" data-date="${ds}">
      <span class="n">${day.getDate()}</span>${items.slice(0, 3).join('')}${items.length > 3 ? `<div class="more">+${items.length - 3}</div>` : ''}</div>`;
  }
  $('#cal').innerHTML = html;
  $('#dayTitle').innerHTML = `${fmt(selDate)} 일정 <button class="btn sm" data-new="event" data-date="${selDate}">+ 추가</button>`;
  const dayEv = events.filter((e) => e.date === selDate).sort(byDate);
  const dayMt = meetings.filter((x) => x.date === selDate);
  $('#dayEvents').innerHTML = listHTML(dayEv, eventItem, '', '이 날은 일정이 없습니다.') +
    (dayMt.length ? `<div class="group-title">이 날의 회의메모</div>${dayMt.map((x) => meetingItem(x)).join('')}` : '');
}
$('#cal').addEventListener('click', (e) => {
  const d = e.target.closest('[data-date]');
  if (!d) return;
  if (selDate === d.dataset.date) openEvent(null, d.dataset.date);
  selDate = d.dataset.date;
  renderCalendar();
});
$('#calPrev').onclick = () => { calMonth.setMonth(calMonth.getMonth() - 1); renderCalendar(); };
$('#calNext').onclick = () => { calMonth.setMonth(calMonth.getMonth() + 1); renderCalendar(); };
$('#calToday').onclick = () => { calMonth = new Date(); calMonth.setDate(1); selDate = todayStr(); renderCalendar(); };

/* ---------- 회의메모 목록 ---------- */
const meetingText = (m) => [m.title, m.attendees, m.tags, m.content, m.summary, (m.decisions || []).join(' '), (m.keywords || []).join(' ')].join(' ').toLowerCase();
function renderMeetings() {
  const q = $('#meetingFilter').value.trim();
  const words = q.toLowerCase().split(/\s+/).filter(Boolean);
  const list = store.get().meetings.filter((m) => words.every((w) => meetingText(m).includes(w))).sort((a, b) => b.date.localeCompare(a.date));
  $('#meetingList').innerHTML = listHTML(list, meetingItem, q, q ? '검색 결과가 없습니다.' : '아직 회의메모가 없습니다.');
}
$('#meetingFilter').addEventListener('input', renderMeetings);

/* ---------- 할일 ---------- */
function renderTasks() {
  const f = $('#taskFilter').value;
  const owner = $('#taskOwner').value.trim();
  const today = todayStr();
  let list = store.get().tasks;
  if (f === 'open') list = list.filter((t) => !t.done);
  if (f === 'done') list = list.filter((t) => t.done);
  if (f === 'overdue') list = list.filter((t) => !t.done && t.due && t.due < today);
  if (owner) list = list.filter((t) => (t.owner || '').includes(owner));
  $('#taskList').innerHTML = listHTML([...list].sort(byDue), taskItem, owner);
}
$('#taskFilter').addEventListener('change', renderTasks);
$('#taskOwner').addEventListener('input', renderTasks);

/* ---------- 조회 · 리포트 ---------- */
let lastResult = { events: [], meetings: [], tasks: [] };

function itemText(kind, x) {
  if (kind === 'events') return [x.title, x.place, x.attendees, x.note, x.category].join(' ');
  if (kind === 'meetings') return meetingText(x);
  return [x.text, x.owner, x.source].join(' ');
}

function query({ q = '', kind = 'all', from = '', to = '' }) {
  const data = store.get();
  const words = q.toLowerCase().split(/\s+/).filter(Boolean);
  const out = { events: [], meetings: [], tasks: [] };
  for (const k of Object.keys(out)) {
    if (kind !== 'all' && kind !== k) continue;
    out[k] = data[k].filter((x) => {
      const d = k === 'tasks' ? x.due || (x.createdAt || '').slice(0, 10) : x.date;
      if ((from || to) && !inRange(d, from, to)) return false;
      const t = itemText(k, x).toLowerCase();
      return words.every((w) => t.includes(w));
    });
  }
  out.events.sort(byDate); out.meetings.sort(byDate); out.tasks.sort(byDue);
  return out;
}

function renderResults(title, res, q = '', opts = {}) {
  lastResult = res;
  const n = res.events.length + res.meetings.length + res.tasks.length;
  $('#resultTitle').textContent = title;
  $('#resultCount').textContent = `${n}건`;
  let html = '';
  if (res.events.length) html += `<div class="group-title">일정 (${res.events.length})</div>${res.events.map((x) => eventItem(x, q)).join('')}`;
  if (res.meetings.length) {
    html += `<div class="group-title">회의메모 (${res.meetings.length})</div>`;
    html += opts.decisions
      ? res.meetings.map((m) => `<div class="item" data-open="meetingView" data-id="${m.id}"><div class="when">${fmt(m.date)}</div><div class="body"><div class="title">${esc(m.title)}</div><ul class="sub" style="display:block">${m.decisions.map((d) => `<li>${esc(d)}</li>`).join('')}</ul></div></div>`).join('')
      : res.meetings.map((x) => meetingItem(x, q)).join('');
  }
  if (res.tasks.length) html += `<div class="group-title">할일 (${res.tasks.length})</div>${res.tasks.map((x) => taskItem(x, q)).join('')}`;
  $('#results').innerHTML = html || '<div class="empty">조건에 맞는 기록이 없습니다.</div>';
}

function runPreset(p) {
  const today = todayStr();
  const d = new Date();
  const monthFrom = todayStr(new Date(d.getFullYear(), d.getMonth(), 1));
  const monthTo = todayStr(new Date(d.getFullYear(), d.getMonth() + 1, 0));
  const none = { events: [], meetings: [], tasks: [] };
  const { tasks, meetings } = store.get();
  $$('#presets .chip').forEach((c) => c.classList.toggle('active', c.dataset.preset === p));
  if (p === 'week') { const [f, t] = weekRange(0); renderResults(`이번 주 일정 (${fmt(f)} ~ ${fmt(t)})`, { ...none, ...query({ kind: 'events', from: f, to: t }) }); }
  if (p === 'nextweek') { const [f, t] = weekRange(1); renderResults(`다음 주 일정 (${fmt(f)} ~ ${fmt(t)})`, { ...none, ...query({ kind: 'events', from: f, to: t }) }); }
  if (p === 'open') renderResults('미완료 할일', { ...none, tasks: tasks.filter((t) => !t.done).sort(byDue) });
  if (p === 'overdue') renderResults('기한 지난 할일', { ...none, tasks: tasks.filter((t) => !t.done && t.due && t.due < today).sort(byDue) });
  if (p === 'monthMeetings') renderResults(`${d.getMonth() + 1}월 회의 요약`, { ...none, meetings: meetings.filter((m) => inRange(m.date, monthFrom, monthTo)).sort(byDate) });
  if (p === 'decisions') renderResults('최근 결정사항 (최근 30일)', { ...none, meetings: meetings.filter((m) => m.decisions?.length && m.date >= addDays(today, -30)).sort((a, b) => b.date.localeCompare(a.date)) }, '', { decisions: true });
}

$('#presets').addEventListener('click', (e) => { const b = e.target.closest('[data-preset]'); if (b) runPreset(b.dataset.preset); });
const runSearch = () => {
  const params = { q: $('#q').value.trim(), kind: $('#qKind').value, from: $('#qFrom').value, to: $('#qTo').value };
  $$('#presets .chip').forEach((c) => c.classList.remove('active'));
  renderResults(params.q ? `“${params.q}” 검색 결과` : '검색 결과', query(params), params.q);
};
$('#qRun').onclick = runSearch;
$('#q').addEventListener('keydown', (e) => { if (e.key === 'Enter') runSearch(); });
$('#qReset').onclick = () => { $('#q').value = ''; $('#qKind').value = 'all'; $('#qFrom').value = ''; $('#qTo').value = ''; runSearch(); };
$('#qExport').onclick = async () => {
  try { await exportExcel(lastResult, `조회결과_${todayStr()}.xlsx`); toast('엑셀로 저장했습니다.'); } catch (e) { alert(e.message); }
};

/* ---------- AI 질문 ---------- */
function mdToHtml(md) {
  const out = [];
  const rows = esc(md).split('\n');
  for (let i = 0; i < rows.length; i++) {
    const l = rows[i];
    if (/^\s*\|.*\|\s*$/.test(l)) {
      const tbl = [];
      while (i < rows.length && /^\s*\|.*\|\s*$/.test(rows[i])) tbl.push(rows[i++]);
      i--;
      const cells = (r) => r.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
      const body = tbl.filter((r) => !/^\s*\|[\s:|-]+\|\s*$/.test(r));
      out.push(`<table>${body.map((r, j) => `<tr>${cells(r).map((c) => (j ? `<td>${c}</td>` : `<th>${c}</th>`)).join('')}</tr>`).join('')}</table>`);
      continue;
    }
    if (/^#{1,4}\s/.test(l)) { out.push(`<b>${l.replace(/^#+\s/, '')}</b><br>`); continue; }
    if (/^\s*[-*]\s/.test(l)) { out.push(`• ${l.replace(/^\s*[-*]\s/, '')}<br>`); continue; }
    out.push(l + '<br>');
  }
  return out.join('').replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
}

function buildContext(question) {
  const data = store.get();
  // 질문 단어와 겹치는 기록을 우선, 이후 최근 기록 순으로 담습니다.
  const words = question.toLowerCase().match(/[가-힣a-z0-9]{2,}/g) || [];
  const score = (t) => words.reduce((s, w) => s + (t.toLowerCase().includes(w.replace(/(은|는|이|가|을|를|에|의|와|과|도|로)$/, '')) ? 1 : 0), 0);
  const recs = [
    ...data.meetings.map((m) => ({ d: m.date, t: `[회의메모] ${m.date} | ${m.title} | 참석: ${m.attendees || '-'} | 태그: ${m.tags || '-'}\n요약: ${m.summary || '-'}\n결정: ${(m.decisions || []).join(' / ') || '-'}\n할일: ${(m.actions || []).map((a) => `${a.text}(${a.owner || '담당 미정'}, ${a.due || '기한 없음'})`).join(' / ') || '-'}\n원문: ${(m.content || '').slice(0, 3000)}` })),
    ...data.events.map((e) => ({ d: e.date, t: `[일정] ${e.date} ${e.start || ''}${e.end ? '~' + e.end : ''} | ${e.title} | ${e.category || ''} | 장소: ${e.place || '-'} | 참석: ${e.attendees || '-'} | ${e.done ? '완료' : ''} ${e.note || ''}` })),
    ...data.tasks.map((t) => ({ d: t.due || '', t: `[할일] ${t.text} | 담당: ${t.owner || '-'} | 마감: ${t.due || '-'} | ${t.done ? '완료' : '미완료'} | 출처: ${t.source || '-'}` })),
  ].map((r) => ({ ...r, s: score(r.t) }))
    .sort((a, b) => b.s - a.s || (b.d || '').localeCompare(a.d || ''));
  let total = 0;
  const picked = [];
  for (const r of recs) {
    if (total + r.t.length > 300000) break;
    picked.push(r.t);
    total += r.t.length;
  }
  return `오늘 날짜: ${todayStr()}\n\n` + picked.join('\n\n');
}

$('#aiAsk').onclick = async () => {
  const q = $('#aiQ').value.trim();
  if (!q) return;
  if (!aiEnabled()) {
    $('#aiOut').innerHTML = '<span class="muted">AI 질문은 설정 탭에서 Claude API 키를 입력하면 사용할 수 있습니다. 키 없이도 위의 빠른 리포트와 상세 검색은 바로 사용할 수 있어요.</span>';
    return;
  }
  const btn = $('#aiAsk');
  btn.disabled = true; btn.textContent = '생각 중…';
  $('#aiOut').innerHTML = '<span class="muted">기록을 살펴보는 중입니다…</span>';
  try {
    const ans = await aiAnswer(q, buildContext(q));
    $('#aiOut').innerHTML = mdToHtml(ans);
  } catch (e) {
    $('#aiOut').innerHTML = `<span style="color:var(--danger)">${esc(e.message)}</span>`;
  } finally {
    btn.disabled = false; btn.textContent = '질문';
  }
};
$('#aiQ').addEventListener('keydown', (e) => { if (e.key === 'Enter') $('#aiAsk').click(); });

/* ---------- 폼: 일정 ---------- */
const dlgEvent = $('#dlgEvent'), fEvent = $('#formEvent');
function fill(form, obj) {
  for (const el of form.elements) {
    if (!el.name) continue;
    if (el.type === 'checkbox') el.checked = !!obj[el.name];
    else el.value = obj[el.name] ?? '';
  }
}
function read(form) {
  const o = {};
  for (const el of form.elements) {
    if (!el.name) continue;
    o[el.name] = el.type === 'checkbox' ? el.checked : el.value.trim();
  }
  return o;
}

function openEvent(id, date) {
  const e = id ? store.find('events', id) : { date: date || selDate || todayStr(), category: '회의' };
  fill(fEvent, e);
  $('[data-del]', fEvent).style.display = id ? '' : 'none';
  dlgEvent.showModal();
}
fEvent.addEventListener('submit', (ev) => {
  if (ev.submitter?.value !== 'save') return;
  const o = read(fEvent);
  if (!o.id) delete o.id;
  store.upsert('events', o);
  toast('일정을 저장했습니다.');
});
$('[data-del]', fEvent).onclick = () => {
  const id = fEvent.elements.id.value;
  if (id && confirm('이 일정을 삭제할까요?')) { store.remove('events', id); dlgEvent.close(); toast('삭제했습니다.'); }
};
$('[data-meeting]', fEvent).onclick = () => {
  const o = read(fEvent);
  if (!o.title) { fEvent.reportValidity(); return; }
  if (!o.id) delete o.id;
  const saved = store.upsert('events', o);
  dlgEvent.close();
  const existing = store.get().meetings.find((m) => m.eventId === saved.id);
  openMeeting(existing?.id, { title: saved.title, date: saved.date, attendees: saved.attendees, tags: saved.category, eventId: saved.id });
};

/* ---------- 폼: 회의메모 ---------- */
const dlgMeeting = $('#dlgMeeting'), fMeeting = $('#formMeeting');
const actionsToText = (a = []) => a.map((x) => [x.text, x.owner, x.due].filter((v, i) => v || i === 0).join(' / ')).join('\n');
const textToActions = (s) => lines(s).map((l) => { const [text, owner = '', due = ''] = l.split(/\s*\/\s*/); return { text, owner, due }; });

function openMeeting(id, preset = {}) {
  const m = id ? store.find('meetings', id) : { date: todayStr(), ...preset };
  fill(fMeeting, { ...m, decisions: (m.decisions || []).join('\n'), actions: actionsToText(m.actions), keywords: (m.keywords || []).join(', ') });
  $('[data-del]', fMeeting).style.display = id ? '' : 'none';
  $('#sumStatus').textContent = m.summarizedBy ? `${m.summarizedBy}으로 요약됨` : '';
  $('#btnAiSum').title = aiEnabled() ? '' : '설정에서 API 키를 넣으면 사용할 수 있습니다';
  dlgMeeting.showModal();
}
function applySummary(r, by) {
  fMeeting.elements.summary.value = r.summary || '';
  fMeeting.elements.decisions.value = (r.decisions || []).join('\n');
  fMeeting.elements.actions.value = actionsToText(r.actions);
  fMeeting.elements.keywords.value = (r.keywords || []).join(', ');
  fMeeting.elements.summarizedBy.value = by;
  $('#sumStatus').textContent = `${by}으로 요약했습니다. 확인 후 수정하고 저장하세요.`;
}
$('#btnLocalSum').onclick = () => {
  const content = fMeeting.elements.content.value;
  if (!content.trim()) { toast('메모 내용을 먼저 입력해주세요.'); return; }
  applySummary(localSummarize(content, fMeeting.elements.date.value), '자동 요약');
};
$('#btnAiSum').onclick = async () => {
  const o = read(fMeeting);
  if (!o.content) { toast('메모 내용을 먼저 입력해주세요.'); return; }
  if (!aiEnabled()) {
    applySummary(localSummarize(o.content, o.date), '자동 요약');
    $('#sumStatus').textContent = 'API 키가 없어 자동 요약을 사용했습니다. (설정에서 키 입력 시 AI 요약)';
    return;
  }
  const btn = $('#btnAiSum');
  btn.disabled = true;
  $('#sumStatus').textContent = 'AI가 요약하는 중…';
  try {
    applySummary(await aiSummarize(o), 'AI 요약');
  } catch (e) {
    $('#sumStatus').textContent = e.message;
  } finally {
    btn.disabled = false;
  }
};
fMeeting.addEventListener('submit', (ev) => {
  if (ev.submitter?.value !== 'save') return;
  const o = read(fMeeting);
  if (!o.id) delete o.id;
  // 요약을 안 했으면 저장할 때 자동 요약을 채워 데이터가 항상 쌓이도록 함
  if (!o.summary && o.content) {
    const r = localSummarize(o.content, o.date);
    Object.assign(o, { summary: r.summary, decisions: r.decisions.join('\n'), actions: actionsToText(r.actions), keywords: r.keywords.join(', '), summarizedBy: '자동 요약' });
  }
  const m = store.upsert('meetings', {
    ...o,
    decisions: lines(o.decisions),
    actions: textToActions(o.actions),
    keywords: splitList(o.keywords),
  });
  store.syncMeetingTasks(m);
  toast(m.actions.length ? `회의메모 저장 · 할일 ${m.actions.length}건 등록` : '회의메모를 저장했습니다.');
});
$('[data-del]', fMeeting).onclick = () => {
  const id = fMeeting.elements.id.value;
  if (id && confirm('이 회의메모와 연결된 할일을 삭제할까요?')) { store.remove('meetings', id); dlgMeeting.close(); toast('삭제했습니다.'); }
};

/* ---------- 회의 상세 ---------- */
const dlgView = $('#dlgView');
let viewId = null;
function summaryText(m) {
  return [`[${m.date}] ${m.title}`, m.attendees && `참석: ${m.attendees}`, '', '■ 요약', m.summary || '-',
    '', '■ 결정사항', ...(m.decisions?.length ? m.decisions.map((d) => '- ' + d) : ['-']),
    '', '■ 할일', ...(m.actions?.length ? m.actions.map((a) => `- ${a.text}${a.owner ? ' (' + a.owner + ')' : ''}${a.due ? ' ~' + a.due : ''}`) : ['-'])].filter((x) => x !== undefined && x !== false).join('\n');
}
function openMeetingView(id) {
  const m = store.find('meetings', id);
  if (!m) return;
  viewId = id;
  $('#viewBody').innerHTML = `<div class="detail">
    <h2>${esc(m.title)}</h2>
    <div class="muted sm">${fmt(m.date)} · ${esc(m.attendees || '참석자 미기재')} ${m.summarizedBy ? `· <span class="tag gray">${esc(m.summarizedBy)}</span>` : ''}</div>
    <div>${splitList(m.tags).map((t) => `<span class="tag">${esc(t)}</span>`).join('')}${(m.keywords || []).map((k) => `<span class="tag gray">#${esc(k)}</span>`).join('')}</div>
    <h3>요약</h3><div>${esc(m.summary || '요약 없음')}</div>
    <h3>결정사항</h3>${m.decisions?.length ? `<ul>${m.decisions.map((d) => `<li>${esc(d)}</li>`).join('')}</ul>` : '<div class="muted">없음</div>'}
    <h3>할일</h3>${m.actions?.length ? `<ul>${m.actions.map((a) => `<li>${esc(a.text)} ${a.owner ? `<span class="tag">${esc(a.owner)}</span>` : ''}${a.due ? `<span class="tag gray">~${fmt(a.due)}</span>` : ''}</li>`).join('')}</ul>` : '<div class="muted">없음</div>'}
    <h3>원문</h3><div class="raw">${esc(m.content || '')}</div></div>`;
  dlgView.showModal();
}
$('#viewEdit').onclick = () => { dlgView.close(); openMeeting(viewId); };
$('#viewCopy').onclick = async () => {
  try { await navigator.clipboard.writeText(summaryText(store.find('meetings', viewId))); toast('요약을 복사했습니다. 메일·메신저에 붙여넣으세요.'); } catch { toast('복사할 수 없습니다.'); }
};

/* ---------- 폼: 할일 ---------- */
const dlgTask = $('#dlgTask'), fTask = $('#formTask');
function openTask(id) {
  const t = id ? store.find('tasks', id) : { due: '' };
  fill(fTask, t);
  $('[data-del]', fTask).style.display = id ? '' : 'none';
  dlgTask.showModal();
}
fTask.addEventListener('submit', (ev) => {
  if (ev.submitter?.value !== 'save') return;
  const o = read(fTask);
  if (!o.id) delete o.id;
  store.upsert('tasks', o);
  toast('할일을 저장했습니다.');
});
$('[data-del]', fTask).onclick = () => {
  const id = fTask.elements.id.value;
  if (id && confirm('이 할일을 삭제할까요?')) { store.remove('tasks', id); dlgTask.close(); }
};

/* ---------- 전역 클릭 처리 ---------- */
document.addEventListener('click', (e) => {
  const tog = e.target.closest('[data-toggle]');
  if (tog) {
    const t = store.find('tasks', tog.dataset.toggle);
    store.upsert('tasks', { id: t.id, done: tog.checked });
    if (tog.checked) toast('완료 처리했습니다 👍');
    return;
  }
  const close = e.target.closest('[data-close]');
  if (close) { close.closest('dialog').close(); return; }
  const nw = e.target.closest('[data-new]');
  if (nw) {
    const k = nw.dataset.new;
    if (k === 'event') openEvent(null, nw.dataset.date);
    if (k === 'meeting') openMeeting(null);
    if (k === 'task') openTask(null);
    return;
  }
  const op = e.target.closest('[data-open]');
  if (op) {
    const { open, id } = op.dataset;
    if (open === 'event') openEvent(id);
    if (open === 'meetingView') openMeetingView(id);
    if (open === 'task') openTask(id);
    return;
  }
  const go = e.target.closest('.stat[data-go]');
  if (go) { show(go.dataset.go); return; }
  const pr = e.target.closest('.stat[data-preset]');
  if (pr) { show('search'); runPreset(pr.dataset.preset); }
});
// 다이얼로그 바깥 클릭 시 닫기
$$('dialog').forEach((d) => d.addEventListener('click', (e) => { if (e.target === d) d.close(); }));

/* ---------- 설정 · 데이터 ---------- */
function renderSettings() {
  const s = settings.get();
  $('#apiKey').value = s.apiKey;
  $('#model').value = s.model;
  const on = aiEnabled();
  $('#aiBadge').textContent = on ? 'AI 사용 가능' : 'API 키 필요';
  $('#aiBadge').classList.toggle('on', on);
}
$('#saveSettings').onclick = () => {
  settings.set({ apiKey: $('#apiKey').value.trim(), model: $('#model').value });
  renderSettings();
  toast('설정을 저장했습니다.');
};
$('#xlsxExport').onclick = async () => {
  try { await exportExcel(store.get(), `업무기록_${todayStr()}.xlsx`); toast('엑셀 파일을 저장했습니다.'); } catch (e) { alert(e.message); }
};
$('#xlsxImport').onchange = async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  try {
    const d = await importExcel(f);
    store.merge(d);
    toast(`가져오기 완료: 일정 ${d.events.length} · 회의 ${d.meetings.length} · 할일 ${d.tasks.length}`);
  } catch (err) { alert('가져오기 실패: ' + err.message); }
  e.target.value = '';
};
$('#jsonExport').onclick = () => {
  const blob = new Blob([JSON.stringify(store.get(), null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `업무노트_백업_${todayStr()}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
};
$('#jsonImport').onchange = async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  try {
    const d = JSON.parse(await f.text());
    if (!d.events || !d.meetings) throw new Error('백업 파일 형식이 아닙니다.');
    if (confirm('현재 데이터를 백업 파일 내용으로 바꿀까요? (취소 시 합치기)')) store.replaceAll(d); else store.merge(d);
    toast('복원했습니다.');
  } catch (err) { alert('복원 실패: ' + err.message); }
  e.target.value = '';
};
$('#wipe').onclick = () => {
  if (confirm('모든 일정·회의메모·할일을 삭제합니다. 먼저 백업하셨나요?') && confirm('정말 삭제할까요? 되돌릴 수 없습니다.')) {
    store.replaceAll({ events: [], meetings: [], tasks: [] });
    toast('모두 삭제했습니다.');
  }
};
$('#loadSample').onclick = () => {
  const t = todayStr();
  const sample = [
    { title: '주간 업무회의', date: addDays(t, -7), attendees: '김팀장, 이대리, 박주임', tags: '주간회의, 프로젝트A',
      content: '프로젝트A 1차 시안 진행 상황 공유.\n디자인 시안은 이번 주 금요일까지 완료 예정.\n예산은 기존 안대로 3천만원으로 확정.\n@이대리 고객사 요구사항 정리해서 ' + addDays(t, -2) + '까지 공유\n@박주임 일정표 업데이트 담당\n다음 회의에서 QA 일정 논의 필요.' },
    { title: '고객사 B 미팅', date: addDays(t, -2), attendees: '김팀장, 고객사 담당자', tags: '고객사B, 영업',
      content: '고객사 B 신규 계약 범위 논의.\n유지보수 포함 1년 계약으로 합의.\n견적서는 @김팀장 이 ' + addDays(t, 3) + '까지 송부.\n추가 기능 요청 2건은 검토 필요.' },
  ];
  sample.forEach((s) => {
    const r = localSummarize(s.content, s.date);
    const m = store.upsert('meetings', { ...s, summary: r.summary, decisions: r.decisions, actions: r.actions, keywords: r.keywords, summarizedBy: '자동 요약' });
    store.syncMeetingTasks(m);
  });
  [
    { title: '팀 주간회의', date: t, start: '10:00', end: '11:00', category: '회의', place: '3층 회의실', attendees: '팀 전체' },
    { title: '월간 보고서 제출', date: addDays(t, 2), category: '마감' },
    { title: '고객사 B 방문', date: addDays(t, 4), start: '14:00', category: '외근', place: '고객사 B 본사' },
  ].forEach((e) => store.upsert('events', e));
  toast('예시 데이터를 넣었습니다.');
};

/* ---------- 시작 ---------- */
function renderAll() {
  renderDashboard();
  renderCalendar();
  renderMeetings();
  renderTasks();
  renderSettings();
}
store.subscribe(renderAll);
renderAll();
runPreset('week');
const startView = location.hash.slice(1);
if (startView && $('#view-' + startView)) show(startView);

if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}
