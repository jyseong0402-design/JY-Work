import { store, settings, todayStr } from './store.js';
import { localSummarize } from './summarize.js';
import { aiSummarize, aiAnswer, aiTranscribe, aiEnabled } from './ai.js';
import { InkPad, renderInkPreview } from './ink.js';
import { inkdb } from './inkdb.js';
import { sync, loadInk, saveInk, flushInk, parseConfig, getConfig } from './sync.js';
import { exportExcel, importExcel } from './excel.js';
import { APP_VERSION, CHANGES, cmpVersion } from './version.js';

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
  const sum = m.summary || (m.content || '').slice(0, 140) || (m.inkPages ? '필기 메모' : '');
  return `<div class="item" data-open="meetingView" data-id="${m.id}">
    <div class="when">${fmt(m.date)}</div>
    <div class="body"><div class="title">${m.series ? `<span class="tag series" style="--c:${seriesColor(m.series)}">📒 ${esc(m.series)}</span> ` : ''}${highlight(m.title, q)}</div>
      <div class="sub">${highlight(sum, q)}</div>
      <div>${tags}${m.decisions?.length ? `<span class="tag green">결정 ${m.decisions.length}</span>` : ''}${m.actions?.length ? `<span class="tag red">할일 ${m.actions.length}</span>` : ''}${m.inkPages ? `<span class="tag yellow">✍️ 필기 ${m.inkPages}쪽</span>` : ''}${m.summary ? '' : '<span class="tag gray">요약 전</span>'}</div>
    </div></div>`;
}

/* 누적 노트(시리즈): 같은 노트 이름의 기록을 한 권처럼 묶어서 날짜별로 쌓음 (정기회의, 프로젝트 등) */
const PALETTE = ['#3f8fb5', '#c9b51a', '#d0663f', '#4f9a62', '#8a63c7', '#c2577a', '#3d7bd1', '#9a8a3a', '#2f9e9e', '#b5653f'];
function seriesColor(name) {
  let h = 2166136261; // FNV-1a: 이름마다 고르게 다른 색
  for (const ch of String(name)) { h ^= ch.charCodeAt(0); h = Math.imul(h, 16777619) >>> 0; }
  return PALETTE[h % PALETTE.length];
}
function seriesList() {
  const map = new Map();
  for (const m of store.get().meetings) {
    if (!m.series) continue;
    const g = map.get(m.series) || { name: m.series, items: [] };
    g.items.push(m);
    map.set(m.series, g);
  }
  const tasks = store.get().tasks;
  return [...map.values()].map((g) => {
    g.items.sort((x, y) => y.date.localeCompare(x.date));
    const ids = new Set(g.items.map((x) => x.id));
    g.open = tasks.filter((t) => !t.done && ids.has(t.meetingId));
    g.last = g.items[0];
    return g;
  }).sort((x, y) => y.last.date.localeCompare(x.last.date));
}
function coverHTML({ color, title, sub, foot, attrs, ink }) {
  return `<div class="nb" ${attrs}>
    <div class="nb-cover" style="--c:${color}">
      <div class="nb-paper paper">${ink ? '<span class="nb-ink">✍️</span>' : ''}</div>
      <div class="nb-band"><b>${esc(title)}</b>${sub ? `<small>${esc(sub)}</small>` : ''}</div>
    </div>
    <div class="nb-foot">${foot}</div>
  </div>`;
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

/* ---------- 회의메모 목록 (노트 표지 / 목록) ---------- */
const meetingText = (m) => [m.series, m.title, m.attendees, m.tags, m.content, m.summary, (m.decisions || []).join(' '), (m.keywords || []).join(' ')].join(' ').toLowerCase();
const pref = (k, d) => { try { return localStorage.getItem(k) || d; } catch { return d; } };
const setPref = (k, v) => { try { localStorage.setItem(k, v); } catch { /* 무시 */ } };
let mFilter = pref('jy-mfilter', 'all');
let mView = pref('jy-mview', 'grid');

function renderMeetings() {
  $$('#mFilter [data-mf]').forEach((b) => b.classList.toggle('on', b.dataset.mf === mFilter));
  $$('#mView [data-mv]').forEach((b) => b.classList.toggle('on', b.dataset.mv === mView));
  const q = $('#meetingFilter').value.trim();
  const words = q.toLowerCase().split(/\s+/).filter(Boolean);
  const all = store.get().meetings;
  const pick = (m) => (mFilter === 'all' || (mFilter === 'series' ? !!m.series : !m.series));
  // 검색 중이거나 목록 보기면 회차까지 펼친 목록
  if (words.length || mView === 'list') {
    const list = all.filter((m) => pick(m) && words.every((w) => meetingText(m).includes(w))).sort((a, b) => b.date.localeCompare(a.date));
    $('#meetingList').className = 'list';
    $('#meetingList').innerHTML = listHTML(list, meetingItem, q, q ? '검색 결과가 없습니다.' : '아직 회의메모가 없습니다.');
    return;
  }
  const cards = [`<button type="button" class="nb nb-new" data-new="${mFilter === 'series' ? 'series' : 'meeting'}"><span>＋</span><small>${mFilter === 'series' ? '누적 노트' : '새 회의메모'}</small></button>`];
  if (mFilter !== 'single') {
    for (const g of seriesList()) {
      cards.push(coverHTML({
        color: seriesColor(g.name), title: g.name, sub: `기록 ${g.items.length}개`,
        attrs: `data-series="${esc(g.name)}"`, ink: g.items.some((x) => x.inkPages),
        foot: `<b>📒 누적 노트</b><span>최근 ${fmt(g.last.date)}${g.open.length ? ` · <em>할일 ${g.open.length}</em>` : ''}</span>`,
      }));
    }
  }
  if (mFilter !== 'series') {
    for (const m of all.filter((x) => !x.series).sort((a, b) => b.date.localeCompare(a.date))) {
      cards.push(coverHTML({
        color: '#7d8597', title: m.title || '제목 없음', sub: '',
        attrs: `data-open="meetingView" data-id="${m.id}"`, ink: !!m.inkPages,
        foot: `<b>${m.summary ? '요약됨' : '낱장 메모'}</b><span>${fmt(m.date)}</span>`,
      }));
    }
  }
  $('#meetingList').className = 'nb-grid';
  $('#meetingList').innerHTML = cards.join('') + (cards.length === 1 ? '<div class="empty nb-empty">정기회의나 계속 이어지는 업무는 “📒 누적 노트”로 만들면 날짜별 기록이 한 권에 계속 쌓입니다.</div>' : '');
}
$('#meetingFilter').addEventListener('input', renderMeetings);
$('#mFilter').addEventListener('click', (e) => { const b = e.target.closest('[data-mf]'); if (b) { mFilter = b.dataset.mf; setPref('jy-mfilter', mFilter); renderMeetings(); } });
$('#mView').addEventListener('click', (e) => { const b = e.target.closest('[data-mv]'); if (b) { mView = b.dataset.mv; setPref('jy-mview', mView); renderMeetings(); } });

/* ---------- 누적 노트 (날짜별 기록) ---------- */
const dlgSeries = $('#dlgSeries');
let curSeries = null;
function openSeries(name) {
  const g = seriesList().find((x) => x.name === name);
  if (!g) return;
  curSeries = name;
  const first = g.items[g.items.length - 1];
  const n = g.items.length;
  $('#seriesBody').innerHTML = `<div class="detail series-detail">
    <div class="series-head" style="--c:${seriesColor(name)}"><span>📒</span><div><h2>${esc(name)}</h2>
      <div class="muted sm">기록 ${n}개 · ${fmt(first.date)} ~ ${fmt(g.last.date)}${g.last.attendees ? ' · ' + esc(g.last.attendees) : ''}</div></div></div>
    <h3>아직 안 끝난 할일 (${g.open.length})</h3>
    ${g.open.length ? g.open.sort(byDue).map((t) => taskItem(t)).join('') : '<div class="muted">모두 완료했습니다 🎉</div>'}
    <h3>날짜별 기록</h3>
    <div class="timeline">${g.items.map((m, i) => `<div class="tl" data-open="meetingView" data-id="${m.id}">
      <div class="tl-no">#${n - i}<br><span>${fmt(m.date)}</span></div>
      <div class="tl-body"><b>${esc(m.title || name)}</b>${m.inkPages ? ` <span class="tag yellow">✍️ ${m.inkPages}쪽</span>` : ''}
        <div class="sub">${esc(m.summary || (m.content || '').slice(0, 160) || (m.inkPages ? '필기 메모' : '내용 없음'))}</div>
        ${m.decisions?.length ? `<ul class="dec">${m.decisions.map((d) => `<li>✔ ${esc(d)}</li>`).join('')}</ul>` : ''}
      </div></div>`).join('')}</div></div>`;
  if (!dlgSeries.open) dlgSeries.showModal();
}
// 오늘 회차 시작: 오늘 기록이 이미 있으면 이어서 쓰고, 없으면 지난 회차의 참석자·태그를 가져와 새로 시작
function startSeriesMeeting(name) {
  const items = store.get().meetings.filter((m) => m.series === name).sort((a, b) => b.date.localeCompare(a.date));
  const today = items.find((m) => m.date === todayStr());
  if (today) return openMeeting(today.id);
  const last = items[0];
  return openMeeting(null, { series: name, title: name, attendees: last?.attendees || '', tags: last?.tags || '' });
}
$('#seriesStart').onclick = () => { dlgSeries.close(); startSeriesMeeting(curSeries); };
$('#seriesExport').onclick = async () => {
  const ms = store.get().meetings.filter((m) => m.series === curSeries);
  const ids = new Set(ms.map((m) => m.id));
  try {
    await exportExcel({ meetings: ms, tasks: store.get().tasks.filter((t) => ids.has(t.meetingId)) }, `${curSeries}_${todayStr()}.xlsx`, ['meetings', 'tasks']);
    toast('엑셀로 저장했습니다.');
  } catch (e) { alert(e.message); }
};
$('#seriesRename').onclick = () => {
  const next = (prompt('노트 새 이름', curSeries) || '').trim();
  if (!next || next === curSeries) return;
  for (const m of store.get().meetings.filter((x) => x.series === curSeries)) {
    store.upsert('meetings', { id: m.id, series: next, title: m.title === curSeries ? next : m.title });
  }
  curSeries = next;
  openSeries(next);
  toast('이름을 바꿨습니다.');
};
function newSeries() {
  const name = (prompt('누적 노트 이름을 입력하세요 (예: 주간업무회의, 프로젝트A, 고객사B)') || '').trim();
  if (!name) return;
  if (seriesList().some((g) => g.name === name)) return openSeries(name);
  openMeeting(null, { series: name, title: name });
}

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
    ...data.meetings.map((m) => ({ d: m.date, t: `[회의메모] ${m.date} | ${m.series ? '노트: ' + m.series + ' | ' : ''}${m.title} | 참석: ${m.attendees || '-'} | 태그: ${m.tags || '-'}\n요약: ${m.summary || '-'}\n결정: ${(m.decisions || []).join(' / ') || '-'}\n할일: ${(m.actions || []).map((a) => `${a.text}(${a.owner || '담당 미정'}, ${a.due || '기한 없음'})`).join(' / ') || '-'}\n원문: ${(m.content || '').slice(0, 3000)}` })),
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

/* ---------- 회의메모 편집기 (펜 필기 + 텍스트 + 요약) ---------- */
const dlgMeeting = $('#dlgMeeting'), fMeeting = $('#formMeeting');
const actionsToText = (a = []) => a.map((x) => [x.text, x.owner, x.due].filter((v, i) => v || i === 0).join(' / ')).join('\n');
const textToActions = (s) => lines(s).map((l) => { const [text, owner = '', due = ''] = l.split(/\s*\/\s*/); return { text, owner, due }; });
const pad = new InkPad($('#inkHost'), { onChange: () => scheduleSave() });
let saveTimer = null;
let lastSaved = '';
fMeeting.addEventListener('input', (e) => { if (!e.target.closest('.ink-host')) scheduleSave(); });

/* ----- 자동 저장: 쓰는 동안 1.2초 쉬면 저장, 닫을 때·앱을 내릴 때도 저장 ----- */
function setSaveState(state) {
  const el = $('#saveState');
  const t = new Date();
  const hm = `${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')}`;
  el.className = 'save-state ' + state;
  el.textContent = { pending: '✎ 입력 중…', saved: `✓ 자동 저장됨 ${hm}`, idle: '자동 저장', error: '⚠ 저장 실패' }[state] || '';
}
function scheduleSave(delay = 1200) {
  if (!dlgMeeting.open) return;
  setSaveState('pending');
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => saveMeeting(), delay);
}
const autoTitle = (o) => o.series || `회의메모 ${fmt(o.date)}`;

function saveMeeting({ final = false } = {}) {
  clearTimeout(saveTimer);
  saveTimer = null;
  const o = read(fMeeting);
  const empty = !o.content && pad.isEmpty() && !o.summary && !o.decisions && !o.actions && !o.title;
  if (!o.id && empty) { setSaveState('idle'); return null; } // 아무것도 안 쓴 새 메모는 저장하지 않음
  if (!o.id) delete o.id;
  if (!o.title) o.title = autoTitle(o);
  // 직접 요약하지 않았으면, 회의를 마칠 때 최신 내용으로 자동 요약을 다시 만듦
  if (final && o.content && (!o.summary || o.summarizedBy === '자동')) {
    const r = localSummarize(o.content, o.date);
    Object.assign(o, { summary: r.summary, decisions: r.decisions.join('\n'), actions: actionsToText(r.actions), keywords: r.keywords.join(', '), summarizedBy: '자동' });
  }
  const prev = o.id ? store.find('meetings', o.id) : null;
  // 필기는 이번에 실제로 바뀌었을 때만 저장 (불러오는 중이거나 못 불러왔을 때 지우는 일 방지)
  const inkChanged = pad.dirty;
  const inkPages = inkChanged ? pad.pageCount() : (prev?.inkPages || 0);
  const inkRev = inkChanged ? (inkPages ? new Date().toISOString() : '') : (prev?.inkRev || '');
  const item = { ...o, inkPages, inkRev, decisions: lines(o.decisions), actions: textToActions(o.actions), keywords: splitList(o.keywords) };
  const sig = JSON.stringify(item);
  let m = prev;
  if (sig !== lastSaved || inkChanged) {
    m = store.upsert('meetings', item);
    lastSaved = JSON.stringify({ ...item, id: m.id });
    if (!fMeeting.elements.id.value) {
      fMeeting.elements.id.value = m.id;
      $('[data-del]', fMeeting).style.display = '';
    }
    store.syncMeetingTasks(m);
  }
  if (inkChanged && m) {
    pad.dirty = false;
    saveInk(m.id, inkPages ? { ...pad.getData(), rev: inkRev } : null, { immediate: final })
      .catch(() => { setSaveState('error'); alert('필기를 저장하지 못했습니다. 저장 공간을 확인해주세요.'); });
  }
  if (final) flushInk();
  if (o.title !== fMeeting.elements.title.value && final) fMeeting.elements.title.value = o.title;
  setSaveState('saved');
  return m;
}
// 앱을 내리거나 다른 앱으로 전환할 때도 저장
document.addEventListener('visibilitychange', () => { if (document.hidden && dlgMeeting.open) { saveMeeting(); flushInk(); } });
window.addEventListener('pagehide', () => { if (dlgMeeting.open) { saveMeeting(); flushInk(); } });

function setTab(tab) {
  fMeeting.dataset.tab = tab;
  $$('.ed-tabs [data-tab]').forEach((b) => b.classList.toggle('on', b.dataset.tab === tab));
  if (tab !== 'sum') setPref('jy-ed-tab', tab);
  if (tab === 'ink') requestAnimationFrame(() => pad.layout(true));
}
$('.ed-tabs').addEventListener('click', (e) => { const b = e.target.closest('[data-tab]'); if (b) setTab(b.dataset.tab); });

/* ----- 요약창 접기/펴기 (가로 화면) ----- */
function setSumCollapsed(v) {
  fMeeting.classList.toggle('sum-collapsed', v);
  $('#sumToggle').classList.toggle('on', !v);
  $('#sumToggle').textContent = v ? '◂ 요약' : '요약 ▸';
  $('#sumToggle').title = v ? '요약창 펴기' : '요약창 접기';
  $('#sumToggle').setAttribute('aria-expanded', String(!v));
  setPref('jy-sum-collapsed', v ? '1' : '0');
  requestAnimationFrame(() => pad.layout(true));
}
$('#sumToggle').onclick = () => setSumCollapsed(!fMeeting.classList.contains('sum-collapsed'));
setSumCollapsed(pref('jy-sum-collapsed', '0') === '1');

/* ----- 정보줄(날짜·노트·참석자·태그)은 필요할 때만 펼침 ----- */
function setInfoOpen(v) {
  fMeeting.classList.toggle('info-open', v);
  $('#infoToggle').classList.toggle('on', v);
  setPref('jy-info-open', v ? '1' : '0');
  requestAnimationFrame(() => pad.layout(true));
}
$('#infoToggle').onclick = () => setInfoOpen(!fMeeting.classList.contains('info-open'));

/* ----- 전체 화면: 위쪽 줄과 요약창을 숨기고 메모장만, 기기 상태바도 숨김 ----- */
function setFocus(v) {
  fMeeting.classList.toggle('focus', v);
  try {
    if (v && !document.fullscreenElement) document.documentElement.requestFullscreen?.({ navigationUI: 'hide' })?.catch(() => {});
    if (!v && document.fullscreenElement) document.exitFullscreen?.().catch(() => {});
  } catch { /* 지원 안 하는 브라우저 */ }
  requestAnimationFrame(() => pad.layout(true));
}
$('#focusToggle').onclick = () => setFocus(true);
$('#focusExit').onclick = () => setFocus(false);
document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && fMeeting.classList.contains('focus')) setFocus(false); });

/* ----- 누적 노트: 지난 기록 요약과 남은 할일을 옆에 보여줌 ----- */
function renderPrevBox() {
  const name = fMeeting.elements.series.value.trim();
  const box = $('#prevBox');
  const curId = fMeeting.elements.id.value;
  const items = name ? store.get().meetings.filter((m) => m.series === name && m.id !== curId).sort((a, b) => b.date.localeCompare(a.date)) : [];
  box.hidden = !items.length;
  if (!items.length) return;
  const ids = new Set(items.map((m) => m.id));
  const open = store.get().tasks.filter((t) => !t.done && ids.has(t.meetingId)).sort(byDue);
  $('#prevTitle').textContent = `📚 이 노트의 지난 기록 · ${name} (${items.length}개)`;
  $('#prevList').innerHTML =
    (open.length ? `<div class="group-title">남은 할일 ${open.length}건</div>${open.map((t) => taskItem(t)).join('')}` : '') +
    `<div class="group-title">최근 기록</div>` +
    items.slice(0, 4).map((m) => `<div class="prev-item" data-open="meetingView" data-id="${m.id}"><b>${fmt(m.date)}</b> ${esc(m.summary || (m.content || '').slice(0, 100) || (m.inkPages ? '필기 메모' : '-'))}
      ${m.decisions?.length ? `<ul>${m.decisions.slice(0, 3).map((d) => `<li>${esc(d)}</li>`).join('')}</ul>` : ''}</div>`).join('') +
    (items.length > 4 ? `<button type="button" class="btn sm" data-series="${esc(name)}">전체 ${items.length}개 기록 보기</button>` : '');
}
fMeeting.elements.series.addEventListener('change', () => {
  const name = fMeeting.elements.series.value.trim();
  const last = store.get().meetings.filter((m) => m.series === name).sort((a, b) => b.date.localeCompare(a.date))[0];
  if (last) { // 같은 노트의 지난 참석자·태그를 비어 있을 때만 채움
    if (!fMeeting.elements.attendees.value) fMeeting.elements.attendees.value = last.attendees || '';
    if (!fMeeting.elements.tags.value) fMeeting.elements.tags.value = last.tags || '';
  }
  renderPrevBox();
});

async function openMeeting(id, preset = {}) {
  if (dlgMeeting.open) saveMeeting({ final: true });
  const m = id ? store.find('meetings', id) : { date: todayStr(), ...preset };
  fill(fMeeting, { ...m, decisions: (m.decisions || []).join('\n'), actions: actionsToText(m.actions), keywords: (m.keywords || []).join(', ') });
  $('#seriesList').innerHTML = seriesList().map((g) => `<option value="${esc(g.name)}"></option>`).join('');
  $('[data-del]', fMeeting).style.display = id ? '' : 'none';
  $('#sumStatus').textContent = m.summarizedBy ? `${m.summarizedBy}으로 요약됨` : '';
  $('#sumDot').classList.toggle('on', !!m.summary);
  $('#btnAiSum').title = aiEnabled() ? '' : '설정에서 API 키를 넣으면 사용할 수 있습니다';
  pad.load(null);
  lastSaved = id ? JSON.stringify(read(fMeeting)) : '';
  setSaveState(id ? 'idle' : 'idle');
  renderPrevBox();
  setInfoOpen(!id ? true : pref('jy-info-open', '0') === '1');
  setTab(id && m.content && !m.inkPages ? 'text' : pref('jy-ed-tab', 'ink'));
  dlgMeeting.showModal();
  if (id && m.inkPages) {
    try { pad.load(await loadInk(m)); } catch { toast('필기를 불러오지 못했습니다.'); }
    if (!m.content) setTab('ink');
  }
}

function closeMeeting() {
  if (fMeeting.classList.contains('focus')) setFocus(false);
  const m = saveMeeting({ final: true });
  dlgMeeting.close();
  if (m) toast(m.actions?.length ? `저장됨 · 할일 ${m.actions.length}건` : '저장했습니다.');
  return true;
}
$('[data-cancel]', fMeeting).onclick = closeMeeting;
dlgMeeting.addEventListener('cancel', (e) => { e.preventDefault(); closeMeeting(); });

function applySummary(r, by) {
  fMeeting.elements.summary.value = r.summary || '';
  fMeeting.elements.decisions.value = (r.decisions || []).join('\n');
  fMeeting.elements.actions.value = actionsToText(r.actions);
  fMeeting.elements.keywords.value = (r.keywords || []).join(', ');
  fMeeting.elements.summarizedBy.value = by;
  $('#sumDot').classList.add('on');
  $('#sumStatus').textContent = `${by}으로 요약했습니다. 필요하면 고치세요. (자동 저장됩니다)`;
  if (fMeeting.classList.contains('sum-collapsed')) setSumCollapsed(false);
  if (window.matchMedia('(max-width: 1099px)').matches) setTab('sum');
  scheduleSave(300);
}
$('#btnLocalSum').onclick = () => {
  const content = fMeeting.elements.content.value;
  if (!content.trim()) {
    $('#sumStatus').textContent = pad.isEmpty() ? '메모 내용을 먼저 입력해주세요.' : '자동 요약은 텍스트만 읽습니다. 펜 필기는 “✨ AI 요약” 또는 “필기 텍스트 변환”을 사용하세요.';
    return;
  }
  applySummary(localSummarize(content, fMeeting.elements.date.value), '자동 요약');
};

async function busy(btn, msg, fn) {
  btn.disabled = true;
  $('#sumStatus').textContent = msg;
  try { await fn(); } catch (e) { $('#sumStatus').textContent = e.message; } finally { btn.disabled = false; }
}
$('#btnAiSum').onclick = () => {
  const o = read(fMeeting);
  if (!o.content && pad.isEmpty()) { $('#sumStatus').textContent = '메모 내용을 먼저 입력해주세요.'; return; }
  if (!aiEnabled()) {
    if (o.content) {
      applySummary(localSummarize(o.content, o.date), '자동 요약');
      $('#sumStatus').textContent = 'API 키가 없어 자동 요약을 사용했습니다. (필기 인식·AI 요약은 설정에서 키 입력 후 사용)';
    } else {
      $('#sumStatus').textContent = '펜 필기를 요약하려면 설정에서 Claude API 키를 입력해주세요.';
    }
    return;
  }
  busy($('#btnAiSum'), 'AI가 메모와 필기를 읽고 요약하는 중…', async () => applySummary(await aiSummarize(o, pad.toImages()), 'AI 요약'));
};
$('#btnTranscribe').onclick = () => {
  if (pad.isEmpty()) { $('#sumStatus').textContent = '변환할 펜 필기가 없습니다.'; return; }
  if (!aiEnabled()) { $('#sumStatus').textContent = '필기 텍스트 변환은 설정에서 Claude API 키를 입력하면 사용할 수 있습니다.'; return; }
  busy($('#btnTranscribe'), '필기를 읽는 중…', async () => {
    const text = (await aiTranscribe(pad.toImages())).trim();
    const ta = fMeeting.elements.content;
    ta.value = ta.value.trim() ? `${ta.value.trim()}\n\n[필기 변환]\n${text}` : text;
    scheduleSave(300);
    $('#sumStatus').textContent = '필기를 텍스트로 옮겼습니다. 텍스트 탭에서 확인하세요. (필기 원본도 그대로 보관됩니다)';
    setTab('text');
  });
};

fMeeting.addEventListener('submit', (ev) => {
  ev.preventDefault();
  closeMeeting();
});
$('[data-del]', fMeeting).onclick = () => {
  const id = fMeeting.elements.id.value;
  if (id && confirm('이 회의메모(필기 포함)와 연결된 할일을 삭제할까요?')) {
    clearTimeout(saveTimer);
    store.remove('meetings', id);
    saveInk(id, null).catch(() => {});
    pad.dirty = false;
    dlgMeeting.close();
    toast('삭제했습니다.');
  }
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
    ${m.series ? `<button type="button" class="tag series" style="--c:${seriesColor(m.series)}" data-series="${esc(m.series)}">📒 ${esc(m.series)} 노트 열기</button>` : ''}
    <h2>${esc(m.title)}</h2>
    <div class="muted sm">${fmt(m.date)} · ${esc(m.attendees || '참석자 미기재')} ${m.summarizedBy ? `· <span class="tag gray">${esc(m.summarizedBy)}</span>` : ''}</div>
    <div>${splitList(m.tags).map((t) => `<span class="tag">${esc(t)}</span>`).join('')}${(m.keywords || []).map((k) => `<span class="tag gray">#${esc(k)}</span>`).join('')}</div>
    <h3>요약</h3><div>${esc(m.summary || '요약 없음')}</div>
    <h3>결정사항</h3>${m.decisions?.length ? `<ul>${m.decisions.map((d) => `<li>${esc(d)}</li>`).join('')}</ul>` : '<div class="muted">없음</div>'}
    <h3>할일</h3>${m.actions?.length ? `<ul>${m.actions.map((a) => `<li>${esc(a.text)} ${a.owner ? `<span class="tag">${esc(a.owner)}</span>` : ''}${a.due ? `<span class="tag gray">~${fmt(a.due)}</span>` : ''}</li>`).join('')}</ul>` : '<div class="muted">없음</div>'}
    ${m.inkPages ? `<h3>펜 필기 (${m.inkPages}쪽)</h3><div id="viewInk" class="ink-preview"></div>` : ''}
    ${m.content ? `<h3>텍스트 메모</h3><div class="raw">${esc(m.content)}</div>` : ''}</div>`;
  dlgView.showModal();
  if (m.inkPages) loadInk(m).then((d) => renderInkPreview($('#viewInk'), d)).catch(() => {});
}
$('#viewEdit').onclick = () => { dlgView.close(); openMeeting(viewId); };
// 아무 메모나 누적 노트로: 이 메모를 노트의 첫 기록으로 묶고, 오늘 기록을 이어서 씀
$('#viewAppend').onclick = () => {
  const m = store.find('meetings', viewId);
  if (!m) return;
  let name = m.series;
  if (!name) {
    name = m.title || `회의메모 ${fmt(m.date)}`;
    store.upsert('meetings', { id: m.id, series: name });
    toast(`“${name}” 누적 노트로 묶었습니다.`);
  }
  dlgView.close();
  startSeriesMeeting(name);
};
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
    if (k === 'series') newSeries();
    if (k === 'task') openTask(null);
    return;
  }
  const sr = e.target.closest('[data-series]');
  if (sr) { if (dlgMeeting.open) saveMeeting(); if (dlgView.open) dlgView.close(); openSeries(sr.dataset.series); return; }
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
$('#jsonExport').onclick = async () => {
  let ink = {};
  try { ink = await inkdb.all(); } catch { /* 필기 없음 */ }
  const blob = new Blob([JSON.stringify({ ...store.get(), ink })], { type: 'application/json' });
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
    const { ink = {}, ...rest } = d;
    if (confirm('현재 데이터를 백업 파일 내용으로 바꿀까요? (취소 시 합치기)')) { await inkdb.clear(); store.replaceAll(rest); } else store.merge(rest);
    for (const [id, v] of Object.entries(ink)) await inkdb.put(id, { ...v, synced: false });
    sync.syncNow();
    toast('복원했습니다.');
  } catch (err) { alert('복원 실패: ' + err.message); }
  e.target.value = '';
};
$('#wipe').onclick = () => {
  const cloud = sync.signedIn() ? '\n(로그인 중이라 다른 기기와 클라우드에서도 삭제됩니다)' : '';
  if (confirm('모든 일정·회의메모·할일을 삭제합니다. 먼저 백업하셨나요?' + cloud) && confirm('정말 삭제할까요? 되돌릴 수 없습니다.')) {
    store.replaceAll({ events: [], meetings: [], tasks: [] });
    inkdb.clear().catch(() => {});
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

/* ---------- 탭·PC 동기화 ---------- */
const timeLabel = (iso) => {
  if (!iso) return '';
  const d = new Date(iso);
  return `${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
};
sync.subscribe((s) => {
  const label = {
    off: '동기화 꺼짐 (이 기기에만 저장)',
    signedout: '로그인하면 탭·PC가 동기화됩니다',
    syncing: '동기화 중…',
    synced: `동기화됨${s.lastSync ? ' · ' + timeLabel(s.lastSync) : ''}`,
    error: '동기화 오류',
  }[s.status] || '';
  const chip = $('#syncChip');
  chip.textContent = { off: '💾 이 기기에만 저장', signedout: '☁️ 로그인 필요', syncing: '🔄 동기화 중', synced: '☁️ 동기화됨', error: '⚠️ 동기화 오류' }[s.status];
  chip.className = 'sync-chip ' + s.status;
  chip.title = s.error || label;
  $('#syncStatus').innerHTML = `<b>${esc(label)}</b>${s.email ? `<br><span class="muted sm">계정: ${esc(s.email)}</span>` : ''}${s.error ? `<br><span class="err sm">${esc(s.error)}</span>` : ''}`;
  $('#syncLogin').hidden = s.status !== 'signedout';
  $('#syncLogout').hidden = !s.email;
  $('#syncNow').hidden = !s.email;
  $('#syncSetup').open = s.status === 'off';
});
$('#syncChip').onclick = () => show('settings');
$('#syncLogin').onclick = async () => {
  try { await sync.signIn(); } catch (e) { alert(e.message); }
};
$('#syncLogout').onclick = async () => {
  if (confirm('로그아웃할까요? 이 기기의 기록은 그대로 남고, 동기화만 멈춥니다.')) await sync.signOut();
};
$('#syncNow').onclick = async () => {
  const r = await sync.syncNow();
  if (r) toast(`동기화 완료 (받음 ${r.down} · 보냄 ${r.up})`);
};
$('#fbSave').onclick = () => {
  const text = $('#fbConfig').value.trim();
  const cfg = parseConfig(text);
  if (text && !cfg) { alert('설정 형식을 인식하지 못했습니다. Firebase 콘솔의 firebaseConfig 내용을 그대로 붙여넣어주세요.'); return; }
  settings.set({ firebase: cfg });
  toast('저장했습니다. 다시 불러옵니다…');
  setTimeout(() => location.reload(), 600);
};
const curCfg = getConfig();
if (curCfg) $('#fbConfig').value = JSON.stringify(curCfg, null, 2);
sync.init();

/* ---------- 버전 표시 · 업데이트 알림 ---------- */
$('#appVer').textContent = `v${APP_VERSION}`;
$('#setVer').textContent = `버전 v${APP_VERSION} · 업데이트 내역`;
function showChanges(list, title) {
  $('#updateBody').innerHTML = `<div class="detail"><h2>${title}</h2>` + list.map((c) => `
    <h3>v${esc(c.v)} <span class="muted sm">${esc(c.date)}</span></h3>
    <ul>${c.items.map((i) => `<li>${esc(i)}</li>`).join('')}</ul>`).join('') + '</div>';
  const d = $('#dlgUpdate');
  if (!d.open) d.showModal();
}
$('#appVer').onclick = $('#setVer').onclick = () => showChanges(CHANGES.slice(0, 5), '업데이트 내역');
// 접속 시: 지난번에 본 버전보다 새 버전이면 바뀐 내용을 알려줌
(() => {
  const seen = pref('jy-seen-version', '');
  setPref('jy-seen-version', APP_VERSION);
  if (seen === APP_VERSION) return;
  const news = seen ? CHANGES.filter((c) => cmpVersion(c.v, seen) > 0) : CHANGES.slice(0, 1);
  if (news.length) setTimeout(() => showChanges(news, `🎉 새 버전 v${APP_VERSION}이 적용되었습니다`), 500);
})();
// 앱을 켜둔 동안 새 버전이 올라오면 '지금 적용' 안내
let latestRemote = APP_VERSION;
async function checkUpdate() {
  try {
    const res = await fetch('js/version.js', { cache: 'no-store' });
    const v = ((await res.text()).match(/APP_VERSION\s*=\s*'([^']+)'/) || [])[1];
    if (v && cmpVersion(v, APP_VERSION) > 0 && v !== latestRemote) {
      latestRemote = v;
      $('#updateBarText').textContent = `새 버전 v${v}이 있습니다`;
      $('#updateBar').hidden = false;
    }
  } catch { /* 오프라인 */ }
}
$('#updateApply').onclick = () => {
  if (dlgMeeting.open) saveMeeting({ final: true });
  flushInk();
  setTimeout(() => location.reload(), 400);
};
$('#updateLater').onclick = () => { $('#updateBar').hidden = true; };
document.addEventListener('visibilitychange', () => { if (!document.hidden) checkUpdate(); });
setInterval(checkUpdate, 30 * 60 * 1000);
setTimeout(checkUpdate, 5000);

/* ---------- 시작 ---------- */
function renderAll() {
  if (dlgSeries.open && curSeries) openSeries(curSeries);
  if (dlgMeeting.open) renderPrevBox();
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
