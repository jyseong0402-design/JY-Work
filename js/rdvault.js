// 📥 RD 보관함: 이번 달 RD를 이 기기(IndexedDB)에 월별로 보관하고, 그 RD로 만들 수 있는 백업자료를 한꺼번에 만듦
// 외부 서버를 쓰지 않음(요금 없음). 파일은 이 기기 브라우저 안에만 저장됨
import { loadXLSX } from './excel.js';
import { readPartnerLedgers, readAccountLedgers, readReceipts, buildArAp, buildInternal, buildPacking, detectBook } from './closing.js';
import { readAggregate, buildVerification } from './freight.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ym = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
const md = (t) => { const d = new Date(t); return `${d.getMonth() + 1}/${d.getDate()}`; };
const kb = (n) => `${Math.max(1, Math.round(n / 1024)).toLocaleString()}KB`;

const RD_TYPES = { partner: '거래처원장', account: '계정별원장', receipts: '기간별입고현황', agg: '운반비 집계' };
// 자료별: 필요한 RD 종류, 지난달 파일 보관 위치
const OUTPUTS = [
  { kind: 'arap', label: '💳 채권·채무', need: ['partner'] },
  { kind: 'prepay', label: '💰 선급금', need: ['partner'] },
  { kind: 'internal', label: '🔁 내부거래 집계', need: ['account'] },
  { kind: 'freight', label: '🚚 운반비 검증', need: ['account', 'agg'] },
  { kind: 'packing', label: '📦 포장소모품비', need: ['account', 'receipts'] },
];

/* ---------- IndexedDB ---------- */
function openDb(name, store, opts) {
  return new Promise((res, rej) => {
    const r = indexedDB.open(name, 1);
    r.onupgradeneeded = () => { if (!r.result.objectStoreNames.contains(store)) r.result.createObjectStore(store, opts); };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
const tx = async (db, store, mode, fn) => { const d = await db; return new Promise((res, rej) => { const t = d.transaction(store, mode); const out = fn(t.objectStore(store)); t.oncomplete = () => res(out?.result ?? out); t.onerror = () => rej(t.error); }); };
const rdDb = () => openDb('jywork-rd', 'files', { keyPath: 'id' });
const rdAll = async () => { const d = await rdDb(); return new Promise((res, rej) => { const q = d.transaction('files').objectStore('files').getAll(); q.onsuccess = () => res(q.result || []); q.onerror = () => rej(q.error); }); };
const rdPut = (v) => tx(rdDb(), 'files', 'readwrite', (s) => s.put(v));
const rdDel = (id) => tx(rdDb(), 'files', 'readwrite', (s) => s.delete(id));
// 지난번 만든 파일: 운반비는 jywork-freight/'last', 나머지는 jywork-closing/kind (자료생성 화면과 같은 곳)
const savedDb = (kind) => (kind === 'freight' ? ['jywork-freight', 'last'] : ['jywork-closing', kind]);
async function getSaved(kind) {
  const [db, key] = savedDb(kind);
  const d = await openDb(db, 'files');
  return new Promise((res) => { const q = d.transaction('files').objectStore('files').get(key); q.onsuccess = () => res(q.result || null); q.onerror = () => res(null); });
}
async function putSaved(kind, v) { const [db, key] = savedDb(kind); return tx(openDb(db, 'files'), 'files', 'readwrite', (s) => s.put(v, key)); }

/* ---------- RD 종류 알아보기 ---------- */
async function classify(XLSX, name, buf) {
  const f = [{ name, buf }];
  const types = [];
  let period = '', month = null;
  try { if (readPartnerLedgers(XLSX, f).rows.length) types.push('partner'); } catch { /* 아님 */ }
  try {
    const a = readAccountLedgers(XLSX, f);
    if (a.rows.length) {
      types.push('account');
      const ds = a.rows.map((r) => r.date).sort((x, y) => x - y);
      const p = (d) => `${d.getMonth() + 1}/${d.getDate()}`;
      period = `${ds[0].getFullYear()}.${p(ds[0])}~${p(ds[ds.length - 1])}`;
      month = ym(ds[ds.length - 1]);
    }
  } catch { /* 아님 */ }
  try {
    const r = readReceipts(XLSX, f);
    if (r.rows.length) {
      types.push('receipts');
      const ds = r.rows.map((x) => x.date).sort((x, y) => x - y);
      if (!month) month = ym(ds[ds.length - 1]);
      if (!period) period = `${ds[0].getMonth() + 1}/${ds[0].getDate()}~${ds[ds.length - 1].getMonth() + 1}/${ds[ds.length - 1].getDate()}`;
    }
  } catch { /* 아님 */ }
  try {
    const g = readAggregate(XLSX, buf);
    if (g && g.cols?.length) { types.push('agg'); if (!month) month = `${g.year}-${String(g.month).padStart(2, '0')}`; if (!period) period = `${g.year}년 ${g.month}월`; }
  } catch { /* 아님 */ }
  return { types, period, month };
}

/* ---------- 화면 ---------- */
const st = { month: ym(new Date(new Date().getFullYear(), new Date().getMonth() - 1, 1)), files: [], saved: {}, outs: [], busy: false };
// 앱이 만든 파일 이름(ARAP_2611.xlsx 등)에서 그 파일의 달 → 보관함 달보다 같거나 최근이면 '지난달 파일'로 쓸 수 없음
const fileYm = (name) => { const m = String(name || '').match(/_(\d{2})(\d{2})\.xlsx$/); return m ? `20${m[1]}-${m[2]}` : null; };
const tooNew = (sv) => { const f = sv && fileYm(sv.name); return !!(f && f >= st.month); };
const monthLabel = (s) => { const [y, m] = s.split('-'); return `${y}년 ${+m}월`; };

async function load() {
  try { st.files = await rdAll(); } catch { st.files = []; }
  for (const o of OUTPUTS) { try { st.saved[o.kind] = await getSaved(o.kind); } catch { st.saved[o.kind] = null; } }
  render();
}

function render() {
  if (!$('#view-rd')) return;
  $('#rdMonth').textContent = `${monthLabel(st.month)} 데이터`;
  const others = [...new Set(st.files.map((f) => f.month))].filter((m) => m !== st.month).sort();
  $('#rdOther').innerHTML = others.length ? `다른 달에 보관된 RD: ${others.map((m) => `<button class="btn sm" data-rd-month="${m}">${monthLabel(m)} (${st.files.filter((f) => f.month === m).length})</button>`).join(' ')}` : '';
  const files = st.files.filter((f) => f.month === st.month).sort((a, b) => a.addedAt - b.addedAt);
  $('#rdFiles').innerHTML = files.length
    ? `<table class="fr-table"><tr><th>파일</th><th>자동 인식</th><th>기간</th><th class="r">크기</th><th>올린 날</th><th></th></tr>
      ${files.map((f) => `<tr><td>${esc(f.name)}</td><td>${f.types.length ? f.types.map((t) => `<span class="tag">${RD_TYPES[t]}</span>`).join('') : '<span class="tag red">알 수 없음</span>'}</td>
        <td class="sm">${esc(f.period || '')}</td><td class="r sm">${kb(f.size)}</td><td class="sm">${md(f.addedAt)}</td>
        <td class="r"><button class="bd-ib" data-rd-dl="${f.id}" title="내려받기">⬇</button> <button class="bd-ib" data-rd-del="${f.id}" title="삭제">🗑</button></td></tr>`).join('')}</table>`
    : '<div class="empty">이 달에 올린 RD가 없습니다. 위 칸에 RD 파일을 넣어 주세요.</div>';
  const have = new Set(files.flatMap((f) => f.types));
  let ready = 0;
  $('#rdCan').innerHTML = `<table class="fr-table"><tr><th>자료</th><th>필요한 RD</th><th>지난달 파일</th><th>상태</th></tr>
    ${OUTPUTS.map((o) => {
      const miss = o.need.filter((t) => !have.has(t));
      const sv = st.saved[o.kind];
      const ok = !miss.length && sv && !tooNew(sv);
      if (ok) ready++;
      return `<tr><td><b>${o.label}</b></td>
        <td class="sm">${o.need.map((t) => `${RD_TYPES[t]} ${have.has(t) ? '<span class="ok">✓</span>' : '<span class="plus">✗</span>'}`).join(' · ')}</td>
        <td class="sm">${sv ? `↺ ${esc(sv.name)}` : '<span class="muted">없음</span>'} <button class="btn sm" data-rd-prev="${o.kind}">${sv ? '바꾸기' : '지난달 파일 넣기'}</button></td>
        <td class="sm">${ok ? '<b class="ok">만들 수 있음</b>' : sv && tooNew(sv) ? (fileYm(sv.name) === st.month ? `<b class="ok">✓ 이 달 자료 만듦</b> <span class="muted">(${esc(sv.name)})</span>` : `<span class="plus">보관된 파일이 ${monthLabel(fileYm(sv.name))} 것이라 이 달 자료는 만들 수 없음 — 다시 만들려면 '바꾸기'로 그 전달 완성본을 넣어 주세요</span>`) : miss.length ? `<span class="muted">${miss.map((t) => RD_TYPES[t]).join('·')} 필요</span>` : '<span class="muted">지난달 파일 필요</span>'}</td></tr>`;
    }).join('')}</table>`;
  const btn = $('#rdRun');
  btn.disabled = st.busy || !ready;
  btn.textContent = st.busy ? '만드는 중…' : `⚡ 가능한 자료 ${ready}개 한꺼번에 만들기`;
}

async function addFiles(list) {
  if (!list.length) return;
  $('#rdMsg').textContent = '파일 종류를 확인하는 중…';
  const XLSX = await loadXLSX();
  const msgs = [];
  const months = new Set();
  for (const f of list) {
    const buf = await f.arrayBuffer();
    const c = await classify(XLSX, f.name, buf);
    // 날짜가 있는 RD(계정별원장·입고현황·집계)는 그 달에, 날짜가 없는 거래처원장은 보고 있는 달에 보관
    const month = c.month || st.month;
    const rec = { id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, month, name: f.name, size: f.size, types: c.types, period: c.period, addedAt: Date.now(), data: buf };
    // 같은 달·같은 이름은 바꿔 넣기
    for (const old of st.files.filter((x) => x.month === month && x.name === f.name)) await rdDel(old.id);
    await rdPut(rec);
    months.add(month);
    msgs.push(`${f.name}: ${c.types.length ? c.types.map((t) => RD_TYPES[t]).join('+') : '알 수 없는 파일'}${month !== st.month ? ` → ${monthLabel(month)}에 보관` : ''}`);
  }
  // 넣은 RD가 다른 달 데이터면 그 달 화면으로 바로 이동
  const moved = months.size === 1 && !months.has(st.month);
  if (moved) st.month = [...months][0];
  $('#rdMsg').textContent = msgs.join(' / ') + (moved ? ` — ${monthLabel(st.month)} 데이터 화면으로 이동했습니다` : '');
  st.files = await rdAll();
  render();
}

async function setPrev(kind, file) {
  const { default: JSZip } = await import('./vendor/jszip.mjs');
  const buf = await file.arrayBuffer();
  if (kind !== 'freight') {
    const d = await detectBook(JSZip, buf).catch(() => ({ kinds: [] }));
    const okKind = kind === 'arap' ? d.kinds.some((k) => k === 'ar' || k === 'ap') : d.kinds.includes(kind);
    if (!okKind) { alert(`'${file.name}'은(는) ${OUTPUTS.find((o) => o.kind === kind).label.slice(2)} 파일이 아닙니다.`); return; }
  }
  await putSaved(kind, { name: file.name, data: new Uint8Array(buf), at: Date.now() });
  st.saved[kind] = { name: file.name };
  st.saved[kind] = await getSaved(kind);
  render();
}

function download(name, data) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

async function runAll() {
  st.busy = true; render();
  $('#rdOut').innerHTML = '<div class="card"><p class="muted">RD로 자료를 만드는 중입니다… (자료당 몇 초씩 걸려요)</p></div>';
  st.outs = [];
  const cards = [];
  try {
    const [XLSX, { default: JSZip }] = await Promise.all([loadXLSX(), import('./vendor/jszip.mjs')]);
    const files = st.files.filter((f) => f.month === st.month);
    const of = (t) => files.filter((f) => f.types.includes(t)).map((f) => ({ name: f.name, buf: f.data }));
    for (const o of OUTPUTS) {
      const sv = st.saved[o.kind];
      if (!sv || tooNew(sv) || o.need.some((t) => !of(t).length)) continue;
      try {
        let out;
        const prevBuf = sv.data;
        if (o.kind === 'freight') {
          out = await buildVerification({ XLSX, JSZip, aggBuf: of('agg')[0].buf, ledgerBufs: of('account'), prevBuf });
          out.report.warnings = out.report.warnings || [];
        } else {
          const fn = o.kind === 'internal' ? buildInternal : o.kind === 'packing' ? buildPacking : buildArAp;
          const rd = o.kind === 'internal' ? of('account') : o.kind === 'packing' ? [...of('account'), ...of('receipts')] : of('partner');
          out = await fn({ XLSX, JSZip, prevBuf, rdFiles: rd });
        }
        st.outs.push({ kind: o.kind, label: o.label, out });
        await putSaved(o.kind, { name: out.fileName, data: out.data, at: Date.now() }).catch(() => {});
        const w = out.report.warnings || [];
        const scopeBad = w.filter((x) => /범위 확인/.test(x));
        const scopeOk = (out.report.info || []).filter((x) => /범위 확인 ✓/.test(x));
        cards.push(`<div class="rd-res"><div class="rd-res-h"><span>✅ <b>${o.label}</b> → ${esc(out.fileName)}</span><button class="btn sm primary" data-rd-out="${st.outs.length - 1}">⬇ 내려받기</button></div>
          ${scopeBad.map((x) => `<div class="rd-scope bad">${esc(x)}</div>`).join('')}${scopeOk.map((x) => `<div class="rd-scope ok sm">${esc(x)}</div>`).join('')}
          ${w.length ? `<details><summary class="sm">확인할 것 ${w.length}건</summary><ul class="fr-notes">${w.map((x) => `<li>⚠ ${esc(x)}</li>`).join('')}</ul></details>` : '<div class="muted sm">확인할 것 없음</div>'}</div>`);
      } catch (e) {
        console.error(e);
        cards.push(`<div class="rd-res bad"><b>${o.label}</b>: 만들지 못했습니다 — ${esc(e.message || e)}</div>`);
      }
      $('#rdOut').innerHTML = `<div class="card">${cards.join('')}</div>`;
    }
    const zipBtn = st.outs.length > 1 ? '<button class="btn primary" id="rdZip">⬇ 전체 내려받기 (zip)</button>' : '';
    $('#rdOut').innerHTML = `<div class="card"><div class="card-head"><h2>${monthLabel(st.month)} 데이터로 만든 자료 ${st.outs.length}개</h2>${zipBtn}</div>${cards.join('')}
      <p class="muted sm">만든 파일은 이 기기에 '지난번 만든 파일'로 보관되어 다음 달에 자동으로 쓰입니다. 월마감 현황에서 완료 체크를 해 주세요.</p></div>`;
    for (const o of OUTPUTS) { try { st.saved[o.kind] = await getSaved(o.kind); } catch { /* 무시 */ } }
  } catch (e) {
    $('#rdOut').innerHTML = `<div class="card err-card"><h2>만들지 못했습니다</h2><p>${esc(e.message || e)}</p></div>`;
  } finally {
    st.busy = false; render();
  }
}

async function zipAll() {
  const { default: JSZip } = await import('./vendor/jszip.mjs');
  const z = new JSZip();
  for (const { out } of st.outs) z.file(out.fileName, out.data);
  download(`closing_${st.month.replace('-', '')}.zip`, await z.generateAsync({ type: 'blob' }));
}

export function initRdVault() {
  if (!$('#view-rd')) return;
  const zone = $('#rdDrop'), inp = $('#rdIn');
  zone.addEventListener('click', () => inp.click());
  inp.addEventListener('change', () => { const f = [...inp.files]; inp.value = ''; addFiles(f); });
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('over'); addFiles([...e.dataTransfer.files].filter((f) => /\.xls[xm]?$/i.test(f.name))); });
  $('#rdPrev').onclick = () => { const [y, m] = st.month.split('-').map(Number); st.month = ym(new Date(y, m - 2, 1)); render(); };
  $('#rdNext').onclick = () => { const [y, m] = st.month.split('-').map(Number); st.month = ym(new Date(y, m, 1)); render(); };
  $('#rdRun').onclick = runAll;
  const prevIn = $('#rdPrevIn');
  let prevKind = null;
  prevIn.addEventListener('change', () => { const f = prevIn.files[0]; prevIn.value = ''; if (f && prevKind) setPrev(prevKind, f); });
  $('#view-rd').addEventListener('click', async (e) => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.rdDel) {
      const f = st.files.find((x) => x.id === b.dataset.rdDel);
      if (f && confirm(`'${f.name}'을(를) 보관함에서 지울까요?`)) { await rdDel(f.id); st.files = await rdAll(); render(); }
    } else if (b.dataset.rdDl) {
      const f = st.files.find((x) => x.id === b.dataset.rdDl);
      if (f) download(f.name, f.data);
    } else if (b.dataset.rdPrev) { prevKind = b.dataset.rdPrev; prevIn.click(); }
    else if (b.dataset.rdOut !== undefined) { const o = st.outs[+b.dataset.rdOut]; if (o) download(o.out.fileName, o.out.data); }
    else if (b.id === 'rdZip') zipAll();
    else if (b.dataset.rdMonth) { st.month = b.dataset.rdMonth; render(); }
  });
  load();
}
