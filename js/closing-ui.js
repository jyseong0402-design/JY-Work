// 📊 월마감 자동화 화면: 지난달 보고서 + ERP RD → 이번 달을 붙인 보고서
import { detectBook, buildArAp, buildInternal, buildPacking } from './closing.js';
import { loadXLSX } from './excel.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const won = (n) => (Math.round(n || 0)).toLocaleString('ko-KR');

const KIND_LABEL = { arap: '채권·채무', prepay: '선급금', internal: '내부거래 집계', packing: '포장소모품비' };
const kindOf = (kinds) => (kinds.includes('internal') ? 'internal' : kinds.includes('packing') ? 'packing' : kinds.includes('prepay') && !kinds.includes('ar') && !kinds.includes('ap') ? 'prepay' : kinds.some((k) => ['ar', 'ap', 'prepay'].includes(k)) ? 'arap' : null);

// 만든 파일은 기기에 보관 → 다음 달에 '지난번 파일'로 바로 사용
const DB = 'jywork-closing', ST = 'files';
const idb = () => new Promise((res, rej) => { const r = indexedDB.open(DB, 1); r.onupgradeneeded = () => r.result.createObjectStore(ST); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const idbPut = async (k, v) => { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction(ST, 'readwrite'); t.objectStore(ST).put(v, k); t.oncomplete = res; t.onerror = () => rej(t.error); }); };
const idbAll = async () => { const d = await idb(); return new Promise((res, rej) => { const out = {}; const c = d.transaction(ST).objectStore(ST).openCursor(); c.onsuccess = () => { const cur = c.result; if (cur) { out[cur.key] = cur.value; cur.continue(); } else res(out); }; c.onerror = () => rej(c.error); }); };

const state = { prevs: [], rds: [], saved: {} };

function fileLabel(f) { return `${esc(f.name)} <span class="muted">(${Math.round((f.size || f.data?.length || 0) / 1024).toLocaleString()}KB)</span>`; }
function render() {
  $('#clPrevName').innerHTML = state.prevs.length
    ? state.prevs.map((p, i) => `<span class="chipfile">${p.kind ? `<b>${KIND_LABEL[p.kind]}</b> · ` : p.kind === null ? '<b class="minus">알 수 없는 파일</b> · ' : ''}${fileLabel(p.file)} <button type="button" data-rm-prev="${i}" title="빼기">✕</button></span>`).join('')
    : '<span class="muted">채권채무 · 선급금 · 내부거래 집계 · 포장소모품비 파일 (여러 개 가능)</span>';
  $('#clRdName').innerHTML = state.rds.length
    ? state.rds.map((f, i) => `<span class="chipfile">${fileLabel(f)} <button type="button" data-rm-rd="${i}" title="빼기">✕</button></span>`).join('')
    : '<span class="muted">거래처원장 · 계정별원장 · 기간별입고현황 (전체계정 그대로, 여러 개 가능)</span>';
  const saved = Object.entries(state.saved);
  $('#clSaved').innerHTML = saved.length ? `<span class="muted sm">지난번 만든 파일:</span> ${saved.map(([k, v]) => `<button type="button" class="btn sm" data-use="${k}">↺ ${esc(KIND_LABEL[k] || k)} · ${esc(v.name)}</button>`).join(' ')}` : '';
  $('#clRun').disabled = !(state.prevs.some((p) => p.kind) && state.rds.length);
}

async function addPrev(files) {
  const { default: JSZip } = await import('./vendor/jszip.mjs');
  for (const f of files) {
    const item = { file: f, kind: undefined };
    state.prevs.push(item);
    render();
    try {
      const d = await detectBook(JSZip, await f.arrayBuffer());
      item.kind = kindOf(d.kinds);
    } catch { item.kind = null; }
    render();
  }
}

function bindDrop(zoneSel, inputSel, onFiles) {
  const z = $(zoneSel), inp = $(inputSel);
  z.addEventListener('click', (e) => { if (!e.target.closest('button')) inp.click(); });
  inp.addEventListener('change', () => { onFiles([...inp.files]); inp.value = ''; render(); });
  z.addEventListener('dragover', (e) => { e.preventDefault(); z.classList.add('over'); });
  z.addEventListener('dragleave', () => z.classList.remove('over'));
  z.addEventListener('drop', (e) => { e.preventDefault(); z.classList.remove('over'); onFiles([...e.dataTransfer.files].filter((f) => /\.xls[xm]?$/i.test(f.name))); render(); });
}

function download(out) {
  const blob = new Blob([out.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = out.fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

function partRows(kind, r) {
  if (kind === 'internal') {
    const p = r.parts[0];
    const line = (n, x) => (x ? `<tr><td>${n} 시트</td><td class="r">${x.rows}행</td><td class="r">${won(x.total)}</td><td class="sm">${x.removed ? `기존 ${p.m}월 ${x.removed}행 교체` : ''}${x.byRule.length ? ` · 새 거래 ${x.byRule.length}건 확인 필요` : ''}</td></tr>` : `<tr><td>${n} 시트</td><td colspan="3" class="muted">해당 행 없음</td></tr>`);
    return `<table class="fr-table"><tr><th>항목</th><th class="r">추가</th><th class="r">금액</th><th>비고</th></tr>
      ${line('비용', p.cost)}${line('수익', p.income)}
      <tr><td>내부거래 집계</td><td class="r">${p.aggRows ? `${p.aggRows}행` : '-'}</td><td></td><td class="sm">${p.y}년 ${p.m}월 묶음</td></tr></table>`;
  }
  if (kind === 'packing') {
    return `<table class="fr-table"><tr><th>항목</th><th class="r">추가</th><th class="r">금액</th><th>비고</th></tr>
      ${r.parts.map((p) => `<tr><td>${esc(p.label)}</td><td class="r">${p.rows !== null && p.rows !== undefined ? `${p.rows}행` : `${esc(p.col)}열~`}</td><td class="r">${p.total !== undefined ? won(p.total) : ''}</td><td class="sm">${p.removed ? `기존 ${p.removed}행 교체` : ''}${p.summaryCol ? `요약 ${esc(p.summaryCol)}열 추가` : ''}</td></tr>`).join('')}</table>`;
  }
  return `<table class="fr-table"><tr><th>항목</th><th>새 시트</th><th class="r">거래처</th><th class="r">잔액 합계</th><th>비교표</th></tr>
    ${r.parts.map((p) => `<tr><td>${esc(p.label)}</td><td>${esc(p.sheet)}</td><td class="r">${p.rows}</td><td class="r">${won(p.balance)}</td><td class="sm">${esc(p.cmp)} ${esc(p.col)}열${p.fresh ? ` · 새 거래처 ${p.fresh}` : ''}</td></tr>`).join('')}</table>`;
}

function resultCard(kind, src, out, i) {
  const r = out.report;
  return `<div class="card">
    <div class="fr-head"><div><h2>${esc(KIND_LABEL[kind])} 완성</h2><p class="muted sm">${esc(src)} → <b>${esc(out.fileName)}</b></p></div>
      <button class="btn primary" data-dl="${i}">⬇ 내려받기</button></div>
    ${partRows(kind, r)}
    ${r.warnings.length ? `<h3 class="sm">확인할 것</h3><ul class="fr-notes">${r.warnings.map((w) => `<li>⚠ ${esc(w)}</li>`).join('')}</ul>` : ''}
    ${r.info.length ? `<ul class="fr-notes muted">${r.info.map((w) => `<li>${esc(w)}</li>`).join('')}</ul>` : ''}
  </div>`;
}

const outs = [];
async function run() {
  const btn = $('#clRun');
  btn.disabled = true; btn.textContent = '만드는 중…';
  $('#clResult').innerHTML = '<div class="card"><p class="muted">RD를 읽고 보고서를 만드는 중입니다… (파일이 크면 10~20초 걸릴 수 있어요)</p></div>';
  outs.length = 0;
  const cards = [];
  try {
    const [XLSX, { default: JSZip }] = await Promise.all([loadXLSX(), import('./vendor/jszip.mjs')]);
    const rdFiles = await Promise.all(state.rds.map(async (f) => ({ name: f.name, buf: await f.arrayBuffer() })));
    for (const p of state.prevs) {
      if (!p.kind) continue;
      try {
        const prevBuf = await p.file.arrayBuffer();
        const fn = p.kind === 'internal' ? buildInternal : p.kind === 'packing' ? buildPacking : buildArAp;
        const out = await fn({ XLSX, JSZip, prevBuf, rdFiles });
        outs.push(out);
        cards.push(resultCard(p.kind, p.file.name, out, outs.length - 1));
        try { await idbPut(p.kind, { name: out.fileName, data: out.data, at: Date.now() }); state.saved[p.kind] = { name: out.fileName, data: out.data }; } catch { /* 저장 공간 */ }
      } catch (e) {
        console.error(e);
        cards.push(`<div class="card err-card"><h2>${esc(KIND_LABEL[p.kind])}: 만들지 못했습니다</h2><p>${esc(e.message || e)}</p></div>`);
      }
      $('#clResult').innerHTML = cards.join('');
    }
  } catch (e) {
    $('#clResult').innerHTML = `<div class="card err-card"><h2>만들지 못했습니다</h2><p>${esc(e.message || e)}</p></div>`;
  } finally {
    btn.textContent = '보고서 만들기';
    render();
  }
}

export function initClosing() {
  if (!$('#view-closing')) return;
  bindDrop('#clPrev', '#clPrevIn', (files) => addPrev(files));
  bindDrop('#clRd', '#clRdIn', (files) => { state.rds.push(...files); });
  document.addEventListener('click', (e) => {
    const t = e.target.closest('[data-rm-prev],[data-rm-rd],[data-use],[data-dl]');
    if (!t || !t.closest('#view-closing')) return;
    if (t.dataset.rmPrev !== undefined) { state.prevs.splice(+t.dataset.rmPrev, 1); render(); }
    else if (t.dataset.rmRd !== undefined) { state.rds.splice(+t.dataset.rmRd, 1); render(); }
    else if (t.dataset.use) {
      const v = state.saved[t.dataset.use];
      if (!v) return;
      const file = new File([v.data], v.name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      state.prevs = state.prevs.filter((p) => p.kind !== t.dataset.use);
      state.prevs.push({ file, kind: t.dataset.use });
      render();
    } else if (t.dataset.dl !== undefined) download(outs[+t.dataset.dl]);
  });
  $('#clRun').addEventListener('click', run);
  idbAll().then((all) => { state.saved = all || {}; render(); }).catch(() => render());
  render();
}
