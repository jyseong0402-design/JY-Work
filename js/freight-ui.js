// 🚚 운반비 검증 화면
import { buildVerification } from './freight.js';
import { loadXLSX } from './excel.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const won = (n) => (Math.round(n || 0)).toLocaleString('ko-KR');
const sign = (n) => (Math.abs(n) < 1 ? '<span class="ok">0</span>' : `<span class="${n > 0 ? 'plus' : 'minus'}">${n > 0 ? '+' : ''}${won(n)}</span>`);

// 지난번에 만든 검증파일을 기기에 보관해 두고 다음 달에 바로 쓰기
const DB = 'jywork-freight', ST = 'files';
const idb = () => new Promise((res, rej) => { const r = indexedDB.open(DB, 1); r.onupgradeneeded = () => r.result.createObjectStore(ST); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const idbPut = async (k, v) => { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction(ST, 'readwrite'); t.objectStore(ST).put(v, k); t.oncomplete = res; t.onerror = () => rej(t.error); }); };
const idbGet = async (k) => { const d = await idb(); return new Promise((res, rej) => { const q = d.transaction(ST).objectStore(ST).get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); };

const state = { agg: null, ledgers: [], prev: null, lastOut: null, mapping: {}, overrides: {} };

function fileLabel(f) { return `${esc(f.name)} <span class="muted">(${Math.round(f.size / 1024).toLocaleString()}KB)</span>`; }
function render() {
  $('#frAggName').innerHTML = state.agg ? fileLabel(state.agg) : '<span class="muted">파일을 끌어다 놓거나 눌러서 선택</span>';
  $('#frLedName').innerHTML = state.ledgers.length ? state.ledgers.map((f, i) => `<span class="chipfile">${fileLabel(f)} <button type="button" data-rm="${i}" title="빼기">✕</button></span>`).join('') : '<span class="muted">여러 개 선택 가능 (대흥·군자·송이애 등)</span>';
  $('#frPrevName').innerHTML = state.prev ? fileLabel(state.prev) : '<span class="muted">지난달까지 작업한 검증파일</span>';
  $('#frRun').disabled = !(state.agg && state.ledgers.length && state.prev);
}

function bindDrop(zoneSel, inputSel, onFiles) {
  const z = $(zoneSel), inp = $(inputSel);
  z.addEventListener('click', (e) => { if (!e.target.closest('button')) inp.click(); });
  inp.addEventListener('change', () => { onFiles([...inp.files]); inp.value = ''; render(); });
  z.addEventListener('dragover', (e) => { e.preventDefault(); z.classList.add('over'); });
  z.addEventListener('dragleave', () => z.classList.remove('over'));
  z.addEventListener('drop', (e) => { e.preventDefault(); z.classList.remove('over'); onFiles([...e.dataTransfer.files].filter((f) => /\.xls[xm]?$/i.test(f.name))); render(); });
}

async function run() {
  const btn = $('#frRun');
  btn.disabled = true; btn.textContent = '만드는 중…';
  $('#frResult').innerHTML = '<div class="card"><p class="muted">파일을 읽고 대사하는 중입니다… (원장이 크면 10초 정도 걸릴 수 있어요)</p></div>';
  try {
    const [XLSX, { default: JSZip }] = await Promise.all([loadXLSX(), import('./vendor/jszip.mjs')]);
    const buf = (f) => f.arrayBuffer();
    const extraMapping = Object.entries(state.mapping).map(([vendor, v]) => ({ vendor, ...v })).filter((x) => x.payee);
    const out = await buildVerification({
      XLSX, JSZip,
      aggBuf: await buf(state.agg),
      ledgerBufs: await Promise.all(state.ledgers.map(async (f) => ({ name: f.name, buf: await buf(f) }))),
      prevBuf: await buf(state.prev),
      extraMapping,
      overrides: state.overrides,
    });
    state.lastOut = out;
    showResult(out);
    try { await idbPut('last', { name: out.fileName, data: out.data, at: Date.now() }); } catch { /* 저장 공간 부족 등 */ }
  } catch (e) {
    console.error(e);
    $('#frResult').innerHTML = `<div class="card err-card"><h2>만들지 못했습니다</h2><p>${esc(e.message || e)}</p></div>`;
  } finally {
    btn.textContent = '검증파일 만들기';
    render();
  }
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

function showResult(out) {
  const r = out.report;
  const diffs = r.recon.filter((x) => Math.abs(x.diff) >= 1);
  const totDiff = r.companies.reduce((s, c) => s + c.diff, 0);
  const html = [];
  html.push(`<div class="card fr-head">
    <div><h2>${r.year}년 ${r.month}월 검증파일 완성</h2>
      <p class="muted sm">지난달 '${esc(r.template)}' 시트를 본떠 운반비${r.month}월·현황${r.month}월을 만들고, 계정원장에 ${r.month}월 ${r.ledgerCount}건을 붙였습니다. 엑셀에서 열면 수식이 다시 계산됩니다.</p></div>
    <button class="btn primary" id="frDown">⬇ ${esc(out.fileName)}</button></div>`);
  html.push(`<div class="stats fr-stats">
    <div class="stat"><b>${won(r.opsTotal)}</b><span>운영 집계 합계</span></div>
    <div class="stat"><b>${won(r.ledgerTotal)}</b><span>전표(원장 차변) ${r.ledgerCount}건</span></div>
    <div class="stat ${Math.abs(totDiff) >= 1 ? 'warn' : ''}"><b>${sign(totDiff)}</b><span>운영 − 전표 (선행)</span></div>
    <div class="stat ${r.unmapped.length ? 'warn' : ''}"><b>${r.unmapped.length}</b><span>매핑 없는 거래처</span></div></div>`);
  const notes = [...r.warnings.map((w) => `<li class="w">⚠️ ${esc(w)}</li>`), ...r.info.map((w) => `<li>ℹ️ ${esc(w)}</li>`)];
  if (notes.length) html.push(`<div class="card"><ul class="fr-notes">${notes.join('')}</ul></div>`);
  if (r.unmapped.length) {
    html.push(`<div class="card"><h2>매핑 없는 거래처 (${r.unmapped.length})</h2>
      <p class="muted sm">mapping 시트에 없는 거래처입니다. 지급처 이름을 넣고 <b>다시 만들기</b>를 누르면 mapping 시트에 추가됩니다. (운반비 시트 I열의 지급처 이름과 같게: 예) 스타, 이그린, 기타)</p>
      <table class="fr-table"><tr><th>회사</th><th>거래처(관리항목1)</th><th class="r">금액</th><th>적요</th><th>지급처</th></tr>
      ${r.unmapped.map((u) => `<tr><td>${esc(u.co)}</td><td>${esc(u.vendor)}</td><td class="r">${won(u.amount)}</td><td class="sm muted">${esc(u.memo).slice(0, 50)}</td>
        <td><input class="fr-map" data-vendor="${esc(u.vendor)}" data-co="${esc(u.co)}" data-lead="${esc(u.lead)}" value="${esc(state.mapping[u.vendor]?.payee || '')}" placeholder="예) 기타" list="frPayees"/></td></tr>`).join('')}
      </table>
      <datalist id="frPayees">${[...new Set(r.recon.map((x) => x.payee).filter((p) => p && !p.startsWith('(')))].map((p) => `<option value="${esc(p)}">`).join('')}</datalist>
      <div class="row-end"><button class="btn primary" id="frRemap">매핑 넣고 다시 만들기</button></div></div>`);
  }
  // 납품처별 금액 (지난달 대비 크게 바뀐 곳 ⚠, 집계파일 오류는 여기서 수기 수정)
  const big = (p) => p.prev !== null && p.prev !== undefined && Math.abs(p.amount - p.prev) >= 500000 && (p.prev === 0 || Math.abs(p.amount - p.prev) / Math.abs(p.prev) >= 0.3);
  const flagged = r.places.filter(big).length;
  html.push(`<details class="card fr-places" ${flagged || r.places.some((p) => p.overridden) ? 'open' : ''}>
    <summary><b>납품처별 금액</b> <span class="muted sm">${r.places.length}곳 · 지난달 대비 30%·50만 원 이상 변동 ⚠ ${flagged}곳</span></summary>
    <p class="muted sm">집계파일에 오류가 있으면 <b>이번 달</b> 칸을 고친 뒤 <b>수정값 반영해서 다시 만들기</b>를 누르세요. (엑소푸레쉬 등 0 처리 규칙 대상은 자동 0원)</p>
    <table class="fr-table"><tr><th>행</th><th>회사</th><th>납품처 · 운송사</th><th class="r">지난달</th><th class="r">집계파일</th><th class="r">이번 달 (수정 가능)</th></tr>
    ${r.places.map((p) => `<tr class="${big(p) ? 'flag' : ''}"><td class="muted">${p.row}</td><td>${esc(p.co)}</td><td>${big(p) ? '⚠ ' : ''}${esc(p.place)} · <span class="muted">${esc(p.carrier)}</span>${p.zeroed ? ' <span class="tag gray">0 처리 규칙</span>' : ''}${p.overridden ? ' <span class="tag">수기 수정</span>' : ''}</td>
      <td class="r muted">${p.prev === null || p.prev === undefined ? '신규' : won(p.prev)}</td><td class="r">${p.agg === null || p.agg === undefined ? '-' : won(p.agg)}</td>
      <td class="r"><input class="fr-ov" inputmode="numeric" data-key="${esc(p.key)}" value="${won(p.amount)}" data-orig="${Math.round(p.amount)}" /></td></tr>`).join('')}
    </table>
    <div class="row-end"><button class="btn" id="frOvReset">수기 수정 모두 취소</button><button class="btn primary" id="frOvApply">수정값 반영해서 다시 만들기</button></div></details>`);
  html.push(`<div class="card"><h2>회사별 운영 집계 vs 전표</h2>
    <table class="fr-table"><tr><th>회사</th><th class="r">운영 집계</th><th class="r">전표(선행)</th><th class="r">차이</th></tr>
    ${r.companies.map((c) => `<tr><td>${esc(c.co)}</td><td class="r">${won(c.ops)}</td><td class="r">${won(c.ledger)}</td><td class="r">${sign(c.diff)}</td></tr>`).join('')}</table></div>`);
  html.push(`<div class="card"><h2>차이 나는 지급처 (${diffs.length})</h2>
    <p class="muted sm">운반비${r.month}월 시트 U열(사유)에 아래 '제안'을 적어 두었습니다. 엑셀에서 확인하고 고쳐 쓰세요.</p>
    ${diffs.length ? `<table class="fr-table"><tr><th>회사</th><th>거래처 → 지급처</th><th class="r">운영</th><th class="r">전표</th><th class="r">차이</th><th>제안 (지난달 사유 · 원장 적요)</th></tr>
    ${diffs.map((x) => `<tr><td>${esc(x.co)}</td><td>${esc(x.vendor)}<br><span class="muted sm">→ ${esc(x.payee)} · U${x.row}</span></td><td class="r">${won(x.ops)}</td><td class="r">${won(x.ledger)}</td><td class="r">${sign(x.diff)}</td>
      <td class="sm">${x.prevNote ? `<div><b>지난달</b> ${esc(x.prevNote)}</div>` : ''}${x.lines.slice(0, 3).map((l) => `<div class="muted">${won(l.amount)} · ${esc(l.memo).slice(0, 60)}</div>`).join('')}</td></tr>`).join('')}</table>` : '<p>모든 지급처가 일치합니다 🎉</p>'}</div>`);
  if (r.listChanges.length) {
    const icon = { change: '✏️', new: '🆕', missing: '➖', 'new-fail': '⛔' };
    html.push(`<div class="card"><h2>납품처 목록 변경 (지난달 대비)</h2><p class="muted sm">운반비${r.month}월 시트 C~J열에 반영했습니다. 지급처(I열)가 맞는지 확인하세요.</p>
      <ul class="fr-notes">${r.listChanges.map((c) => `<li>${icon[c.type] || '•'} ${esc(c.text)}</li>`).join('')}</ul></div>`);
  }
  html.push(`<div class="card"><h2>넣은 원장</h2><table class="fr-table"><tr><th>파일</th><th>사업장</th><th class="r">${r.month}월 차변 행</th><th class="r">제외(대변·결산)</th><th class="r">매출차감 0처리</th></tr>
    ${r.ledgerFiles.map((f) => `<tr><td>${esc(f.name)}</td><td class="sm">${esc((f.sites || []).join(', ') || f.error || '')}</td><td class="r">${f.taken ?? '-'}</td><td class="r">${f.credit ?? '-'}</td><td class="r">${f.deduct ?? '-'}</td></tr>`).join('')}</table></div>`);
  $('#frResult').innerHTML = html.join('');
  $('#frDown').onclick = () => download(out);
  $('#frOvApply').onclick = () => {
    document.querySelectorAll('.fr-ov').forEach((i) => {
      const v = Number(String(i.value).replace(/[^\d.-]/g, ''));
      if (String(i.value).trim() === '' || isNaN(v)) return;
      if (Math.round(v) !== +i.dataset.orig) state.overrides[i.dataset.key] = v;
    });
    run();
  };
  $('#frOvReset').onclick = () => { state.overrides = {}; run(); };
  $('#frRemap')?.addEventListener('click', () => {
    document.querySelectorAll('.fr-map').forEach((i) => { state.mapping[i.dataset.vendor] = { payee: i.value.trim(), co: i.dataset.co, lead: i.dataset.lead }; });
    run();
  });
}

export function initFreight() {
  bindDrop('#frAgg', '#frAggIn', (fs) => { if (fs[0]) state.agg = fs[0]; });
  bindDrop('#frLed', '#frLedIn', (fs) => { state.ledgers.push(...fs); });
  bindDrop('#frPrev', '#frPrevIn', (fs) => { if (fs[0]) state.prev = fs[0]; });
  $('#frLedName').addEventListener('click', (e) => { const b = e.target.closest('[data-rm]'); if (b) { e.stopPropagation(); state.ledgers.splice(+b.dataset.rm, 1); render(); } });
  $('#frRun').onclick = run;
  $('#frUseLast').onclick = async (e) => {
    e.stopPropagation();
    const last = await idbGet('last').catch(() => null);
    if (!last) { alert('이 기기에서 만든 검증파일이 아직 없습니다.'); return; }
    state.prev = new File([last.data], last.name);
    render();
  };
  idbGet('last').then((last) => { if (last) { $('#frUseLast').hidden = false; $('#frUseLast').textContent = `↺ 지난번 만든 파일 (${last.name})`; } }).catch(() => {});
  render();
}
