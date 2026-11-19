// 📊 월마감 자동화 화면: 지난달 보고서 + ERP RD → 이번 달을 붙인 보고서
import { detectBook, buildArAp, buildInternal, buildPacking, buildEnergy, buildRawMaterial, buildCapex } from './closing.js';
import { loadXLSX } from './excel.js';

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const won = (n) => (Math.round(n || 0)).toLocaleString('ko-KR');

const KIND_LABEL = { arap: '채권·채무', prepay: '선급금', internal: '내부거래 집계', packing: '포장소모품비', energy: '전력비·수도광열비', rawmat: 'Raw Material', capex: 'CAPEX 재분류 검증' };
const kindOf = (kinds) => (kinds.includes('capex') ? 'capex' : kinds.includes('energy') ? 'energy' : kinds.includes('rawmat') ? 'rawmat' : kinds.includes('internal') ? 'internal' : kinds.includes('packing') ? 'packing' : kinds.includes('prepay') && !kinds.includes('ar') && !kinds.includes('ap') ? 'prepay' : kinds.some((k) => ['ar', 'ap', 'prepay'].includes(k)) ? 'arap' : null);

// 만든 파일은 기기에 보관 → 다음 달에 '지난번 파일'로 바로 사용
const DB = 'jywork-closing', ST = 'files';
const idb = () => new Promise((res, rej) => { const r = indexedDB.open(DB, 1); r.onupgradeneeded = () => r.result.createObjectStore(ST); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const idbPut = async (k, v) => { const d = await idb(); return new Promise((res, rej) => { const t = d.transaction(ST, 'readwrite'); t.objectStore(ST).put(v, k); t.oncomplete = res; t.onerror = () => rej(t.error); }); };
const idbAll = async () => { const d = await idb(); return new Promise((res, rej) => { const out = {}; const c = d.transaction(ST).objectStore(ST).openCursor(); c.onsuccess = () => { const cur = c.result; if (cur) { out[cur.key] = cur.value; cur.continue(); } else res(out); }; c.onerror = () => rej(c.error); }); };

// ── 백업자료 종류 목록 ──────────────────────────────────────────
// 새 자료를 추가할 때: 여기 한 줄 + closing.js에 만드는 함수(build) + index.html 드롭다운 <option> 하나
const TYPES = {
  arap: { build: () => buildArAp, prev: '① 지난달 채권·채무 파일', prevHint: '매출비교·매입비교 시트가 있는 파일', rd: '② ERP 거래처원장', rdHint: '전체계정 그대로 OK — 외상매출금·외상매입금만 골라 씀 · ⚠ 법인 범위는 지난달과 같게 (기초 잔액으로 자동 대조) · 대흥(비교표형)·송이애/군자(원장형) 모두 가능',
    desc: '매출채권·매입채무 월 시트 + 매출비교·매입비교 이번 달 열',
    rules: ['거래처원장에서 외상매출금(11110401)·외상매입금(12100101) 행만 골라 \'YY년MM월매출채권/매입채무\' 시트를 새로 만듦 (지난달 시트 서식 그대로)', '매출비교·매입비교에 이번 달 열을 끼워 넣음 — 기존 열·수식은 그대로, 오른쪽 차액 열만 이번 달 기준으로 옮김', '목록에 없는 거래처는 끝에 B열 \'신규\'로 추가 → 구분을 바꿔 주세요', '거래처원장에는 날짜가 없어 지난달 시트의 다음 달로 만듦', '이번 RD의 기초 잔액이 지난달 시트의 잔액과 다르면 \'법인 범위가 다를 수 있음\'으로 경고 (차이 큰 거래처 표시)'] },
  prepay: { build: () => buildArAp, prev: '① 지난달 선급금 파일', prevHint: '선급금증감 시트가 있는 파일', rd: '② ERP 거래처원장', rdHint: '전체계정 그대로 OK — 선급금만 골라 씀 · ⚠ 법인 범위는 지난달과 같게 (기초 잔액으로 자동 대조)',
    desc: '선급금 월 시트 + 선급금증감 이번 달 열',
    rules: ['거래처원장에서 선급금(11110800) 행만 골라 \'YY년MM월선급금\' 시트를 새로 만듦', '선급금증감에 이번 달 열 추가 (기존 열·수식 그대로), 새 거래처는 끝에 \'신규\'로 추가'] },
  internal: { build: () => buildInternal, prev: '① 지난달 내부거래 집계 파일', prevHint: '비용·수익·내부거래 집계 시트가 있는 파일', rd: '② ERP 계정별원장', rdHint: '전체계정 그대로 OK — 비용·수익 시트에 쓰던 계정만 골라 씀',
    desc: '비용·수익 시트 이번 달 행 + 내부거래 집계 이번 달 묶음',
    rules: ['비용·수익 시트 맨 아래에 이번 달 행만 붙임 — 최근 쓰던 계정 중 상대가 계열사(대흥·송이애·군자·중부/이천팜·SPC)인 행', 'D열(내용)·상대회사는 지난 기록에서 같은 거래를 찾아 채우고, 처음 보는 거래는 적요로 판단해 알려 줌', '\'내부거래 집계\' 맨 아래에 이번 달 묶음(25행) 추가', 'RD 날짜로 이번 달을 정하고, 같은 달을 다시 만들면 그 달 행만 새로 바꿈'] },
  packing: { build: () => buildPacking, prev: '① 지난달 포장소모품비 파일', prevHint: '수불o·수불x 시트가 있는 파일', rd: '② ERP RD', rdHint: '기간별입고현황 + 계정별원장 (전체계정 그대로 OK)',
    desc: '수불o·수불x 이번 달 행 + 회사별 시트 이번 달 열',
    rules: ['기간별입고현황 → 수불o, 계정별원장의 포장 계정(Mapping P열) → 수불x 맨 아래에 이번 달 행만 붙임', '회사별 시트에 이번 달 3칸(수불o·수불x 판관·제조) + 월별 요약 열 추가 (기존 열은 오른쪽으로 밀릴 뿐 수식·서식 그대로)', '\'포장소모품비 예산정리\' 시트는 손으로 넣는 칸이라 비워 둠', 'RD 날짜로 이번 달을 정하고, 같은 달을 다시 만들면 그 달 행만 새로 바꿈'] },
  energy: { build: () => buildEnergy, prev: '① 지난달 전력비·수도광열비 파일', prevHint: '전력비·가스 등유 벙커c유 등 시트가 있는 \'계획대비계정증감\' 파일', rd: '② 전력비 및 에너지 사용현황', rdHint: '담당자에게 받는 사용현황 파일 2개 (대흥 / 송이애·새송이·용인) 를 같이 넣기',
    desc: '전력비·가스(등유·벙커C유) 시트 이번 달 행 — 사용량·요금 칸 채움, 단가·합계는 기존 수식',
    rules: ['RD 사용현황 파일의 계량기 열을 지난 달 값이 같은 열끼리 자동으로 짝지어 이번 달 사용량·요금을 옮김 (열 순서·이름이 달라도 OK)', '단가·합계·Total 칸은 원래 수식 그대로 (값을 덮지 않음)', '이번 달 행이 없으면(새해 1월 등) 블록마다 맨 위에 한 줄 끼우고 위 행 수식을 복사', 'RD에서 짝을 못 찾은 계량기 열은 경고로 알려 줌 — 손으로 확인', '계약전력·bgt 시트는 그대로'] },
  rawmat: { build: () => buildRawMaterial, prev: '① 지난달 Raw Material 파일', prevHint: "'raw materials P' 시트가 있는 파일", rd: '② ERP 기간별입고현황', rdHint: '법인별 3개 (대흥·송이애·군자) 를 같이 넣기 — 구분 \'원재료\'만 골라 씀',
    desc: '물량(kg)·매입액 이번 달 열 채움 (단가·합계는 기존 수식)',
    rules: ['구분이 \'원재료\'인 입고를 입고공장으로 나눔: 팽이=디에이치아그로·디에이치팜·대흥 북부, 느타리(대흥)=대흥 남부, 느타리(송이애)=송이애(이천), 새송이=이천팜', '품목명으로 행을 맞춰 물량·금액 합계(반품 포함)를 넣음 — 톱밥처럼 \'포\' 단위는 C열 kg(16kg)로 =포수×16 수식', '표고(원/봉): 물량은 미배양 배지(이천 생산입고+원삼 구매)로 넣고, 금액은 생산분 원가가 RD에 없어 비워 둠(경고로 구매액 안내)', '파일에 없는 입고(군자 이천지점 톱밥·소맥피 등, 송이애 배지 대체)는 목록으로 알려 줌', '이번 달 열이 이미 있어야 함 — 새해(1월)는 엑셀에서 열을 먼저 만들어 주세요'] },
  capex: { build: () => buildCapex, prev: '① 지난달 CAPEX 재분류 검증 파일', prevHint: "'닥터애그(CAPEX Actual)' 시트가 있는 파일", rd: '② 대흥 통합 Capex Actual 파일', rdHint: "대흥에서 받는 '통합 Capex Actual' 파일 ('CAPEX 닥터애그' 시트를 읽음)",
    desc: '닥터애그(CAPEX Actual) 맨 아래에 이번 달 투자 행 추가 (송이애·군자는 KPI 보고 수기)',
    rules: ["대흥 파일 'CAPEX 닥터애그' 시트에서 이번 달 금액이 있는 행만 골라 검증파일 맨 아래에 붙임 (A 연도·B 월·E 구분1·F 구분2·I 투자내역)", "금액은 원 → 천원으로 바꿔 '{연도} Actual' 블록의 그 달 열에 넣음", '위쪽 합계·다른 시트의 수식은 그대로 (새 행까지 자동으로 계산)', '송이애·군자 Actual 시트는 각 법인 KPI 파일을 보고 손으로', '같은 달을 다시 만들면 그 달 행만 새로 바꿈'] },
  freight: { desc: '집계파일 + 계정별원장 + 지난 검증파일 → 이번 달 검증 시트(현황·운반비)를 붙인 검증파일' },
};
const pref = (k, d) => { try { return localStorage.getItem(k) || d; } catch { return d; } };
const setPref = (k, v) => { try { localStorage.setItem(k, v); } catch { /* 무시 */ } };

const state = { mode: 'freight', prev: null, rds: [], saved: {} };

function fileLabel(f) { return `${esc(f.name)} <span class="muted">(${Math.round((f.size || f.data?.length || 0) / 1024).toLocaleString()}KB)</span>`; }
function render() {
  const T = TYPES[state.mode];
  $('#autoDesc').textContent = T.desc || '';
  if (state.mode === 'freight') return;
  $('#clPrevTitle').textContent = T.prev; $('#clPrevHint').textContent = T.prevHint;
  $('#clRdTitle').textContent = T.rd; $('#clRdHint').textContent = T.rdHint;
  $('#clRules').innerHTML = T.rules.map((r) => `<li>${esc(r)}</li>`).join('');
  const p = state.prev;
  const bad = p && p.kind !== undefined && p.kind !== state.mode;
  $('#clPrevName').innerHTML = p
    ? `<span class="chipfile">${p.kind === undefined ? '<span class="muted">확인 중…</span> ' : bad ? '<b class="minus">⚠ 고른 자료의 파일이 아님</b> · ' : '✅ '}${fileLabel(p.file)} <button type="button" data-x-prev="1" title="빼기">✕</button></span>`
    : '<span class="muted">📂 끌어다 놓거나 눌러서 선택</span>';
  $('#clRdName').innerHTML = state.rds.length
    ? state.rds.map((f, i) => `<span class="chipfile">${fileLabel(f)} <button type="button" data-x-rd="${i}" title="빼기">✕</button></span>`).join('')
    : '<span class="muted">📂 여러 개 선택 가능</span>';
  const sv = state.saved[state.mode];
  $('#clSaved').innerHTML = sv ? `<button type="button" class="btn sm" data-use="${state.mode}">↺ 지난번 만든 파일 · ${esc(sv.name)}</button>` : '';
  const okPrev = !!(p && p.kind === state.mode);
  $('#clRun').disabled = !(okPrev && state.rds.length);
  $('#clPrev').classList.toggle('has', okPrev);
  $('#clPrev').classList.toggle('bad', !!bad);
  $('#clRd').classList.toggle('has', state.rds.length > 0);
  const miss = [!okPrev && '① 지난달 파일', !state.rds.length && '② RD'].filter(Boolean);
  $('#clHint').innerHTML = p && p.kind === undefined ? '파일 종류 확인 중…' : bad ? '<span class="minus">① 칸에 고른 자료의 파일을 넣어 주세요</span>' : miss.length ? `${miss.join(' · ')}을(를) 넣어 주세요` : '<span class="ok">✓ 준비 완료 — 파일 만들기를 누르세요</span>';
}

function setMode(m) {
  if (!TYPES[m]) m = 'freight';
  if (m !== state.mode) { state.prev = null; $('#clResult').innerHTML = ''; }
  state.mode = m;
  setPref('autoType', m);
  $('#autoType').value = m;
  $$('.auto-panel').forEach((el) => { el.hidden = el.dataset.panel !== (m === 'freight' ? 'freight' : 'closing'); });
  render();
}

async function setPrev(file) {
  const { default: JSZip } = await import('./vendor/jszip.mjs');
  const item = { file, kind: undefined };
  state.prev = item;
  render();
  try {
    const d = await detectBook(JSZip, await file.arrayBuffer());
    item.kind = kindOf(d.kinds);
  } catch { item.kind = null; }
  render();
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
  if (kind === 'energy' || kind === 'rawmat' || kind === 'capex') {
    return `<table class="fr-table"><tr><th>항목</th><th class="r">채운 칸</th><th>비고</th></tr>
      ${r.parts.map((p) => `<tr><td>${esc(p.label)}</td><td class="r">${p.rows}</td><td class="sm">${esc(p.note || (p.col ? `${p.col}열 · 금액 합계 ${won(p.total)}` : ''))}</td></tr>`).join('')}</table>`;
  }
  if (r.parts[0]?.kind === 'arapList') {
    return `<table class="fr-table"><tr><th>시트</th><th>달</th><th class="r">추가 행</th><th class="r">잔액 합계</th><th>비고</th></tr>
      ${r.parts.map((p) => `<tr><td>${esc(p.sheet)}</td><td>${esc(p.month)}</td><td class="r">${p.rows}</td><td class="r">${won(p.balance)}</td><td class="sm">${p.removed ? `기존 ${p.removed}행 교체` : ''}</td></tr>`).join('')}</table>`;
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
  $('#clResult').innerHTML = '<div class="card"><p class="muted">RD를 읽고 파일을 만드는 중입니다… (파일이 크면 10~20초 걸릴 수 있어요)</p></div>';
  outs.length = 0;
  const kind = state.mode;
  try {
    const [XLSX, { default: JSZip }] = await Promise.all([loadXLSX(), import('./vendor/jszip.mjs')]);
    const rdFiles = await Promise.all(state.rds.map(async (f) => ({ name: f.name, buf: await f.arrayBuffer() })));
    const prevBuf = await state.prev.file.arrayBuffer();
    const out = await TYPES[kind].build()({ XLSX, JSZip, prevBuf, rdFiles });
    outs.push(out);
    $('#clResult').innerHTML = resultCard(kind, state.prev.file.name, out, 0);
    window.dispatchEvent(new CustomEvent('backup:made', { detail: { type: kind, fileName: out.fileName } }));
    // 송이애 원장형은 대흥 채권·채무와 따로 보관 (RD 보관함의 'arapList' 칸)
    const slot = out.report.parts[0]?.kind === 'arapList' ? 'arapList' : kind;
    try { await idbPut(slot, { name: out.fileName, data: out.data, at: Date.now() }); state.saved[slot] = { name: out.fileName, data: out.data }; } catch { /* 저장 공간 */ }
  } catch (e) {
    console.error(e);
    $('#clResult').innerHTML = `<div class="card err-card"><h2>만들지 못했습니다</h2><p>${esc(e.message || e)}</p></div>`;
  } finally {
    btn.textContent = '파일 만들기';
    render();
  }
}

export function openAutoType(m) { setMode(m); }

export function initClosing() {
  if (!$('#view-auto')) return;
  $('#autoType').addEventListener('change', (e) => setMode(e.target.value));
  bindDrop('#clPrev', '#clPrevIn', (files) => { if (files[0]) setPrev(files[0]); });
  bindDrop('#clRd', '#clRdIn', (files) => { state.rds.push(...files); });
  document.addEventListener('click', (e) => {
    const t = e.target.closest('[data-x-prev],[data-x-rd],[data-use],[data-dl]');
    if (!t || !t.closest('#view-auto, #dlgMake')) return;
    if (t.dataset.xPrev !== undefined) { state.prev = null; render(); }
    else if (t.dataset.xRd !== undefined) { state.rds.splice(+t.dataset.xRd, 1); render(); }
    else if (t.dataset.use) {
      const v = state.saved[t.dataset.use];
      if (!v) return;
      state.prev = { file: new File([v.data], v.name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), kind: t.dataset.use };
      render();
    } else if (t.dataset.dl !== undefined) download(outs[+t.dataset.dl]);
  });
  $('#clRun').addEventListener('click', run);
  idbAll().then((all) => { state.saved = all || {}; render(); }).catch(() => render());
  setMode(pref('autoType', 'freight'));
}
