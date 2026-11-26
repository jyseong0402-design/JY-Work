// 펜 필기 엔진 (갤럭시 탭 S펜 대응)
// - 필압에 따라 굵기 변화, 펜/형광펜/지우개, 실행취소·다시실행
// - 손바닥 인식 방지: 펜이 한 번 감지되면 손가락은 스크롤 전용
// - S펜 측면 버튼을 누른 채 쓰면 지우개
// - 좌표는 가로 1000 기준 논리 좌표로 저장 → 화면 크기/방향이 바뀌어도 그대로 유지

import { icon } from './icons.js';

export const PAGE_W = 1000;
export const PAGE_H = 1300;
export const LINE = 50; // 예전 기본 줄 간격(논리 단위) — 줄 간격 정보가 없는 기존 필기용
export const TARGET_LS_PX = 34; // 화면에서 보이는 줄 간격 목표(px): 메모장이 넓어져도 줄이 벌어지지 않음

// 쪽마다 줄 간격(ls, 논리 단위)을 기억합니다.
// - 빈 쪽: 지금 화면 폭에 맞춰 약 34px 간격이 되도록 계산
// - 글씨가 있는 쪽: 처음 정해진 간격을 유지 → 다른 기기·화면 방향에서도 글씨와 줄이 어긋나지 않음
export function pageLs(page, scale) {
  if (!page.strokes.length || !page.ls) {
    if (!page.strokes.length) page.ls = Math.round((TARGET_LS_PX / scale) * 10) / 10;
    else page.ls = LINE;
  }
  return page.ls;
}

const COLORS = [
  ['#1d2433', '검정'], ['#1f4fd1', '파랑'], ['#d23b3b', '빨강'], ['#1d8a4e', '초록'],
];
const HL_COLORS = [['#ffe14d', '노랑'], ['#7be495', '연두'], ['#ff9ec7', '분홍']];
const SIZES = [[2, '가늘게'], [3.5, '보통'], [6, '굵게']];

/* ---- 쪽에 넣은 사진: 획 목록에 { t:'img', pid, x, y, w, h } 로 들어감 (사진 자체는 photos.js) ---- */
let imgSource = null; // pid → dataURL (app.js가 연결)
const imgCache = new Map(); // pid → { img, fail, cbs }
export function setImageSource(fn) { imgSource = fn; }
function imgOf(pid, onload) {
  let e = imgCache.get(pid);
  if (e?.fail) { imgCache.delete(pid); e = null; } // 못 받았던 사진은 다음에 다시 시도
  if (!e) {
    e = { img: null, fail: false, cbs: new Set() };
    imgCache.set(pid, e);
    Promise.resolve(imgSource?.(pid)).then((src) => {
      if (!src) { e.fail = true; return; }
      const im = new Image();
      im.onload = () => { e.img = im; e.cbs.forEach((f) => f()); e.cbs.clear(); };
      im.onerror = () => { e.fail = true; };
      im.src = src;
    }).catch(() => { e.fail = true; });
  }
  if (!e.img && onload) e.cbs.add(onload);
  return e.img;
}
export const isImg = (s) => s.t === 'img';
function drawImg(ctx, s, scale, onload) {
  const img = imgOf(s.pid, onload);
  const [x, y, w, h] = [s.x * scale, s.y * scale, s.w * scale, s.h * scale];
  if (img) { ctx.drawImage(img, x, y, w, h); return; }
  ctx.save();
  ctx.fillStyle = 'rgba(120,120,120,.15)'; ctx.fillRect(x, y, w, h);
  ctx.strokeStyle = 'rgba(120,120,120,.5)'; ctx.setLineDash([6, 4]); ctx.strokeRect(x, y, w, h);
  ctx.fillStyle = '#888'; ctx.font = `${Math.max(12, 16 * scale)}px sans-serif`; ctx.textAlign = 'center';
  ctx.fillText('📷 사진 불러오는 중…', x + w / 2, y + h / 2);
  ctx.restore();
}
/* ---- 글상자: { t:'txt', x, y, w, h, text, c, fs } — 펜 필기 중에도 'T' 도구로 쪽 위에 글자 입력 ---- */
export const isTxt = (s) => s.t === 'txt';
export const isBox = (s) => s.t === 'img' || s.t === 'txt'; // 옮기기·크기 조절 도구로 다루는 것
export const TXT_LH = 1.35;
const TXT_FONT = (px) => `${px}px "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif`;
let measureCtx = null;
// 줄 나눔은 논리 좌표(가로 1000)에서 계산 → 화면 크기와 상관없이 같은 위치에서 줄이 바뀜
export function txtLines(s) {
  measureCtx ||= document.createElement('canvas').getContext('2d');
  measureCtx.font = TXT_FONT(s.fs);
  const out = [];
  for (const para of String(s.text || '').split('\n')) {
    let line = '';
    for (const ch of para) {
      if (line && measureCtx.measureText(line + ch).width > s.w) { out.push(line); line = ch.trim() ? ch : ''; } else line += ch;
    }
    out.push(line);
  }
  s.h = Math.max(s.fs * TXT_LH, out.length * s.fs * TXT_LH); // 고르기·지우개용 높이
  return out;
}
function drawText(ctx, s, scale) {
  const lines = txtLines(s);
  ctx.save();
  ctx.fillStyle = s.c;
  ctx.font = TXT_FONT(s.fs * scale);
  ctx.textBaseline = 'middle';
  lines.forEach((l, i) => ctx.fillText(l, s.x * scale, (s.y + (i + 0.5) * s.fs * TXT_LH) * scale));
  ctx.restore();
}

// 사진을 먼저(아래) 그리고 그 위에 필기·글상자(쓴 순서대로)
export function drawAll(ctx, strokes, scale, onload) {
  for (const s of strokes) if (isImg(s)) drawImg(ctx, s, scale, onload);
  for (const s of strokes) if (!isImg(s)) drawStroke(ctx, s, scale);
}

export function drawStroke(ctx, s, scale) {
  if (isImg(s)) return;
  if (isTxt(s)) { if (!s.editing) drawText(ctx, s, scale); return; }
  const p = s.p;
  if (!p || p.length < 3) return;
  ctx.save();
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = s.c;
  if (s.t === 'hl') {
    ctx.globalAlpha = 0.38;
    ctx.lineWidth = s.w * scale;
    ctx.beginPath();
    ctx.moveTo(p[0] * scale, p[1] * scale);
    for (let i = 3; i < p.length; i += 3) ctx.lineTo(p[i] * scale, p[i + 1] * scale);
    if (p.length === 3) ctx.lineTo(p[0] * scale + 0.1, p[1] * scale);
    ctx.stroke();
  } else {
    if (p.length === 3) {
      ctx.fillStyle = s.c;
      ctx.beginPath();
      ctx.arc(p[0] * scale, p[1] * scale, (s.w * (0.35 + p[2])) * scale / 2, 0, Math.PI * 2);
      ctx.fill();
    }
    for (let i = 3; i < p.length; i += 3) {
      const pr = (p[i - 1] + p[i + 2]) / 2;
      ctx.lineWidth = Math.max(0.6, s.w * (0.35 + pr) * scale);
      ctx.beginPath();
      ctx.moveTo(p[i - 3] * scale, p[i - 2] * scale);
      ctx.lineTo(p[i] * scale, p[i + 1] * scale);
      ctx.stroke();
    }
  }
  ctx.restore();
}

// 읽기 전용 미리보기(상세 화면)용
export function renderInkPreview(container, data) {
  container.innerHTML = '';
  for (const page of data?.pages || []) {
    if (!page.strokes.length) continue;
    const wrap = document.createElement('div');
    wrap.className = 'paper ink-page preview';
    const cv = document.createElement('canvas');
    wrap.appendChild(cv);
    container.appendChild(wrap);
    const paint = () => {
      const w = wrap.clientWidth;
      if (!w) return;
      const scale = w / PAGE_W;
      const dpr = window.devicePixelRatio || 1;
      wrap.style.height = PAGE_H * scale + 'px';
      setPaperVars(wrap, scale, page.ls || LINE);
      cv.width = Math.round(w * dpr);
      cv.height = Math.round(PAGE_H * scale * dpr);
      const ctx = cv.getContext('2d');
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      drawAll(ctx, page.strokes, scale, paint);
    };
    requestAnimationFrame(paint);
  }
}

export function setPaperVars(el, scale, ls = LINE) {
  el.style.setProperty('--ls', ls * scale + 'px');
  el.style.setProperty('--mg', ls * 2.2 * scale + 'px'); // 빨간 여백선도 줄 간격에 비례
}

export class InkPad {
  constructor(host, opts = {}) {
    this.host = host;
    this.opts = opts;
    this.onChange = opts.onChange;
    this.pages = [{ strokes: [] }];
    this.undoStack = [];
    this.redoStack = [];
    this.tool = { t: 'pen', c: COLORS[0][0], w: SIZES[1][0], hc: HL_COLORS[0][0] };
    this.penSeen = localStorage.getItem('jy-pen-seen') === '1';
    this.fingerDraw = !this.penSeen;
    this.dirty = false;
    this.build();
  }

  // 도구는 상단 도구줄(left: 쪽추가·되돌리기, center: 펜·형광펜·지우개·색, menu: 더보기)에 그립니다.
  build() {
    const { left, center, menu } = this.opts;
    this.host.innerHTML = '<div class="ink-pages"></div>';
    this.pagesEl = this.host.querySelector('.ink-pages');
    left.innerHTML = `
      <button type="button" class="tbb" data-act="page" title="쪽 추가">${icon('plus')}</button>
      <button type="button" class="tbb" data-act="undo" title="실행취소">${icon('undo')}</button>
      <button type="button" class="tbb" data-act="redo" title="다시실행">${icon('redo')}</button>`;
    center.innerHTML = `
      <button type="button" class="tbb tool" data-tool="pen" title="펜">${icon('pen')}</button>
      <button type="button" class="tbb tool" data-tool="hl" title="형광펜">${icon('hl')}</button>
      <button type="button" class="tbb tool" data-tool="eraser" title="지우개 (S펜 버튼을 누르고 문질러도 지워짐)">${icon('eraser')}</button>
      <button type="button" class="tbb tool" data-tool="text" title="글자 넣기 (쪽을 누르고 입력, 글상자를 누르면 고치기)">${icon('tbox')}</button>
      <button type="button" class="tbb tool" data-tool="select" title="사진·글상자 옮기기·크기 조절 (누르세요)">${icon('select')}</button>
      <span class="pal-wrap">
        <button type="button" class="tbb cur-color" data-act="palette" title="색·굵기"><i></i></button>
        <div class="palette" role="dialog" aria-label="색·굵기">
          <div class="pal-row pen-only">${COLORS.map(([c, n]) => `<button type="button" class="sw" data-color="${c}" title="${n}" style="--sw:${c}"></button>`).join('')}</div>
          <div class="pal-row pen-only">${SIZES.map(([w, n]) => `<button type="button" class="sz" data-size="${w}" title="${n}"><i style="--d:${w * 2 + 2}px"></i></button>`).join('')}</div>
          <div class="pal-row hl-only">${HL_COLORS.map(([c, n]) => `<button type="button" class="sw" data-hcolor="${c}" title="${n}" style="--sw:${c}"></button>`).join('')}</div>
        </div>
      </span>`;
    this.toolEls = [left, center, menu].filter(Boolean);
    this.toolEls.forEach((el) => el.addEventListener('click', (e) => this.onTool(e)));
    this.palette = center.querySelector('.palette');
    // 팔레트 바깥을 누르면 닫기
    document.addEventListener('pointerdown', (e) => {
      if (!e.target.closest('.pal-wrap')) this.palette.classList.remove('open');
    });
    this.ro = new ResizeObserver(() => this.layout());
    this.ro.observe(this.pagesEl);
    this.syncTools();
  }

  onTool(e) {
    const b = e.target.closest('button');
    if (!b) return;
    this.commitText();
    const d = b.dataset;
    if (d.tool) this.tool.t = d.tool;
    if (d.color) { this.tool.c = d.color; if (this.tool.t !== 'text') this.tool.t = 'pen'; }
    if (d.hcolor) { this.tool.hc = d.hcolor; this.tool.t = 'hl'; }
    if (d.size) this.tool.w = +d.size;
    if (this.tool.t !== 'select') this.deselect();
    if (d.act === 'palette') {
      if (this.tool.t === 'eraser' || this.tool.t === 'select') this.tool.t = 'pen'; // 글자 도구는 그대로(글자 색·크기)
      this.palette.classList.toggle('open');
    }
    if (d.act === 'undo') this.undo();
    if (d.act === 'redo') this.redo();
    if (d.act === 'page') { this.addPage(true); this.pagesEl.lastElementChild?.scrollIntoView({ behavior: 'smooth' }); }
    if (d.act === 'finger') this.fingerDraw = !this.fingerDraw;
    if (d.act === 'clear' && !this.isEmpty() && confirm('이 메모의 필기를 모두 지울까요?')) {
      this.push({ type: 'clear', before: this.pages });
      this.pages = [{ strokes: [] }];
      this.renderPages();
      this.changed();
    }
    this.syncTools();
  }

  syncTools() {
    const q = (sel) => this.toolEls.flatMap((el) => [...el.querySelectorAll(sel)]);
    q('[data-tool]').forEach((b) => b.classList.toggle('on', b.dataset.tool === this.tool.t));
    q('[data-color]').forEach((b) => b.classList.toggle('on', b.dataset.color === this.tool.c));
    q('[data-hcolor]').forEach((b) => b.classList.toggle('on', b.dataset.hcolor === this.tool.hc));
    q('[data-size]').forEach((b) => b.classList.toggle('on', +b.dataset.size === this.tool.w));
    this.opts.center.dataset.tool = this.tool.t;
    if (this.pagesEl) this.pagesEl.dataset.tool = this.tool.t;
    const cur = q('.cur-color i')[0];
    if (cur) cur.style.background = this.tool.t === 'hl' ? this.tool.hc : this.tool.c;
    q('[data-act=finger]').forEach((f) => {
      f.classList.toggle('on', this.fingerDraw);
      const lbl = f.querySelector('.lbl');
      if (lbl) lbl.textContent = this.fingerDraw ? '켜짐' : '꺼짐 (손가락은 스크롤)';
    });
    q('[data-act=undo]').forEach((x) => { x.disabled = !this.undoStack.length; });
    q('[data-act=redo]').forEach((x) => { x.disabled = !this.redoStack.length; });
  }

  /* ---- 데이터 ---- */
  load(data) {
    if (this.editing) { this.editing.ta.remove(); this.editing.grip?.remove(); this.editing = null; } // 다른 메모로 바뀜: 입력 중이던 글상자는 버림
    this.pages = data?.pages?.length ? data.pages.map((p) => ({ strokes: [...p.strokes], ls: p.ls || (p.strokes.length ? LINE : 0) })) : [{ strokes: [] }];
    this.undoStack = [];
    this.redoStack = [];
    this.dirty = false;
    this.renderPages();
    this.syncTools();
    this.pagesEl.scrollTop = 0;
  }
  getData() {
    // 뒤쪽의 빈 페이지는 저장하지 않음
    const pages = this.pages.map((p) => ({ strokes: p.strokes, ls: p.ls }));
    while (pages.length > 1 && !pages[pages.length - 1].strokes.length) pages.pop();
    return { pages };
  }
  isEmpty() { return this.pages.every((p) => !p.strokes.length); }
  pageCount() { return this.getData().pages.filter((p) => p.strokes.length).length; }
  changed() { this.dirty = true; this.syncTools(); this.onChange?.(); }

  // AI 인식용 이미지(흰 배경, PNG base64)
  toImages() {
    const scale = 1.2;
    return this.pages.filter((p) => p.strokes.length).map((p) => {
      const cv = document.createElement('canvas');
      cv.width = PAGE_W * scale;
      cv.height = PAGE_H * scale;
      const ctx = cv.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, cv.width, cv.height);
      drawAll(ctx, p.strokes, scale);
      return cv.toDataURL('image/png').split(',')[1];
    });
  }

  /* ---- 실행취소 ---- */
  push(action) { this.undoStack.push(action); this.redoStack = []; if (this.undoStack.length > 300) this.undoStack.shift(); }
  apply(a, reverse) {
    this.deselect();
    if (a.type === 'imgmove' || a.type === 'txtedit') {
      Object.assign(a.s, reverse ? a.from : a.to);
      this.redraw(a.page);
    } else if (a.type === 'imgpage') {
      if (reverse) {
        const t = this.pages[a.to].strokes, j = t.indexOf(a.s);
        if (j >= 0) t.splice(j, 1);
        this.pages[a.from].strokes.splice(a.i, 0, a.s);
        if (a.created && a.to === this.pages.length - 1 && !this.pages[a.to].strokes.length) { this.pages.pop(); this.renderPages(); }
      } else {
        if (a.created && a.to >= this.pages.length) { this.pages.push({ strokes: [] }); this.renderPages(); }
        const f = this.pages[a.from].strokes, j = f.indexOf(a.s);
        if (j >= 0) f.splice(j, 1);
        this.pages[a.to].strokes.push(a.s);
      }
      this.redraw(a.from); if (this.pages[a.to]) this.redraw(a.to);
    } else if (a.type === 'add') {
      const list = this.pages[a.page].strokes;
      if (reverse) list.splice(list.indexOf(a.stroke), 1); else list.push(a.stroke);
      this.redraw(a.page);
    } else if (a.type === 'erase') {
      const list = this.pages[a.page].strokes;
      if (reverse) a.removed.slice().sort((x, y) => x.i - y.i).forEach(({ s, i }) => list.splice(i, 0, s));
      else a.removed.forEach(({ s }) => list.splice(list.indexOf(s), 1));
      this.redraw(a.page);
    } else if (a.type === 'page') {
      // at 없는 예전 기록 = 맨 끝 쪽
      if (reverse) this.pages.splice(a.at ?? this.pages.length - 1, 1);
      else this.pages.splice(a.at ?? this.pages.length, 0, { strokes: [] });
      this.renderPages();
    } else if (a.type === 'clear') {
      if (reverse) { a.after = this.pages; this.pages = a.before; } else { this.pages = a.after || [{ strokes: [] }]; }
      this.renderPages();
    }
  }
  undo() { this.commitText(); const a = this.undoStack.pop(); if (!a) return; this.apply(a, true); this.redoStack.push(a); this.changed(); }
  redo() { this.commitText(); const a = this.redoStack.pop(); if (!a) return; this.apply(a, false); this.undoStack.push(a); this.changed(); }

  // at 위치(0=맨 앞)에 빈 쪽을 끼움. 되돌리기는 순서대로라 다른 기록의 쪽 번호와 어긋나지 않음
  addPage(record, at = this.pages.length) {
    this.commitText();
    at =Math.max(0, Math.min(at, this.pages.length));
    this.pages.splice(at, 0, { strokes: [] });
    if (record) this.push({ type: 'page', at });
    this.renderPages();
    this.changed();
    return at;
  }

  /* ---- 화면 ---- */
  /* ---- 사진 넣기·고르기 ---- */
  imageIds() {
    const ids = new Set();
    for (const p of this.pages) for (const s of p.strokes) if (isImg(s)) ids.add(s.pid);
    return ids;
  }
  // 지금 보고 있는 쪽의 보이는 위치에 넣고, 바로 옮기기·크기 조절 상태로
  insertImage(pid, w0 = 800, h0 = 600, { select = true } = {}) {
    const els = [...this.pagesEl.children];
    const box = this.pagesEl.getBoundingClientRect();
    let page = 0;
    els.forEach((el, i) => { const r = el.getBoundingClientRect(); if (r.top <= box.top + box.height / 2) page = i; });
    const r = els[page]?.getBoundingClientRect();
    const k = Math.min(1, 700 / w0, 700 / h0);
    let w = w0 * k, h = h0 * k;
    if (!this.scale) this.scale = 1;
    let y = r ? Math.max(40, (box.top - r.top) / this.scale + 40) : 40;
    if (y + h > PAGE_H - 20) y = Math.max(20, PAGE_H - 20 - h);
    if (y + h > PAGE_H - 20) { const f = (PAGE_H - 40) / h; w *= f; h *= f; y = 20; }
    const s = { t: 'img', pid, x: Math.round((PAGE_W - w) / 2), y: Math.round(y), w: Math.round(w), h: Math.round(h) };
    this.pages[page].strokes.push(s);
    this.push({ type: 'add', page, stroke: s });
    this.redraw(page);
    this.changed();
    if (select) { this.tool.t = 'select'; this.syncTools(); this.selectImg(page, s); }
    return s;
  }
  deselect() { this.sel?.el.remove(); this.sel = null; }

  /* ---- 글상자 찾기 ---- */
  // 찾기용 목록: [{ p: 쪽 번호(0부터), t: 글자 }]
  textBoxes() {
    const out = [];
    this.pages.forEach((pg, p) => pg.strokes.forEach((s) => { if (isTxt(s) && s.text) out.push({ p, t: s.text }); }));
    return out;
  }
  // 찾는 말이 든 글상자로 이동해 잠깐 반짝임 (띄어쓰기 무시)
  reveal(q) {
    const words = String(q || '').toLowerCase().split(/\s+/).filter(Boolean);
    if (!words.length) return false;
    const has = (t) => { const a = t.toLowerCase(), b = a.replace(/\s+/g, ''); return words.some((w) => a.includes(w) || b.includes(w.replace(/\s+/g, ''))); };
    for (let p = 0; p < this.pages.length; p++) {
      const s = this.pages[p].strokes.find((x) => isTxt(x) && x.text && has(x.text));
      if (!s) continue;
      const go = () => {
        const host = this.pagesEl.children[p];
        if (!host || !this.scale) return;
        txtLines(s);
        this.pagesEl.scrollTop = Math.max(0, host.offsetTop + s.y * this.scale - this.pagesEl.clientHeight / 3);
        const el = document.createElement('div');
        el.className = 'txt-hit';
        Object.assign(el.style, { left: s.x * this.scale - 6 + 'px', top: s.y * this.scale - 4 + 'px', width: s.w * this.scale + 12 + 'px', height: s.h * this.scale + 8 + 'px' });
        host.appendChild(el);
        setTimeout(() => el.remove(), 2600);
      };
      requestAnimationFrame(() => requestAnimationFrame(go)); // 쪽 크기가 정해진 뒤
      return true;
    }
    return false;
  }

  /* ---- 글상자 입력·고치기 ---- */
  // 펜 굵기(가늘게·보통·굵게)를 글자 크기로 사용
  txtSize() { return { 2: 22, 3.5: 28, 6: 40 }[this.tool.w] || 28; }
  // 글자 도구로 누름: 글상자 위 = 고치기, 입력 중에 바깥 = 입력 끝, 빈 곳 = 새 글상자
  txtHit(page, [x, y]) {
    return [...(this.pages[page]?.strokes || [])].reverse().find((s) => isTxt(s) && x >= s.x && x <= s.x + s.w && y >= s.y && y <= s.y + (s.h || s.fs * TXT_LH));
  }
  // 글상자 위치 (쪽 밖으로 너무 나가지 않게)
  placeTxt(s, x, y) {
    s.x = Math.round(Math.max(-s.w / 2, Math.min(PAGE_W - 40, x)));
    s.y = Math.round(Math.max(0, Math.min(PAGE_H - s.fs, y)));
  }
  textAt(page, [x, y], wasEditing = !!this.editing) {
    const hit = this.txtHit(page, [x, y]);
    if (hit && hit === this.editing?.s) return; // 입력 중인 상자 안
    if (hit) { this.editText(page, hit); return; }
    if (wasEditing) { this.commitText(); return; }
    this.newText(page, x, y);
  }
  newText(page, x, y) {
    const fs = this.txtSize();
    let w = Math.min(560, PAGE_W - x - 20);
    if (w < 160) { w = 160; x = PAGE_W - 180; }
    const s = { t: 'txt', x: Math.round(Math.max(0, x)), y: Math.round(Math.max(0, Math.min(PAGE_H - fs * 2, y - fs * 0.7))), w: Math.round(w), h: fs * TXT_LH, text: '', c: this.tool.c, fs };
    this.editText(page, s, true);
  }
  editText(page, s, isNew = false) {
    this.commitText();
    this.deselect();
    const host = this.pagesEl.children[page];
    if (!host) return;
    const ta = document.createElement('textarea');
    ta.className = 'ink-txt';
    ta.value = s.text || '';
    ta.placeholder = '글자 입력';
    ta.spellcheck = false;
    const sc = this.scale;
    Object.assign(ta.style, {
      left: s.x * sc + 'px', top: s.y * sc + 'px', width: s.w * sc + 'px',
      fontSize: s.fs * sc + 'px', lineHeight: String(TXT_LH), color: s.c,
    });
    const fit = () => { ta.style.height = 'auto'; ta.style.height = Math.max(s.fs * TXT_LH * sc, ta.scrollHeight) + 'px'; };
    ta.addEventListener('input', fit);
    ta.addEventListener('pointerdown', (e) => e.stopPropagation());
    ta.addEventListener('keydown', (e) => { if (e.key === 'Escape') ta.blur(); e.stopPropagation(); });
    ta.addEventListener('blur', () => this.commitText());
    host.appendChild(ta);
    // 입력 중 왼쪽 위 ✥ 손잡이를 끌면 글상자 옮기기 (손가락도 됨)
    const grip = document.createElement('span');
    grip.className = 'ink-txt-grip';
    grip.textContent = '✥';
    grip.title = '끌어서 옮기기';
    const placeGrip = () => Object.assign(grip.style, { left: s.x * this.scale + 'px', top: s.y * this.scale + 'px' });
    placeGrip();
    grip.addEventListener('pointerdown', (e) => {
      e.preventDefault(); e.stopPropagation(); // 입력칸 포커스 유지
      grip.setPointerCapture(e.pointerId);
      const from = { x: s.x, y: s.y }, sx = e.clientX, sy = e.clientY;
      const mv = (ev) => {
        this.placeTxt(s, from.x + (ev.clientX - sx) / this.scale, from.y + (ev.clientY - sy) / this.scale);
        Object.assign(ta.style, { left: s.x * this.scale + 'px', top: s.y * this.scale + 'px' });
        placeGrip();
      };
      const end = () => {
        grip.removeEventListener('pointermove', mv);
        grip.removeEventListener('pointerup', end);
        grip.removeEventListener('pointercancel', end);
        // 이미 쪽에 있던 글상자만 되돌리기 기록 (새 글상자는 입력 끝날 때 최종 위치로 들어감)
        if (!isNew && (s.x !== from.x || s.y !== from.y)) { this.push({ type: 'imgmove', page, s, from, to: { x: s.x, y: s.y } }); this.changed(); }
        ta.focus({ preventScroll: true });
      };
      grip.addEventListener('pointermove', mv);
      grip.addEventListener('pointerup', end);
      grip.addEventListener('pointercancel', end);
    });
    host.appendChild(grip);
    s.editing = true;
    this.redraw(page);
    this.editing = { page, s, ta, grip, isNew, from: s.text || '' };
    fit();
    ta.focus({ preventScroll: true });
    if (!isNew) ta.setSelectionRange(ta.value.length, ta.value.length);
  }
  commitText() {
    const e = this.editing;
    if (!e) return;
    this.editing = null;
    const { page, s, ta, grip, isNew, from } = e;
    ta.remove();
    grip?.remove();
    delete s.editing;
    const text = ta.value.replace(/\s+$/, '');
    const list = this.pages[page]?.strokes;
    if (!list) return;
    if (isNew) {
      if (text) { s.text = text; list.push(s); this.push({ type: 'add', page, stroke: s }); this.changed(); }
    } else if (!text) {
      const i = list.indexOf(s);
      if (i >= 0) { list.splice(i, 1); this.push({ type: 'erase', page, removed: [{ s, i }] }); this.changed(); }
    } else if (text !== from) {
      s.text = text;
      this.push({ type: 'txtedit', page, s, from: { text: from }, to: { text } });
      this.changed();
    }
    this.redraw(page);
  }
  // 사진을 다른 쪽으로: 마지막 쪽에서 아래로 보내면 새 쪽을 만듦
  moveImgPage(from, s, to) {
    if (to < 0) return;
    const list = this.pages[from].strokes, i = list.indexOf(s);
    if (i < 0) return;
    const created = to >= this.pages.length;
    if (created) { this.pages.push({ strokes: [] }); this.renderPages(); }
    list.splice(i, 1);
    this.pages[to].strokes.push(s);
    this.push({ type: 'imgpage', s, from, to, i, created });
    this.deselect();
    this.redraw(from); this.redraw(to);
    this.changed();
    this.selectImg(to, s);
    this.pagesEl.children[to]?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  selectImg(page, s) {
    this.deselect();
    const host = this.pagesEl.children[page];
    if (!host) return;
    const el = document.createElement('div');
    el.className = 'img-sel';
    const txt = isTxt(s);
    if (txt) txtLines(s); // 높이 최신으로
    el.innerHTML = `<span class="img-bar">${page > 0 ? '<button type="button" data-a="up" title="이전 쪽으로 옮기기">⬆</button>' : ''}<button type="button" data-a="down" title="다음 쪽으로 옮기기 (마지막 쪽이면 새 쪽)">⬇</button>${txt ? '<button type="button" data-a="edit" title="글자 고치기">✏️</button>' : '<button type="button" data-a="zoom" title="크게 보기·저장">🔍</button>'}<button type="button" data-a="del" title="${txt ? '글상자 지우기' : '사진 빼기'}">🗑</button></span><i class="rs" title="${txt ? '폭 조절' : '크기 조절'}"></i>`;
    host.appendChild(el);
    this.sel = { el, page, s };
    const place = () => Object.assign(el.style, { left: s.x * this.scale + 'px', top: s.y * this.scale + 'px', width: s.w * this.scale + 'px', height: s.h * this.scale + 'px' });
    this.sel.place = place;
    place();
    const zb = el.querySelector('[data-a=zoom]');
    if (zb) zb.onclick = () => this.opts.onImageOpen?.(s.pid);
    const eb = el.querySelector('[data-a=edit]');
    if (eb) eb.onclick = () => this.editText(page, s);
    // 다른 쪽으로 옮기기(위치·크기는 그대로)
    for (const [a, d] of [['up', -1], ['down', 1]]) {
      const b = el.querySelector(`[data-a=${a}]`);
      if (b) b.onclick = () => this.moveImgPage(page, s, page + d);
    }
    el.querySelector('[data-a=del]').onclick = () => {
      const list = this.pages[page].strokes;
      const i = list.indexOf(s);
      if (i < 0) return;
      list.splice(i, 1);
      this.push({ type: 'erase', page, removed: [{ s, i }] });
      this.deselect();
      this.redraw(page);
      this.changed();
    };
    // 끌어서 옮기기 / 오른쪽 아래 모서리로 크기 조절(비율 유지)
    el.addEventListener('pointerdown', (e) => {
      if (e.target.closest('button')) return;
      e.preventDefault(); e.stopPropagation();
      el.setPointerCapture(e.pointerId);
      const resize = !!e.target.closest('.rs');
      const from = { x: s.x, y: s.y, w: s.w, h: s.h };
      const sx = e.clientX, sy = e.clientY;
      const mv = (ev) => {
        const dx = (ev.clientX - sx) / this.scale, dy = (ev.clientY - sy) / this.scale;
        if (resize && txt) { // 글상자는 폭만 바꾸고 높이는 줄 수에 맞춤
          s.w = Math.round(Math.max(80, Math.min(PAGE_W - from.x, from.w + dx)));
          txtLines(s);
        } else if (resize) {
          const w = Math.max(60, Math.min(PAGE_W - from.x, from.w + Math.max(dx, dy * from.w / from.h)));
          s.w = Math.round(w); s.h = Math.round(w * from.h / from.w);
        } else {
          s.x = Math.round(Math.max(-s.w / 2, Math.min(PAGE_W - s.w / 2, from.x + dx)));
          s.y = Math.round(Math.max(0, Math.min(PAGE_H - 40, from.y + dy)));
        }
        place();
        this.redraw(page);
      };
      const end = () => {
        el.removeEventListener('pointermove', mv);
        el.removeEventListener('pointerup', end);
        el.removeEventListener('pointercancel', end);
        const to = { x: s.x, y: s.y, w: s.w, h: s.h };
        if (JSON.stringify(to) !== JSON.stringify(from)) { this.push({ type: 'imgmove', page, s, from, to }); this.changed(); }
      };
      el.addEventListener('pointermove', mv);
      el.addEventListener('pointerup', end);
      el.addEventListener('pointercancel', end);
    });
  }

  renderPages() {
    this.commitText();
    this.deselect();
    this.pagesEl.innerHTML = '';
    this.pages.forEach((_, i) => {
      const el = document.createElement('div');
      el.className = 'paper ink-page';
      const last = i === this.pages.length - 1;
      el.innerHTML = `<canvas class="base"></canvas><canvas class="live"></canvas><span class="pno">${i + 1}</span>`
        + `<button type="button" class="pg-ins" data-at="${i}" title="${i ? `${i}쪽과 ${i + 1}쪽 사이에` : '맨 앞에'} 빈 쪽 끼우기">＋ 쪽</button>`
        + (last ? `<button type="button" class="pg-ins end" data-at="${i + 1}" title="맨 뒤에 빈 쪽 추가">＋ 쪽</button>` : '');
      el.querySelectorAll('.pg-ins').forEach((b) => {
        b.addEventListener('pointerdown', (e) => e.stopPropagation());
        b.onclick = (e) => {
          e.stopPropagation();
          const at = this.addPage(true, +b.dataset.at);
          this.pagesEl.children[at]?.scrollIntoView({ behavior: 'smooth', block: 'center' });
        };
      });
      this.pagesEl.appendChild(el);
      this.bindPage(el, i);
    });
    this.layout(true);
  }

  layout(force) {
    const w = this.pagesEl.clientWidth - 16;
    if (w <= 0) return;
    const landscape = window.innerWidth > window.innerHeight;
    const cssW = Math.round(Math.min(w, window.innerWidth * (landscape ? 0.8 : 0.97), 1600));
    if (!force && cssW === this.cssW) return;
    this.cssW = cssW;
    this.scale = cssW / PAGE_W;
    this.dpr = window.devicePixelRatio || 1;
    [...this.pagesEl.children].forEach((el, i) => {
      el.style.width = cssW + 'px';
      el.style.height = PAGE_H * this.scale + 'px';
      setPaperVars(el, this.scale, pageLs(this.pages[i], this.scale));
      el.querySelectorAll('canvas').forEach((cv) => {
        cv.width = Math.round(cssW * this.dpr);
        cv.height = Math.round(PAGE_H * this.scale * this.dpr);
        cv.style.width = cssW + 'px';
        cv.style.height = PAGE_H * this.scale + 'px';
        cv.getContext('2d').setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
      });
      this.redraw(i);
    });
    this.sel?.place();
  }

  ctxOf(i, live) {
    const el = this.pagesEl.children[i];
    return el?.querySelector(live ? 'canvas.live' : 'canvas.base').getContext('2d');
  }

  redraw(i) {
    const ctx = this.ctxOf(i);
    if (!ctx || !this.scale) return;
    ctx.clearRect(0, 0, PAGE_W * this.scale + 1, PAGE_H * this.scale + 1);
    if (this.pages[i]) drawAll(ctx, this.pages[i].strokes, this.scale, () => this.redraw(i));
  }

  /* ---- 입력 ---- */
  bindPage(el, pageIndex) {
    const live = el.querySelector('canvas.live');
    let cur = null; // 그리는 중인 획
    let mode = null; // 'draw' | 'erase' | 'scroll'
    let lastY = 0;
    let removed = [];
    let tapAt = null; // 글자 도구에서 손가락으로 누른 곳(움직이지 않고 떼면 입력)
    let drag = null; // 글자 도구에서 펜·마우스로 끄는 글상자

    const pt = (e) => {
      const r = live.getBoundingClientRect();
      const x = Math.round(((e.clientX - r.left) / this.scale) * 10) / 10;
      const y = Math.round(((e.clientY - r.top) / this.scale) * 10) / 10;
      const pr = e.pointerType === 'pen' ? Math.round((e.pressure || 0.5) * 100) / 100 : 0.5;
      return [x, y, pr];
    };

    const eraseAt = ([x, y]) => {
      const list = this.pages[pageIndex].strokes;
      const r = 14;
      let hit = false;
      for (let i = list.length - 1; i >= 0; i--) {
        const s = list[i];
        if (isImg(s)) continue; // 사진은 지우개로 지우지 않음 (사진 옮기기 도구의 🗑)
        if (isTxt(s)) { // 글상자는 문지른 곳이 상자 안이면 통째로 지움
          if (x >= s.x && x <= s.x + s.w && y >= s.y && y <= s.y + (s.h || s.fs * TXT_LH)) { removed.push({ s, i }); list.splice(i, 1); hit = true; }
          continue;
        }
        const rr = r + s.w / 2;
        for (let k = 0; k < s.p.length; k += 3) {
          if (Math.abs(s.p[k] - x) < rr && Math.abs(s.p[k + 1] - y) < rr) {
            removed.push({ s, i });
            list.splice(i, 1);
            hit = true;
            break;
          }
        }
      }
      if (hit) this.redraw(pageIndex);
    };

    const down = (e) => {
      this.palette?.classList.remove('open');
      if (e.pointerType === 'pen' && !this.penSeen) {
        this.penSeen = true;
        this.fingerDraw = false;
        try { localStorage.setItem('jy-pen-seen', '1'); } catch { /* 무시 */ }
        this.syncTools();
      }
      if (mode) return; // 이미 다른 포인터로 작업 중 (멀티터치 무시)
      if (this.tool.t === 'text') { // 글자 넣기: 글상자를 누르면 고치기, 빈 곳은 새 글상자(손가락은 살짝 누르면 입력, 끌면 스크롤)
        const p = pt(e);
        if (e.pointerType === 'touch') {
          mode = 'scroll';
          lastY = e.clientY;
          tapAt = { p, x: e.clientX, y: e.clientY, was: !!this.editing };
          live.setPointerCapture(e.pointerId);
          this.activeId = e.pointerId;
          return;
        }
        e.preventDefault();
        // 펜·마우스로 글상자를 누른 채 끌면 옮기기, 그냥 떼면 고치기
        const hit = this.txtHit(pageIndex, p);
        if (hit) {
          if (this.editing) this.commitText();
          mode = 'txtdrag';
          drag = { s: hit, from: { x: hit.x, y: hit.y }, sx: e.clientX, sy: e.clientY, p, moved: false };
          live.setPointerCapture(e.pointerId);
          this.activeId = e.pointerId;
          return;
        }
        this.textAt(pageIndex, p);
        return;
      }
      if (this.tool.t === 'select') { // 사진·글상자 고르기: 누르면 선택, 빈 곳은 선택 해제(손가락은 스크롤)
        const [x, y] = pt(e);
        const list = this.pages[pageIndex].strokes;
        const hit = [...list].reverse().find((s) => isBox(s) && x >= s.x && x <= s.x + s.w && y >= s.y && y <= s.y + s.h);
        if (hit) { e.preventDefault(); this.selectImg(pageIndex, hit); return; }
        this.deselect();
        if (e.pointerType !== 'touch') return;
      }
      if (e.pointerType === 'touch' && (!this.fingerDraw || this.tool.t === 'select')) {
        mode = 'scroll';
        lastY = e.clientY;
        live.setPointerCapture(e.pointerId);
        this.activeId = e.pointerId;
        return;
      }
      if (e.button !== 0 && e.button !== 2 && e.button !== 5) return;
      e.preventDefault();
      live.setPointerCapture(e.pointerId);
      this.activeId = e.pointerId;
      // S펜 측면 버튼(buttons=2) 또는 지우개 끝(buttons=32)은 지우개
      const erasing = this.tool.t === 'eraser' || (e.pointerType === 'pen' && (e.buttons & 2 || e.buttons & 32));
      if (erasing) {
        mode = 'erase';
        removed = [];
        eraseAt(pt(e));
        return;
      }
      mode = 'draw';
      const isHl = this.tool.t === 'hl';
      cur = { t: isHl ? 'hl' : 'pen', c: isHl ? this.tool.hc : this.tool.c, w: isHl ? 22 : this.tool.w, p: pt(e) };
      this.liveDraw(live, cur);
    };

    const move = (e) => {
      if (!mode || e.pointerId !== this.activeId) return;
      if (mode === 'scroll') {
        this.pagesEl.scrollTop -= e.clientY - lastY;
        lastY = e.clientY;
        if (tapAt && Math.hypot(e.clientX - tapAt.x, e.clientY - tapAt.y) > 10) tapAt = null;
        return;
      }
      e.preventDefault();
      if (mode === 'txtdrag') {
        const dx = e.clientX - drag.sx, dy = e.clientY - drag.sy;
        if (!drag.moved && Math.hypot(dx, dy) < 5) return;
        drag.moved = true;
        this.placeTxt(drag.s, drag.from.x + dx / this.scale, drag.from.y + dy / this.scale);
        this.redraw(pageIndex);
        return;
      }
      const evs = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
      for (const ev of evs.length ? evs : [e]) {
        const p = pt(ev);
        if (mode === 'erase') { eraseAt(p); continue; }
        const n = cur.p.length;
        if (Math.hypot(p[0] - cur.p[n - 3], p[1] - cur.p[n - 2]) < 0.8) continue;
        cur.p.push(...p);
      }
      if (mode === 'draw') this.liveDraw(live, cur);
    };

    const up = (e) => {
      if (!mode || e.pointerId !== this.activeId) return;
      if (mode === 'scroll' && tapAt && e.type === 'pointerup') { const { p, was } = tapAt; tapAt = null; mode = null; this.activeId = null; this.textAt(pageIndex, p, was); return; }
      tapAt = null;
      if (mode === 'txtdrag') {
        const d = drag; drag = null; mode = null; this.activeId = null;
        if (d.moved) { this.push({ type: 'imgmove', page: pageIndex, s: d.s, from: d.from, to: { x: d.s.x, y: d.s.y } }); this.changed(); } else this.editText(pageIndex, d.s);
        return;
      }
      if (mode === 'erase' && removed.length) {
        this.push({ type: 'erase', page: pageIndex, removed });
        this.changed();
      }
      if (mode === 'draw' && cur) {
        this.pages[pageIndex].strokes.push(cur);
        this.push({ type: 'add', page: pageIndex, stroke: cur });
        live.getContext('2d').clearRect(0, 0, live.width, live.height);
        drawStroke(this.ctxOf(pageIndex), cur, this.scale);
        // 마지막 쪽 아래쪽까지 쓰면 새 쪽을 자동으로 추가
        const maxY = Math.max(...cur.p.filter((_, k) => k % 3 === 1));
        if (pageIndex === this.pages.length - 1 && maxY > PAGE_H - 140) this.addPage(true);
        this.changed();
      }
      cur = null;
      mode = null;
      this.activeId = null;
    };

    live.addEventListener('pointerdown', down);
    live.addEventListener('pointermove', move);
    live.addEventListener('pointerup', up);
    live.addEventListener('pointercancel', up);
    live.addEventListener('contextmenu', (e) => e.preventDefault());
  }

  liveDraw(live, s) {
    const ctx = live.getContext('2d');
    ctx.clearRect(0, 0, PAGE_W * this.scale + 1, PAGE_H * this.scale + 1);
    drawStroke(ctx, s, this.scale);
  }
}
