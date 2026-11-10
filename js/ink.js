// 펜 필기 엔진 (갤럭시 탭 S펜 대응)
// - 필압에 따라 굵기 변화, 펜/형광펜/지우개, 실행취소·다시실행
// - 손바닥 인식 방지: 펜이 한 번 감지되면 손가락은 스크롤 전용
// - S펜 측면 버튼을 누른 채 쓰면 지우개
// - 좌표는 가로 1000 기준 논리 좌표로 저장 → 화면 크기/방향이 바뀌어도 그대로 유지

export const PAGE_W = 1000;
export const PAGE_H = 1300;
export const LINE = 50; // 줄 간격(논리 단위)
export const MARGIN = 110; // 왼쪽 빨간 여백선 위치(논리 단위)

const COLORS = [
  ['#1d2433', '검정'], ['#1f4fd1', '파랑'], ['#d23b3b', '빨강'], ['#1d8a4e', '초록'],
];
const HL_COLORS = [['#ffe14d', '노랑'], ['#7be495', '연두'], ['#ff9ec7', '분홍']];
const SIZES = [[2, '가늘게'], [3.5, '보통'], [6, '굵게']];

export function drawStroke(ctx, s, scale) {
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
      setPaperVars(wrap, scale);
      cv.width = Math.round(w * dpr);
      cv.height = Math.round(PAGE_H * scale * dpr);
      const ctx = cv.getContext('2d');
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      page.strokes.forEach((s) => drawStroke(ctx, s, scale));
    };
    requestAnimationFrame(paint);
  }
}

export function setPaperVars(el, scale) {
  el.style.setProperty('--ls', LINE * scale + 'px');
  el.style.setProperty('--mg', MARGIN * scale + 'px');
}

export class InkPad {
  constructor(host, { onChange } = {}) {
    this.host = host;
    this.onChange = onChange;
    this.pages = [{ strokes: [] }];
    this.undoStack = [];
    this.redoStack = [];
    this.tool = { t: 'pen', c: COLORS[0][0], w: SIZES[1][0], hc: HL_COLORS[0][0] };
    this.penSeen = localStorage.getItem('jy-pen-seen') === '1';
    this.fingerDraw = !this.penSeen;
    this.dirty = false;
    this.build();
  }

  build() {
    this.host.innerHTML = `
      <div class="ink-tools" role="toolbar" aria-label="필기 도구">
        <div class="tg">
          <button type="button" data-tool="pen" title="펜">✒️<span>펜</span></button>
          <button type="button" data-tool="hl" title="형광펜">🖍️<span>형광펜</span></button>
          <button type="button" data-tool="eraser" title="지우개 (S펜 버튼 누르고 쓰기)">🧽<span>지우개</span></button>
        </div>
        <div class="tg colors pen-only">${COLORS.map(([c, n]) => `<button type="button" class="sw" data-color="${c}" title="${n}" style="--sw:${c}"></button>`).join('')}</div>
        <div class="tg colors hl-only">${HL_COLORS.map(([c, n]) => `<button type="button" class="sw" data-hcolor="${c}" title="${n}" style="--sw:${c}"></button>`).join('')}</div>
        <div class="tg pen-only">${SIZES.map(([w, n]) => `<button type="button" class="sz" data-size="${w}" title="${n}"><i style="--d:${w * 2 + 2}px"></i></button>`).join('')}</div>
        <div class="tg">
          <button type="button" data-act="undo" title="실행취소">↶</button>
          <button type="button" data-act="redo" title="다시실행">↷</button>
        </div>
        <span class="grow"></span>
        <div class="tg">
          <button type="button" data-act="finger" class="finger" title="손가락으로 쓰기">☝️<span>손가락</span></button>
          <button type="button" data-act="page" title="페이지 추가">＋<span>쪽</span></button>
          <button type="button" data-act="clear" title="이 메모의 필기 모두 지우기">🗑️</button>
        </div>
      </div>
      <div class="ink-pages"></div>`;
    this.toolsEl = this.host.querySelector('.ink-tools');
    this.pagesEl = this.host.querySelector('.ink-pages');
    this.toolsEl.addEventListener('click', (e) => this.onTool(e));
    this.ro = new ResizeObserver(() => this.layout());
    this.ro.observe(this.pagesEl);
    this.syncTools();
  }

  onTool(e) {
    const b = e.target.closest('button');
    if (!b) return;
    const d = b.dataset;
    if (d.tool) this.tool.t = d.tool;
    if (d.color) { this.tool.c = d.color; this.tool.t = 'pen'; }
    if (d.hcolor) { this.tool.hc = d.hcolor; this.tool.t = 'hl'; }
    if (d.size) this.tool.w = +d.size;
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
    const q = (s) => this.toolsEl.querySelectorAll(s);
    q('[data-tool]').forEach((b) => b.classList.toggle('on', b.dataset.tool === this.tool.t));
    q('[data-color]').forEach((b) => b.classList.toggle('on', b.dataset.color === this.tool.c));
    q('[data-hcolor]').forEach((b) => b.classList.toggle('on', b.dataset.hcolor === this.tool.hc));
    q('[data-size]').forEach((b) => b.classList.toggle('on', +b.dataset.size === this.tool.w));
    this.toolsEl.dataset.tool = this.tool.t;
    const f = this.toolsEl.querySelector('.finger');
    f.classList.toggle('on', this.fingerDraw);
    f.title = this.fingerDraw ? '손가락으로 쓰기: 켜짐' : '손가락으로 쓰기: 꺼짐 (손가락은 스크롤)';
    q('[data-act=undo]')[0].disabled = !this.undoStack.length;
    q('[data-act=redo]')[0].disabled = !this.redoStack.length;
  }

  /* ---- 데이터 ---- */
  load(data) {
    this.pages = data?.pages?.length ? data.pages.map((p) => ({ strokes: [...p.strokes] })) : [{ strokes: [] }];
    this.undoStack = [];
    this.redoStack = [];
    this.dirty = false;
    this.renderPages();
    this.syncTools();
    this.pagesEl.scrollTop = 0;
  }
  getData() {
    // 뒤쪽의 빈 페이지는 저장하지 않음
    const pages = this.pages.map((p) => ({ strokes: p.strokes }));
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
      p.strokes.forEach((s) => drawStroke(ctx, s, scale));
      return cv.toDataURL('image/png').split(',')[1];
    });
  }

  /* ---- 실행취소 ---- */
  push(action) { this.undoStack.push(action); this.redoStack = []; if (this.undoStack.length > 300) this.undoStack.shift(); }
  apply(a, reverse) {
    if (a.type === 'add') {
      const list = this.pages[a.page].strokes;
      if (reverse) list.splice(list.indexOf(a.stroke), 1); else list.push(a.stroke);
      this.redraw(a.page);
    } else if (a.type === 'erase') {
      const list = this.pages[a.page].strokes;
      if (reverse) a.removed.slice().sort((x, y) => x.i - y.i).forEach(({ s, i }) => list.splice(i, 0, s));
      else a.removed.forEach(({ s }) => list.splice(list.indexOf(s), 1));
      this.redraw(a.page);
    } else if (a.type === 'page') {
      if (reverse) this.pages.pop(); else this.pages.push({ strokes: [] });
      this.renderPages();
    } else if (a.type === 'clear') {
      if (reverse) { a.after = this.pages; this.pages = a.before; } else { this.pages = a.after || [{ strokes: [] }]; }
      this.renderPages();
    }
  }
  undo() { const a = this.undoStack.pop(); if (!a) return; this.apply(a, true); this.redoStack.push(a); this.changed(); }
  redo() { const a = this.redoStack.pop(); if (!a) return; this.apply(a, false); this.undoStack.push(a); this.changed(); }

  addPage(record) {
    this.pages.push({ strokes: [] });
    if (record) this.push({ type: 'page' });
    this.renderPages();
    this.changed();
  }

  /* ---- 화면 ---- */
  renderPages() {
    this.pagesEl.innerHTML = '';
    this.pages.forEach((_, i) => {
      const el = document.createElement('div');
      el.className = 'paper ink-page';
      el.innerHTML = `<canvas class="base"></canvas><canvas class="live"></canvas><span class="pno">${i + 1}</span>`;
      this.pagesEl.appendChild(el);
      this.bindPage(el, i);
    });
    this.layout(true);
  }

  layout(force) {
    const w = this.pagesEl.clientWidth - 24;
    if (w <= 0) return;
    const cssW = Math.min(w, 1100);
    if (!force && cssW === this.cssW) return;
    this.cssW = cssW;
    this.scale = cssW / PAGE_W;
    this.dpr = window.devicePixelRatio || 1;
    [...this.pagesEl.children].forEach((el, i) => {
      el.style.width = cssW + 'px';
      el.style.height = PAGE_H * this.scale + 'px';
      setPaperVars(el, this.scale);
      el.querySelectorAll('canvas').forEach((cv) => {
        cv.width = Math.round(cssW * this.dpr);
        cv.height = Math.round(PAGE_H * this.scale * this.dpr);
        cv.style.width = cssW + 'px';
        cv.style.height = PAGE_H * this.scale + 'px';
        cv.getContext('2d').setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
      });
      this.redraw(i);
    });
  }

  ctxOf(i, live) {
    const el = this.pagesEl.children[i];
    return el?.querySelector(live ? 'canvas.live' : 'canvas.base').getContext('2d');
  }

  redraw(i) {
    const ctx = this.ctxOf(i);
    if (!ctx || !this.scale) return;
    ctx.clearRect(0, 0, PAGE_W * this.scale + 1, PAGE_H * this.scale + 1);
    this.pages[i].strokes.forEach((s) => drawStroke(ctx, s, this.scale));
  }

  /* ---- 입력 ---- */
  bindPage(el, pageIndex) {
    const live = el.querySelector('canvas.live');
    let cur = null; // 그리는 중인 획
    let mode = null; // 'draw' | 'erase' | 'scroll'
    let lastY = 0;
    let removed = [];

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
      if (e.pointerType === 'pen' && !this.penSeen) {
        this.penSeen = true;
        this.fingerDraw = false;
        try { localStorage.setItem('jy-pen-seen', '1'); } catch { /* 무시 */ }
        this.syncTools();
      }
      if (mode) return; // 이미 다른 포인터로 작업 중 (멀티터치 무시)
      if (e.pointerType === 'touch' && !this.fingerDraw) {
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
        return;
      }
      e.preventDefault();
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
