// 메모 공유·내보내기: 필기 쪽을 노트 종이 그대로 이미지(PNG)·긴 이미지 한 장·PDF로 만들어
// 📤 공유(탭·폰 공유 창: 카톡·메일·드라이브 …) 또는 ⬇ 다운로드. 모두 기기 안에서 만들고 서버로 보내지 않음.
import { PAGE_W, PAGE_H, LINE, drawAll, drawPaper, loadImages } from './ink.js';

const SCALE = 1.6; // 1600×2080px (선명하게)

export async function renderPage(page, scale = SCALE) {
  await loadImages(page.strokes);
  const cv = document.createElement('canvas');
  cv.width = Math.round(PAGE_W * scale);
  cv.height = Math.round(PAGE_H * scale);
  const ctx = cv.getContext('2d');
  drawPaper(ctx, scale, page.ls || LINE);
  drawAll(ctx, page.strokes, scale);
  return cv;
}

const toBlob = (cv, type, q) => new Promise((res, rej) => cv.toBlob((b) => (b ? res(b) : rej(new Error('이미지를 만들지 못했습니다'))), type, q));

// 여러 쪽을 세로로 이어 붙인 한 장 (너무 길면 줄여서 높이 16000px 이하)
async function longImage(pages) {
  const gap = 24;
  const scale = Math.min(SCALE, (16000 + gap) / (pages.length * (PAGE_H + gap)));
  const cvs = [];
  for (const p of pages) cvs.push(await renderPage(p, scale));
  const out = document.createElement('canvas');
  out.width = cvs[0].width;
  out.height = cvs.reduce((t, c) => t + c.height, 0) + gap * (cvs.length - 1);
  const ctx = out.getContext('2d');
  ctx.fillStyle = '#d9d4b0';
  ctx.fillRect(0, 0, out.width, out.height);
  let y = 0;
  for (const c of cvs) { ctx.drawImage(c, 0, y); y += c.height + gap; }
  return out;
}

// JPEG 쪽들로 PDF 만들기 (외부 라이브러리 없이) — 쪽 크기 595×773.5pt(종이 비율 1000:1300)
export async function buildPdf(jpegs, title = '') {
  const enc = new TextEncoder();
  const parts = [];
  const offs = [];
  let len = 0;
  const add = (x) => { const b = typeof x === 'string' ? enc.encode(x) : x; parts.push(b); len += b.length; };
  const obj = (n, body) => { offs[n] = len; add(`${n} 0 obj\n`); body(); add('\nendobj\n'); };
  const W = 595, H = 773.5;
  const n = jpegs.length;
  // 1 카탈로그, 2 쪽 목록, 3 정보, 그다음 쪽마다 [쪽, 내용, 그림]
  const pageNo = (i) => 4 + i * 3;
  add('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');
  obj(1, () => add('<< /Type /Catalog /Pages 2 0 R >>'));
  obj(2, () => add(`<< /Type /Pages /Count ${n} /Kids [${jpegs.map((_, i) => `${pageNo(i)} 0 R`).join(' ')}] >>`));
  // 제목은 UTF-16BE(한글)로
  const hex = [...String(title)].map((ch) => { const c = ch.codePointAt(0); return c > 0xffff ? '003F' : c.toString(16).padStart(4, '0'); }).join('');
  obj(3, () => add(`<< /Title <FEFF${hex}> /Producer (Dr.AG) >>`));
  jpegs.forEach((j, i) => {
    const p = pageNo(i);
    obj(p, () => add(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${W} ${H}] /Resources << /XObject << /Im0 ${p + 2} 0 R >> >> /Contents ${p + 1} 0 R >>`));
    const cs = `q ${W} 0 0 ${H} 0 0 cm /Im0 Do Q`;
    obj(p + 1, () => add(`<< /Length ${cs.length} >>\nstream\n${cs}\nendstream`));
    obj(p + 2, () => { add(`<< /Type /XObject /Subtype /Image /Width ${j.w} /Height ${j.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${j.bytes.length} >>\nstream\n`); add(j.bytes); add('\nendstream'); });
  });
  const total = 4 + n * 3;
  const xref = len;
  add(`xref\n0 ${total}\n0000000000 65535 f \n`);
  for (let k = 1; k < total; k++) add(`${String(offs[k]).padStart(10, '0')} 00000 n \n`);
  add(`trailer\n<< /Size ${total} /Root 1 0 R /Info 3 0 R >>\nstartxref\n${xref}\n%%EOF`);
  return new Blob(parts, { type: 'application/pdf' });
}

const safe = (s) => String(s || '메모').replace(/[\\/:*?"<>|\n\r]+/g, '_').trim().slice(0, 60) || '메모';

// kind: 'png'(쪽마다) | 'long'(한 장) | 'pdf' → File[]
export async function makeFiles(pages, kind, base) {
  const name = safe(base);
  if (kind === 'pdf') {
    const jpegs = [];
    for (const p of pages) {
      const cv = await renderPage(p);
      const b = await toBlob(cv, 'image/jpeg', 0.9);
      jpegs.push({ w: cv.width, h: cv.height, bytes: new Uint8Array(await b.arrayBuffer()) });
    }
    return [new File([await buildPdf(jpegs, base)], `${name}.pdf`, { type: 'application/pdf' })];
  }
  if (kind === 'long') return [new File([await toBlob(await longImage(pages), 'image/png')], `${name}.png`, { type: 'image/png' })];
  const out = [];
  for (let i = 0; i < pages.length; i++) {
    out.push(new File([await toBlob(await renderPage(pages[i]), 'image/png')], pages.length > 1 ? `${name}_${i + 1}쪽.png` : `${name}.png`, { type: 'image/png' }));
  }
  return out;
}

export const canShareFiles = (files) => {
  try { return !!navigator.canShare && navigator.canShare({ files }); } catch { return false; }
};
export const shareSupported = () => canShareFiles([new File(['x'], 'x.png', { type: 'image/png' })]);

export async function downloadFiles(files) {
  for (const f of files) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(f);
    a.download = f.name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 30000);
    if (files.length > 1) await new Promise((r) => setTimeout(r, 350)); // 여러 장 연속 저장
  }
}
