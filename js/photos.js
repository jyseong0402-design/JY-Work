// 메모 사진: 기기에는 IndexedDB 'jywork-photos'(key=사진 id), 클라우드는 sync.js가 users/{uid}/photos 에 조각으로 저장
// 사진 기록 = { id, mid(메모 id), name, data(JPEG dataURL), w, h, at, synced }
const DB = 'jywork-photos';
const STORE = 'p';
let dbp = null;

function db() {
  if (!dbp) {
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  return dbp;
}
async function tx(mode, fn) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const t = d.transaction(STORE, mode);
    const req = fn(t.objectStore(STORE));
    t.oncomplete = () => resolve(req?.result);
    t.onerror = () => reject(t.error);
  });
}
export const photodb = {
  get: (id) => tx('readonly', (s) => s.get(id)),
  put: (rec) => tx('readwrite', (s) => s.put(rec, rec.id)),
  del: (id) => tx('readwrite', (s) => s.delete(id)),
  all: () => tx('readonly', (s) => s.getAll()),
};

function loadImg(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('사진을 열 수 없습니다: ' + file.name)); };
    img.src = url;
  });
}

// 동기화 용량을 줄이려고 긴 변 1600px JPEG로 줄임 (글씨가 있는 사진도 읽을 수 있는 크기)
export async function shrinkImage(file, max = 1600, quality = 0.82) {
  let src;
  try { src = await createImageBitmap(file, { imageOrientation: 'from-image' }); } catch { src = await loadImg(file); }
  const s = Math.min(1, max / Math.max(src.width, src.height));
  const w = Math.max(1, Math.round(src.width * s)), h = Math.max(1, Math.round(src.height * s));
  const cv = document.createElement('canvas');
  cv.width = w; cv.height = h;
  const g = cv.getContext('2d');
  g.fillStyle = '#fff'; g.fillRect(0, 0, w, h); // 투명 배경(PNG 캡처)은 흰색으로
  g.drawImage(src, 0, 0, w, h);
  src.close?.();
  return { data: cv.toDataURL('image/jpeg', quality), w, h };
}
