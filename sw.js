// 오프라인 지원용 서비스워커: 앱 파일은 네트워크 우선(최신 유지), 끊기면 캐시 사용
const CACHE = 'jywork-v4';
const ASSETS = ['./', './index.html', './css/style.css', './js/app.js', './js/store.js',
  './js/summarize.js', './js/ai.js', './js/excel.js', './js/ink.js', './js/inkdb.js', './js/sync.js',
  './js/firebase-config.js', './js/vendor/firebase.mjs', './js/vendor/xlsx.full.min.js',
  './js/vendor/anthropic-sdk.mjs', './manifest.webmanifest', './icons/icon.svg'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  // 외부 요청(AI, 로그인, 동기화 서버)은 서비스워커가 건드리지 않음
  if (new URL(req.url).origin !== location.origin) return;
  e.respondWith(fetch(req).then((res) => {
    const copy = res.clone();
    caches.open(CACHE).then((c) => c.put(req, copy));
    return res;
  }).catch(() => caches.match(req, { ignoreSearch: true })));
});
