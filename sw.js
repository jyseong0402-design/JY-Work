// 오프라인 지원용 서비스워커: 앱 파일은 캐시 우선, 외부 CDN은 네트워크 우선
const CACHE = 'jywork-v2';
const ASSETS = ['./', './index.html', './css/style.css', './js/app.js', './js/store.js',
  './js/summarize.js', './js/ai.js', './js/excel.js', './js/ink.js', './js/inkdb.js', './js/vendor/xlsx.full.min.js', './js/vendor/anthropic-sdk.mjs', './manifest.webmanifest', './icons/icon.svg'];

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
  const url = new URL(req.url);
  if (url.hostname === 'api.anthropic.com') return;
  if (url.origin === location.origin) {
    e.respondWith(fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req)));
  } else {
    e.respondWith(fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req)));
  }
});
