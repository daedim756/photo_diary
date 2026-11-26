// 앱 화면을 저장해 두고, 인터넷이 느리거나 끊겨도 열리게 함 (항상 새 버전을 먼저 시도)
const CACHE = 'photo-diary-v4';
const SHELL = ['./', 'index.html', 'style.css', 'app.js', 'onedrive.js', 'config.js', 'icon.svg', 'manifest.webmanifest'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return; // 지도·API 요청은 건드리지 않음
  e.respondWith(
    fetch(e.request, { cache: 'no-cache' }) // 브라우저에 보관된 옛 파일 대신 항상 최신 버전 확인
      .then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return res; })
      .catch(() => caches.match(e.request)),
  );
});
