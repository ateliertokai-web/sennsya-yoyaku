/* ホーム画面に追加したときのためのサービスワーカー。
   画面のファイルは常に最新を取りに行き、圏外のときだけ前回の内容を使う。予約のやり取り(Google)は触らない。 */
const CACHE = 'wash-v1';
self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(fetch(req).then(res => {
    const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res;
  }).catch(() => caches.match(req).then(r => r || caches.match('./'))));
});
