/* ホーム画面に追加したときのためのサービスワーカー。
   - 画面のファイルは常に最新を取りに行き、圏外のときだけ前回の内容を使う。予約のやり取り(Google)は触らない。
   - 通知(プッシュ): 中身なしの通知が届いたら、予約サーバーに「何のお知らせか」を聞いて表示し、アイコンのバッジも更新する。 */
const CACHE = 'wash-v1';
self.window = self; // config.js(window.WASH_CONFIG = …)をここでも読めるように
try { importScripts('config.js'); } catch (e) {}
const API = (self.WASH_CONFIG && self.WASH_CONFIG.api) || 'https://script.google.com/macros/s/AKfycbzHj9e15-2k4kkNa63ctM0P2KBXVbKfzGJGvhBNKXl0Er2u5_75qR7Kmev8Z0TzTdcg/exec';

self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(fetch(req).then(res => {
    const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res;
  }).catch(() => caches.match(req).then(r => r || caches.match('./'))));
});

/** アイコンのバッジ(数字)。0 なら消す。対応していない端末では何もしない */
function setBadge(n) {
  try {
    if (n > 0 && self.navigator.setAppBadge) return self.navigator.setAppBadge(n);
    if (self.navigator.clearAppBadge) return self.navigator.clearAppBadge();
  } catch (e) {}
  return Promise.resolve();
}

/** アプリが画面に見えているか(見えている間に届いたお知らせは、すぐ既読にして数字を付けない) */
async function visibleApps() {
  const list = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
  return list.filter(c => c.visibilityState === 'visible');
}
self.addEventListener('push', e => e.waitUntil((async () => {
  let messages = [], badge = 0;
  const open = await visibleApps().catch(() => []);
  try {
    const sub = await self.registration.pushManager.getSubscription();
    if (sub) {
      const r = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'pushInbox', endpoint: sub.endpoint, seen: open.length > 0 }) });
      const j = await r.json();
      if (j.ok) { messages = j.data.messages || []; badge = j.data.badge || 0; }
    }
  } catch (err) { /* 聞けなかったときも、下の決まった文で必ず通知を出す(出さないと、端末が通知を止めることがある) */ }
  if (!messages.length) messages = [{ title: '洗車予約', body: '新しいお知らせがあります。アプリを開いてご確認ください。', tag: 'wash' }];
  for (const m of messages) {
    await self.registration.showNotification(m.title || '洗車予約', { body: m.body || '', tag: m.tag || 'wash', icon: 'icon-192.png', badge: 'icon-192.png', data: { url: './' } });
  }
  if (open.length) {
    // アプリを見ている最中: 通知はすぐ片付け、数字も付けず、画面にお知らせして案内を出してもらう
    const shown = await self.registration.getNotifications().catch(() => []);
    shown.forEach(n => n.close());
    open.forEach(c => c.postMessage({ type: 'push-arrived' }));
    await setBadge(0);
    return;
  }
  await setBadge(badge);
})()));

/* 通知を押したら、アプリを開く(開いていれば前に出す) */
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil((async () => {
    const list = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    const app = list.find(c => new URL(c.url).pathname.replace(/index\.html$/, '') === new URL('./', self.registration.scope).pathname);
    if (app) { await app.focus(); app.postMessage({ type: 'push-open' }); return; }
    await self.clients.openWindow('./');
  })());
});
