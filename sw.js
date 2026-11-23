// 讓「加入主畫面」後在沒有網路時也能開啟。
// 策略：先用快取立即開啟，有網路時在背景更新（下次開啟就是新版）。
// 修改程式後請把版本號 +1，舊快取才會被清掉。
const CACHE = 'street-signup-v4';
const ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', (e) => {
  // cache: 'reload' 確保抓到伺服器上的最新版，而不是瀏覽器 HTTP 快取裡的舊版
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(ASSETS.map((u) => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;

  e.respondWith(
    caches.open(CACHE).then(async (cache) => {
      const cached = await cache.match(req, { ignoreSearch: true }) ||
        (req.mode === 'navigate' ? await cache.match('./index.html') : undefined);
      const fresh = fetch(new Request(req.url, { cache: 'no-cache' })).then((res) => {
        if (res.ok) cache.put(req, res.clone());
        return res;
      }).catch(() => undefined);
      if (cached) {
        e.waitUntil(fresh);
        return cached;
      }
      return (await fresh) || new Response('離線中，且尚未快取此頁面', { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
    })
  );
});
