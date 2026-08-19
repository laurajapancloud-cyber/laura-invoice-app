const CACHE_NAME = 'laura-v10';
// ローカル資産のみプリキャッシュ（CDN依存を排除。CDNが不通だと install 全体が失敗していた）
const ASSETS_TO_CACHE = [
  '/static/icons/icon-192.png',
  '/static/icons/icon-512.png',
  '/static/icons/maskable-512.png',
  '/static/icons/apple-touch-icon.png',
  '/static/vendor/xlsx.full.min.js',
  '/static/vendor/alpine.min.js',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      // 1件の失敗でインストール全体が失敗しないよう個別に追加
      Promise.allSettled(ASSETS_TO_CACHE.map((url) => cache.add(url)))
    )
  );
  // NOTE: skipWaiting() は廃止。作業中のタブの資産をデプロイ途中で差し替えない。
  // 新バージョンは全タブを閉じた後に有効化される。
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) =>
      Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return; // 外部リソースには介入しない

  // API と HTML ナビゲーション: Network-first
  if (url.pathname.startsWith('/api/') || event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request).then(response => {
        // 成功した 200 OK のHTMLのみキャッシュ（リダイレクトはキャッシュしない）
        if (event.request.mode === 'navigate' && response.status === 200) {
          const resClone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, resClone));
        }
        return response;
      }).catch(() => caches.match(event.request))
    );
    return;
  }

  // 静的資産: Stale-While-Revalidate（表示は即時、裏で更新。古い資産が永久に残る問題を解消）
  event.respondWith(
    caches.match(event.request).then((cached) => {
      const fetchAndUpdate = fetch(event.request).then((fetchResponse) => {
        if (fetchResponse && fetchResponse.status === 200) {
          const clone = fetchResponse.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return fetchResponse;
      }).catch(() => cached);
      return cached || fetchAndUpdate;
    })
  );
});
