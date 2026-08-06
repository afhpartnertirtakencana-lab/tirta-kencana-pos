/* Tirta Kencana POS - Service Worker
   Strategy:
   - Same-origin app shell (html/css/js/icons): network-first, falling back to cache.
     Network-first (not cache-first) is deliberate: this app is updated frequently
     (v14+ and counting), so a stale cache-first SW would silently serve old code.
   - Cross-origin requests (Google Apps Script backend, fonts, CDN libs): always
     network, never intercepted. Transaction/stock data must never be served from
     cache, and CDN libs are already versioned in their URLs.
   - Bump CACHE_VERSION on every deploy so old caches are purged on activate.
*/

const CACHE_VERSION = 'tk-pos-v2';
const APP_SHELL = [
  './index.html',
  './css/style.css',
  './js/app.js',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-192-maskable.png',
  './icons/icon-512-maskable.png',
  './icons/apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Only handle GET requests on our own origin (app shell). Everything else
  // (GAS backend calls, Google Fonts, CDN libraries, cross-origin images)
  // passes straight through to the network untouched.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) {
    return;
  }

  event.respondWith(
    fetch(req)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE_VERSION).then((cache) => cache.put(req, copy)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(req).then((cached) => cached || caches.match('./index.html')))
  );
});
