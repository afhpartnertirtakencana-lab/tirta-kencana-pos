// Service Worker minimal untuk Tirta Kencana POS
// Sengaja TIDAK melakukan caching agresif (pernah menyebabkan masalah cache sebelumnya).
// Tujuan utama: memenuhi syarat "installability" PWA di Chrome/Android
// sehingga tombol "Install" / "Add to Home Screen" bisa muncul.

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

// Passthrough sederhana - tidak menyimpan cache apa pun,
// jadi tidak akan menyebabkan halaman menampilkan versi lama.
self.addEventListener('fetch', (event) => {
  event.respondWith(
    fetch(event.request).catch(() => new Response('', { status: 503, statusText: 'Network error' }))
  );
});
