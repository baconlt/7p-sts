// Minimal service worker — required for PWA install prompt
// Caches the shell so the app icon works offline

const CACHE = '7p-sts-v3';
const SHELL = ['/', '/index.html', '/manifest.json', '/icon-192.png', '/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Only cache shell files — let the iframe load fresh from Apps Script
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});
