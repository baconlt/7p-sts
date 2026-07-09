// Minimal service worker — required for PWA install prompt
// Caches the shell so the app icon works offline

const CACHE = '7p-sts-v6';
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
  const req = e.request;

  // Navigations / the HTML shell: network-first so a fresh wrapper always wins
  // when online (prevents devices getting stuck on a stale shell). Fall back to
  // cache only when offline.
  if (req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html')) {
    e.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match(req).then(cached => cached || caches.match('/')))
    );
    return;
  }

  // Other shell assets (icons, manifest): cache-first — let the iframe load
  // fresh from Apps Script (it's cross-origin and not intercepted here).
  e.respondWith(
    caches.match(req).then(cached => cached || fetch(req))
  );
});
