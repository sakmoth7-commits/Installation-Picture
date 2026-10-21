// Installation Tracker service worker: lets the app install and open fast. API calls are never cached.
const CACHE = 'it-tracker-v2';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // POSTs to the Apps Script API go straight to the network
  const url = new URL(req.url);
  const same = url.origin === self.location.origin;
  const cdn = url.hostname === 'cdn.jsdelivr.net' || url.hostname === 'unpkg.com';
  if (!same && !cdn) return;

  // Pages: network first, so a new index.html is picked up as soon as you are online.
  if (req.mode === 'navigate' || (same && url.pathname.endsWith('.html'))) {
    e.respondWith(
      fetch(req).then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return r; })
        .catch(() => caches.match(req).then(r => r || caches.match('index.html')))
    );
    return;
  }
  // Everything else (libraries, icons): cache first.
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
    if (r.ok || r.type === 'opaque') { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return r;
  })));
});
