// Offline support: app files come from the network when online, from the cache when not.
const CACHE = 't1d-__VERSION__';
const SHELL = ['/', '/index.html', '/main.js', '/backend.js', '/app.js', '/styles.css', '/config.js', '/manifest.webmanifest', '/icon-192.png', '/icon.svg'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== 't1d-cdn').map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const u = new URL(e.request.url);
  if (u.hostname.endsWith('supabase.co')) return;                       // his data: always live
  if (u.origin === location.origin) {
    e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; })
      .catch(() => caches.match(e.request).then(m => m || caches.match('/index.html'))));
  } else if (/cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|fonts\.(googleapis|gstatic)\.com/.test(u.hostname)) {
    e.respondWith(caches.open('t1d-cdn').then(c => c.match(e.request).then(m => {
      const f = fetch(e.request).then(r => { c.put(e.request, r.clone()); return r; }).catch(() => m);
      return m || f; })));
  }
});
