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
// phone notifications sent by the GitHub reminder job
self.addEventListener('push', e => {
  let d = {}; try { d = e.data.json(); } catch { d = { title: 'T1D Care Hub', body: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.title || 'T1D Care Hub', { body: d.body || '', tag: d.tag || undefined, icon: '/icon-192.png', badge: '/icon-192.png', data: { url: d.url || '/' } }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close(); const url = (e.notification.data && e.notification.data.url) || '/';
  e.waitUntil(clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => { for (const c of cs) { if ('focus' in c) { c.navigate && c.navigate(url).catch(() => {}); return c.focus(); } } return clients.openWindow(url); }));
});
