// Come With PWA service worker.
// Deliberately conservative: the dashboard is a live, frequently-deployed app
// against Supabase, so we NEVER want to serve a stale dashboard.html or hijack
// API calls. Strategy:
//   • Navigations (the HTML) → network-first, fall back to cache only offline.
//   • Same-origin static assets (icons, manifest) → cache-first.
//   • Everything cross-origin (Supabase, fonts, DICE, etc.) → straight to
//     network, untouched.
const CACHE = 'cw-shell-v2';
const SHELL = ['/dashboard.html', '/manifest.webmanifest', '/icons/icon-192.png', '/icons/apple-touch-icon.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;                 // never touch Supabase/APIs/fonts

  if (req.mode === 'navigate') {                              // HTML: fresh first, cache offline
    e.respondWith(
      fetch(req).then((r) => { const cp = r.clone(); caches.open(CACHE).then((c) => c.put('/dashboard.html', cp)); return r; })
        .catch(() => caches.match('/dashboard.html'))
    );
    return;
  }
  // Static same-origin assets: cache-first, then network (and cache it).
  e.respondWith(
    caches.match(req).then((hit) => hit || fetch(req).then((r) => {
      if (r.ok) { const cp = r.clone(); caches.open(CACHE).then((c) => c.put(req, cp)); }
      return r;
    }).catch(() => hit))
  );
});

// ---- Web Push (opt-in) ------------------------------------------------------
self.addEventListener('push', (e) => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (_) { d = { title: 'Come With', body: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.title || 'Come With', {
    body: d.body || '', icon: '/icons/icon-192.png', badge: '/icons/favicon-32.png',
    data: { url: d.url || '/dashboard.html' }, tag: d.tag || undefined,
  }));
});
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  const url = (e.notification.data && e.notification.data.url) || '/dashboard.html';
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((cs) => {
    for (const c of cs) { if (c.url.includes('/dashboard.html') && 'focus' in c) return c.focus(); }
    return self.clients.openWindow(url);
  }));
});
