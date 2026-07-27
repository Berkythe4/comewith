// Come With PWA service worker — v4.
// Chrome only treats the app as INSTALLABLE if the SW actually handles the
// start_url navigation with respondWith(). v3 returned early (no respondWith),
// which silently disabled the install prompt. So navigations are now
// NETWORK-FIRST via respondWith: fresh from the network when online (never
// stale), cached copy only as an offline fallback. Icons/manifest cache-first.
const CACHE = 'cw-static-v4';
const STATIC = ['/manifest.webmanifest', '/icons/icon-192.png', '/icons/apple-touch-icon.png', '/icons/favicon-32.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(STATIC)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Navigations (the app HTML) → network-first, cache as offline fallback. The
  // respondWith is what makes Chrome consider the app installable.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req).then((r) => { const cp = r.clone(); caches.open(CACHE).then((c) => c.put('/dashboard.html', cp)); return r; })
        .catch(() => caches.match('/dashboard.html').then((h) => h || caches.match(req)))
    );
    return;
  }
  if (url.origin !== location.origin) return;   // Supabase/APIs/fonts untouched
  if (!STATIC.includes(url.pathname)) return;
  e.respondWith(caches.match(req).then((hit) => hit || fetch(req)));
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
