// Come With PWA service worker — v3.
// LESSON: caching the app HTML risks serving stale, broken app code. So this SW
// now NEVER caches dashboard.html — navigations are network-only. It caches only
// the icons + manifest (which never change) so install still works offline. This
// guarantees the app code is always fresh from the network.
const CACHE = 'cw-static-v3';
const STATIC = ['/manifest.webmanifest', '/icons/icon-192.png', '/icons/apple-touch-icon.png', '/icons/favicon-32.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(STATIC)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  // Delete EVERY old cache (including v1/v2 that cached HTML) and take over now.
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  // Navigations / any HTML → always network. Never serve app code from cache.
  if (req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html')) return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;   // Supabase/APIs/fonts untouched
  // Only the small static icon/manifest set is cached.
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
