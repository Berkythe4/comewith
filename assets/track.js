/* Come With — first-party analytics beacon (see migration 108 + functions/track)
 *
 * Counts pageviews and link clicks on the public pages. No cookie, no
 * localStorage, no cross-site id: the session id is random per TAB and dies
 * with the tab, which is enough to count visitors and useless for following a
 * person around. Nothing here can read or write the database directly — it
 * POSTs to the `track` edge function, which is the only thing holding a key.
 *
 * Drop-in: <script src="assets/track.js" defer></script>
 * Anything with data-track="label" is counted as a click even when it's not a
 * link (buttons, tabs); outbound links are counted automatically.
 */
(function () {
  var FN = 'https://yaytdosxfhcqatmhctzk.supabase.co/functions/v1/track';
  // Local dev and Netlify deploy previews would otherwise pollute the numbers.
  var host = location.hostname;
  if (host === 'localhost' || host === '127.0.0.1' || host === '' || /--.*\.netlify\.app$/.test(host)) return;

  // Page identity = pathname + only the params that name a THING. Keeping
  // ?s=<episode> means episode pages count separately; dropping everything else
  // means a hundred utm variants don't shatter one page into a hundred rows.
  // 'p' names a links page. links.html normalises itself to /links/<slug>, so
  // this only matters for the ?p= fallback form (a bio link written before the
  // rewrite existed) — without it those visits would all collapse into one row.
  var KEEP = ['s', 'slug', 'id', 'e', 'a', 'p'];
  function pagePath() {
    try {
      var q = new URLSearchParams(location.search), keep = [];
      KEEP.forEach(function (k) { var v = q.get(k); if (v) keep.push(k + '=' + v.slice(0, 80)); });
      return (location.pathname || '/') + (keep.length ? '?' + keep.join('&') : '');
    } catch (e) { return location.pathname || '/'; }
  }

  function sid() {
    try {
      var s = sessionStorage.getItem('cw_sid');
      if (!s) { s = Math.random().toString(36).slice(2) + Date.now().toString(36); sessionStorage.setItem('cw_sid', s); }
      return s;
    } catch (e) { return null; }  // private mode — the view still counts, just unattributed
  }

  function send(o) {
    try {
      o.sid = sid();
      var b = JSON.stringify(o);
      // text/plain keeps sendBeacon preflight-free; sendBeacon itself survives
      // the page being closed, which a fetch on click would not.
      if (navigator.sendBeacon) navigator.sendBeacon(FN, new Blob([b], { type: 'text/plain;charset=UTF-8' }));
      else fetch(FN, { method: 'POST', body: b, headers: { 'Content-Type': 'text/plain' }, keepalive: true });
    } catch (e) { /* never let analytics break a page */ }
  }

  var q = new URLSearchParams(location.search);
  var refHost = '';
  try { refHost = document.referrer ? new URL(document.referrer).hostname : ''; } catch (e) {}
  if (refHost === location.hostname) refHost = '';   // internal navigation isn't a referrer

  send({
    k: 'view', p: pagePath(), r: refHost || null,
    us: q.get('utm_source'), um: q.get('utm_medium'), uc: q.get('utm_campaign'),
  });

  document.addEventListener('click', function (e) {
    var el = e.target && e.target.closest && e.target.closest('a[href], [data-track]');
    if (!el) return;
    var href = el.getAttribute('href') || '';
    var label = (el.getAttribute('data-track') || el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 120);
    var outbound = false, url = href;
    if (href) {
      try { var u = new URL(href, location.href); outbound = u.hostname !== location.hostname; url = u.href; }
      catch (err) { return; }
      // mailto:/tel:/# aren't traffic
      if (/^(mailto|tel|javascript):/i.test(href) || href.charAt(0) === '#') { if (!el.hasAttribute('data-track')) return; }
    } else if (!el.hasAttribute('data-track')) return;
    send({ k: 'click', p: pagePath(), l: url || null, t: label || null, o: outbound });
  }, true);
})();
