// Come With prototypes — shared data layer (live, read-only public endpoints).
// Safe to ship client-side: the publishable key is a public anon-tier key; all
// access is governed by RLS (public can read v_public_events + insert inquiries).
window.CW = (function () {
  const URL = 'https://yaytdosxfhcqatmhctzk.supabase.co';
  const KEY = 'sb_publishable_IkigzWOTU3ZSMK9DxqwwJw_AaQkShCi';
  const H = { apikey: KEY, Authorization: 'Bearer ' + KEY };

  // Upcoming public events (returns [] on any failure so the UI can fall back).
  async function fetchEvents() {
    try {
      const r = await fetch(URL + '/rest/v1/v_public_events?select=*&order=event_date.asc', { headers: H });
      if (!r.ok) return [];
      return await r.json();
    } catch (_) { return []; }
  }

  // Submit a booking/inquiry to the CRM (anon insert allowed). Returns {ok, error}.
  async function submitInquiry(data) {
    try {
      const body = {
        full_name: data.full_name, email: data.email, phone: data.phone || null,
        event_type: data.event_type || null, event_date: data.event_date || null,
        venue: data.venue || null,
        services_selected: Array.isArray(data.services) ? data.services : [],
        message: data.message || null, source: 'website',
      };
      const r = await fetch(URL + '/rest/v1/inquiries', {
        method: 'POST',
        headers: { ...H, 'Content-Type': 'application/json', Prefer: 'return=minimal' },
        body: JSON.stringify(body),
      });
      return r.ok ? { ok: true } : { ok: false, error: 'HTTP ' + r.status };
    } catch (e) { return { ok: false, error: String(e) }; }
  }

  const fmtDate = (s) => {
    if (!s) return '';
    const d = new Date(s + 'T00:00:00');
    return d.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
  };
  const fmtDay = (s) => { const d = new Date((s || '') + 'T00:00:00'); return { mon: d.toLocaleDateString('en-US', { month: 'short' }).toUpperCase(), day: d.getDate() }; };

  // Static past-events highlights (real Come With events; v_public_events only
  // exposes upcoming public ones, so past highlights are curated here).
  const PAST = [
    { name: 'Dance Infusion #2', date: '2026-05-09', venue: 'Signal', tag: 'Benefit · National MS Society', note: '117 guests · ~$3,000 to mission' },
    { name: 'Knicks G5 Watch Party', date: '2026-06-13', venue: 'Crossroads Café', tag: 'Party', note: '~60 in the room' },
    { name: 'DI Artist Showcase — Kristen London & 32LVS', date: '2026-04-18', venue: 'Studio', tag: 'Showcase', note: 'recorded set' },
    { name: 'Dance Infusion #1', date: '2025-09-08', venue: 'Signal', tag: 'Benefit · National MS Society', note: '42 tickets · $1,140 to mission' },
  ];
  const IMPACT = { raised: '$4,140+', beneficiary: 'National MS Society', benefits: 2, pctToMission: '39%+' };
  const DJS = ['Berky', 'KRNeY', 'SPF 50', 'Kristen London', '32LVS', 'Just Martin', 'Henry', 'Kloud9'];

  return { fetchEvents, submitInquiry, fmtDate, fmtDay, PAST, IMPACT, DJS, IG: 'https://instagram.com/comewithnyc', EMAIL: 'berky@comewith.org' };
})();
