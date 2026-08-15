// pull-ticketmaster
//
// Widens the Market data with Ticketmaster's official Discovery API — EDM ONLY.
// classificationName=Dance/Electronic is a loose net (it lets Pop through), so we
// STRICTLY keep only events whose genre === "Dance/Electronic". Upserts into
// ra_events with source='tm' (no RSVP — TM has none; the demand metric stays RA).
// Admin JWT OR service-role. Secret: TM_API_KEY.
//
// Body: { from?: "YYYY-MM-DD" (default today), to?: "YYYY-MM-DD",
//          days?: number (default 42, used when `to` is absent),
//          city?: string, cities?: string[] }
//   cities default = all five boroughs. TM matches the city name literally, so
//   "New York" alone is Manhattan only — see the note at the cities list below.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const SUPA = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const admin = createClient(SUPA, SRK);
  const auth = req.headers.get("Authorization") || "";
  const bearer = auth.replace(/^Bearer\s+/i, "");
  const roleOf = (t: string) => { try { return JSON.parse(atob(t.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))).role || null; } catch { return null; } };
  let authed = bearer === SRK || roleOf(bearer) === "service_role";
  if (!authed && bearer) {
    const uc = createClient(SUPA, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
    const { data: { user } } = await uc.auth.getUser();
    if (user) { const { data: p } = await admin.from("profiles").select("role").eq("id", user.id).single(); authed = !!p && ["master_admin", "sub_admin"].includes(p.role); }
  }
  if (!authed) return err(401, "admin only");

  const key = Deno.env.get("TM_API_KEY");
  if (!key) return err(400, "Ticketmaster isn't configured (TM_API_KEY not set).");

  try {
    const b = await req.json().catch(() => ({}));
    const days = Math.min(180, Math.max(7, Number(b.days) || 42));
    // `city` is a LITERAL city-name match at Ticketmaster's end, and NYC is five
    // of them. Asking only for "New York" returned Manhattan and nothing else —
    // 27 future shows across 11 venues, not one in Brooklyn or Queens, so Brooklyn
    // Steel / Kings Theatre / Avant Gardner never existed as far as this pull was
    // concerned. Query each borough.
    const cities: string[] = Array.isArray(b.cities) && b.cities.length
      ? b.cities.map((c: unknown) => String(c))
      : b.city ? [String(b.city)]
      : ["New York", "Brooklyn", "Queens", "Bronx", "Staten Island"];
    // The window can start in the future — an episode planned for October wants
    // October's shows, and asking from today just burns pages on the near term.
    const now = new Date();
    const iso = (d: Date) => d.toISOString().slice(0, 19) + "Z";
    const dayRe = /^\d{4}-\d{2}-\d{2}$/;
    const todayIso = now.toISOString().slice(0, 10);
    const fromDate = typeof b.from === "string" && dayRe.test(b.from) && b.from > todayIso
      ? new Date(b.from + "T00:00:00Z") : now;
    const toDate = typeof b.to === "string" && dayRe.test(b.to) && b.to > fromDate.toISOString().slice(0, 10)
      ? new Date(b.to + "T23:59:59Z")
      : new Date(fromDate.getTime() + days * 86400000);
    const start = iso(fromDate), end = iso(toDate);

    const rows: Record<string, unknown>[] = [];
    const artistMap = new Map<string, Record<string, unknown>>();
    let grandTotal = 0;
    const perCity: Record<string, number> = {};
    let firstCall = true;
    for (const city of cities) {
    let total = 0, cityKept = 0;
    for (let page = 0; page < 5; page++) {
      const params = new URLSearchParams({
        apikey: key, classificationName: "Dance/Electronic", city, startDateTime: start, endDateTime: end,
        size: "100", page: String(page), sort: "date,asc",
      });
      const r = await fetch(`https://app.ticketmaster.com/discovery/v2/events.json?${params}`, { headers: { "Accept": "application/json" } });
      // Only a dead FIRST call is fatal — one borough 404ing must not throw away
      // the four that answered.
      if (!r.ok) { if (firstCall) return err(502, `Ticketmaster responded ${r.status}.`); break; }
      firstCall = false;
      const j = await r.json();
      total = j.page?.totalElements || 0;
      const events = j._embedded?.events || [];
      for (const e of events) {
        const cls = (e.classifications || [])[0] || {};
        const genre = cls.genre?.name || null;
        // STRICT EDM filter — drop the Pop/other false positives the loose query lets in.
        if (genre !== "Dance/Electronic") continue;
        const venue = (e._embedded?.venues || [])[0] || {};
        const date = e.dates?.start?.localDate || null;
        const pr = (e.priceRanges || [])[0];
        const cost = pr ? (pr.min === pr.max ? `$${Math.round(pr.min)}` : `$${Math.round(pr.min)}-${Math.round(pr.max)}`) : null;
        const genres = [genre, cls.subGenre?.name].filter(Boolean);
        const flyer = (e.images || []).sort((a: any, z: any) => (z.width || 0) - (a.width || 0))[0]?.url || null;
        const attractions = (e._embedded?.attractions || []) as any[];
        rows.push({
          ra_id: `tm:${e.id}`, source: "tm", title: e.name, event_date: date,
          start_time: e.dates?.start?.dateTime || null, venue_name: venue.name || null, area_id: null,
          attending: null, interested_count: null, is_ticketed: true, is_pick: false,
          genres, flyer_url: flyer, content_url: e.url || null,
          lineup: attractions.map((a) => ({ name: a.name, soundcloud: null })),
          next_cost: cost, fetched_at: new Date().toISOString(),
        });
        cityKept++;
        // TM performers → ra_artists (so they appear in the artist views). No socials/RSVP.
        for (const a of attractions) {
          if (!a?.id || !a?.name) continue;
          const key = `tm:${a.id}`;
          const prev = artistMap.get(key);
          if (!prev || (date && (prev.next_event_date as string) > date)) {
            artistMap.set(key, {
              ra_id: key, source: "tm", name: a.name, soundcloud: null, instagram: null,
              follower_count: null, image: (a.images || [])[0]?.url || null, content_url: a.url || null,
              next_event_date: date, next_event_title: e.name, next_venue: venue.name || null,
              next_cost: cost, next_event_url: e.url || null, genres, fetched_at: new Date().toISOString(),
            });
          }
        }
      }
      if (events.length < 100 || (page + 1) * 100 >= total) break;
    }
    grandTotal += total; perCity[city] = cityKept;
    }

    // Dedupe by ra_id (TM can repeat), then replace the TM slice of the window.
    const map = new Map<string, Record<string, unknown>>();
    rows.forEach((r) => map.set(r.ra_id as string, r));
    const finalRows = [...map.values()].map(({ next_cost, ...r }) => r); // next_cost isn't an ra_events col
    // Bounded at BOTH ends. This used to delete every tm row from `start` forward
    // and re-insert only what this pull returned — fine while the pull was always
    // [today, today+90], fatal once the window can be narrower or start later:
    // pulling a 4-week window would have deleted every tm show beyond it.
    await admin.from("ra_events").delete().eq("source", "tm")
      .gte("event_date", start.slice(0, 10)).lte("event_date", end.slice(0, 10));
    let saved = 0;
    if (finalRows.length) {
      const { error } = await admin.from("ra_events").upsert(finalRows, { onConflict: "ra_id" });
      if (error) { console.error("tm upsert:", error.message); return err(500, "Could not save Ticketmaster events."); }
      saved = finalRows.length;
    }
    // TM artists
    await admin.from("ra_artists").delete().eq("source", "tm")
      .gte("next_event_date", start.slice(0, 10)).lte("next_event_date", end.slice(0, 10));
    const artistRows = [...artistMap.values()];
    if (artistRows.length) {
      const { error: ae } = await admin.from("ra_artists").upsert(artistRows, { onConflict: "ra_id" });
      if (ae) console.error("tm artist upsert:", ae.message);
    }
    return new Response(JSON.stringify({
      success: true, source: "ticketmaster", cities, per_city: perCity,
      total_from_tm: grandTotal, edm_saved: saved,
      from: start.slice(0, 10), to: end.slice(0, 10),
    }), { headers: JH });
  } catch (e) {
    console.error("pull-ticketmaster:", e instanceof Error ? e.message : String(e));
    return err(502, "Ticketmaster pull failed — try again shortly.");
  }
});
