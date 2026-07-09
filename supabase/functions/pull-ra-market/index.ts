// pull-ra-market
//
// Fetches PUBLIC Resident Advisor event data (ra.co/graphql — no auth, no
// private/ticketing data) for a metro area + date window and caches it into
// ra_events / ra_artists. Server-side so there's no browser CORS and the
// scraping stays off the client. Admin JWT OR service-role bearer (so pg_cron
// can call it later). Public promoter data only — money/guestlist is never here.
//
// Body (all optional): { area?: number, days?: number, maxPages?: number }
//   area default = site_content 'ops.ra_area_id' (8 = New York)
//   days default = 28, maxPages default = 12 (×50 = up to 600 events)

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

const RA_URL = "https://ra.co/graphql";
const RA_HEADERS = {
  "Content-Type": "application/json",
  "Referer": "https://ra.co/events",
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
};

const LISTINGS_QUERY = `query($f: FilterInputDtoInput, $ps: Int, $pg: Int) {
  eventListings(filters: $f, pageSize: $ps, page: $pg) {
    totalResults
    data { event {
      id title date startTime attending interestedCount isTicketed
      pick { id } genres { name } flyerFront contentUrl
      venue { name }
      artists { id name soundcloud instagram followerCount image contentUrl }
    } }
  }
}`;

async function raFetch(area: number, dayFrom: string, dayTo: string, page: number) {
  const body = JSON.stringify({
    query: LISTINGS_QUERY,
    variables: { f: { areas: { eq: area }, listingDate: { gte: dayFrom, lte: dayTo } }, ps: 50, pg: page },
  });
  const res = await fetch(RA_URL, { method: "POST", headers: RA_HEADERS, body });
  if (!res.ok) throw new Error(`RA responded ${res.status}`);
  const json = await res.json();
  if (json.errors) throw new Error("RA query error: " + JSON.stringify(json.errors).slice(0, 200));
  return json.data?.eventListings;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const URLS = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const admin = createClient(URLS, SRK);

  // Auth: service-role bearer (any key format) OR admin JWT. We check the JWT's
  // `role` claim rather than string-matching the env key, so it's robust to
  // Supabase's legacy vs new key formats.
  const auth = req.headers.get("Authorization") || "";
  const bearer = auth.replace(/^Bearer\s+/i, "");
  const jwtRole = (tok: string): string | null => {
    try { return JSON.parse(atob(tok.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))).role || null; } catch { return null; }
  };
  let authed = bearer === SRK || jwtRole(bearer) === "service_role";
  if (!authed && bearer) {
    const uc = createClient(URLS, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
    const { data: { user } } = await uc.auth.getUser();
    if (user) {
      const { data: prof } = await admin.from("profiles").select("role").eq("id", user.id).single();
      authed = !!prof && ["master_admin", "sub_admin"].includes(prof.role);
    }
  }
  if (!authed) return err(401, "admin only");

  try {
    const b = await req.json().catch(() => ({}));
    let area = Number(b.area) || 0;
    if (!area) {
      const { data: s } = await admin.from("site_content").select("value").eq("key", "ops.ra_area_id").maybeSingle();
      area = Number(s?.value) || 8;
    }
    const days = Math.min(90, Math.max(7, Number(b.days) || 28));
    const maxPages = Math.min(20, Math.max(1, Number(b.maxPages) || 12));

    const today = new Date();
    const iso = (d: Date) => d.toISOString().slice(0, 10);
    const dayFrom = iso(today);
    const dayTo = iso(new Date(today.getTime() + days * 86400000));

    const eventMap = new Map<string, Record<string, unknown>>(); // keyed by ra_id — RA repeats events across pages
    const artistMap = new Map<string, Record<string, unknown>>();
    let total = 0, pages = 0;

    for (let page = 1; page <= maxPages; page++) {
      const listings = await raFetch(area, dayFrom, dayTo, page);
      if (!listings) break;
      total = listings.totalResults || 0;
      const items = listings.data || [];
      pages = page;
      for (const it of items) {
        const e = it.event;
        if (!e?.id) continue;
        const lineup = (e.artists || []).map((a: any) => ({
          ra_id: a.id, name: a.name, soundcloud: a.soundcloud || null,
          follower_count: a.followerCount ?? null, content_url: a.contentUrl || null,
        }));
        const evDate = (e.date || "").slice(0, 10) || null;
        eventMap.set(e.id, {
          ra_id: e.id, title: e.title, event_date: evDate,
          start_time: e.startTime || null, venue_name: e.venue?.name || null, area_id: area,
          attending: e.attending ?? 0, interested_count: e.interestedCount ?? 0,
          is_ticketed: !!e.isTicketed, is_pick: !!e.pick,
          genres: (e.genres || []).map((g: any) => g.name),
          flyer_url: e.flyerFront || null, content_url: e.contentUrl ? `https://ra.co${e.contentUrl}` : null,
          lineup, fetched_at: new Date().toISOString(),
        }); // eventMap.set(...)
        // Dedupe artists — keep each artist's SOONEST upcoming show.
        for (const a of (e.artists || [])) {
          if (!a?.id) continue;
          const prev = artistMap.get(a.id);
          if (!prev || (evDate && (prev.next_event_date as string) > evDate)) {
            artistMap.set(a.id, {
              ra_id: a.id, name: a.name, soundcloud: a.soundcloud || null, instagram: a.instagram || null,
              follower_count: a.followerCount ?? null, image: a.image || null,
              content_url: a.contentUrl ? `https://ra.co${a.contentUrl}` : null,
              next_event_date: evDate, next_event_title: e.title, next_venue: e.venue?.name || null,
              genres: (e.genres || []).map((g: any) => g.name),
              fetched_at: new Date().toISOString(),
            });
          }
        }
      }
      if (items.length < 50 || page * 50 >= total) break;
    }

    // Replace the window: clear old cache, then insert fresh (keeps the table
    // to "what's upcoming now" and avoids stale past events lingering).
    await admin.from("ra_events").delete().gte("event_date", dayFrom);
    await admin.from("ra_artists").delete().gte("next_event_date", dayFrom);
    let evN = 0, arN = 0;
    const eventRows = [...eventMap.values()];
    if (eventRows.length) {
      const { error: e1 } = await admin.from("ra_events").upsert(eventRows, { onConflict: "ra_id" });
      if (e1) { console.error("ra_events upsert:", e1.message); return err(500, "Could not save events."); }
      evN = eventRows.length;
    }
    const artistRows = [...artistMap.values()];
    if (artistRows.length) {
      const { error: e2 } = await admin.from("ra_artists").upsert(artistRows, { onConflict: "ra_id" });
      if (e2) { console.error("ra_artists upsert:", e2.message); return err(500, "Could not save artists."); }
      arN = artistRows.length;
    }

    return new Response(JSON.stringify({
      success: true, area, days, pages, total_available: total,
      events_saved: evN, artists_saved: arN,
      artists_with_soundcloud: artistRows.filter((a) => a.soundcloud).length,
    }), { headers: JH });
  } catch (e) {
    console.error("pull-ra-market:", e instanceof Error ? e.message : String(e));
    return err(502, "Couldn't reach Resident Advisor — it may be rate-limiting. Try again shortly.");
  }
});
