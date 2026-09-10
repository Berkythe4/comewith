// pull-dice
//
// Widens the Come With Radio / Market data with DICE (dice.fm) EDM shows in NYC.
// DICE has no official public API, so this uses the SAME endpoints their own
// website calls (reverse-engineered, no auth):
//   • POST https://api.dice.fm/unified_search  {tag:"gig:<genre>", lat, lng, cursor?}
//       → event listings near a point (id, name, dates, venues, images).
//         PAGED: the response's `next_page_cursor` goes back as `cursor`.
//         Page 1 alone reaches ~16 days out — see the note on search().
//   • GET  https://api.dice.fm/events/<id>
//       → detail: perm_name (public URL) + summary_lineup.top_artists
//
// Writes ra_events + ra_artists with source='dice' ONLY. It deletes just its own
// source='dice' rows before upserting, so it can never touch the RA ('ra') or
// Ticketmaster ('tm') data — those pulls are untouched. The Come With Radio reads
// ra_events regardless of source, so DICE shows flow in automatically.
//
// Admin JWT OR service-role. No secret required (DICE endpoints are open).
// Body: { from?: "YYYY-MM-DD" (default today), to?: "YYYY-MM-DD",
//          days?: number (default 42, used when `to` is absent),
//          maxDetail?: number (default 600, max 900) }
//   maxDetail bounds the detail-fetch pass. The response reports `status`
//   (OK | PARTIAL), dropped_over_cap, timed_out, not_reached, search_pages and
//   last_date, so a truncated pull can never pass for a full one.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Safari/605.1.15";

// NYC point (matches RA area 8) + the electronic/EDM genre tags DICE exposes.
const NYC = { lat: 40.7128, lng: -74.006 };
const TAGS = [
  "music:dj", "music:party", "gig:techno", "gig:house", "gig:tech-house",
  "gig:drum-bass", "gig:drum-and-bass", "gig:disco", "gig:garage", "gig:trance",
  "gig:dubstep", "gig:electronic", "gig:minimal", "gig:afro-house", "gig:amapiano",
];

const MAX_PAGES = 12;

// unified_search is PAGED. Every response carries `next_page_cursor`, and the
// next page is fetched by sending it back as `cursor`. Reading only the first
// page — which is what this did until 2026-09-10 — is a query with an
// undeclared cap, the same shape as an unranged PostgREST select: it answers
// confidently and truncates in silence.
//
// It cost us real shows. Page 1 of `music:dj` reached 16 days out; Lane 8's
// Cross Pollination at Brooklyn Storehouse was 17 days out, on page 2, and
// never entered the candidate list. The event was correctly tagged `music:dj`
// the whole time — nothing was wrong with the event or the tag list. Paged,
// `music:dj` alone yields 379 events where page 1 gave 94, and the full 15-tag
// sweep goes 343 → 829 for six extra HTTP calls.
//
// Pages come back date-ascending, so once a whole page sits beyond the window
// there is nothing later worth asking for.
async function search(tag: string, cutoff: string, stats: { pages: number }): Promise<any[]> {
  const out: any[] = [];
  const seen = new Set<string>();
  let cursor: string | null = null;

  for (let page = 0; page < MAX_PAGES; page++) {
    let j: any;
    try {
      const body: Record<string, unknown> = { tag, lat: NYC.lat, lng: NYC.lng };
      if (cursor) body.cursor = cursor;
      const r = await fetch("https://api.dice.fm/unified_search", {
        method: "POST",
        headers: { "Content-Type": "application/json", "User-Agent": UA },
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(12000),
      });
      if (!r.ok) break;
      j = await r.json();
    } catch { break; }
    stats.pages++;

    const found: any[] = [];
    const walk = (o: any) => {
      if (Array.isArray(o)) { for (const v of o) walk(v); return; }
      if (o && typeof o === "object") {
        if (o.type === "event" && o.event?.id) found.push(o.event);
        for (const v of Object.values(o)) walk(v);
      }
    };
    walk(j);

    let fresh = 0, earliest = "9999-99-99";
    for (const e of found) {
      if (!e?.id || seen.has(e.id)) continue;
      seen.add(e.id);
      out.push(e);
      fresh++;
      const d = (e?.dates?.event_start_date || "").slice(0, 10);
      if (d && d < earliest) earliest = d;
    }

    cursor = typeof j?.next_page_cursor === "string" ? j.next_page_cursor : null;
    if (!cursor || !fresh) break;
    // This page is entirely past the window; every later page is later still.
    if (earliest !== "9999-99-99" && earliest > cutoff) break;
  }
  return out;
}
async function detail(id: string): Promise<any | null> {
  try {
    const r = await fetch(`https://api.dice.fm/events/${id}`, {
      headers: { "User-Agent": UA }, signal: AbortSignal.timeout(12000),
    });
    if (!r.ok) return null;
    return await r.json();
  } catch { return null; }
}
const isNYC = (v: any) => {
  const city = (v?.city?.name || "").toLowerCase();
  const addr = (v?.address || "").toLowerCase();
  return city.includes("new york") || addr.includes("new york") || addr.includes("brooklyn") || addr.includes("queens") || /\bny\b/.test(addr);
};

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

  try {
    const b = await req.json().catch(() => ({}));
    const days = Math.min(180, Math.max(7, Number(b.days) || 42));
    const today = new Date().toISOString().slice(0, 10);
    // The window can START in the future. Planning an episode two months out and
    // pulling [today, today+90] spent the whole detail budget on the near weeks
    // and never reached the window at all — soonest-first guarantees it. Anchor
    // on `from` so the cap is spent where the episode actually is.
    const from = typeof b.from === "string" && /^\d{4}-\d{2}-\d{2}$/.test(b.from) && b.from > today ? b.from : today;
    const cutoff = typeof b.to === "string" && /^\d{4}-\d{2}-\d{2}$/.test(b.to) && b.to > from
      ? b.to
      : new Date(new Date(from + "T00:00:00Z").getTime() + days * 86400000).toISOString().slice(0, 10);

    // 1. Collect candidate events across the genre tags (dedup by id).
    const cand = new Map<string, any>();
    const searchStats = { pages: 0 };
    for (const tag of TAGS) {
      const evs = await search(tag, cutoff, searchStats);
      for (const e of evs) {
        if (!cand.has(e.id)) cand.set(e.id, { ...e, _tag: tag.split(":")[1] });
      }
    }

    // 2. Detail-fetch each (bounded) for perm_name + lineup + a firm venue city.
    //    The cap keeps us under the function time budget — but WHICH events it
    //    spends itself on is the whole game. It used to take the first 160 ids in
    //    tag order, so on the 2026-08-14 pull DICE stopped dead at 8/21: weeks 2-4
    //    of a 4-week window had zero DICE shows and nothing reported it. Now we
    //    drop anything the search already dates outside the window (no detail
    //    fetch needed for those), take what's left SOONEST-FIRST, and return the
    //    overflow count so a truncated pull says so out loud.
    const dateOf = (e: any) => (e?.dates?.event_start_date || "").slice(0, 10);
    const inRange = [...cand.values()].filter((e) => { const d = dateOf(e); return !d || (d >= from && d <= cutoff); });
    inRange.sort((a, b) => (dateOf(a) || "9999-99-99").localeCompare(dateOf(b) || "9999-99-99"));
    // Default raised 240 → 600, IN THE SAME CHANGE as pagination, because fixing
    // only the paging would move the blind spot rather than remove it: paging
    // takes a 42-day window from 277 candidates to 555, and the Lane 8 show that
    // prompted all this lands at position 319 — still dropped under a 240 cap.
    // Ceiling 900 for wide windows; the deadline below is what keeps that safe.
    const maxDetail = Math.min(900, Math.max(40, Number(b.maxDetail) || 600));
    const picked = inRange.slice(0, maxDetail);
    const droppedOverCap = inRange.length - picked.length;
    const ids = picked.map((e) => e.id);
    const rows: Record<string, unknown>[] = [];
    const artistMap = new Map<string, Record<string, unknown>>();
    let scanned = 0, kept = 0;

    // Refuse to start work we cannot finish. The edge runtime's wall clock is
    // ~150s; a full 900-id pass is ~113 rounds of 8. Stopping at a deadline and
    // SAYING SO beats being killed mid-pass, which writes nothing and leaves no
    // trace of the attempt. Candidates are soonest-first, so what a short run
    // drops is always the far end of the window, never a random hole.
    const DEADLINE = Date.now() + 110_000;
    let timedOut = false;
    for (let i = 0; i < ids.length; i += 8) {
      if (Date.now() > DEADLINE) { timedOut = true; break; }
      const batch = ids.slice(i, i + 8);
      const details = await Promise.all(batch.map((id) => detail(id)));
      for (let k = 0; k < batch.length; k++) {
        scanned++;
        const id = batch[k];
        const base = cand.get(id);
        const d = details[k] || base;
        const start = (d.dates?.event_start_date || base.dates?.event_start_date || "");
        const date = start.slice(0, 10);
        if (!date || date < today || date > cutoff) continue;
        const venue = (d.venues || base.venues || [])[0] || {};
        if (!isNYC(venue)) continue;
        const url = d.perm_name ? `https://dice.fm/event/${d.perm_name}` : null;
        const flyer = d.images?.landscape || d.images?.square || base.images?.square || null;
        const genres = base._tag ? [base._tag.replace(/-/g, " ")] : [];
        const top = d.summary_lineup?.top_artists || [];
        const lineup = top.map((a: any) => ({ name: a.name })).filter((a: any) => a.name);
        kept++;
        rows.push({
          ra_id: `dice:${id}`, source: "dice", title: d.name || base.name,
          event_date: date, start_time: start || null, venue_name: venue.name || null,
          area_id: null, attending: null, interested_count: null,
          is_ticketed: true, is_pick: false, genres,
          flyer_url: flyer, content_url: url, lineup, fetched_at: new Date().toISOString(),
        });
        for (const a of top) {
          if (!a.artist_id || !a.name) continue;
          const akey = `dice:${a.artist_id}`;
          const prev = artistMap.get(akey);
          if (!prev || (date && (prev.next_event_date as string) > date)) {
            artistMap.set(akey, {
              ra_id: akey, source: "dice", name: a.name, soundcloud: null, instagram: null,
              follower_count: null, image: a.image?.url || null, content_url: url,
              next_event_date: date, next_event_title: d.name || base.name, next_venue: venue.name || null,
              next_event_url: url, genres, fetched_at: new Date().toISOString(),
            });
          }
        }
      }
    }

    // 3. Replace ONLY the dice-sourced rows (never touch ra / tm).
    // Bounded at BOTH ends — this pull now covers [from, cutoff], which may be a
    // narrow window well into the future. Deleting everything from `today` on and
    // re-inserting only the window would throw away every dice show outside it.
    await admin.from("ra_events").delete().eq("source", "dice")
      .gte("event_date", from).lte("event_date", cutoff);
    if (rows.length) {
      const { error } = await admin.from("ra_events").upsert(rows, { onConflict: "ra_id" });
      if (error) { console.error("dice ra_events:", error.message); return err(500, "Could not save DICE events: " + error.message); }
    }
    const artistRows = [...artistMap.values()];
    if (artistRows.length) {
      const { error: ae } = await admin.from("ra_artists").upsert(artistRows, { onConflict: "ra_id" });
      if (ae) console.error("dice ra_artists:", ae.message);
    }
    // last_date makes a short pull obvious at a glance: if DICE only reaches a
    // week out, that's the search's own horizon, not a filter you can widen.
    const lastDate = rows.reduce((m, r) => (r.event_date as string) > m ? (r.event_date as string) : m, "");
    return new Response(JSON.stringify({
      // PARTIAL is its own state, not a quieter kind of success: a run that hit
      // the cap or the clock has a known blind spot at the far end of the window
      // and the UI must show that as a problem, not a tick.
      success: true, source: "dice",
      status: (droppedOverCap > 0 || timedOut) ? "PARTIAL" : "OK",
      candidates: cand.size, in_window_candidates: inRange.length,
      search_pages: searchStats.pages,
      detailed: ids.length, dropped_over_cap: droppedOverCap,
      timed_out: timedOut, not_reached: timedOut ? ids.length - scanned : 0,
      scanned, saved: kept,
      last_date: lastDate || null, artists: artistRows.length,
      // Echo the window actually pulled, so a caller can tell "DICE has nothing
      // there" apart from "I asked for the wrong dates".
      from, to: cutoff,
    }), { headers: JH });
  } catch (e) {
    return err(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
