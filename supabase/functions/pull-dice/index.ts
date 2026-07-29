// pull-dice
//
// Widens the Come With Radio / Market data with DICE (dice.fm) EDM shows in NYC.
// DICE has no official public API, so this uses the SAME endpoints their own
// website calls (reverse-engineered, no auth):
//   • POST https://api.dice.fm/unified_search  {tag:"gig:<genre>", lat, lng}
//       → event listings near a point (id, name, dates, venues, images)
//   • GET  https://api.dice.fm/events/<id>
//       → detail: perm_name (public URL) + summary_lineup.top_artists
//
// Writes ra_events + ra_artists with source='dice' ONLY. It deletes just its own
// source='dice' rows before upserting, so it can never touch the RA ('ra') or
// Ticketmaster ('tm') data — those pulls are untouched. The Come With Radio reads
// ra_events regardless of source, so DICE shows flow in automatically.
//
// Admin JWT OR service-role. No secret required (DICE endpoints are open).
// Body: { days?: number (default 42) }

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

async function search(tag: string): Promise<any[]> {
  try {
    const r = await fetch("https://api.dice.fm/unified_search", {
      method: "POST",
      headers: { "Content-Type": "application/json", "User-Agent": UA },
      body: JSON.stringify({ tag, lat: NYC.lat, lng: NYC.lng }),
    });
    if (!r.ok) return [];
    const j = await r.json();
    const out: any[] = [];
    const walk = (o: any) => {
      if (Array.isArray(o)) { for (const v of o) walk(v); return; }
      if (o && typeof o === "object") {
        if (o.type === "event" && o.event?.id) out.push(o.event);
        for (const v of Object.values(o)) walk(v);
      }
    };
    walk(j);
    return out;
  } catch { return []; }
}
async function detail(id: string): Promise<any | null> {
  try {
    const r = await fetch(`https://api.dice.fm/events/${id}`, { headers: { "User-Agent": UA } });
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
    const days = Math.min(90, Math.max(7, Number(b.days) || 42));
    const today = new Date().toISOString().slice(0, 10);
    const cutoff = new Date(Date.now() + days * 86400000).toISOString().slice(0, 10);

    // 1. Collect candidate events across the genre tags (dedup by id).
    const cand = new Map<string, any>();
    for (const tag of TAGS) {
      const evs = await search(tag);
      for (const e of evs) {
        if (!cand.has(e.id)) cand.set(e.id, { ...e, _tag: tag.split(":")[1] });
      }
    }

    // 2. Detail-fetch each (bounded) for perm_name + lineup + a firm venue city.
    //    Cap to keep well under the function time budget.
    const ids = [...cand.keys()].slice(0, 160);
    const rows: Record<string, unknown>[] = [];
    const artistMap = new Map<string, Record<string, unknown>>();
    let scanned = 0, kept = 0;

    for (let i = 0; i < ids.length; i += 6) {
      const batch = ids.slice(i, i + 6);
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
    await admin.from("ra_events").delete().eq("source", "dice").gte("event_date", today);
    if (rows.length) {
      const { error } = await admin.from("ra_events").upsert(rows, { onConflict: "ra_id" });
      if (error) { console.error("dice ra_events:", error.message); return err(500, "Could not save DICE events: " + error.message); }
    }
    const artistRows = [...artistMap.values()];
    if (artistRows.length) {
      const { error: ae } = await admin.from("ra_artists").upsert(artistRows, { onConflict: "ra_id" });
      if (ae) console.error("dice ra_artists:", ae.message);
    }
    return new Response(JSON.stringify({ success: true, source: "dice", candidates: cand.size, scanned, saved: kept, artists: artistRows.length }), { headers: JH });
  } catch (e) {
    return err(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
