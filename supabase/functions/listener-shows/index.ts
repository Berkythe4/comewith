// listener-shows  (LISTENER — verify_jwt on; any signed-in user, incl. customers)
//
// "The artists you saved are playing NYC soon." This is the piece that turns the
// radio from a listening page into something that sells tickets: a listener hearts
// a track, and later finds out that artist has a show — with the date, venue,
// price and ticket link.
//
// Runs server-side on purpose. ra_artists is admin-RLS'd and stays that way: a
// listener must never be able to read the whole market-intel table, only the
// slice that matches artists THEY saved. We look up their saves with the service
// role, scoped hard to their own user_id from the verified JWT.
//
// Body: {}  →  { shows: [{ artist, date, venue, cost, url, saved_titles[] }] }

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

const SUPA = Deno.env.get("SUPABASE_URL")!;
const DIACRITICS = new RegExp("[\\u0300-\\u036f]", "g");
// Exact-after-normalization only. A fuzzy match here would tell someone their
// favourite artist is playing when it's a different act with a similar name —
// worse than saying nothing, because they might buy a ticket on it.
const norm = (s: string) =>
  String(s || "").toLowerCase().normalize("NFD").replace(DIACRITICS, "")
    .replace(/&/g, " and ").replace(/[^a-z0-9]+/g, " ").trim();

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const admin = createClient(SUPA, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  const auth = req.headers.get("Authorization") || "";
  const uc = createClient(SUPA, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
  const { data: { user } } = await uc.auth.getUser();
  if (!user) return err(401, "Sign in first.");

  try {
    // Their saved tracks — scoped to playlists they own. Never trust a body param
    // for identity; the user id comes from the verified token.
    const { data: pls } = await admin.from("listener_playlists").select("id").eq("user_id", user.id);
    const ids = (pls || []).map((p) => p.id);
    if (!ids.length) return new Response(JSON.stringify({ shows: [] }), { headers: JH });

    const { data: saves } = await admin.from("listener_playlist_tracks")
      .select("artist_name, title").in("playlist_id", ids);
    if (!saves?.length) return new Response(JSON.stringify({ shows: [] }), { headers: JH });

    const byArtist = new Map<string, { name: string; titles: string[] }>();
    for (const s of saves) {
      const k = norm(s.artist_name || "");
      if (!k) continue;
      const e = byArtist.get(k) || { name: s.artist_name || "", titles: [] };
      if (s.title && e.titles.length < 3) e.titles.push(s.title);
      byArtist.set(k, e);
    }
    if (!byArtist.size) return new Response(JSON.stringify({ shows: [] }), { headers: JH });

    const today = new Date().toISOString().slice(0, 10);
    const { data: artists } = await admin.from("ra_artists")
      .select("name, next_event_date, next_venue, next_cost, next_event_url")
      .gte("next_event_date", today).order("next_event_date").limit(600);

    const shows = [];
    const seen = new Set<string>();
    for (const a of artists || []) {
      const k = norm(a.name || "");
      const hit = byArtist.get(k);
      if (!hit || seen.has(k)) continue;
      seen.add(k);
      shows.push({
        artist: a.name,
        date: a.next_event_date,
        venue: a.next_venue || null,
        cost: a.next_cost || null,
        url: a.next_event_url || null,
        saved_titles: hit.titles,
      });
    }
    return new Response(JSON.stringify({ shows }), { headers: JH });
  } catch (e) {
    console.error("listener-shows:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load your artists' shows.");
  }
});
