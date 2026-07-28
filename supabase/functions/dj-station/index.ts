// dj-station  (TOKEN-gated, no login)
//
// A DJ assigned to an episode gets a link (dj.html?ep=<dj_token>). This returns a
// SCOPED research view for that episode: only NYC artists playing within the
// episode's window (+ optional genre filter from dj_search_params), each with
// their songs, plus the episode's current tracklist. Read-only. The token is the
// only credential; revoking it (clearing dj_token) kills access immediately.
//
// Body: { token: string }
import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");
  const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  try {
    const { token } = await req.json().catch(() => ({}));
    if (!token || typeof token !== "string") return err(400, "missing token");

    const { data: ep } = await admin.from("sc_playlists")
      .select("id, station_no, name, drop_date, dj_search_params")
      .eq("dj_token", token).maybeSingle();
    if (!ep) return err(404, "This link is invalid or has been revoked.");

    const params = ep.dj_search_params || {};
    const weeks = Math.max(1, Math.min(12, Number(params.weeks) || 4));
    const genres: string[] = Array.isArray(params.genres) ? params.genres.filter(Boolean) : [];
    const today = new Date().toISOString().slice(0, 10);
    const to = new Date(Date.now() + weeks * 7 * 86400000).toISOString().slice(0, 10);

    // Scoped artist pool: playing NYC within the window, optional genre overlap.
    let q = admin.from("ra_artists")
      .select("name, soundcloud, follower_count, genres, city, next_event_date, next_venue, next_event_url")
      .gte("next_event_date", today).lte("next_event_date", to)
      .order("next_event_date", { ascending: true }).limit(160);
    if (genres.length) q = q.overlaps("genres", genres);
    const { data: artists } = await q;

    // Their songs, from the scan cache (keyed by normalized soundcloud URL).
    const norm = (u: string) => (u || "").trim().toLowerCase().replace("://www.", "://").replace(/\/+$/, "").split("?")[0];
    const scUrls = [...new Set((artists || []).map((a) => a.soundcloud).filter(Boolean).map(norm))];
    const songByUrl: Record<string, any[]> = {};
    for (let i = 0; i < scUrls.length; i += 200) {
      const { data: cache } = await admin.from("sc_artist_cache")
        .select("soundcloud, songs, is_producer, followers").in("soundcloud", scUrls.slice(i, i + 200));
      (cache || []).forEach((c) => { songByUrl[norm(c.soundcloud)] = (c.songs || []).slice(0, 12).map((s: any) => ({ title: s.title, url: s.permalink_url })); });
    }
    const out = (artists || []).map((a) => ({
      name: a.name, soundcloud: a.soundcloud, followers: a.follower_count || 0,
      genres: a.genres || [], city: a.city || null,
      next_event_date: a.next_event_date, next_venue: a.next_venue, next_event_url: a.next_event_url,
      songs: a.soundcloud ? (songByUrl[norm(a.soundcloud)] || []) : [],
    }));

    // The episode's current tracklist (what's already in).
    const { data: tracks } = await admin.from("sc_playlist_tracks")
      .select("artist_name, title, permalink_url, show_date, show_venue, sort")
      .eq("playlist_id", ep.id).order("sort");

    return new Response(JSON.stringify({
      ok: true,
      episode: { no: ep.station_no, name: ep.name, drop_date: ep.drop_date },
      scope: { weeks, genres, from: today, to },
      artists: out,
      tracklist: tracks || [],
    }), { headers: JH });
  } catch (e) {
    console.error("dj-station:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load the episode.");
  }
});
