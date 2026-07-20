// get-station  (PUBLIC — deploy with --no-verify-jwt)
//
// Serves Come With Radio to the public radio.html page:
//   ?list=1     → all PUBLISHED episodes (hub view): meta + track count/length.
//   ?s=<slug>   → one PUBLISHED episode by its pretty slug (episode page).
//   ?t=<token>  → any episode by its secret public_token (unlisted preview —
//                 how a station is shared before Keith flips it live).
// Returns station meta (mix track / YouTube / descriptions) + the tracklist with
// each track's show (date/venue/price) + bpm/key. Read-only; service role under
// the hood, so nothing here relies on anon grants.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "GET, POST, OPTIONS" };
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

const STATION_COLS = "id, slug, name, note, desc_public, published, published_at, status, station_no, drop_date, sc_playlist_url, mix_sc_track_url, mix_youtube_url, cover_url";
const TRACK_COLS = "title, artist_name, permalink_url, duration_ms, playback_count, artwork_url, show_date, show_venue, show_cost, show_url, bpm, song_key, camelot, genres, sort";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  try {
    const url = new URL(req.url);
    let token = url.searchParams.get("t") || url.searchParams.get("token") || "";
    let slug = url.searchParams.get("s") || url.searchParams.get("slug") || "";
    let list = url.searchParams.get("list") || "";
    if (!token && !slug && !list && req.method === "POST") {
      const b = await req.json().catch(() => ({}));
      token = (b.token || "").toString(); slug = (b.slug || "").toString(); list = (b.list || "").toString();
    }
    token = token.trim(); slug = slug.trim().toLowerCase();

    const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

    // Hub: every live episode, newest first.
    if (list) {
      const { data: pls } = await admin.from("sc_playlists").select(STATION_COLS)
        .eq("published", true).order("published_at", { ascending: false, nullsFirst: false });
      const ids = (pls || []).map((p) => p.id);
      const agg: Record<string, { n: number; ms: number; art: string | null }> = {};
      if (ids.length) {
        const { data: trs } = await admin.from("sc_playlist_tracks")
          .select("playlist_id, duration_ms, artwork_url").in("playlist_id", ids).order("sort");
        for (const t of trs || []) {
          const a = (agg[t.playlist_id] ||= { n: 0, ms: 0, art: null });
          a.n++; a.ms += t.duration_ms || 0; if (!a.art && t.artwork_url) a.art = t.artwork_url;
        }
      }
      // Tease the next scheduled drop (drops are dated in advance): the nearest
      // future-dated station that isn't live yet. Number + date + name only.
      const today = new Date().toISOString().slice(0, 10);
      const { data: nd } = await admin.from("sc_playlists").select("station_no, name, drop_date")
        .eq("published", false).gte("drop_date", today).order("drop_date").limit(1).maybeSingle();
      return new Response(JSON.stringify({
        stations: (pls || []).map(({ id, ...p }) => ({
          ...p,
          track_count: agg[id]?.n || 0,
          total_min: Math.round((agg[id]?.ms || 0) / 60000),
          artwork_url: p.cover_url || agg[id]?.art || null,
        })),
        next_drop: nd || null,
      }), { headers: JH });
    }

    if (!token && !slug) return err(400, "token or slug required");

    let q = admin.from("sc_playlists").select(STATION_COLS);
    q = token ? q.eq("public_token", token) : q.eq("slug", slug).eq("published", true);
    const { data: pl } = await q.maybeSingle();
    if (!pl) return err(404, "Station not found.");

    const { data: tracks } = await admin.from("sc_playlist_tracks").select(TRACK_COLS)
      .eq("playlist_id", pl.id).order("sort");

    const { id: _id, ...station } = pl;
    return new Response(JSON.stringify({
      station: { ...station, soundcloud_url: pl.sc_playlist_url },
      tracks: tracks || [],
    }), { headers: JH });
  } catch (e) {
    console.error("get-station:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load the station.");
  }
});
