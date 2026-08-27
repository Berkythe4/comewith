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

// NOTE: sc_playlist_url (the private test-listening playlist Keith pushes in
// step ①) is deliberately NOT selected. The public page must never link the
// SOURCE PLAYLIST — the only SoundCloud/YouTube links a listener gets are the
// FINAL MIX (mix_sc_track_url / mix_youtube_url). To get the songs themselves
// they come to the episode page and export the tracklist.
// mix_sc_track_id lets the page embed the mix by numeric id — the widget then
// survives the mix being moved between accounts / re-uploaded (which orphaned
// EP1's stored permalink and silently broke its player).
// edition_name / edition_seq are what tell the public page it belongs to a
// special run (Elements) rather than the weekly NYC show. The page uses them for
// the theme and to show the number the AUDIENCE knows — Elements Ep1, not SHOW 3.
// Still no sc_playlist_url here: listeners get the final mix, not the source
// playlist. That omission is deliberate, see the note above.
const STATION_COLS = "id, slug, name, note, desc_public, published, published_at, status, station_no, drop_date, mix_sc_track_url, mix_sc_track_id, mix_youtube_url, cover_url, mix_by, edition_name, edition_seq";
// sample_url is Beatport's own public preview clip — it lets a track that
// isn't on SoundCloud still be auditioned, including on the phone via a
// preview link. energy/comment are private working notes and stay OUT.
// release_date is here so the render tool can build a full cues file from this
// endpoint alone. That is what lets someone with no database credentials make
// an episode video: an episode token is enough, and nothing else has to be
// handed out. It is the same year already printed on every track card.
const TRACK_COLS = "title, artist_name, permalink_url, sample_url, duration_ms, playback_count, artwork_url, show_date, show_venue, show_cost, show_url, bpm, song_key, camelot, genres, release_date, sort";

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

    // Show-level artwork ("Come With Radio" itself). Used for any episode that
    // hasn't set its own cover, so a new episode is never artwork-less.
    const { data: artRow } = await admin.from("site_content")
      .select("value").eq("key", "ops.radio_artwork").maybeSingle();
    const stationArt = (artRow?.value || "").toString().trim() || null;

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
          // Episode cover -> SHOW artwork -> a track's cover, in that order. Track
          // art is the last resort only: an episode without its own cover must fall
          // back to "Come With Radio" branding, not to whatever song happens to sit
          // at sort 10. Getting this order wrong put a SoundCloud single's cover on
          // the homepage lead card for EP1 (2026-07-30). radio.html's episode page
          // already used cover_url || station_artwork — this makes the hub, the
          // all-episodes grid and the homepage agree with it.
          artwork_url: p.cover_url || stationArt || agg[id]?.art || null,
          // Exposed so a client can tell a real cover from the branded fallback.
          has_own_cover: !!p.cover_url,
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
      station: { ...station, station_artwork: stationArt },
      tracks: tracks || [],
    }), { headers: JH });
  } catch (e) {
    console.error("get-station:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load the station.");
  }
});
