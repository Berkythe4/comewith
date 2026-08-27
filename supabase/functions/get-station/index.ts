// get-station  (PUBLIC — deploy with --no-verify-jwt)
//
// Serves Come With Radio to the public radio.html page:
//   ?list=1     → all PUBLISHED episodes (hub view): meta + track count/length.
//   ?s=<slug>   → one PUBLISHED episode by its pretty slug (episode page).
//   ?t=<token>  → any episode by its secret public_token (unlisted preview —
//                 how a station is shared before Keith flips it live).
//   ?artist=<id>→ the PUBLISHED episodes credited to one public artist, for the
//                 "Radio" section on artist.html. Station reads stay function-only
//                 (sc_playlists is anon-revoked, 103), so this is a mode here
//                 rather than a new anon view.
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
const STATION_COLS = "id, slug, name, note, desc_public, published, published_at, status, station_no, drop_date, mix_sc_track_url, mix_sc_track_id, mix_youtube_url, cover_url, mix_by, edition_name, edition_seq, assigned_actor_id, mix_duration_ms";
// sample_url is Beatport's own public preview clip — it lets a track that
// isn't on SoundCloud still be auditioned, including on the phone via a
// preview link. energy/comment are private working notes and stay OUT.
// release_date is here so the render tool can build a full cues file from this
// endpoint alone. That is what lets someone with no database credentials make
// an episode video: an episode token is enough, and nothing else has to be
// handed out. It is the same year already printed on every track card.
const TRACK_COLS = "title, artist_name, permalink_url, sample_url, duration_ms, playback_count, artwork_url, show_date, show_venue, show_cost, show_url, bpm, song_key, camelot, genres, release_date, sort";

// The episode's runtime, from the published mix itself (206). Every surface used
// to sum sc_playlist_tracks.duration_ms, which is the length of the SOURCE
// TRACKS: a DJ set cuts and overlaps them, so the sum over-reports — SHOW 7 read
// 109 min for a 56-min mix. (Beatport-sourced episodes sum from PREVIEW CLIPS
// and land near the truth by coincidence, which is not a measurement either.)
// total_min is still sent — it is a
// truthful "total length of the source tracks" — but it is NOT the runtime and
// no page may render it as one.
const mixMin = (ms: unknown) => (ms == null ? null : Math.round(Number(ms) / 60000));

const norm = (s: unknown) => String(s ?? "").trim().toLowerCase();

// The collective — the only artists a listener can click through to, because
// artist.html?id= renders nothing for an actor without public_profile.
// A name shared by two public actors is dropped rather than guessed: linking a
// credit to the wrong person's profile is worse than not linking it.
async function publicActors(admin: any) {
  const { data } = await admin.from("actors").select("id, display_name")
    .eq("public_profile", true).is("deleted_at", null);
  const byName = new Map<string, string>(), byId = new Map<string, string>(), dupe = new Set<string>();
  for (const a of data || []) {
    byId.set(a.id, a.display_name);
    const n = norm(a.display_name);
    if (!n) continue;
    if (byName.has(n)) dupe.add(n); else byName.set(n, a.id);
  }
  for (const n of dupe) byName.delete(n);
  return { byName, byId };
}

// Which artist an episode is credited to, or null.
//
// The credit is `mix_by` — that is the name PRINTED on the page, so it is the
// name that has to resolve. `assigned_actor_id` is only who was given access to
// build the episode (130); linking on it while a different name is credited
// would point "Mixed by <guest>" at somebody else's profile, so it is the
// fallback only when nothing is credited at all. One rule, used in both
// directions, so the set of episodes linking to an artist is exactly the set
// that artist's page lists back.
function creditedArtist(st: { mix_by?: string | null; assigned_actor_id?: string | null },
                        pa: { byName: Map<string, string>; byId: Map<string, string> }) {
  const credited = norm(st.mix_by);
  if (credited) return pa.byName.get(credited) || null;
  return st.assigned_actor_id && pa.byId.has(st.assigned_actor_id) ? st.assigned_actor_id : null;
}

// Track count / running time / fallback art per station. Paged by PRIMARY KEY:
// PostgREST caps a select at 1000 rows and truncates SILENTLY, and a weekly show
// crosses 1000 tracks somewhere around its fiftieth episode — at which point an
// unpaged read would quietly start reporting short episodes (§18).
async function trackAgg(admin: any, ids: string[]) {
  const agg: Record<string, { n: number; ms: number; art: string | null }> = {};
  if (!ids.length) return agg;
  const PAGE = 1000;
  for (let from = 0; ; from += PAGE) {
    const { data: trs } = await admin.from("sc_playlist_tracks")
      .select("id, playlist_id, duration_ms, artwork_url").in("playlist_id", ids)
      .order("id").range(from, from + PAGE - 1);
    for (const t of trs || []) {
      const a = (agg[t.playlist_id] ||= { n: 0, ms: 0, art: null });
      a.n++; a.ms += t.duration_ms || 0; if (!a.art && t.artwork_url) a.art = t.artwork_url;
    }
    if (!trs || trs.length < PAGE) break;
  }
  return agg;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  try {
    const url = new URL(req.url);
    let token = url.searchParams.get("t") || url.searchParams.get("token") || "";
    let slug = url.searchParams.get("s") || url.searchParams.get("slug") || "";
    let list = url.searchParams.get("list") || "";
    let artist = url.searchParams.get("artist") || "";
    if (!token && !slug && !list && !artist && req.method === "POST") {
      const b = await req.json().catch(() => ({}));
      token = (b.token || "").toString(); slug = (b.slug || "").toString();
      list = (b.list || "").toString(); artist = (b.artist || "").toString();
    }
    token = token.trim(); slug = slug.trim().toLowerCase(); artist = artist.trim();

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
      const agg = await trackAgg(admin, (pls || []).map((p) => p.id));
      const pa = await publicActors(admin);
      // Tease the next scheduled drop (drops are dated in advance): the nearest
      // future-dated station that isn't live yet. Number + date + name only.
      const today = new Date().toISOString().slice(0, 10);
      const { data: nd } = await admin.from("sc_playlists").select("station_no, name, drop_date")
        .eq("published", false).gte("drop_date", today).order("drop_date").limit(1).maybeSingle();
      return new Response(JSON.stringify({
        stations: (pls || []).map(({ id, assigned_actor_id, ...p }) => ({
          ...p,
          // Who "Mixed by <name>" points at on the hub card. Null = the credit is
          // somebody outside the collective, and the name stays plain text.
          mix_by_id: creditedArtist({ ...p, assigned_actor_id }, pa),
          mix_min: mixMin(p.mix_duration_ms),
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

    // An artist's own episodes, for the Radio section on artist.html. Published
    // only — an unlisted preview stays behind its token, which is the whole point
    // of the token. An actor who isn't public gets the same empty answer as an
    // unknown id: this endpoint must never confirm that a private profile exists.
    if (artist) {
      const { data: a } = await admin.from("actors").select("id, display_name")
        .eq("id", artist).eq("public_profile", true).is("deleted_at", null).maybeSingle();
      if (!a) return new Response(JSON.stringify({ artist: null, stations: [] }), { headers: JH });
      const { data: pls } = await admin.from("sc_playlists").select(STATION_COLS)
        .eq("published", true).order("published_at", { ascending: false, nullsFirst: false });
      const pa = await publicActors(admin);
      const mine = (pls || []).filter((p) => creditedArtist(p, pa) === a.id);
      const agg = await trackAgg(admin, mine.map((p) => p.id));
      return new Response(JSON.stringify({
        artist: { id: a.id, display_name: a.display_name },
        stations: mine.map((p) => ({
          slug: p.slug, name: p.name, station_no: p.station_no,
          edition_name: p.edition_name, edition_seq: p.edition_seq,
          published_at: p.published_at, mix_by: p.mix_by,
          mix_min: mixMin(p.mix_duration_ms),
          track_count: agg[p.id]?.n || 0,
          total_min: Math.round((agg[p.id]?.ms || 0) / 60000),
          // Same fallback order as the hub: the episode's own cover, then the
          // show's artwork, then a track's — never a song cover ahead of branding.
          artwork_url: p.cover_url || stationArt || agg[p.id]?.art || null,
        })),
      }), { headers: JH });
    }

    if (!token && !slug) return err(400, "token, slug or artist required");

    let q = admin.from("sc_playlists").select(STATION_COLS);
    q = token ? q.eq("public_token", token) : q.eq("slug", slug).eq("published", true);
    const { data: pl } = await q.maybeSingle();
    if (!pl) return err(404, "Station not found.");

    const { data: tracks } = await admin.from("sc_playlist_tracks").select(TRACK_COLS)
      .eq("playlist_id", pl.id).order("sort");

    const pa = await publicActors(admin);
    const { id: _id, assigned_actor_id, ...station } = pl;
    return new Response(JSON.stringify({
      station: { ...station, station_artwork: stationArt,
                 mix_by_id: creditedArtist({ ...station, assigned_actor_id }, pa),
                 mix_min: mixMin(pl.mix_duration_ms) },
      tracks: tracks || [],
    }), { headers: JH });
  } catch (e) {
    console.error("get-station:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load the station.");
  }
});
