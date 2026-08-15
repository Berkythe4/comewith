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

// A "song" is at most this long; anything above it is a DJ set / mix / livestream.
// Same 15-minute contract sc-enrich and sc-tracks document, enforced here at read
// time because the cache holds whatever cap was in force when each artist was
// scanned. Keep these in step if the contract ever changes.
const SONG_MAX_MS = 15 * 60 * 1000;

// Window-mode paging. There is no silent cap here: MAX_ARTISTS is a safety stop,
// and if it ever binds the response says so (`scope.capped`) rather than handing
// the DJ a short list that looks complete.
const EV_PAGE = 1000;
// Kept well under PostgREST's URL budget — an .in() list is a query string, and
// 200 names of arbitrary length is close enough to the limit to be a hazard.
const ART_CHUNK = 100;
const MAX_ARTISTS = 1500;

// Normalized SoundCloud URL — the join key for the scan cache and the dedupe key
// for artists. Defined once, up here, because both the window pool and the song
// lookup need it.
const norm = (u: string) => (u || "").trim().toLowerCase().replace("://www.", "://").replace(/\/+$/, "").split("?")[0];

// The same artist name can exist under several source rows — Brainrack and Flash
// Gea are on the Elements bill AND have a thin RA row with no soundcloud. Taking
// whichever arrived first silently served those acts an EMPTY crate (35 and 25
// songs cached, 0 delivered). Prefer the row that actually has a profile, then
// the better-followed one.
const better = (a: any, cur: any) => {
  if (!cur) return true;
  const rank = (x: any) => [x.soundcloud ? 1 : 0, x.follower_count || 0];
  const [as, af] = rank(a), [cs, cf] = rank(cur);
  return as > cs || (as === cs && af > cf);
};

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
    const { token, action, track } = await req.json().catch(() => ({}));
    if (!token || typeof token !== "string") return err(400, "missing token");

    const { data: ep } = await admin.from("sc_playlists")
      .select("id, station_no, name, drop_date, dj_search_params, status")
      .eq("dj_token", token).maybeSingle();
    if (!ep) return err(404, "This link is invalid or has been revoked.");

    // The DJ adds a pick to the episode (source='dj' so it's reviewable + only the
    // DJ's own adds are theirs to remove). Token is the only credential.
    if (action === "add") {
      const id = track && (track.sc_track_id != null) ? String(track.sc_track_id) : "";
      if (!id) return err(400, "no track");
      const { data: dup } = await admin.from("sc_playlist_tracks").select("id").eq("playlist_id", ep.id).eq("sc_track_id", id).maybeSingle();
      if (dup) return new Response(JSON.stringify({ ok: true, already: true }), { headers: JH });
      const { data: mx } = await admin.from("sc_playlist_tracks").select("sort").eq("playlist_id", ep.id).order("sort", { ascending: false }).limit(1).maybeSingle();
      const { error } = await admin.from("sc_playlist_tracks").insert({
        playlist_id: ep.id, sc_track_id: id, title: track.title || null, artist_name: track.artist_name || null,
        permalink_url: track.url || track.permalink_url || null, duration_ms: track.duration_ms || null,
        playback_count: track.playback_count || null, artwork_url: track.artwork_url || null,
        source: "dj", sort: (mx?.sort || 0) + 10,
      });
      if (error) return err(500, "Could not add that song.");
      return new Response(JSON.stringify({ ok: true, added: true }), { headers: JH });
    }
    if (action === "remove") {
      const id = track && (track.sc_track_id != null) ? String(track.sc_track_id) : "";
      if (!id) return err(400, "no track");
      // Only the DJ's own picks (source='dj') can be pulled — never the curated set.
      await admin.from("sc_playlist_tracks").delete().eq("playlist_id", ep.id).eq("sc_track_id", id).eq("source", "dj");
      return new Response(JSON.stringify({ ok: true, removed: true }), { headers: JH });
    }

    const params = ep.dj_search_params || {};
    const genres: string[] = Array.isArray(params.genres) ? params.genres.filter(Boolean) : [];
    const artistNames: string[] = Array.isArray(params.artists) ? params.artists.filter(Boolean) : [];
    const weeks = Math.max(1, Math.min(12, Number(params.weeks) || 4));
    const today = new Date().toISOString().slice(0, 10);
    const to = new Date(Date.now() + weeks * 7 * 86400000).toISOString().slice(0, 10);
    const SEL = "name, soundcloud, follower_count, genres, city, next_event_date, next_venue, next_event_url";

    // Two scope modes:
    //  • FIXED LINEUP (params.artists set, e.g. a festival edition) → exactly those
    //    artists, in the given order, no date window.
    //  • DEFAULT → NYC artists playing within the window (+ optional genre filter).
    let artists;
    let poolTotal: number | null = null, capped = false;
    if (artistNames.length) {
      const { data } = await admin.from("ra_artists").select(SEL).in("name", artistNames);
      // The SAME artist name can exist under several sources — Brainrack and
      // Flash Gea are on the Elements bill AND have a thin RA row with no
      // soundcloud. Taking whichever row arrived first silently served those
      // acts an EMPTY crate (35 and 25 songs cached, 0 delivered). Prefer the
      // row that actually has a profile, then the better-followed one.
      const byName: Record<string, any> = {};
      (data || []).forEach((a) => { if (better(a, byName[a.name])) byName[a.name] = a; });
      artists = artistNames.map((n) => byName[n]).filter(Boolean);
    } else {
      // WINDOW MODE — built from ra_events.lineup, NOT ra_artists.next_event_date.
      //
      // `ra_artists` collapses an artist to ONE row carrying next_event_date =
      // their SOONEST show. That is a summary of the pull, not a fact about the
      // artist, so filtering on it drops anyone playing just BEFORE the window
      // and again INSIDE it. It cost the dashboard 77 artists (70 with a
      // SoundCloud link) on the 2026-08-18 window before raWindowPool() was
      // rebuilt on the lineups; this path carried the identical bug, plus a
      // silent .limit(160) against an ~879-artist window. Mirror of
      // raWindowPool() in dashboard.html — keep the two in step.
      // Keyed lowercase (names on a bill drift in casing); `name` keeps the
      // original spelling, which is what ra_artists.name is queried on.
      const showByName: Record<string, { date: string; venue: string | null; genres: string[]; name: string }> = {};
      for (let from = 0; ; from += EV_PAGE) {
        const { data: evs, error: evErr } = await admin.from("ra_events")
          .select("event_date, venue_name, genres, lineup")
          .gte("event_date", today).lte("event_date", to)
          .order("event_date", { ascending: true })
          .range(from, from + EV_PAGE - 1);
        if (evErr) return err(500, "Could not read the show listings.");
        if (!evs || !evs.length) break;
        for (const e of evs) {
          if (!e.event_date) continue;
          for (const a of ((e.lineup as any[]) || [])) {
            const n = ((a && a.name) || "").trim();
            if (!n) continue;
            // Events arrive date-ascending, so the FIRST hit is the soonest show
            // this artist plays inside the window — that's what the DJ should see.
            const k = n.toLowerCase();
            if (!showByName[k]) showByName[k] = { date: e.event_date, venue: e.venue_name || null, genres: e.genres || [], name: n };
          }
        }
        if (evs.length < EV_PAGE) break;
      }

      // Their ra_artists rows. Names come off the bills, which is where those rows
      // were built from, so the casing lines up; the next_event_date sweep below
      // is the backstop that also catches partners and manual adds carrying no
      // event row at all.
      const billed = [...new Set(Object.values(showByName).map((s) => s.name))];
      const rows: any[] = [];
      for (let i = 0; i < billed.length; i += ART_CHUNK) {
        const chunk = billed.slice(i, i + ART_CHUNK);
        const { data } = await admin.from("ra_artists").select(SEL).in("name", chunk);
        if (data) rows.push(...data);
      }
      // Case-insensitive retry is not free, so instead sweep the artists whose own
      // stamped date lands in the window — cheap, and it recovers anyone the name
      // match missed as well as the no-event-row cases.
      // Paged explicitly: PostgREST caps an unbounded select at 1000 rows and says
      // nothing, which is the same silent-shortfall shape this fix exists to kill.
      for (let from = 0; ; from += EV_PAGE) {
        const { data } = await admin.from("ra_artists").select(SEL)
          .gte("next_event_date", today).lte("next_event_date", to)
          .order("next_event_date", { ascending: true })
          .range(from, from + EV_PAGE - 1);
        if (!data || !data.length) break;
        rows.push(...data);
        if (data.length < EV_PAGE) break;
      }

      // Dedupe on the SoundCloud profile (falling back to the name), preferring
      // the row that actually carries a profile.
      const byKey: Record<string, any> = {};
      for (const a of rows) {
        const k = a.soundcloud ? norm(a.soundcloud) : "n:" + (a.name || "").trim().toLowerCase();
        if (better(a, byKey[k])) byKey[k] = a;
      }

      // Re-point each artist at the show they play IN the window (date, venue,
      // that bill's genres) rather than whichever earlier show is on their row.
      let pool = Object.values(byKey).map((a: any) => {
        const s = showByName[(a.name || "").trim().toLowerCase()];
        const own = a.next_event_date >= today && a.next_event_date <= to ? a.next_event_date : null;
        const useShow = s && (!own || s.date <= own);
        return {
          ...a,
          next_event_date: useShow ? s.date : (own || a.next_event_date),
          next_venue: useShow ? (s.venue || a.next_venue) : a.next_venue,
          genres: (useShow && s.genres.length) ? [...new Set([...(a.genres || []), ...s.genres])] : (a.genres || []),
        };
      }).filter((a: any) => a.next_event_date >= today && a.next_event_date <= to);

      // Genre filter runs AFTER the bill's genres are merged in, so an artist
      // carrying no genres of their own still matches on the night they play.
      if (genres.length) pool = pool.filter((a: any) => (a.genres || []).some((g: string) => genres.includes(g)));
      pool.sort((a: any, b: any) => String(a.next_event_date).localeCompare(String(b.next_event_date)));

      poolTotal = pool.length;
      capped = pool.length > MAX_ARTISTS;
      artists = capped ? pool.slice(0, MAX_ARTISTS) : pool;
    }

    // Their songs, from the scan cache (keyed by normalized soundcloud URL).
    const scUrls = [...new Set((artists || []).map((a) => a.soundcloud).filter(Boolean).map(norm))];
    const songByUrl: Record<string, any[]> = {};
    for (let i = 0; i < scUrls.length; i += 200) {
      const { data: cache } = await admin.from("sc_artist_cache")
        .select("soundcloud, songs, is_producer, followers").in("soundcloud", scUrls.slice(i, i + 200));
      // SONGS ONLY — never DJ sets. sc-enrich already applies this cut when it
      // scans, but the cap it used is whatever maxMinutes the operator had set at
      // the time: the Elements lineup was scanned on 2026-07-28 with it wide open,
      // so 73 artists have multi-HOUR uploads sitting in `songs` (Otternonsense
      // 279 min, Cloonee 238). Re-apply the cut at READ time so a bad scan can
      // never put a mix in front of the DJ. Duration is the real distinguisher —
      // mixes are still kind=track on SoundCloud.
      (cache || []).forEach((c) => {
        // NO per-artist cap. There was a .slice(0, 12) here, and it was invisible:
        // twelve songs reads as "that's their catalogue" whether the artist has
        // twelve or a hundred and twelve, so the DJ never knew to look further.
        // The length filter stays — that removes DJ sets, which is a different
        // thing from hiding songs.
        songByUrl[norm(c.soundcloud)] = (c.songs || [])
          .filter((s: any) => !s.duration_ms || Number(s.duration_ms) <= SONG_MAX_MS)
          .map((s: any) => ({ sc_track_id: s.sc_track_id, title: s.title, url: s.permalink_url, duration_ms: s.duration_ms, playback_count: s.playback_count, artwork_url: s.artwork_url }));
      });
    }
    // A sub-group (e.g. the festival's Disco Den stage) so the DJ can sort it out.
    const discoSet = new Set((Array.isArray(params.disco) ? params.disco : []).map((n: string) => n));
    // Which festival day each artist actually plays. Only set when an edition is
    // scoped WIDER than its own day (Ep1 carries the whole festival because it is
    // the early slot) — without it 138 artists arrive as one undifferentiated
    // list and the DJ can't tell their own night from the rest of the weekend.
    const dayOf: Record<string, string> = (params.day_of && typeof params.day_of === "object") ? params.day_of : {};
    const out = (artists || []).map((a) => ({
      name: a.name, soundcloud: a.soundcloud, followers: a.follower_count || 0,
      genres: a.genres || [], city: a.city || null,
      group: discoSet.has(a.name) ? "disco" : "main",
      day: dayOf[a.name] || null,
      next_event_date: a.next_event_date, next_venue: a.next_venue, next_event_url: a.next_event_url,
      songs: a.soundcloud ? (songByUrl[norm(a.soundcloud)] || []) : [],
    }));

    // The episode's current tracklist (what's already in).
    const { data: tracks } = await admin.from("sc_playlist_tracks")
      .select("sc_track_id, artist_name, title, permalink_url, source, sort")
      .eq("playlist_id", ep.id).order("sort");

    return new Response(JSON.stringify({
      ok: true,
      episode: { no: ep.station_no, name: ep.name, drop_date: ep.drop_date },
      scope: {
        weeks, genres, from: today, to, pool: params.pool || null, day: params.day || null,
        // 'all-producers' = this edition deliberately reaches past its own day.
        reach: params.scope || null,
        count: (artists || []).length,
        // A cap that binds is reported, never silent: pool_total is what the
        // window actually holds, capped says the list was cut.
        pool_total: poolTotal,
        capped,
      },
      artists: out,
      tracklist: tracks || [],
    }), { headers: JH });
  } catch (e) {
    console.error("dj-station:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load the episode.");
  }
});
