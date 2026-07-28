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
    if (artistNames.length) {
      const { data } = await admin.from("ra_artists").select(SEL).in("name", artistNames);
      const byName: Record<string, any> = {}; (data || []).forEach((a) => { byName[a.name] = byName[a.name] || a; });
      artists = artistNames.map((n) => byName[n]).filter(Boolean);
    } else {
      let q = admin.from("ra_artists").select(SEL)
        .gte("next_event_date", today).lte("next_event_date", to)
        .order("next_event_date", { ascending: true }).limit(160);
      if (genres.length) q = q.overlaps("genres", genres);
      artists = (await q).data;
    }

    // Their songs, from the scan cache (keyed by normalized soundcloud URL).
    const norm = (u: string) => (u || "").trim().toLowerCase().replace("://www.", "://").replace(/\/+$/, "").split("?")[0];
    const scUrls = [...new Set((artists || []).map((a) => a.soundcloud).filter(Boolean).map(norm))];
    const songByUrl: Record<string, any[]> = {};
    for (let i = 0; i < scUrls.length; i += 200) {
      const { data: cache } = await admin.from("sc_artist_cache")
        .select("soundcloud, songs, is_producer, followers").in("soundcloud", scUrls.slice(i, i + 200));
      (cache || []).forEach((c) => { songByUrl[norm(c.soundcloud)] = (c.songs || []).slice(0, 12).map((s: any) => ({ sc_track_id: s.sc_track_id, title: s.title, url: s.permalink_url, duration_ms: s.duration_ms, playback_count: s.playback_count, artwork_url: s.artwork_url })); });
    }
    const out = (artists || []).map((a) => ({
      name: a.name, soundcloud: a.soundcloud, followers: a.follower_count || 0,
      genres: a.genres || [], city: a.city || null,
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
      scope: { weeks, genres, from: today, to, pool: params.pool || null, day: params.day || null, count: (artists || []).length },
      artists: out,
      tracklist: tracks || [],
    }), { headers: JH });
  } catch (e) {
    console.error("dj-station:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load the episode.");
  }
});
