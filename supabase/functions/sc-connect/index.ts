// sc-connect  (ADMIN — verify_jwt on)
//
// Admin-side SoundCloud actions:
//   action=status     → { configured, connected, username }
//   action=start      → { authorize_url }  (generates PKCE state+verifier, stores them)
//   action=export     → creates/updates a real SoundCloud playlist from an in-app
//                       station's tracks; refreshes the access token if expired.
//   action=sync       → pulls the exported playlist's final track order back in.
//   action=upload_mix → attach the recorded MIX to a station as a SoundCloud track:
//                       either resolve a pasted track link (recommended for big
//                       HQ files uploaded via SoundCloud's own uploader), or
//                       stream the mix file from the radio-mixes bucket to
//                       POST /tracks (private) — the API-upload path.
//   action=finalize   → go live: slug + descriptions + published on the station,
//                       and push the short description (with the site link) to
//                       the SoundCloud mix track + flip it public.
// The public sc-oauth function handles the browser redirect + token exchange.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const ok = (o: unknown) => new Response(JSON.stringify(o), { headers: JH });

const b64url = (buf: ArrayBuffer | Uint8Array) => {
  const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  return btoa(String.fromCharCode(...bytes)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
};
const SC_ACCEPT = "application/json; charset=utf-8";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const SUPA = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const SITE = (Deno.env.get("SITE_URL") || "https://comewith.org").replace(/\/+$/, "");
  const admin = createClient(SUPA, SRK);

  // Admin auth.
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

  const clientId = Deno.env.get("SC_CLIENT_ID"), clientSecret = Deno.env.get("SC_CLIENT_SECRET");
  const redirectUri = `${SUPA}/functions/v1/sc-oauth`;
  const body = await req.json().catch(() => ({}));
  const action = (body.action || "status").toString();

  const { data: row } = await admin.from("sc_oauth").select("*").eq("id", "singleton").maybeSingle();

  // Valid access token, refreshing when within a minute of expiry (refresh
  // tokens are single-use). Returns null when the connection is dead.
  async function freshToken(): Promise<string | null> {
    if (!row?.access_token) return null;
    if (!(row.expires_at && new Date(row.expires_at).getTime() < Date.now() + 60000 && row.refresh_token)) return row.access_token;
    const rr = await fetch("https://secure.soundcloud.com/oauth/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded", "accept": SC_ACCEPT },
      body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId!, client_secret: clientSecret!, refresh_token: row.refresh_token }),
    });
    const rt = await rr.json();
    if (rr.ok && rt.access_token) {
      await admin.from("sc_oauth").update({ access_token: rt.access_token, refresh_token: rt.refresh_token || row.refresh_token, expires_at: new Date(Date.now() + (rt.expires_in || 3600) * 1000).toISOString(), updated_at: new Date().toISOString() }).eq("id", "singleton");
      return rt.access_token;
    }
    console.error("sc refresh:", JSON.stringify(rt).slice(0, 160));
    return null;
  }

  if (action === "status") {
    return ok({ configured: !!(clientId && clientSecret), connected: !!row?.access_token, username: row?.sc_username || null });
  }

  if (!clientId || !clientSecret) return err(400, "SoundCloud isn't configured yet — the app's Client ID/Secret still need to be set.");

  // Play counts for the finished mixes. SoundCloud has no webhook and nothing
  // else in the app ever read a track's stats back, so the episode's reach was
  // invisible — the KPI cards (110) run on what this stores.
  if (action === "mix_stats") {
    if (!row?.access_token) return err(400, "Connect SoundCloud first.");
    const token = await freshToken();
    if (!token) return err(400, "SoundCloud connection expired — reconnect and try again.");
    const { data: stations } = await admin.from("sc_playlists")
      .select("id, station_no, mix_sc_track_id").eq("status", "live").not("mix_sc_track_id", "is", null);
    if (!stations?.length) return ok({ success: true, updated: 0, note: "No live episode has a SoundCloud mix linked yet." });

    let updated = 0; const failed: number[] = [];
    for (const s of stations) {
      const r = await fetch(`https://api.soundcloud.com/tracks/${s.mix_sc_track_id}`, {
        headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT },
      });
      if (!r.ok) { failed.push(s.station_no); continue; }
      const j = await r.json().catch(() => null);
      if (!j) { failed.push(s.station_no); continue; }
      const num = (v: unknown) => (v == null ? null : Number(v));
      await admin.from("sc_playlists").update({
        mix_sc_plays: num(j.playback_count),
        // The API has moved from favoritings_count to likes_count and still
        // returns the old name on some tracks; take whichever is present.
        mix_sc_likes: num(j.likes_count ?? j.favoritings_count),
        mix_sc_reposts: num(j.reposts_count),
        mix_sc_comments: num(j.comment_count),
        mix_stats_at: new Date().toISOString(),
      }).eq("id", s.id);
      updated++;
    }
    return ok({ success: true, updated, failed });
  }

  if (action === "start") {
    const verifier = b64url(crypto.getRandomValues(new Uint8Array(48)));
    const challenge = b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
    const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
    // Upsert only these columns — existing tokens (if any) are preserved.
    await admin.from("sc_oauth").upsert({ id: "singleton", state, code_verifier: verifier, updated_at: new Date().toISOString() });
    const authorize = `https://secure.soundcloud.com/authorize?response_type=code&client_id=${encodeURIComponent(clientId)}&redirect_uri=${encodeURIComponent(redirectUri)}&code_challenge=${challenge}&code_challenge_method=S256&state=${state}`;
    return ok({ authorize_url: authorize });
  }

  if (action === "export") {
    if (!row?.access_token) return err(400, "Connect SoundCloud first.");
    const playlistId = (body.playlist_id || "").toString();
    if (!playlistId) return err(400, "playlist_id required");

    const token = await freshToken();
    if (!token) return err(401, "SoundCloud connection expired — reconnect.");

    const { data: pl } = await admin.from("sc_playlists").select("id, name, sc_playlist_id").eq("id", playlistId).single();
    if (!pl) return err(404, "playlist not found");
    const { data: tracks } = await admin.from("sc_playlist_tracks").select("sc_track_id, title").eq("playlist_id", playlistId).order("sort");
    const wanted = (tracks || []).map((t) => ({ id: Number(t.sc_track_id), title: (t.title || "track " + t.sc_track_id).toString() })).filter((t) => t.id);
    if (!wanted.length) return err(400, "This station has no songs to export.");

    // SoundCloud rejects the WHOLE playlist if even one track can't be added via
    // the public API (uploader disabled off-SoundCloud/API embedding, or it went
    // private/deleted since we cached it). Pre-check each track with this user's
    // token and keep only the ones the API will actually accept — reporting the
    // rest instead of failing the whole export.
    const checks = await Promise.all(wanted.map(async (t) => {
      try {
        const cr = await fetch(`https://api.soundcloud.com/tracks/${t.id}`, { headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT } });
        const tj = cr.ok ? await cr.json().catch(() => ({})) : {};
        // Reject deleted/private (non-200) and tracks explicitly blocked from embedding.
        return { ...t, okTrack: cr.ok && tj.embeddable_by !== "none" };
      } catch { return { ...t, okTrack: false }; }
    }));
    const good = checks.filter((t) => t.okTrack);
    const skipped = checks.filter((t) => !t.okTrack).map((t) => t.title);
    if (!good.length) return err(422, `SoundCloud won't let any of these ${wanted.length} track(s) be added via the API — the uploaders have off-SoundCloud/API sharing turned off. Try a station with different songs.`);
    const trackObjs = good.map((t) => ({ id: t.id }));

    const description = (body.description || "").toString().slice(0, 4000) || "Built in the Come With dashboard from upcoming NYC artists.";
    // SoundCloud's playlist endpoint rejects a JSON body ("Could not parse JSON
    // request body") — it wants Rails-style nested form params instead.
    const form = new URLSearchParams();
    form.set("playlist[title]", pl.name || "Come With station");
    form.set("playlist[description]", description);
    form.set("playlist[sharing]", "private");
    for (const t of trackObjs) form.append("playlist[tracks][][id]", String(t.id));
    const isUpdate = !!pl.sc_playlist_id;
    const scRes = await fetch(`https://api.soundcloud.com/playlists${isUpdate ? "/" + pl.sc_playlist_id : ""}`, {
      method: isUpdate ? "PUT" : "POST",
      headers: { "Authorization": "OAuth " + token, "Content-Type": "application/x-www-form-urlencoded", "accept": SC_ACCEPT },
      body: form.toString(),
    });
    const j = await scRes.json().catch(() => ({}));
    if (!scRes.ok) {
      console.error("sc playlist:", scRes.status, JSON.stringify(j).slice(0, 300));
      if (scRes.status === 401) return err(401, "SoundCloud connection expired — reconnect.");
      const scMsg = (j?.errors?.[0]?.error_message || j?.error?.message || j?.error || j?.message || "").toString().slice(0, 160);
      return err(502, "SoundCloud rejected the playlist" + (scMsg ? ": " + scMsg : ". It may not recognize one of the track IDs."));
    }
    const purl = j.permalink_url || null, pid = j.id != null ? String(j.id) : pl.sc_playlist_id;
    await admin.from("sc_playlists").update({ sc_playlist_id: pid, sc_playlist_url: purl, updated_at: new Date().toISOString() }).eq("id", playlistId);
    return ok({ success: true, url: purl, updated: isUpdate, tracks: trackObjs.length, skipped });
  }

  // Pull the exported SoundCloud playlist back — capture the final track order
  // (and any tracks added/removed on SoundCloud) into the tool.
  if (action === "sync") {
    if (!row?.access_token) return err(400, "Connect SoundCloud first.");
    const playlistId = (body.playlist_id || "").toString();
    if (!playlistId) return err(400, "playlist_id required");
    const { data: pl } = await admin.from("sc_playlists").select("id, sc_playlist_id, station_no").eq("id", playlistId).single();
    if (!pl?.sc_playlist_id) return err(400, "Export this station to SoundCloud first, then reorder it there and sync.");

    const token = await freshToken();
    if (!token) return err(401, "SoundCloud connection expired — reconnect.");

    const scRes = await fetch(`https://api.soundcloud.com/playlists/${pl.sc_playlist_id}?access=playable,preview,blocked`, {
      headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT },
    });
    if (!scRes.ok) { if (scRes.status === 401) return err(401, "SoundCloud connection expired — reconnect."); return err(502, "Couldn't read the playlist from SoundCloud."); }
    const j = await scRes.json().catch(() => ({}));
    const scTracks = ((j.tracks || []) as any[]).filter((t) => t?.id != null);
    if (!scTracks.length) return err(400, "That SoundCloud playlist is empty or unreadable.");

    // SoundCloud processes a reorder as a remove-then-re-add; for a moment the API
    // returns the playlist with the in-flight tracks MISSING. `track_count` is the
    // true size — if fewer tracks came back, the snapshot is mid-update, so we
    // reorder/add but NEVER delete (a reorder must never lose songs).
    const trackCount = typeof j.track_count === "number" ? j.track_count : scTracks.length;
    const complete = scTracks.length >= trackCount;

    const { data: existing } = await admin.from("sc_playlist_tracks").select("id, sc_track_id, title, artist_name, permalink_url, artwork_url, duration_ms").eq("playlist_id", playlistId);
    const byId: Record<string, string> = {}; (existing || []).forEach((t) => { byId[t.sc_track_id] = t.id; });
    const scIds = new Set<string>();
    let pos = 0, added = 0;
    for (const t of scTracks) {
      const sid = String(t.id); scIds.add(sid); pos += 10;
      if (byId[sid]) {
        await admin.from("sc_playlist_tracks").update({ sort: pos }).eq("id", byId[sid]);
      } else {
        added++;
        await admin.from("sc_playlist_tracks").insert({
          playlist_id: playlistId, sc_track_id: sid, title: t.title || "(untitled)", artist_name: t.user?.username || null,
          permalink_url: t.permalink_url || null, duration_ms: t.duration || null, playback_count: t.playback_count || null,
          artwork_url: t.artwork_url || null, sort: pos,
        });
      }
    }
    // Only remove locally when the snapshot is trustworthy (complete). A track
    // cut on SoundCloud while test-listening = "considered and PASSED" — log it
    // in the permanent song memory so it (a) shows a ✋ mark when researching
    // future stations and (b) auto-carries into next week's station at finalize.
    const cut = complete ? (existing || []).filter((t) => !scIds.has(t.sc_track_id)) : [];
    if (cut.length) {
      const now = new Date().toISOString();
      for (const t of cut) {
        await admin.from("sc_song_log").upsert({
          sc_track_id: t.sc_track_id, title: t.title, artist_name: t.artist_name,
          permalink_url: t.permalink_url, artwork_url: t.artwork_url, duration_ms: t.duration_ms,
          passed_playlist_id: playlistId, passed_station_no: pl.station_no ?? null, passed_at: now, updated_at: now,
        }, { onConflict: "sc_track_id" });
      }
      await admin.from("sc_playlist_tracks").delete().in("id", cut.map((t) => t.id));
    }
    return ok({ success: true, tracks: scIds.size, added, removed: cut.length, incomplete: !complete });
  }

  // Attach the recorded mix to a station as a SoundCloud TRACK. Two paths:
  //  - body.track_url: Keith uploaded the mix with SoundCloud's own uploader
  //    (best for big high-quality files) — resolve the link to a track id.
  //  - stored file: stream the mix from the radio-mixes bucket to POST /tracks
  //    as a PRIVATE track (flips public at finalize).
  if (action === "upload_mix") {
    if (!row?.access_token) return err(400, "Connect SoundCloud first.");
    const playlistId = (body.playlist_id || "").toString();
    if (!playlistId) return err(400, "playlist_id required");
    const { data: pl } = await admin.from("sc_playlists").select("id, name, status, mix_file_path").eq("id", playlistId).single();
    if (!pl) return err(404, "station not found");

    const token = await freshToken();
    if (!token) return err(401, "SoundCloud connection expired — reconnect.");

    const trackUrl = (body.track_url || "").toString().trim();
    let tj: any = null;
    if (trackUrl) {
      const r = await fetch(`https://api.soundcloud.com/resolve?url=${encodeURIComponent(trackUrl)}`, { headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT } });
      tj = await r.json().catch(() => ({}));
      if (!r.ok || tj.kind !== "track") return err(400, "That link doesn't resolve to a SoundCloud track — paste the track's own page URL.");
    } else {
      if (!pl.mix_file_path) return err(400, "Upload the mix file first, or paste the SoundCloud track link instead.");
      const dl = await admin.storage.from("radio-mixes").download(pl.mix_file_path);
      if (!dl.data) return err(400, "Couldn't read the uploaded mix from storage" + (dl.error ? `: ${dl.error.message}` : "."));
      const fname = pl.mix_file_path.split("/").pop() || "mix.mp3";
      const fd = new FormData();
      fd.set("track[title]", `${pl.name || "Come With station"} — Come With Radio`);
      fd.set("track[sharing]", "private");
      fd.set("track[asset_data]", new File([dl.data], fname));
      const up = await fetch("https://api.soundcloud.com/tracks", {
        method: "POST", headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT }, body: fd,
      });
      tj = await up.json().catch(() => ({}));
      if (!up.ok) {
        console.error("sc track upload:", up.status, JSON.stringify(tj).slice(0, 300));
        if (up.status === 401) return err(401, "SoundCloud connection expired — reconnect.");
        const scMsg = (tj?.errors?.[0]?.error_message || tj?.error || tj?.message || "").toString().slice(0, 200);
        return err(502, "SoundCloud rejected the upload" + (scMsg ? ": " + scMsg : " — you can upload it on soundcloud.com and paste the track link instead."));
      }
    }
    await admin.from("sc_playlists").update({
      mix_sc_track_id: tj.id != null ? String(tj.id) : null,
      mix_sc_track_url: tj.permalink_url || trackUrl || null,
      cover_url: tj.artwork_url || null,
      status: pl.status === "building" ? "testing" : pl.status,
      updated_at: new Date().toISOString(),
    }).eq("id", playlistId);
    return ok({ success: true, url: tj.permalink_url || trackUrl, id: tj.id != null ? String(tj.id) : null, source: trackUrl ? "link" : "upload" });
  }

  // Go live: slug + descriptions + published flag on the station, then push the
  // short description (with the episode page link) to the SoundCloud mix track
  // and flip it public. SoundCloud failures don't block the page going live —
  // they come back as sc_warning.
  if (action === "finalize") {
    const playlistId = (body.playlist_id || "").toString();
    if (!playlistId) return err(400, "playlist_id required");
    const { data: pl } = await admin.from("sc_playlists").select("id, name, slug, status, station_no, published_at, mix_sc_track_id, mix_sc_track_url").eq("id", playlistId).single();
    if (!pl) return err(404, "station not found");
    const { data: plTracks } = await admin.from("sc_playlist_tracks").select("sc_track_id, title, artist_name, permalink_url, artwork_url, duration_ms").eq("playlist_id", playlistId).order("sort");
    if (!plTracks?.length) return err(400, "This station has no tracks — nothing to publish.");

    const slugify = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60);
    let slug = slugify((body.slug || "").toString()) || pl.slug || (slugify(pl.name || "station") + "-" + new Date().toISOString().slice(0, 10));
    const descPublic = (body.desc_public || "").toString().slice(0, 8000) || null;
    const descSc = (body.desc_sc || "").toString().slice(0, 4000) || null;
    const youtube = (body.youtube_url || "").toString().trim() || null;

    // Claim the slug; on a collision (another episode owns it) suffix -2, -3…
    const base = slug;
    for (let i = 0; i < 5; i++) {
      const { error } = await admin.from("sc_playlists").update({
        slug, desc_public: descPublic, desc_sc: descSc, mix_youtube_url: youtube,
        published: true, status: "live",
        published_at: pl.published_at || new Date().toISOString(),
        updated_at: new Date().toISOString(),
      }).eq("id", playlistId);
      if (!error) break;
      if (error.code !== "23505" || i === 4) return err(500, "Couldn't publish: " + error.message);
      slug = `${base}-${i + 2}`;
    }

    let scWarning: string | null = null;
    if (pl.mix_sc_track_id) {
      const token = await freshToken();
      if (!token) scWarning = "SoundCloud connection expired — page is live, but the mix description wasn't pushed. Reconnect and finalize again.";
      else {
        const form = new URLSearchParams();
        if (descSc) form.set("track[description]", descSc);
        form.set("track[sharing]", "public");
        const r = await fetch(`https://api.soundcloud.com/tracks/${pl.mix_sc_track_id}`, {
          method: "PUT", headers: { "Authorization": "OAuth " + token, "Content-Type": "application/x-www-form-urlencoded", "accept": SC_ACCEPT },
          body: form.toString(),
        });
        if (!r.ok) {
          const j = await r.json().catch(() => ({}));
          console.error("sc finalize track:", r.status, JSON.stringify(j).slice(0, 200));
          scWarning = "Page is live, but SoundCloud rejected the track update — set the description/public flag on soundcloud.com manually.";
        } else {
          // The URL we stored at upload time is the PRIVATE share link
          // (.../<track>/s-XXXX) — the embed widget 404s on that shape. Flipping
          // the track public changes its permalink to the clean canonical one,
          // so re-capture it here or the episode page keeps a broken player.
          const tj = await r.json().catch(() => ({} as Record<string, unknown>));
          const fresh = typeof tj?.permalink_url === "string" ? tj.permalink_url : "";
          if (fresh && fresh !== pl.mix_sc_track_url) {
            await admin.from("sc_playlists").update({ mix_sc_track_url: fresh, updated_at: new Date().toISOString() }).eq("id", playlistId);
          }
        }
      }
    } else {
      scWarning = "No SoundCloud mix track linked yet — the page is live without one.";
    }

    // Permanent song memory: everything in a live episode is PLAYED (in EP n, on
    // this date). Clears nothing else — a song's earlier passed history stays.
    const nowIso = new Date().toISOString();
    for (const t of plTracks) {
      await admin.from("sc_song_log").upsert({
        sc_track_id: t.sc_track_id, title: t.title, artist_name: t.artist_name,
        permalink_url: t.permalink_url, artwork_url: t.artwork_url, duration_ms: t.duration_ms,
        played_playlist_id: playlistId, played_station_no: pl.station_no ?? null, played_at: nowIso, updated_at: nowIso,
      }, { onConflict: "sc_track_id" });
    }

    // Start next week's station (next number) and auto-carry every song that was
    // considered + passed but never played and never carried before. carried_to
    // is set once, so cutting a carried song again won't resurrect it forever.
    let nextInfo: { id: string; station_no: number; carried: number } | null = null;
    try {
      const { data: maxRow } = await admin.from("sc_playlists").select("station_no").not("station_no", "is", null).order("station_no", { ascending: false }).limit(1).maybeSingle();
      const nextNo = (maxRow?.station_no || 0) + 1;
      const { data: next } = await admin.from("sc_playlists").insert({ name: "Weekly station", station_no: nextNo }).select("id, station_no").single();
      if (next) {
        const { data: carry } = await admin.from("sc_song_log").select("sc_track_id, title, artist_name, permalink_url, artwork_url, duration_ms, passed_station_no")
          .not("passed_at", "is", null).is("played_at", null).is("carried_to", null).order("passed_at");
        let pos = 0;
        for (const c of carry || []) {
          pos += 10;
          const ins = await admin.from("sc_playlist_tracks").insert({
            playlist_id: next.id, sc_track_id: c.sc_track_id, title: c.title, artist_name: c.artist_name,
            permalink_url: c.permalink_url, artwork_url: c.artwork_url, duration_ms: c.duration_ms,
            sort: pos, carried_from: c.passed_station_no ?? null,
          });
          if (!ins.error) await admin.from("sc_song_log").update({ carried_to: next.id, updated_at: nowIso }).eq("sc_track_id", c.sc_track_id);
        }
        nextInfo = { id: next.id, station_no: next.station_no, carried: (carry || []).length };
      }
    } catch (e) { console.error("finalize next-station:", e instanceof Error ? e.message : String(e)); }

    // Drop a "posted" card on the social calendar so releases show up alongside
    // the rest of the content plan. Best-effort — never blocks going live.
    try {
      await admin.from("social_posts").insert({
        title: `📻 Come With Radio EP ${pl.station_no ?? ""} — ${pl.name || ""}`.trim(),
        caption: (descSc || "").slice(0, 1000) || null,
        channels: ["other"], series: "Come With Radio", content_pillar: "radio episode",
        stage: "posted", scheduled_for: nowIso, posted_at: nowIso,
        link_url: `${SITE}/radio.html?s=${slug}`,
      });
    } catch (e) { console.error("finalize social post:", e instanceof Error ? e.message : String(e)); }

    return ok({ success: true, slug, page_url: `${SITE}/radio.html?s=${slug}`, sc_url: pl.mix_sc_track_url, sc_warning: scWarning, next: nextInfo });
  }

  return err(400, "unknown action");
});
