// sc-connect  (ADMIN — verify_jwt on)
//
// Admin-side SoundCloud actions:
//   action=status → { configured, connected, username }
//   action=start  → { authorize_url }  (generates PKCE state+verifier, stores them)
//   action=export → creates/updates a real SoundCloud playlist from an in-app
//                    station's tracks; refreshes the access token if expired.
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

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const SUPA = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
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

  if (action === "status") {
    return ok({ configured: !!(clientId && clientSecret), connected: !!row?.access_token, username: row?.sc_username || null });
  }

  if (!clientId || !clientSecret) return err(400, "SoundCloud isn't configured yet — the app's Client ID/Secret still need to be set.");

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

    // Refresh the access token if it's within a minute of expiry (refresh tokens are single-use).
    let token = row.access_token;
    if (row.expires_at && new Date(row.expires_at).getTime() < Date.now() + 60000 && row.refresh_token) {
      const rr = await fetch("https://secure.soundcloud.com/oauth/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded", "accept": "application/json; charset=utf-8" },
        body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, client_secret: clientSecret, refresh_token: row.refresh_token }),
      });
      const rt = await rr.json();
      if (rr.ok && rt.access_token) {
        token = rt.access_token;
        await admin.from("sc_oauth").update({ access_token: rt.access_token, refresh_token: rt.refresh_token || row.refresh_token, expires_at: new Date(Date.now() + (rt.expires_in || 3600) * 1000).toISOString(), updated_at: new Date().toISOString() }).eq("id", "singleton");
      } else { console.error("sc refresh:", JSON.stringify(rt).slice(0, 160)); return err(401, "SoundCloud connection expired — reconnect."); }
    }

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
        const cr = await fetch(`https://api.soundcloud.com/tracks/${t.id}`, { headers: { "Authorization": "OAuth " + token, "accept": "application/json; charset=utf-8" } });
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
      headers: { "Authorization": "OAuth " + token, "Content-Type": "application/x-www-form-urlencoded", "accept": "application/json; charset=utf-8" },
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
    const { data: pl } = await admin.from("sc_playlists").select("id, sc_playlist_id").eq("id", playlistId).single();
    if (!pl?.sc_playlist_id) return err(400, "Export this station to SoundCloud first, then reorder it there and sync.");

    let token = row.access_token;
    if (row.expires_at && new Date(row.expires_at).getTime() < Date.now() + 60000 && row.refresh_token) {
      const rr = await fetch("https://secure.soundcloud.com/oauth/token", {
        method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded", "accept": "application/json; charset=utf-8" },
        body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, client_secret: clientSecret, refresh_token: row.refresh_token }),
      });
      const rt = await rr.json();
      if (rr.ok && rt.access_token) { token = rt.access_token; await admin.from("sc_oauth").update({ access_token: rt.access_token, refresh_token: rt.refresh_token || row.refresh_token, expires_at: new Date(Date.now() + (rt.expires_in || 3600) * 1000).toISOString() }).eq("id", "singleton"); }
      else return err(401, "SoundCloud connection expired — reconnect.");
    }

    const scRes = await fetch(`https://api.soundcloud.com/playlists/${pl.sc_playlist_id}`, {
      headers: { "Authorization": "OAuth " + token, "accept": "application/json; charset=utf-8" },
    });
    if (!scRes.ok) { if (scRes.status === 401) return err(401, "SoundCloud connection expired — reconnect."); return err(502, "Couldn't read the playlist from SoundCloud."); }
    const j = await scRes.json().catch(() => ({}));
    const scTracks = (j.tracks || []) as any[];
    if (!scTracks.length) return err(400, "That SoundCloud playlist is empty or unreadable.");

    const { data: existing } = await admin.from("sc_playlist_tracks").select("id, sc_track_id").eq("playlist_id", playlistId);
    const byId: Record<string, string> = {}; (existing || []).forEach((t) => { byId[t.sc_track_id] = t.id; });
    const scIds = new Set<string>();
    let pos = 0, added = 0;
    for (const t of scTracks) {
      if (t?.id == null) continue;
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
    const toRemove = (existing || []).filter((t) => !scIds.has(t.sc_track_id)).map((t) => t.id);
    if (toRemove.length) await admin.from("sc_playlist_tracks").delete().in("id", toRemove);
    return ok({ success: true, tracks: scIds.size, added, removed: toRemove.length });
  }

  return err(400, "unknown action");
});
