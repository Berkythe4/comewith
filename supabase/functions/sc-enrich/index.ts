// sc-enrich  (ADMIN)
//
// For a BATCH of SoundCloud profile URLs: resolve each, pull their tracks, and
// classify producer vs DJ-only, caching the result (+ their songs) in
// sc_artist_cache keyed by the profile URL (survives RA re-pulls).
//
// Producer = has >=1 original SONG: kind === "track", 45s <= duration <= maxMin,
// streamable. Long uploads (DJ sets / mixes) are counted as sets, not songs.
//
// Body: { urls: string[] (<=25), maxMinutes?: number (default 15) }
// The dashboard loops this over unscanned artists with a progress bar.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36";
// Normalize a SoundCloud profile URL: https, no www (SC's resolve 404s on www.), no trailing slash / query.
const norm = (u: string) => u.trim().toLowerCase().replace(/^http:\/\//, "https://").replace("://www.", "://").replace(/\/+$/, "").split("?")[0];

async function extractClientId(): Promise<string | null> {
  const home = await (await fetch("https://soundcloud.com/", { headers: { "User-Agent": UA } })).text();
  const scripts = [...home.matchAll(/<script[^>]+src="([^"]+)"/g)].map((m) => m[1]).reverse();
  for (const s of scripts) {
    try { const js = await (await fetch(s, { headers: { "User-Agent": UA } })).text(); const m = js.match(/client_id[:=]"([a-zA-Z0-9]{20,})"/); if (m) return m[1]; } catch { /* next */ }
  }
  return null;
}

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
    const urls: string[] = Array.isArray(b.urls) ? b.urls.slice(0, 25).map(norm) : [];
    if (!urls.length) return err(400, "urls[] required");
    const maxMs = Math.max(3, Math.min(30, Number(b.maxMinutes) || 15)) * 60000;
    const minMs = 45000; // ignore sub-45s clips/IDs

    const { data: cidRow } = await admin.from("site_content").select("value").eq("key", "ops.sc_client_id").maybeSingle();
    let cid = (cidRow?.value || "").trim();
    if (!cid) { cid = (await extractClientId()) || ""; if (cid) await admin.from("site_content").update({ value: cid }).eq("key", "ops.sc_client_id"); }
    if (!cid) return err(502, "SoundCloud unavailable (no client_id).");

    const api = "https://api-v2.soundcloud.com";
    const call = (path: string) => fetch(`${api}${path}${path.includes("?") ? "&" : "?"}client_id=${cid}`, { headers: { "User-Agent": UA } });

    const rows: Record<string, unknown>[] = [];
    let refreshed = false;
    for (const url of urls) {
      try {
        let r = await call(`/resolve?url=${encodeURIComponent(url)}`);
        if ((r.status === 401 || r.status === 403) && !refreshed) {
          refreshed = true; const fresh = await extractClientId();
          if (fresh) { cid = fresh; await admin.from("site_content").update({ value: cid }).eq("key", "ops.sc_client_id"); r = await call(`/resolve?url=${encodeURIComponent(url)}`); }
        }
        if (!r.ok) { rows.push({ soundcloud: url, ok: false, is_producer: false, song_count: 0, set_count: 0, songs: [], scanned_at: new Date().toISOString() }); continue; }
        const u = await r.json();
        if (u?.kind !== "user" || !u.id) { rows.push({ soundcloud: url, ok: false, scanned_at: new Date().toISOString() }); continue; }
        const tr = await (await call(`/users/${u.id}/tracks?limit=50`)).json();
        const all = (tr.collection || []) as any[];
        const isSong = (t: any) => t.kind === "track" && (t.duration || 0) >= minMs && (t.duration || 0) <= maxMs && t.streamable !== false;
        const songsRaw = all.filter(isSong);
        const setCount = all.filter((t) => t.kind === "track" && (t.duration || 0) > maxMs).length;
        const songs = songsRaw
          .sort((a, b) => (b.created_at || "").localeCompare(a.created_at || ""))
          .slice(0, 30)
          .map((t) => ({ sc_track_id: String(t.id), title: t.title, permalink_url: t.permalink_url, duration_ms: t.duration, playback_count: t.playback_count ?? 0, created_at: t.created_at, artwork_url: t.artwork_url || u.avatar_url || null }));
        rows.push({
          soundcloud: url, sc_user_id: String(u.id), username: u.username, avatar_url: u.avatar_url || null,
          followers: u.followers_count ?? null, is_producer: songs.length > 0, song_count: songs.length,
          set_count: setCount, songs, ok: true, scanned_at: new Date().toISOString(),
        });
      } catch (_) {
        rows.push({ soundcloud: url, ok: false, scanned_at: new Date().toISOString() });
      }
    }

    if (rows.length) {
      const { error } = await admin.from("sc_artist_cache").upsert(rows, { onConflict: "soundcloud" });
      if (error) { console.error("sc-enrich upsert:", error.message); return err(500, "Could not save results."); }
    }
    return new Response(JSON.stringify({
      success: true, scanned: rows.length,
      producers: rows.filter((r) => r.is_producer).length,
      failed: rows.filter((r) => r.ok === false).length,
    }), { headers: JH });
  } catch (e) {
    console.error("sc-enrich:", e instanceof Error ? e.message : String(e));
    return err(502, "SoundCloud scan failed — try again shortly.");
  }
});
