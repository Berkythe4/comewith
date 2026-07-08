// sc-tracks
//
// Admin-only. Given a SoundCloud artist PROFILE url, returns that artist's
// SONGS (not DJ sets) via SoundCloud's internal read API (api-v2.soundcloud.com).
// "Songs not sets": keep kind === "track" AND duration under maxMinutes (mixes
// are still kind=track — duration is the real distinguisher). SoundCloud "sets"
// are kind=playlist and are excluded outright.
//
// The public client_id rotates; we cache the working one in site_content
// (ops.sc_client_id) and re-extract from soundcloud.com's JS bundles on failure.
//
// Body: { url: string, maxMinutes?: number (default 15), limit?: number }

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36";

async function extractClientId(): Promise<string | null> {
  const home = await (await fetch("https://soundcloud.com/", { headers: { "User-Agent": UA } })).text();
  const scripts = [...home.matchAll(/<script[^>]+src="([^"]+)"/g)].map((m) => m[1]).reverse();
  for (const s of scripts) {
    try {
      const js = await (await fetch(s, { headers: { "User-Agent": UA } })).text();
      const m = js.match(/client_id[:=]"([a-zA-Z0-9]{20,})"/);
      if (m) return m[1];
    } catch { /* try next bundle */ }
  }
  return null;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const URLS = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const admin = createClient(URLS, SRK);

  // Auth: admin JWT OR service-role (role-claim check — robust to key format).
  const auth = req.headers.get("Authorization") || "";
  const bearer = auth.replace(/^Bearer\s+/i, "");
  const roleOf = (t: string) => { try { return JSON.parse(atob(t.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))).role || null; } catch { return null; } };
  let authed = bearer === SRK || roleOf(bearer) === "service_role";
  if (!authed && bearer) {
    const uc = createClient(URLS, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
    const { data: { user } } = await uc.auth.getUser();
    if (user) {
      const { data: prof } = await admin.from("profiles").select("role").eq("id", user.id).single();
      authed = !!prof && ["master_admin", "sub_admin"].includes(prof.role);
    }
  }
  if (!authed) return err(401, "admin only");

  try {
    const b = await req.json().catch(() => ({}));
    const url = (b.url || "").toString().trim();
    if (!url || !/soundcloud\.com/i.test(url)) return err(400, "A SoundCloud profile URL is required.");
    const maxMs = Math.max(3, Math.min(60, Number(b.maxMinutes) || 15)) * 60000;
    const limit = Math.max(5, Math.min(100, Number(b.limit) || 50));

    // Get a cached client_id; refresh + persist if missing/stale.
    const { data: cidRow } = await admin.from("site_content").select("value").eq("key", "ops.sc_client_id").maybeSingle();
    let cid = (cidRow?.value || "").trim();

    const api = "https://api-v2.soundcloud.com";
    const call = async (path: string) => {
      const sep = path.includes("?") ? "&" : "?";
      return fetch(`${api}${path}${sep}client_id=${cid}`, { headers: { "User-Agent": UA } });
    };

    // Resolve the profile → user id. On 401/403, refresh the client_id once.
    const resolvePath = `/resolve?url=${encodeURIComponent(url)}`;
    let r = cid ? await call(resolvePath) : new Response(null, { status: 401 });
    if (!cid || r.status === 401 || r.status === 403) {
      const fresh = await extractClientId();
      if (!fresh) return err(502, "SoundCloud is unavailable right now — try again shortly.");
      cid = fresh;
      await admin.from("site_content").update({ value: cid }).eq("key", "ops.sc_client_id");
      r = await call(resolvePath);
    }
    if (!r.ok) return err(404, "Couldn't find that SoundCloud profile.");
    const u = await r.json();
    if (u?.kind !== "user" || !u.id) return err(404, "That link isn't a SoundCloud artist profile.");

    const tr = await (await call(`/users/${u.id}/tracks?limit=${limit}`)).json();
    const all = (tr.collection || []) as any[];
    const songs = all
      .filter((t) => t.kind === "track" && (t.duration || 0) > 0 && (t.duration || 0) <= maxMs && (t.streamable !== false))
      .map((t) => ({
        sc_track_id: String(t.id),
        title: t.title,
        artist_name: u.username,
        permalink_url: t.permalink_url,
        duration_ms: t.duration,
        playback_count: t.playback_count ?? 0,
        created_at: t.created_at,
        artwork_url: t.artwork_url || u.avatar_url || null,
      }));

    return new Response(JSON.stringify({
      success: true,
      artist: { name: u.username, url: u.permalink_url, followers: u.followers_count ?? null, avatar: u.avatar_url || null },
      total_tracks: all.length,
      sets_excluded: all.filter((t) => (t.duration || 0) > maxMs || t.kind === "playlist").length,
      songs,
    }), { headers: JH });
  } catch (e) {
    console.error("sc-tracks:", e instanceof Error ? e.message : String(e));
    return err(502, "Couldn't reach SoundCloud — try again shortly.");
  }
});
