// sc-social  (ADMIN)
//
// Follow / repost on the connected SoundCloud account, one target at a time
// (manual, selective — the dashboard calls this per artist / per track).
//   follow   → PUT    /me/followings/{user_id}
//   unfollow → DELETE /me/followings/{user_id}
//   repost   → POST   /reposts/tracks/{track_id}
//   unrepost → DELETE /reposts/tracks/{track_id}
// Uses the stored OAuth token (refreshes if near expiry). Logs to sc_social_log.
//
// Body: { action, target_id, target_label? }

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

const EP: Record<string, { method: string; path: (id: string) => string; type: string }> = {
  follow: { method: "PUT", path: (id) => `/me/followings/${id}`, type: "user" },
  unfollow: { method: "DELETE", path: (id) => `/me/followings/${id}`, type: "user" },
  repost: { method: "POST", path: (id) => `/reposts/tracks/${id}`, type: "track" },
  unrepost: { method: "DELETE", path: (id) => `/reposts/tracks/${id}`, type: "track" },
};

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
    const action = (b.action || "").toString();
    const target_id = (b.target_id || "").toString().trim();
    const target_label = (b.target_label || "").toString().slice(0, 200) || null;
    const ep = EP[action];
    if (!ep || !target_id) return err(400, "action + target_id required");

    const { data: row } = await admin.from("sc_oauth").select("*").eq("id", "singleton").maybeSingle();
    if (!row?.access_token) return err(400, "Connect SoundCloud first.");

    const clientId = Deno.env.get("SC_CLIENT_ID"), clientSecret = Deno.env.get("SC_CLIENT_SECRET");
    let token = row.access_token;
    if (row.expires_at && new Date(row.expires_at).getTime() < Date.now() + 60000 && row.refresh_token && clientId && clientSecret) {
      const rr = await fetch("https://secure.soundcloud.com/oauth/token", {
        method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded", "accept": "application/json; charset=utf-8" },
        body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, client_secret: clientSecret, refresh_token: row.refresh_token }),
      });
      const rt = await rr.json();
      if (rr.ok && rt.access_token) { token = rt.access_token; await admin.from("sc_oauth").update({ access_token: rt.access_token, refresh_token: rt.refresh_token || row.refresh_token, expires_at: new Date(Date.now() + (rt.expires_in || 3600) * 1000).toISOString() }).eq("id", "singleton"); }
      else return err(401, "SoundCloud connection expired — reconnect.");
    }

    const scRes = await fetch(`https://api.soundcloud.com${ep.path(target_id)}`, {
      method: ep.method, headers: { "Authorization": "OAuth " + token, "accept": "application/json; charset=utf-8" },
    });
    const ok = scRes.ok || scRes.status === 201 || scRes.status === 200 || scRes.status === 404 && action.startsWith("un");
    if (!ok) {
      const txt = (await scRes.text().catch(() => "")).slice(0, 160);
      console.error("sc-social", action, scRes.status, txt);
      await admin.from("sc_social_log").insert({ action, target_type: ep.type, target_id, target_label, ok: false, detail: `HTTP ${scRes.status}` });
      if (scRes.status === 401) return err(401, "SoundCloud connection expired — reconnect.");
      return err(502, `SoundCloud rejected the ${action} (HTTP ${scRes.status}).`);
    }
    await admin.from("sc_social_log").insert({ action, target_type: ep.type, target_id, target_label, ok: true });
    return new Response(JSON.stringify({ success: true, action, target_id }), { headers: JH });
  } catch (e) {
    console.error("sc-social:", e instanceof Error ? e.message : String(e));
    return err(502, "Action failed — try again shortly.");
  }
});
