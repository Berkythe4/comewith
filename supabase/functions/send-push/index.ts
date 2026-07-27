// send-push
//
// Sends a Web Push to a user's subscribed devices. Called by the client's
// notify() right after a notification row is inserted, so a teammate's phone
// buzzes even when the dashboard isn't open — but ONLY if they opted in (there's
// simply no subscription row otherwise, so this is a no-op).
//
// Auth: an admin JWT (the notifier) OR service role. Secrets: VAPID_PUBLIC_KEY,
// VAPID_PRIVATE_KEY, VAPID_SUBJECT.
// Body: { user_id, title, body?, url? }

import { createClient } from "npm:@supabase/supabase-js@2";
import webpush from "npm:web-push@3.6.7";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

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

  const pub = Deno.env.get("VAPID_PUBLIC_KEY"), priv = Deno.env.get("VAPID_PRIVATE_KEY");
  if (!pub || !priv) return err(500, "VAPID keys not configured");
  webpush.setVapidDetails(Deno.env.get("VAPID_SUBJECT") || "mailto:berky@comewith.org", pub, priv);

  try {
    const b = await req.json().catch(() => ({}));
    const userId = (b.user_id || "").toString();
    if (!userId) return err(400, "user_id required");
    const payload = JSON.stringify({ title: (b.title || "Come With").toString(), body: (b.body || "").toString(), url: (b.url || "/dashboard.html").toString() });

    const { data: subs } = await admin.from("push_subscriptions").select("id, subscription").eq("user_id", userId);
    if (!subs?.length) return new Response(JSON.stringify({ sent: 0, note: "no subscriptions" }), { headers: JH });

    let sent = 0; const dead: string[] = [];
    for (const s of subs) {
      try {
        await webpush.sendNotification(s.subscription as any, payload);
        sent++;
      } catch (e: any) {
        // 404/410 = the browser unsubscribed / endpoint is gone → prune it.
        if (e?.statusCode === 404 || e?.statusCode === 410) dead.push(s.id);
      }
    }
    if (dead.length) await admin.from("push_subscriptions").delete().in("id", dead);
    return new Response(JSON.stringify({ sent, pruned: dead.length }), { headers: JH });
  } catch (e) {
    return err(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
