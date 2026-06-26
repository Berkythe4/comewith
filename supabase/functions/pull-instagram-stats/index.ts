// pull-instagram-stats
//
// Fetches Instagram metrics via the Meta/Instagram Graph API and upserts them
// into metric_snapshots (instagram.followers / instagram.reach /
// instagram.profile_views) so the Strategy KPIs + trends auto-update — exactly
// like pull-youtube-stats.
//
// NOT YET DEPLOYED — needs two secrets first (see SETUP below). Once set:
//   SUPABASE_ACCESS_TOKEN=$SBP_PAT supabase functions deploy pull-instagram-stats \
//     --project-ref yaytdosxfhcqatmhctzk --no-verify-jwt
//
// SETUP (one-time, yours):
//   1. Switch the Instagram account to Business/Creator and link it to a
//      Facebook Page.
//   2. Create a Meta app at developers.facebook.com → add "Instagram Graph API".
//   3. Generate a LONG-LIVED access token with instagram_basic +
//      instagram_manage_insights + pages_read_engagement.
//   4. Find the IG Business account id (the numeric ig-user-id).
//   5. Set secrets: IG_USER_ID, IG_ACCESS_TOKEN (+ optional IG_CRON_SECRET).
//
// Auth mirrors pull-youtube-stats: ?secret=<IG_CRON_SECRET> for cron, or a valid
// master/sub admin JWT.

import { createClient } from "npm:@supabase/supabase-js@2";

const HEADERS = {
  "Content-Type": "application/json",
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-cron-secret",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: HEADERS });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: HEADERS });

  // --- auth: cron secret OR a master/sub admin session ---
  const url = new URL(req.url);
  const cronSecret = Deno.env.get("IG_CRON_SECRET");
  const provided = url.searchParams.get("secret") || req.headers.get("x-cron-secret");
  let authed = !!cronSecret && provided === cronSecret;
  if (!authed) {
    const auth = req.headers.get("Authorization");
    if (auth) {
      const userClient = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
      const { data: { user } } = await userClient.auth.getUser();
      if (user) {
        const a = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
        const { data: prof } = await a.from("profiles").select("role").eq("id", user.id).single();
        authed = !!prof && ["master_admin", "sub_admin"].includes(prof.role);
      }
    }
  }
  if (!authed) return err(401, "unauthorized");

  const igUser = Deno.env.get("IG_USER_ID");
  const token = Deno.env.get("IG_ACCESS_TOKEN");
  if (!igUser || !token) return err(500, "IG_USER_ID / IG_ACCESS_TOKEN not set — see SETUP in the function header");

  const G = "https://graph.facebook.com/v21.0";
  // 1) profile-level: followers + media count
  const prof = await (await fetch(`${G}/${igUser}?fields=followers_count,media_count,username&access_token=${token}`)).json();
  if (prof.error) return err(502, "Instagram API: " + (prof.error.message || JSON.stringify(prof.error)));
  const followers = Number(prof.followers_count);
  const media = Number(prof.media_count);

  // 2) day-level insights: reach + profile views
  let reach: number | null = null, profileViews: number | null = null;
  const ins = await (await fetch(`${G}/${igUser}/insights?metric=reach,profile_views&period=day&access_token=${token}`)).json();
  for (const m of (ins.data || [])) {
    const v = m.values?.[m.values.length - 1]?.value;
    if (m.name === "reach") reach = Number(v);
    if (m.name === "profile_views") profileViews = Number(v);
  }

  const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  const today = new Date().toISOString().slice(0, 10);
  async function upsert(metric_key: string, value: number | null) {
    if (value == null || !Number.isFinite(value)) return;
    const { data: ex } = await admin.from("metric_snapshots").select("id").eq("metric_key", metric_key).eq("captured_on", today).is("series_id", null).maybeSingle();
    if (ex) await admin.from("metric_snapshots").update({ value, source: "instagram_api" }).eq("id", ex.id);
    else await admin.from("metric_snapshots").insert({ metric_key, value, captured_on: today, series_id: null, source: "instagram_api" });
  }
  await upsert("instagram.followers", followers);
  await upsert("instagram.media", media);
  await upsert("instagram.reach", reach);
  await upsert("instagram.profile_views", profileViews);

  return new Response(JSON.stringify({ ok: true, captured_on: today, username: prof.username, followers, media, reach, profile_views: profileViews }), { headers: HEADERS });
});
