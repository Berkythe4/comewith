// pull-youtube-stats
//
// Fetches channel statistics from the YouTube Data API v3 and upserts them into
// metric_snapshots (youtube.subscribers, youtube.avg_views) for today, so the
// Strategy KPI cards + trends auto-update. No manual logging.
//
// Auth: a daily pg_cron job passes ?secret=<YT_CRON_SECRET>; the dashboard
// "Refresh YouTube" button passes a valid admin JWT. Either is accepted.
// Secrets: YOUTUBE_API_KEY, YOUTUBE_CHANNEL_ID, YT_CRON_SECRET (+ the SUPABASE_* set).

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

  // --- auth: cron secret OR a valid admin session ---
  const url = new URL(req.url);
  const cronSecret = Deno.env.get("YT_CRON_SECRET");
  const provided = url.searchParams.get("secret") || req.headers.get("x-cron-secret");
  let authed = !!cronSecret && provided === cronSecret;
  if (!authed) {
    const auth = req.headers.get("Authorization");
    if (auth) {
      const userClient = createClient(
        Deno.env.get("SUPABASE_URL")!,
        Deno.env.get("SUPABASE_ANON_KEY")!,
        { global: { headers: { Authorization: auth } } },
      );
      const { data: { user } } = await userClient.auth.getUser();
      if (user) {
        const a = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
        const { data: prof } = await a.from("profiles").select("role").eq("id", user.id).single();
        authed = !!prof && ["master_admin", "sub_admin"].includes(prof.role);
      }
    }
  }
  if (!authed) return err(401, "unauthorized");

  const apiKey = Deno.env.get("YOUTUBE_API_KEY");
  const channelId = Deno.env.get("YOUTUBE_CHANNEL_ID");
  if (!apiKey || !channelId) return err(500, "YOUTUBE_API_KEY / YOUTUBE_CHANNEL_ID not set");

  const yt = await fetch(`https://www.googleapis.com/youtube/v3/channels?part=statistics&id=${channelId}&key=${apiKey}`);
  const d = await yt.json();
  const stat = d.items?.[0]?.statistics;
  if (!stat) return err(502, "YouTube returned no stats: " + (d.error?.message || JSON.stringify(d)));

  const subs = Number(stat.subscriberCount);
  const views = Number(stat.viewCount);
  const videos = Number(stat.videoCount);
  const avg = videos > 0 ? Math.round(views / videos) : null;

  const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  const today = new Date().toISOString().slice(0, 10);
  async function upsert(metric_key: string, value: number | null) {
    if (value == null || !Number.isFinite(value)) return;
    const { data: ex } = await admin.from("metric_snapshots")
      .select("id").eq("metric_key", metric_key).eq("captured_on", today).is("series_id", null).maybeSingle();
    if (ex) await admin.from("metric_snapshots").update({ value, source: "youtube_api" }).eq("id", ex.id);
    else await admin.from("metric_snapshots").insert({ metric_key, value, captured_on: today, series_id: null, source: "youtube_api" });
  }
  await upsert("youtube.subscribers", subs);
  await upsert("youtube.avg_views", avg);

  return new Response(
    JSON.stringify({ ok: true, captured_on: today, subscribers: subs, total_views: views, videos, avg_views: avg }),
    { headers: HEADERS },
  );
});
