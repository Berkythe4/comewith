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

  const YT = "https://www.googleapis.com/youtube/v3";
  // 1) channel stats + uploads playlist
  const ch = await (await fetch(`${YT}/channels?part=statistics,contentDetails&id=${channelId}&key=${apiKey}`)).json();
  const stat = ch.items?.[0]?.statistics;
  if (!stat) return err(502, "YouTube returned no stats: " + (ch.error?.message || JSON.stringify(ch)));
  const uploads = ch.items?.[0]?.contentDetails?.relatedPlaylists?.uploads;
  const subs = Number(stat.subscriberCount);
  const totalViews = Number(stat.viewCount);
  const videos = Number(stat.videoCount);
  const avg = videos > 0 ? Math.round(totalViews / videos) : null;

  // 2) collect video ids from the uploads playlist (paginate, capped)
  const videoIds: string[] = [];
  let pageToken = "";
  while (uploads && videoIds.length < 300) {
    const pl = await (await fetch(`${YT}/playlistItems?part=contentDetails&maxResults=50&playlistId=${uploads}&pageToken=${pageToken}&key=${apiKey}`)).json();
    for (const it of (pl.items || [])) videoIds.push(it.contentDetails.videoId);
    pageToken = pl.nextPageToken || "";
    if (!pageToken) break;
  }

  // 3) per-video statistics (batches of 50)
  let totalLikes = 0, totalComments = 0, lastUpload: string | null = null;
  const videoRows: any[] = [];
  for (let i = 0; i < videoIds.length; i += 50) {
    const v = await (await fetch(`${YT}/videos?part=statistics,snippet&id=${videoIds.slice(i, i + 50).join(",")}&key=${apiKey}`)).json();
    for (const vid of (v.items || [])) {
      const s = vid.statistics || {}, sn = vid.snippet || {};
      const vv = Number(s.viewCount || 0), vl = Number(s.likeCount || 0), vc = Number(s.commentCount || 0);
      totalLikes += vl; totalComments += vc;
      if (!lastUpload || sn.publishedAt > lastUpload) lastUpload = sn.publishedAt;
      videoRows.push({
        video_id: vid.id, title: sn.title || null, published_at: sn.publishedAt || null,
        views: vv, likes: vl, comments: vc,
        thumbnail_url: sn.thumbnails?.medium?.url || sn.thumbnails?.default?.url || null,
        fetched_at: new Date().toISOString(),
      });
    }
  }
  const engagement = totalViews > 0 ? Math.round((totalLikes + totalComments) / totalViews * 10000) / 100 : null;
  const daysSinceUpload = lastUpload ? Math.floor((Date.now() - new Date(lastUpload).getTime()) / 86400000) : null;

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
  await upsert("youtube.total_views", totalViews);
  await upsert("youtube.videos", videos);
  await upsert("youtube.total_likes", totalLikes);
  await upsert("youtube.engagement_rate", engagement);
  await upsert("youtube.days_since_upload", daysSinceUpload);
  if (videoRows.length) await admin.from("youtube_videos").upsert(videoRows, { onConflict: "video_id" });

  return new Response(
    JSON.stringify({ ok: true, captured_on: today, subscribers: subs, total_views: totalViews, videos, avg_views: avg, total_likes: totalLikes, total_comments: totalComments, engagement_rate: engagement, days_since_upload: daysSinceUpload, videos_indexed: videoRows.length }),
    { headers: HEADERS },
  );
});
