// get-station  (PUBLIC — deploy with --no-verify-jwt)
//
// Serves a Come With Radio station for the public radio.html page, by its
// secret public_token. Returns the tracklist + each track's show (date/venue/
// price) + bpm/key. Read-only, tokenized — nothing is discoverable without the
// token, and the page is unlinked/noindex, so a station stays private until
// Keith shares it / publishes it.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "GET, POST, OPTIONS" };
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  try {
    const url = new URL(req.url);
    let token = url.searchParams.get("t") || url.searchParams.get("token") || "";
    if (!token && req.method === "POST") { const b = await req.json().catch(() => ({})); token = (b.token || "").toString(); }
    token = token.trim();
    if (!token) return err(400, "token required");

    const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
    const { data: pl } = await admin.from("sc_playlists").select("id, name, note, published, sc_playlist_url").eq("public_token", token).maybeSingle();
    if (!pl) return err(404, "Station not found.");

    const { data: tracks } = await admin.from("sc_playlist_tracks")
      .select("title, artist_name, permalink_url, duration_ms, playback_count, artwork_url, show_date, show_venue, show_cost, show_url, bpm, song_key, camelot, sort")
      .eq("playlist_id", pl.id).order("sort");

    return new Response(JSON.stringify({
      station: { name: pl.name, note: pl.note, published: pl.published, soundcloud_url: pl.sc_playlist_url },
      tracks: tracks || [],
    }), { headers: JH });
  } catch (e) {
    console.error("get-station:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load the station.");
  }
});
