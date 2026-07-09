// sc-bpm  (ADMIN)
//
// Best-effort BPM + key + Camelot for a station's tracks, via FREE keyless
// sources: MusicBrainz (artist+title → recording MBID) → AcousticBrainz
// (MBID → key/scale + BPM). Coverage is partial — well-known released tracks
// resolve; underground SoundCloud edits often won't. Fills what it can.
//
// Body: { playlist_id }. Respects MusicBrainz's 1 req/sec limit.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const UA = "ComeWithRadio/1.0 ( berky@comewith.org )";
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

const CAMELOT_MAJ: Record<number, string> = { 0: "8B", 1: "3B", 2: "10B", 3: "5B", 4: "12B", 5: "7B", 6: "2B", 7: "9B", 8: "4B", 9: "11B", 10: "6B", 11: "1B" };
const CAMELOT_MIN: Record<number, string> = { 0: "5A", 1: "12A", 2: "7A", 3: "2A", 4: "9A", 5: "4A", 6: "11A", 7: "6A", 8: "1A", 9: "8A", 10: "3A", 11: "10A" };
const PITCH: Record<string, number> = { "C": 0, "C#": 1, "DB": 1, "D": 2, "D#": 3, "EB": 3, "E": 4, "F": 5, "F#": 6, "GB": 6, "G": 7, "G#": 8, "AB": 8, "A": 9, "A#": 10, "BB": 10, "B": 11 };
function toCamelot(key: string, scale: string): { camelot: string | null; label: string | null } {
  const pc = PITCH[(key || "").toUpperCase()];
  if (pc == null) return { camelot: null, label: null };
  const minor = /min/i.test(scale || "");
  return { camelot: (minor ? CAMELOT_MIN : CAMELOT_MAJ)[pc], label: key + (minor ? "m" : "") };
}
// Strip remix/edit/free-DL noise so the title matches a released recording.
function cleanTitle(t: string): string {
  return (t || "")
    .replace(/\[[^\]]*\]|\([^)]*\)/g, " ")
    .replace(/\b(feat|ft|prod|remix|edit|bootleg|free|dl|download|premiere|master|vip|mix)\b.*$/i, " ")
    .replace(/[^\w\s'&-]/g, " ").replace(/\s+/g, " ").trim();
}

async function jget(url: string) {
  try { const r = await fetch(url, { headers: { "User-Agent": UA, "Accept": "application/json" } }); return r.ok ? await r.json() : null; } catch { return null; }
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
    const playlist_id = (b.playlist_id || "").toString();
    if (!playlist_id) return err(400, "playlist_id required");
    const { data: tracks } = await admin.from("sc_playlist_tracks").select("id, title, artist_name, bpm").eq("playlist_id", playlist_id).order("sort");
    const list = (tracks || []).slice(0, 60);

    let matched = 0;
    for (const t of list) {
      if (t.bpm) { continue; } // already have it
      const title = cleanTitle(t.title);
      const artist = (t.artist_name || "").trim();
      if (!title || !artist) continue;
      const q = encodeURIComponent(`artist:"${artist}" AND recording:"${title}"`);
      const mb = await jget(`https://musicbrainz.org/ws/2/recording?query=${q}&fmt=json&limit=1`);
      await sleep(1100); // MusicBrainz: 1 req/sec
      const mbid = mb?.recordings?.[0]?.id;
      if (!mbid) continue;
      const ll = await jget(`https://acousticbrainz.org/api/v1/${mbid}/low-level`);
      const tonal = ll?.tonal || {};
      const bpm = ll?.rhythm?.bpm;
      const { camelot, label } = toCamelot(tonal.key_key, tonal.key_scale);
      if (!bpm && !camelot) continue;
      await admin.from("sc_playlist_tracks").update({
        bpm: bpm ? Math.round(bpm) : null, song_key: label, camelot,
      }).eq("id", t.id);
      matched++;
    }
    return new Response(JSON.stringify({ success: true, checked: list.length, matched }), { headers: JH });
  } catch (e) {
    console.error("sc-bpm:", e instanceof Error ? e.message : String(e));
    return err(502, "BPM lookup failed — try again shortly.");
  }
});
