// track-sources  (ADMIN — verify_jwt on, plus an explicit master_admin/sub_admin check)
//
// "Where can I buy this?" for a station's tracklist. Since the set is now bought
// and arranged in Rekordbox rather than recorded off SoundCloud, the useful
// question before a buy run is: which of these songs actually exist on Beatport
// or Bandcamp, and what are their real BPM / key?
//
//   • Beatport — official v4 catalog API (OAuth bearer). Gives a confident match
//     plus BPM, key, label and price. Tokens live in public.beatport_oauth
//     (admin-RLS'd, anon revoked) because Beatport ROTATES the refresh token on
//     every use, so it has to be written back at runtime — an env secret can't be.
//     BEATPORT_CLIENT_ID + the initial BEATPORT_REFRESH_TOKEN come from secrets
//     and seed the row on first use.
//   • Bandcamp — has NO official API. Best-effort only, via the unofficial
//     autocomplete endpoint their own search box uses. Availability + link only:
//     Bandcamp doesn't carry BPM/key at all. Any failure here is non-fatal and
//     reported as "unknown" rather than "not available" — never let a scraped
//     endpoint's outage look like a definitive answer.
//
// Body: { playlist_id, apply_bpm_key?: boolean, limit?: number }
//   → { results: [{ track_id, title, artist, beatport: {...}|null, bandcamp: {...}|null,
//                   confidence }], applied, beatport_ok, bandcamp_ok, error? }
// Results are cached onto sc_playlist_tracks (beatport_url/price, bandcamp_url,
// sources_checked_at) so reopening the panel doesn't re-hit either API.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36";

const SUPA = Deno.env.get("SUPABASE_URL")!;
const BP_TOKEN_URL = "https://api.beatport.com/v4/auth/o/token/";
const BP_SEARCH = "https://api.beatport.com/v4/catalog/search/";

// ---- matching (mirrors the dashboard's Rekordbox matcher) --------------------
// Combining-mark range via the constructor — a literal U+0300–U+036F class is
// invisible in source and gets mangled by editors/encodings.
const DIACRITICS = new RegExp("[\\u0300-\\u036f]", "g");
const norm = (s: string) =>
  String(s || "").toLowerCase().normalize("NFD").replace(DIACRITICS, "")
    .replace(/&/g, " and ")
    .replace(/\((?:original|extended|radio|club|dub|vocal|instrumental)\s*(?:mix|edit|version|cut)?\)/g, " ")
    .replace(/\[[^\]]*\]/g, " ")
    .replace(/\b(?:feat|ft|featuring)\b.*$/, " ")
    .replace(/[^a-z0-9]+/g, " ").trim();

function sim(a: string, b: string): number {
  const x = norm(a), y = norm(b);
  if (!x || !y) return 0;
  if (x === y) return 1;
  if (x.includes(y) || y.includes(x)) return 0.92;
  const A = new Set(x.split(" ")), B = new Set(y.split(" "));
  let inter = 0;
  for (const t of A) if (B.has(t)) inter++;
  return (2 * inter) / (A.size + B.size);
}
// A remix is a DIFFERENT SONG: if either side names a remix/edit, both must agree
// on the remixer, otherwise the original mix would happily match "(X Remix)".
function remixTag(s: string): string {
  const m = String(s || "").match(/[([]([^)\]]*\b(?:remix|rmx|edit|bootleg|flip|refix)\b[^)\]]*)[)\]]/i);
  return m ? norm(m[1]) : "";
}
function score(title: string, artist: string, candTitle: string, candArtist: string): number {
  const rA = remixTag(title), rB = remixTag(candTitle);
  if ((rA || rB) && sim(rA, rB) < 0.6) return 0;
  return 0.68 * sim(title, candTitle) + 0.32 * sim(artist, candArtist);
}
const MIN_SCORE = 0.62;

// ---- Beatport ---------------------------------------------------------------
async function bpToken(admin: any): Promise<{ token?: string; error?: string }> {
  const { data: row } = await admin.from("beatport_oauth").select("*").eq("id", "singleton").maybeSingle();
  const clientId = Deno.env.get("BEATPORT_CLIENT_ID");
  if (!clientId) return { error: "BEATPORT_CLIENT_ID secret isn't set on the project." };

  if (row?.access_token && row.expires_at && new Date(row.expires_at).getTime() > Date.now() + 60_000) {
    return { token: row.access_token };
  }
  const refresh = row?.refresh_token || Deno.env.get("BEATPORT_REFRESH_TOKEN");
  if (!refresh) return { error: "No Beatport refresh token — paste a fresh one (see the Beatport button's help text)." };

  const body = new URLSearchParams({ grant_type: "refresh_token", refresh_token: refresh, client_id: clientId });
  const r = await fetch(BP_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", "User-Agent": UA },
    body,
  });
  if (!r.ok) {
    const detail = (await r.text()).slice(0, 200);
    // Park the reason so the dashboard can tell you WHY instead of "0 found".
    await admin.from("beatport_oauth").upsert({
      id: "singleton", last_error: `token refresh ${r.status}: ${detail}`, updated_at: new Date().toISOString(),
    });
    return { error: `Beatport token refresh failed (${r.status}). Re-paste a fresh token.` };
  }
  const j = await r.json();
  await admin.from("beatport_oauth").upsert({
    id: "singleton",
    access_token: j.access_token,
    // Rotated — persist it or the NEXT refresh fails.
    refresh_token: j.refresh_token || refresh,
    expires_at: new Date(Date.now() + (j.expires_in || 3600) * 1000).toISOString(),
    last_error: null,
    updated_at: new Date().toISOString(),
  });
  return { token: j.access_token };
}

function bpKey(k: any): { song_key: string | null; camelot: string | null } {
  if (!k || typeof k !== "object") return { song_key: null, camelot: null };
  const camelot = k.camelot_number != null && k.camelot_letter
    ? `${k.camelot_number}${String(k.camelot_letter).toUpperCase()}` : null;
  let song_key: string | null = k.name || null;
  if (!song_key && k.letter) {
    song_key = `${k.letter}${k.is_sharp ? "#" : k.is_flat ? "b" : ""}` +
      (/min/i.test(k.chord_type?.name || "") ? "m" : "");
  }
  return { song_key, camelot };
}

async function bpSearch(token: string, title: string, artist: string) {
  const attempt = async (q: string) => {
    const u = `${BP_SEARCH}?q=${encodeURIComponent(q)}&type=tracks&per_page=10`;
    const r = await fetch(u, { headers: { Authorization: `Bearer ${token}`, "User-Agent": UA } });
    if (!r.ok) throw new Error(`beatport ${r.status}`);
    const j = await r.json();
    return (j.tracks || j.results || []) as any[];
  };
  let cands: any[] = [];
  try { cands = await attempt(`${artist} ${title}`.trim()); } catch (_) { /* fall through */ }
  if (!cands.length) { try { cands = await attempt(title); } catch (_) { /* give up */ } }

  let best: any = null, bestScore = 0;
  for (const c of cands) {
    const cTitle = [c.name, c.mix_name && !/^original( mix)?$/i.test(c.mix_name) ? `(${c.mix_name})` : ""].filter(Boolean).join(" ");
    const cArtist = (c.artists || []).map((a: any) => a.name).join(" ");
    const s = score(title, artist, cTitle, cArtist);
    if (s > bestScore) { bestScore = s; best = { ...c, _title: cTitle, _artist: cArtist }; }
  }
  if (!best || bestScore < MIN_SCORE) return null;
  const k = bpKey(best.key);
  const price = best.price?.display || (best.price?.value != null ? `$${(best.price.value / 100).toFixed(2)}` : null);
  return {
    url: best.slug && best.id ? `https://www.beatport.com/track/${best.slug}/${best.id}` : null,
    title: best._title, artist: best._artist,
    bpm: typeof best.bpm === "number" ? best.bpm : null,
    song_key: k.song_key, camelot: k.camelot,
    label: best.release?.label?.name || null,
    price, score: Number(bestScore.toFixed(2)),
  };
}

// ---- Bandcamp (unofficial, best-effort) --------------------------------------
async function bcSearch(title: string, artist: string) {
  const r = await fetch("https://bandcamp.com/api/fuzzysearch/1/autocomplete_elastic", {
    method: "POST",
    headers: { "Content-Type": "application/json", "User-Agent": UA },
    body: JSON.stringify({ search_text: `${artist} ${title}`.trim(), search_filter: "t", fuzziness: 0 }),
  });
  if (!r.ok) throw new Error(`bandcamp ${r.status}`);
  const j = await r.json();
  const items = (j.auto?.results || []) as any[];
  let best: any = null, bestScore = 0;
  for (const c of items) {
    if (c.type && c.type !== "t") continue;
    const s = score(title, artist, c.name || "", c.band_name || "");
    if (s > bestScore) { bestScore = s; best = c; }
  }
  if (!best || bestScore < MIN_SCORE) return null;
  return { url: best.item_url_path || best.url || null, title: best.name, artist: best.band_name, score: Number(bestScore.toFixed(2)) };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const admin = createClient(SUPA, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

  // Admin gate — verify_jwt only proves SOME user; listeners are customers.
  const auth = req.headers.get("Authorization") || "";
  const uc = createClient(SUPA, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
  const { data: { user } } = await uc.auth.getUser();
  if (!user) return err(401, "Sign in first.");
  const { data: prof } = await admin.from("profiles").select("role, deleted_at").eq("id", user.id).single();
  if (!prof || prof.deleted_at || !["master_admin", "sub_admin"].includes(prof.role)) return err(403, "Admins only.");

  try {
    const b = await req.json().catch(() => ({}));
    const playlistId = (b.playlist_id || "").toString();
    if (!playlistId) return err(400, "playlist_id required");
    const applyBpmKey = !!b.apply_bpm_key;
    const limit = Math.min(Number(b.limit) || 60, 60);

    const { data: tracks } = await admin.from("sc_playlist_tracks")
      .select("id, title, artist_name, bpm, song_key, camelot")
      .eq("playlist_id", playlistId).order("sort").limit(limit);
    if (!tracks?.length) return new Response(JSON.stringify({ results: [], applied: 0 }), { headers: JH });

    const tk = await bpToken(admin);
    const results: any[] = [];
    let applied = 0, bandcampFails = 0;

    // Sequential with a small pause: two third-party APIs, ~20-40 tracks. Politeness
    // beats speed here — a burst is what gets an unofficial endpoint to start 429ing.
    for (const t of tracks) {
      const title = t.title || "", artist = t.artist_name || "";
      let beatport = null, bandcamp = null;
      if (tk.token) {
        try { beatport = await bpSearch(tk.token, title, artist); } catch (_) { /* leave null */ }
      }
      try { bandcamp = await bcSearch(title, artist); } catch (_) { bandcampFails++; }

      const patch: Record<string, unknown> = {
        beatport_url: beatport?.url || null,
        beatport_price: beatport?.price || null,
        bandcamp_url: bandcamp?.url || null,
        sources_checked_at: new Date().toISOString(),
      };
      // Only ever FILL IN missing BPM/key — never overwrite what Rekordbox gave
      // you, since your own analysis of the file you own beats a store's metadata.
      if (applyBpmKey && beatport) {
        if (!t.bpm && beatport.bpm) patch.bpm = beatport.bpm;
        if (!t.song_key && beatport.song_key) patch.song_key = beatport.song_key;
        if (!t.camelot && beatport.camelot) patch.camelot = beatport.camelot;
        if (patch.bpm || patch.song_key || patch.camelot) applied++;
      }
      await admin.from("sc_playlist_tracks").update(patch).eq("id", t.id);

      results.push({ track_id: t.id, title, artist, beatport, bandcamp });
      await new Promise((r) => setTimeout(r, 120));
    }

    return new Response(JSON.stringify({
      results,
      applied,
      beatport_ok: !!tk.token,
      // Distinguish "checked, genuinely not on Bandcamp" from "Bandcamp didn't answer".
      bandcamp_ok: bandcampFails < tracks.length,
      error: tk.error || null,
    }), { headers: JH });
  } catch (e) {
    console.error("track-sources:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not check the stores.");
  }
});
