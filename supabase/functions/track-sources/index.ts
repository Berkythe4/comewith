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
  // Containment must be LENGTH-AWARE. A flat 0.92 meant "If U Need It" scored
  // 0.92 against "Sammy Virji: If U Need It (Callto Speed Garage Dub)" purely
  // because the words appear inside it — so a bootleg dub outscored the real
  // release. The more extra material the container carries, the less it means.
  if (x.includes(y) || y.includes(x)) {
    const ratio = Math.min(x.length, y.length) / Math.max(x.length, y.length);
    return 0.92 * ratio;
  }
  const A = new Set(x.split(" ")), B = new Set(y.split(" "));
  let inter = 0;
  for (const t of A) if (B.has(t)) inter++;
  return (2 * inter) / (A.size + B.size);
}
// A remix is a DIFFERENT SONG: if either side names a remix/edit, both must agree
// on the remixer, otherwise the original mix would happily match "(X Remix)".
// "(Radio Edit)" / "(Extended Mix)" describe the ORIGINAL release, so strip those
// standard qualifiers before looking for a remixer — otherwise every radio edit
// gets treated as somebody's remix and never matches its own release.
const STD_QUALIFIER = /\((?:original|extended|radio|club|vocal|instrumental|album|single)\s*(?:mix|edit|version|cut)?\)/gi;
function remixTag(s: string): string {
  const cleaned = String(s || "").replace(STD_QUALIFIER, " ");
  const m = cleaned.match(/[([]([^)\]]*\b(?:remix|rmx|edit|bootleg|flip|refix|vip|rework)\b[^)\]]*)[)\]]/i);
  return m ? norm(m[1]) : "";
}
// Bracket-only detection isn't enough: uploaders write "Sammy Virji. If U Need
// It. Pat Lok Flip." with no brackets at all, which sailed past the guard and
// matched the ORIGINAL — i.e. it would send you to buy the wrong track. So also
// look for the giveaway word anywhere in the string. "edit"/"vip" stay
// bracket-only because "Radio Edit" is a normal thing to call an original.
const REMIX_ANY = /\b(?:remix|rmx|bootleg|flip|refix|mashup)\b/i;
const isRemix = (s: string) => REMIX_ANY.test(String(s || "")) || remixTag(s) !== "";
function score(title: string, artist: string, candTitle: string, candArtist: string): number {
  // One is a remix and the other isn't → different songs, full stop.
  if (isRemix(title) !== isRemix(candTitle)) return 0;
  const rA = remixTag(title), rB = remixTag(candTitle);
  if (rA && rB && sim(rA, rB) < 0.6) return 0;
  return 0.68 * sim(title, candTitle) + 0.32 * sim(artist, candArtist);
}
const MIN_SCORE = 0.62;

// SoundCloud upload titles are messy in ways that wreck a store search: promo
// junk ("[Free Download]", "(Beatport link in bio)"), and the artist name
// repeated INSIDE the title ("Deeper Purpose - Cigarettes" by Deeper Purpose).
// Searching those verbatim is why a track that IS in the store comes back empty.
function searchTerms(title: string, artist: string) {
  let t = String(title || "")
    .replace(/\[[^\]]*\]/g, " ")
    .replace(/\([^)]*\b(?:free|download|out now|link in bio|premiere|forthcoming|buy now|support)\b[^)]*\)/gi, " ")
    .replace(/\b(?:free\s*d\/?l|free\s*download|out\s*now|premiere)\b/gi, " ");
  const cleanArtist = String(artist || "").replace(/^[\s"'“”‘’]+|[\s"'“”‘’]+$/g, "").trim();
  const a = norm(cleanArtist);
  if (a) {
    const parts = t.split(/\s+[-–—]\s+/);
    if (parts.length >= 2 && norm(parts[0]) === a) t = parts.slice(1).join(" - ");
  }
  t = t.replace(/\s{2,}/g, " ").trim();
  return { title: t || String(title || ""), artist: cleanArtist };
}

// ---- Beatport ---------------------------------------------------------------
// Beatport access tokens live TEN MINUTES and the refresh token is not reachable
// from the browser (not in localStorage, and their site refreshes via a cookie JS
// can't read). So the workable model is: paste a fresh access token when you want
// to run a check. We cache it until its own `exp` so repeat runs inside the same
// ten minutes don't need another paste, and we never persist anything longer-lived.
function jwtExp(token: string): number | null {
  try {
    const p = token.split(".")[1];
    if (!p) return null;
    const json = JSON.parse(atob(p.replace(/-/g, "+").replace(/_/g, "/") + "===".slice((p.length + 3) % 4)));
    return typeof json.exp === "number" ? json.exp : null;
  } catch (_) { return null; }
}
async function storePastedToken(admin: any, token: string) {
  const exp = jwtExp(token);
  await admin.from("beatport_oauth").upsert({
    id: "singleton",
    access_token: token,
    // Trust the token's own expiry; fall back to a conservative 9 minutes.
    expires_at: new Date((exp ? exp * 1000 : Date.now() + 9 * 60_000)).toISOString(),
    last_error: null,
    updated_at: new Date().toISOString(),
  });
}

// `needsSetup` distinguishes "Beatport was never connected" from "the token chain
// broke" — they read very differently to the person clicking the button, and
// telling someone to RE-paste a token they never pasted is just confusing.
async function bpToken(admin: any): Promise<{ token?: string; error?: string; needsSetup?: boolean }> {
  const { data: row } = await admin.from("beatport_oauth").select("*").eq("id", "singleton").maybeSingle();
  const clientId = Deno.env.get("BEATPORT_CLIENT_ID");
  if (!clientId) return { needsSetup: true, error: "Beatport isn't connected yet." };

  if (row?.access_token && row.expires_at && new Date(row.expires_at).getTime() > Date.now() + 60_000) {
    return { token: row.access_token };
  }
  const refresh = row?.refresh_token || Deno.env.get("BEATPORT_REFRESH_TOKEN");
  if (!refresh) return { needsSetup: true, error: "Beatport isn't connected yet." };

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
  const rank = (cands: any[]) => {
    let best: any = null, bestScore = 0;
    for (const c of cands) {
      const cTitle = [c.name, c.mix_name && !/^original( mix)?$/i.test(c.mix_name) ? `(${c.mix_name})` : ""].filter(Boolean).join(" ");
      const cArtist = (c.artists || []).map((a: any) => a.name).join(" ");
      const s = score(title, artist, cTitle, cArtist);
      if (s > bestScore) { bestScore = s; best = { ...c, _title: cTitle, _artist: cArtist }; }
    }
    return { best, bestScore };
  };

  let cands: any[] = [];
  try { cands = await attempt(`${artist} ${title}`.trim()); } catch (_) { /* fall through */ }
  let { best, bestScore } = rank(cands);
  // Retry on the TITLE ALONE when the combined query produced nothing good — not
  // just when it produced nothing at all. Searching "Deeper Purpose Cigarettes"
  // returns "Liquor & Cigarettes" and a track literally named "Deeper Purpose";
  // both are junk, so the old `if (!cands.length)` guard never fired and the real
  // release was never looked for. The artist still decides the winner.
  if (bestScore < MIN_SCORE) {
    try {
      const alt = rank(await attempt(title));
      if (alt.bestScore > bestScore) ({ best, bestScore } = alt);
    } catch (_) { /* keep what we have */ }
  }
  if (!best || bestScore < MIN_SCORE) return null;
  const k = bpKey(best.key);
  // price.value is already in DOLLARS (1.49), not cents — dividing by 100 turned
  // every fallback price into "$0.01". Verified against the live API.
  const price = best.price?.display || (best.price?.value != null ? `$${Number(best.price.value).toFixed(2)}` : null);
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
// Bandcamp has no official API. This is the endpoint their own search box uses.
// It answers HTTP 200 with {"error":true,"error_message":"bad function"} when the
// path is wrong, so status alone is NOT proof of success — an earlier version of
// this checked only r.ok and therefore reported every track as "not on Bandcamp"
// when in fact nothing was ever searched. Validate the PAYLOAD and throw
// otherwise, so the caller can say "couldn't reach Bandcamp" instead of lying.
async function bcSearch(title: string, artist: string) {
  const r = await fetch("https://bandcamp.com/api/bcsearch_public_api/1/autocomplete_elastic", {
    method: "POST",
    headers: { "Content-Type": "application/json", "User-Agent": UA },
    body: JSON.stringify({ search_text: `${artist} ${title}`.trim(), search_filter: "t", full_page: false, fan_id: null }),
  });
  if (!r.ok) throw new Error(`bandcamp ${r.status}`);
  const j = await r.json();
  if (j?.error || !j?.auto || !Array.isArray(j.auto.results)) throw new Error(`bandcamp: ${j?.error_message || "unexpected response"}`);
  const items = j.auto.results as any[];
  let best: any = null, bestScore = 0;
  for (const c of items) {
    if (c.type && c.type !== "t") continue;
    // Bandcamp's `name` is usually "ARTIST - TITLE" and `band_name` is whoever
    // uploaded it (often a label or the remixer, not the original artist), so
    // compare both readings and keep the better one.
    const nm = String(c.name || "");
    const parts = nm.split(/\s+[-–—]\s+/);
    const alt = parts.length >= 2
      ? score(title, artist, parts.slice(1).join(" - "), parts[0])
      : 0;
    const s = Math.max(score(title, artist, nm, c.band_name || ""), alt);
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
    // A token pasted from the browser this run — cached until its own exp.
    const pasted = (b.access_token || "").toString().trim().replace(/^Bearer\s+/i, "");
    if (pasted) {
      const exp = jwtExp(pasted);
      if (exp && exp * 1000 < Date.now()) {
        return new Response(JSON.stringify({
          results: [], applied: 0, beatport_ok: false, bandcamp_ok: true, beatport_expired: true,
          error: "That Beatport token had already expired when it arrived — they only last 10 minutes. Grab a fresh one and paste it again.",
        }), { headers: JH });
      }
      await storePastedToken(admin, pasted);
    }

    // Cheap "do we have a usable Beatport token right now?" check, so the
    // dashboard can ask for one BEFORE making you sit through a full scan.
    if (b.probe) {
      const probe = await bpToken(admin);
      return new Response(JSON.stringify({
        probe: true,
        beatport_ok: !!probe.token,
        needs_setup: !!probe.needsSetup,
      }), { headers: JH });
    }

    // Batched: the dashboard walks through the tracklist a few at a time so it can
    // show real progress. A single 26-track pass took 25-60s of silent spinner,
    // which is indistinguishable from "it's broken".
    const offset = Math.max(0, Number(b.offset) || 0);
    const { count: total } = await admin.from("sc_playlist_tracks")
      .select("id", { count: "exact", head: true }).eq("playlist_id", playlistId);

    const { data: tracks } = await admin.from("sc_playlist_tracks")
      .select("id, title, artist_name, bpm, song_key, camelot")
      .eq("playlist_id", playlistId).order("sort").range(offset, offset + limit - 1);
    if (!tracks?.length) {
      return new Response(JSON.stringify({ results: [], applied: 0, total: total || 0, offset, done: true }), { headers: JH });
    }

    const tk = await bpToken(admin);
    const results: any[] = [];
    let applied = 0, bandcampFails = 0;

    // One track at a time (politeness — a burst is what makes an unofficial
    // endpoint start 429ing) but the two stores are queried CONCURRENTLY, which
    // roughly halves the wall time without adding any burstiness per host.
    for (const t of tracks) {
      const title = t.title || "", artist = t.artist_name || "";
      const q = searchTerms(title, artist);
      const [beatport, bcOut] = await Promise.all([
        tk.token ? bpSearch(tk.token, q.title, q.artist).catch(() => null) : Promise.resolve(null),
        bcSearch(q.title, q.artist).then((r) => ({ ok: true, r })).catch(() => ({ ok: false, r: null })),
      ]);
      const bandcamp = bcOut.r;
      if (!bcOut.ok) bandcampFails++;

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
      await new Promise((r) => setTimeout(r, 60));
    }

    return new Response(JSON.stringify({
      results,
      applied,
      total: total || tracks.length,
      offset,
      done: offset + tracks.length >= (total || tracks.length),
      beatport_ok: !!tk.token,
      // Distinguish "checked, genuinely not on Bandcamp" from "Bandcamp didn't answer".
      bandcamp_ok: bandcampFails < tracks.length,
      beatport_needs_setup: !!tk.needsSetup,
      error: tk.error || null,
    }), { headers: JH });
  } catch (e) {
    console.error("track-sources:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not check the stores.");
  }
});
