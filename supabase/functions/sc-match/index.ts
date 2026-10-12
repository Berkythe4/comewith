// sc-match  (ADMIN)
//
// Resolve a SoundCloud PROFILE from an artist NAME. RA gives most artists a
// SoundCloud link, but DICE / Ticketmaster (and ~339 RA artists) give none — so
// they can never be scanned for songs. This searches SoundCloud's own api-v2
// user search and picks a confident match, then (write:true) fills the empty
// ra_artists.soundcloud so the normal sc-enrich scan can read their tracks.
//
// Matching is CONSERVATIVE on purpose — a wrong profile poisons the radio with
// some stranger's music. We only WRITE a candidate whose handle or display name
// NORMALIZES to the query exactly (diacritics + punctuation stripped) and that
// has tracks; among those prefer verified, then has-tracks, then followers.
// Affix-only or real-name-only hits, and empty profiles, come back as
// needs_review with the candidate and are never written (we'd rather miss than
// mis-assign).
//
// Body: { names: string[] (<=30), write?: boolean (default true) }
// Returns: { results: [{ name, matched, soundcloud, username, followers,
//            verified, track_count, confidence, needs_review?, review_reason? }],
//            matched, needs_review, client_id_ok }

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36";

// Strip diacritics + everything non-alphanumeric, lowercase. "Röyksopp!" -> "royksopp".
const norm = (s: string) =>
  (s || "").normalize("NFKD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
// Common vanity suffixes an official account tacks on — allow "artistmusic",
// "artistofficial", "djartist" to still match "artist".
const SUFFIX = ["official", "music", "sounds", "records", "recordings", "audio", "real", "the", "dj"];
function stripAffix(n: string, q: string): string {
  let x = n;
  for (const s of SUFFIX) {
    if (x.length > q.length && x.startsWith(s) && x.slice(s.length) === q) return q;
    if (x.length > q.length && x.endsWith(s) && x.slice(0, -s.length) === q) return q;
  }
  return x;
}

async function extractClientId(): Promise<string | null> {
  const home = await (await fetch("https://soundcloud.com/", { headers: { "User-Agent": UA } })).text();
  const scripts = [...home.matchAll(/<script[^>]+src="([^"]+)"/g)].map((m) => m[1]).reverse();
  for (const s of scripts) {
    try {
      const js = await (await fetch(s, { headers: { "User-Agent": UA } })).text();
      const m = js.match(/client_id[:=]"([a-zA-Z0-9]{20,})"/);
      if (m) return m[1];
    } catch { /* next */ }
  }
  return null;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const SUPA = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const admin = createClient(SUPA, SRK);

  // admin-only (same gate as sc-enrich)
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
    const names: string[] = Array.isArray(b.names) ? [...new Set(b.names.map((n: unknown) => String(n || "").trim()).filter(Boolean))].slice(0, 30) : [];
    const write = b.write !== false;
    if (!names.length) return err(400, "names[] required");

    const { data: cidRow } = await admin.from("site_content").select("value").eq("key", "ops.sc_client_id").maybeSingle();
    let cid = (cidRow?.value || "").trim();
    if (!cid) { cid = (await extractClientId()) || ""; if (cid) await admin.from("site_content").update({ value: cid }).eq("key", "ops.sc_client_id"); }
    if (!cid) return err(502, "SoundCloud unavailable (no client_id).");

    const api = "https://api-v2.soundcloud.com";
    const search = (q: string) => fetch(`${api}/search/users?q=${encodeURIComponent(q)}&limit=8&client_id=${cid}`, { headers: { "User-Agent": UA } });

    const results: Record<string, unknown>[] = [];
    let refreshed = false;
    for (const name of names) {
      const q = norm(name);
      if (!q) { results.push({ name, matched: false }); continue; }
      try {
        let r = await search(name);
        if ((r.status === 401 || r.status === 403) && !refreshed) {
          refreshed = true; const fresh = await extractClientId();
          if (fresh) { cid = fresh; await admin.from("site_content").update({ value: cid }).eq("key", "ops.sc_client_id"); r = await search(name); }
        }
        if (!r.ok) { results.push({ name, matched: false }); continue; }
        const js = await r.json().catch(() => ({}));
        const cands: any[] = Array.isArray(js.collection) ? js.collection : [];

        // Two tiers (2026-10-12 spot-check, LEARNINGS §93). STRICT: the display name
        // or the handle normalizes to the query as-is. LOOSE: it only matches after
        // stripping an affix ("DJ MINK" for "mink", "artistmusic"), or only the
        // FULL NAME matches - full_name is a person's real name, so "Lil' John
        // Roberts" (full name John Roberts) and an account called "leae" (full name
        // Marc Rebillet) passed as exact. A loose hit is never written; it comes
        // back as needs_review with the candidate, for a person to confirm.
        const strictOf = (c: any) => [c.permalink, c.username].map((f: string) => norm(f || "")).includes(q);
        const looseOf = (c: any) => [c.permalink, c.username, c.full_name].map((f: string) => stripAffix(norm(f || ""), q)).includes(q);
        // Among equals prefer verified, then a profile with TRACKS, then followers:
        // followers-first picked an empty "officialglassjaw" over the band's
        // 52-track account, and an empty profile gives the crate nothing.
        const rank = (a: any, b: any) =>
          (Number(!!b.verified) - Number(!!a.verified)) ||
          (Number((b.track_count || 0) > 0) - Number((a.track_count || 0) > 0)) ||
          ((b.followers_count || 0) - (a.followers_count || 0)) ||
          ((b.track_count || 0) - (a.track_count || 0));
        const strict = cands.filter(strictOf).sort(rank);
        const loose = cands.filter((c) => !strictOf(c) && looseOf(c)).sort(rank);
        const best = strict[0] || loose[0];

        if (!best) { results.push({ name, matched: false }); continue; }
        const soundcloud = (best.permalink_url || "").replace("://www.", "://");
        // Not confident enough to write: a loose-only match, or a profile with no
        // tracks (Martin, 2026-07-30: "low follower count and no songs" = wrong
        // profile). Returned with the candidate so the Add-artist form can offer it.
        // And a small unverified profile is a GUESS on any common name: with the
        // loose hits gone, "John Roberts", "mink" and "Nymph" still each had some
        // stranger whose handle was exactly the name. Below the bar that already
        // defines "high" confidence (verified, or 500+ followers), or with a
        // second exact profile of comparable size, a person decides.
        const fl = best.followers_count || 0;
        const rival = strict[1] && (strict[1].followers_count || 0) >= fl * 0.5 && (strict[1].track_count || 0) > 0;
        const reviewWhy = !strict[0] ? "name matches only loosely (affix or real name)"
          : (best.track_count || 0) === 0 ? "profile has no tracks"
          : !best.verified && fl < 500 ? `small unverified profile (${fl} followers)`
          : !best.verified && rival ? "two exact-name profiles of similar size"
          : null;
        if (reviewWhy) {
          results.push({
            name, matched: false, needs_review: true, review_reason: reviewWhy, soundcloud,
            username: best.username || best.permalink, followers: best.followers_count || 0,
            verified: !!best.verified, track_count: best.track_count || 0,
            city: (best.city || "").toString().trim() || null, confidence: "low",
          });
          continue;
        }
        const city = (best.city || "").toString().trim() || null;   // SoundCloud profile city
        const country = (best.country_code || best.country || "").toString().trim() || null;
        // confidence: verified or a real following reads as a safe, strong match.
        const confidence = best.verified ? "high" : (best.followers_count >= 500 ? "high" : (best.track_count > 0 ? "medium" : "low"));
        results.push({
          name, matched: true, soundcloud,
          username: best.username || best.permalink, followers: best.followers_count || 0,
          verified: !!best.verified, track_count: best.track_count || 0, city, country, confidence,
        });

        if (write && soundcloud) {
          // Fill only EMPTY soundcloud, and only for that exact name — never
          // overwrite a link RA already gave us.
          await admin.from("ra_artists").update({ soundcloud }).ilike("name", name).is("soundcloud", null);
          // City is the NYC-local signal — fill it wherever it's still blank.
          if (city) await admin.from("ra_artists").update({ city }).ilike("name", name).is("city", null);
        }
      } catch { results.push({ name, matched: false }); }
      await new Promise((res) => setTimeout(res, 120)); // be gentle on SC search
    }

    const matched = results.filter((r) => r.matched).length;
    const needsReview = results.filter((r) => r.needs_review).length;
    return new Response(JSON.stringify({ results, matched, needs_review: needsReview, total: names.length, client_id_ok: true }), { headers: JH });
  } catch (e) {
    console.error("sc-match:", e instanceof Error ? e.message : String(e));
    return err(500, "match failed");
  }
});
