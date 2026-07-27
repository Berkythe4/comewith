// sc-match  (ADMIN)
//
// Resolve a SoundCloud PROFILE from an artist NAME. RA gives most artists a
// SoundCloud link, but DICE / Ticketmaster (and ~339 RA artists) give none — so
// they can never be scanned for songs. This searches SoundCloud's own api-v2
// user search and picks a confident match, then (write:true) fills the empty
// ra_artists.soundcloud so the normal sc-enrich scan can read their tracks.
//
// Matching is CONSERVATIVE on purpose — a wrong profile poisons the radio with
// some stranger's music. We only accept a candidate whose handle / display name
// / full name NORMALIZES to the query (diacritics + punctuation stripped), and
// among those prefer verified, then most followers, then has-tracks. No exact
// normalized match → matched:false (we'd rather miss than mis-assign).
//
// Body: { names: string[] (<=30), write?: boolean (default true) }
// Returns: { results: [{ name, matched, soundcloud, username, followers,
//            verified, track_count, confidence }], client_id_ok }

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

        // Keep only candidates whose handle/display/full name normalizes to the
        // query (after allowing an official-style affix).
        const exact = cands.filter((c) => {
          const forms = [c.permalink, c.username, c.full_name].map((f: string) => stripAffix(norm(f || ""), q));
          return forms.includes(q);
        });
        exact.sort((a, b) =>
          (Number(!!b.verified) - Number(!!a.verified)) ||
          ((b.followers_count || 0) - (a.followers_count || 0)) ||
          ((b.track_count || 0) - (a.track_count || 0)));
        const best = exact[0];

        if (!best) { results.push({ name, matched: false }); continue; }
        const soundcloud = (best.permalink_url || "").replace("://www.", "://");
        // confidence: verified or a real following reads as a safe, strong match.
        const confidence = best.verified ? "high" : (best.followers_count >= 500 ? "high" : (best.track_count > 0 ? "medium" : "low"));
        results.push({
          name, matched: true, soundcloud,
          username: best.username || best.permalink, followers: best.followers_count || 0,
          verified: !!best.verified, track_count: best.track_count || 0, confidence,
        });

        if (write && soundcloud) {
          // Fill only EMPTY soundcloud, and only for that exact name — never
          // overwrite a link RA already gave us.
          await admin.from("ra_artists").update({ soundcloud }).ilike("name", name).is("soundcloud", null);
        }
      } catch { results.push({ name, matched: false }); }
      await new Promise((res) => setTimeout(res, 120)); // be gentle on SC search
    }

    const matched = results.filter((r) => r.matched).length;
    return new Response(JSON.stringify({ results, matched, total: names.length, client_id_ok: true }), { headers: JH });
  } catch (e) {
    console.error("sc-match:", e instanceof Error ? e.message : String(e));
    return err(500, "match failed");
  }
});
