// resolve-media  (ADMIN — verify_jwt on)
//
// Root-cause guard for the recurring "I pasted a recap URL and it doesn't show"
// problem. Normalizes + VALIDATES each media URL so only publicly-embeddable
// links get saved:
//   • SoundCloud share/short links (on.soundcloud.com/…) → followed to the
//     canonical track/set URL; tracking params (utm_*, si) stripped.
//   • Public-embeddability verified via the provider's oembed endpoint — a
//     private track, a private/secret set, or a wrong permalink oembed-404s, so
//     we can tell the user WHY instead of the site failing silently.
//   • YouTube links validated + canonicalized the same way.
//
// Body: { urls: string[] }  →  { results: [{ input, ok, url, kind, title, reason }] }
// The dashboard calls this on save (and on demand) to block/clean bad links.

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36";

import { createClient } from "npm:@supabase/supabase-js@2";

const YT_RE = /(?:youtu\.be\/|v=|embed\/|shorts\/)([\w-]{11})/;

// Strip tracking noise but keep meaningful params (e.g. a secret token).
function cleanQuery(u: URL): string {
  const drop = ["utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content", "si", "ref"];
  for (const k of drop) u.searchParams.delete(k);
  u.hash = "";
  return u.toString().replace(/\?$/, "");
}

async function oembedOk(provider: "soundcloud" | "youtube", url: string): Promise<{ ok: boolean; title?: string }> {
  const base = provider === "soundcloud" ? "https://soundcloud.com/oembed" : "https://www.youtube.com/oembed";
  try {
    const r = await fetch(`${base}?format=json&url=${encodeURIComponent(url)}`, { headers: { "User-Agent": UA } });
    if (!r.ok) return { ok: false };
    const j = await r.json().catch(() => ({}));
    return { ok: true, title: j.title || undefined };
  } catch { return { ok: false }; }
}

async function resolveOne(raw: string): Promise<Record<string, unknown>> {
  const input = (raw || "").trim();
  if (!input) return { input, ok: false, kind: "empty", reason: "empty" };
  let host = "";
  try { host = new URL(input).host.toLowerCase(); } catch { return { input, ok: false, kind: "invalid", reason: "Not a valid URL." }; }

  const isSc = /(^|\.)soundcloud\.com$/.test(host) || host === "snd.sc" || host === "on.soundcloud.com";
  const isYt = /(^|\.)youtube\.com$/.test(host) || host === "youtu.be";

  if (isYt) {
    const m = input.match(YT_RE);
    if (!m) return { input, ok: false, kind: "youtube", reason: "Couldn't find a YouTube video id in that link." };
    const url = `https://www.youtube.com/watch?v=${m[1]}`;
    const o = await oembedOk("youtube", url);
    return o.ok
      ? { input, ok: true, kind: "youtube", url, title: o.title }
      : { input, ok: false, kind: "youtube", url, reason: "YouTube won't embed this — the video may be private, deleted, or embedding-disabled." };
  }

  if (isSc) {
    // Follow redirects so on.soundcloud.com / snd.sc short links become canonical.
    let finalUrl = input;
    try {
      const r = await fetch(input, { headers: { "User-Agent": UA }, redirect: "follow" });
      if (r.url) finalUrl = r.url;
    } catch { /* keep input; oembed will judge it */ }
    let clean = finalUrl;
    try { clean = cleanQuery(new URL(finalUrl)); } catch { /* leave as-is */ }

    // Try the cleaned URL; if it fails but there were params, retry with them
    // (covers secret-token share links that only embed WITH the token).
    let o = await oembedOk("soundcloud", clean);
    if (!o.ok && clean !== finalUrl) o = await oembedOk("soundcloud", finalUrl);
    if (o.ok) return { input, ok: true, kind: "soundcloud", url: clean, title: o.title };
    return { input, ok: false, kind: "soundcloud", url: clean, reason: "SoundCloud won't embed this — the track/set may be private or secret, or the link is wrong. Make it public, or paste the track's own page URL." };
  }

  return { input, ok: false, kind: "other", reason: "Not a YouTube or SoundCloud link — only those embed on the site." };
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

  const b = await req.json().catch(() => ({}));
  const urls: string[] = Array.isArray(b.urls) ? b.urls.slice(0, 30) : [];
  if (!urls.length) return err(400, "urls[] required");
  const results = await Promise.all(urls.map(resolveOne));
  return new Response(JSON.stringify({ results }), { headers: JH });
});
