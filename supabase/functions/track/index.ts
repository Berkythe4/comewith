// track
//
// The write end of the first-party analytics beacon (migration 108). Public,
// unauthenticated, called by assets/track.js on the public pages.
//
// Why a function and not a table grant: site_events is anon-revoked like
// everything else (see CLAUDE.md — a blanket anon grant is what caused the
// 016/017 regression). The page can POST here; it can never touch the table.
//
// Deliberately stores NO ip, NO user agent, NO cookie. The service role writes
// exactly the fields below and nothing else, so there's no quiet path from
// "count a pageview" to "profile a person".
//
// Requests arrive via navigator.sendBeacon, which cannot set headers — so the
// body is text/plain and the function must not require an apikey. Deployed
// with --no-verify-jwt.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

// Crawlers, unfurlers and uptime pingers would otherwise read as an audience.
const BOT = /bot|crawl|spider|slurp|preview|fetch|monitor|lighthouse|headless|curl|wget|python-requests|facebookexternalhit|whatsapp|telegram|discord|slack|embedly|vercel|netlify/i;

const cap = (v: unknown, n: number): string | null => {
  if (v == null) return null;
  const s = String(v).trim();
  return s ? s.slice(0, n) : null;
};

const admin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  { auth: { persistSession: false } },
);

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  // Always answer 204. A beacon that argues with the page is worse than a
  // beacon that silently drops an event — analytics must never break a page.
  const ok = () => new Response(null, { status: 204, headers: CORS });
  try {
    if (req.method !== "POST") return ok();
    if (BOT.test(req.headers.get("user-agent") || "")) return ok();

    const body = await req.json().catch(() => null);
    if (!body) return ok();

    const kind = body.k === "click" ? "click" : body.k === "view" ? "view" : null;
    const path = cap(body.p, 300);
    if (!kind || !path) return ok();

    // Referrer arrives as a host from the client; re-strip here so a full URL
    // can never be stored even if the client is out of date or spoofed.
    let referrer = cap(body.r, 200);
    if (referrer) {
      try { referrer = new URL(referrer.includes("//") ? referrer : "https://" + referrer).hostname; }
      catch { /* already a bare host */ }
      referrer = cap(referrer, 200);
    }

    const { error } = await admin.from("site_events").insert({
      kind,
      path,
      referrer,
      utm_source: cap(body.us, 100),
      utm_medium: cap(body.um, 100),
      utm_campaign: cap(body.uc, 100),
      link_url: kind === "click" ? cap(body.l, 500) : null,
      link_label: kind === "click" ? cap(body.t, 120) : null,
      outbound: kind === "click" ? !!body.o : null,
      session_id: cap(body.sid, 40),
    });
    // A silently-dropped beacon is indistinguishable from no traffic, which is
    // the worst possible failure for an analytics table. Log it, and echo it
    // when explicitly asked (?debug=1) so this is diagnosable from a curl.
    if (error) {
      console.error("track insert:", error.message);
      if (new URL(req.url).searchParams.get("debug") === "1") {
        return new Response(JSON.stringify({ error: error.message }), { status: 200, headers: { ...CORS, "Content-Type": "application/json" } });
      }
    }
    return ok();
  } catch (e) {
    console.error("track:", (e as Error).message);
    return ok();
  }
});
