// scan-gear-market
//
// Scans the resale market for the DJ rig stolen from a vehicle (NYPD grand
// larceny case), scores every candidate listing, stores it deduped, and alerts
// on the ones worth a human look. Run by pg_cron three times a day via
// public.gear_watch_kick(), or by hand from the dashboard's Gear Watch tab.
//
// Sources:
//   reverb     — official API, needs REVERB_TOKEN (personal access token)
//   ebay       — Browse API, needs EBAY_CLIENT_ID + EBAY_CLIENT_SECRET
//   craigslist — no public API, and the search RSS is dead (403 on every path,
//                verified 2026-08-18). Uses the internal JSON endpoint the site's
//                own search box calls, sapi.craigslist.org — the same approach
//                track-sources takes with Bandcamp. Area-scoped to NYC.
//
//   facebook   — Meta has NO public Marketplace search API (verified 2026-08-19:
//                every marketplace URL answers 400 unauthenticated; the Commerce
//                Platform API is partner-only, for your own catalog). Goes
//                through Apify, which is billed per result, so it runs only when
//                APIFY_TOKEN is set and reports NOT CONFIGURED otherwise.
//
// OfferUp is still absent — no API and no scraping service worth wiring — so it
// stays a saved-search link in the digest, a 60-second manual check.
//
// Auth: service-role bearer (pg_cron) OR a master/sub admin JWT (dashboard).
// Body (all optional): { trigger?: "cron" | "manual", dryRun?: boolean }
//
// A source that FAILS is reported as failed. It is never reported as "nothing
// found" — a silent outage reading as "your gear isn't listed anywhere" is the
// single most dangerous bug this function could have.

import { createClient } from "npm:@supabase/supabase-js@2";
// The confidence model lives in scoring.ts so it can be unit-tested without
// credentials or network: node --test supabase/functions/scan-gear-market/scoring.test.ts
import { bestMatch, type Listing, type Target } from "./scoring.ts";
import {
  APIFY_FB_ACTOR_DEFAULT, apifyRunUrl, facebookSearchUrl, parseFacebookMarketplace,
  CRAIGSLIST_HEADERS, craigslistSapiUrl, parseCraigslistSapi,
  mergeReverbDetail, parseReverbListings, reverbDetailUrl, reverbHeaders, reverbSearchUrl,
} from "./parsers.ts";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

// ── sources ─────────────────────────────────────────────────────────────────
async function fromReverb(query: string): Promise<Listing[]> {
  const token = Deno.env.get("REVERB_TOKEN");
  if (!token) throw new Error("REVERB_TOKEN not set");
  const res = await fetch(reverbSearchUrl(query), { headers: reverbHeaders(token) });
  if (!res.ok) throw new Error(`Reverb responded ${res.status}`);
  // Validate the payload, not just the status — an authenticated endpoint that
  // answers 200 with an error body would otherwise read as "zero listings".
  return parseReverbListings(await res.json());
}

async function ebayToken(): Promise<string> {
  const id = Deno.env.get("EBAY_CLIENT_ID"), secret = Deno.env.get("EBAY_CLIENT_SECRET");
  if (!id || !secret) throw new Error("EBAY_CLIENT_ID / EBAY_CLIENT_SECRET not set");
  const res = await fetch("https://api.ebay.com/identity/v1/oauth2/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      "Authorization": "Basic " + btoa(`${id}:${secret}`),
    },
    body: "grant_type=client_credentials&scope=" + encodeURIComponent("https://api.ebay.com/oauth/api_scope"),
  });
  const j = await res.json().catch(() => ({}));
  if (!res.ok || !j.access_token) throw new Error(`eBay auth ${res.status}: ${j.error_description || ""}`);
  return j.access_token;
}

async function fromEbay(query: string, token: string): Promise<Listing[]> {
  const url = `https://api.ebay.com/buy/browse/v1/item_summary/search?q=${encodeURIComponent(query)}&limit=50&filter=${encodeURIComponent("buyingOptions:{FIXED_PRICE|AUCTION}")}`;
  const res = await fetch(url, {
    headers: {
      "Authorization": `Bearer ${token}`,
      "X-EBAY-C-MARKETPLACE-ID": "EBAY_US",
      "Content-Type": "application/json",
    },
  });
  if (!res.ok) throw new Error(`eBay responded ${res.status}`);
  const j = await res.json();
  const items = j?.itemSummaries;
  if (!Array.isArray(items)) return [];   // a genuinely empty result set is [] / absent
  return items.map((x: Record<string, any>) => ({
    source: "ebay" as const,
    listing_id: String(x.itemId),
    url: x.itemWebUrl || "",
    title: x.title || "",
    price: x.price?.value ? Number(x.price.value) : null,
    currency: x.price?.currency || "USD",
    location: [x.itemLocation?.city, x.itemLocation?.stateOrProvince, x.itemLocation?.postalCode]
      .filter(Boolean).join(", ") || null,
    seller: x.seller?.username || null,
    seller_feedback: x.seller?.feedbackScore != null ? Number(x.seller.feedbackScore) : null,
    posted_at: x.itemCreationDate || null,
    image_url: x.image?.imageUrl || null,
    description: x.shortDescription || null,
    raw: x,
  }));
}

async function fromFacebook(query: string): Promise<Listing[]> {
  const token = Deno.env.get("APIFY_TOKEN");
  if (!token) throw new Error("APIFY_TOKEN not set");
  const actor = Deno.env.get("APIFY_FB_ACTOR") || APIFY_FB_ACTOR_DEFAULT;
  const loc = Deno.env.get("FB_MARKETPLACE_LOCATION") || "nyc";
  const res = await fetch(apifyRunUrl(actor, token), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      startUrls: [{ url: facebookSearchUrl(query, loc) }],
      resultsLimit: Number(Deno.env.get("FB_RESULTS_LIMIT") || 15),
      // Detail mode is REQUIRED, not a nicety: the summary shape has no date at
      // all, so every listing would skip the theft gate, earn no recency points,
      // and cap below the alert threshold. It also returns the description,
      // which is where a serial number would appear.
      includeListingDetails: true,
    }),
  });
  if (!res.ok) {
    const b = await res.text().catch(() => "");
    throw new Error(`Apify responded ${res.status}${res.status === 402 ? " — out of credit" : ""}${b ? ": " + b.slice(0, 120) : ""}`);
  }
  return parseFacebookMarketplace(await res.json());
}

async function fromCraigslist(query: string): Promise<Listing[]> {
  // The undocumented endpoint Craigslist's own search box calls. The decode
  // lives in parsers.ts so it can be re-checked against the live response
  // without a deploy: node scripts/gearwatch_live_test.mjs
  const res = await fetch(craigslistSapiUrl(query), { headers: CRAIGSLIST_HEADERS });
  if (!res.ok) throw new Error(`Craigslist responded ${res.status}`);
  // Validate the decoded payload, not the status — undocumented means it can
  // change shape without warning, and a shape change must read as FAILED.
  return parseCraigslistSapi(await res.json());
}

// ── manual-check links for the sites with no usable API ─────────────────────
function manualLinks(queries: string[]): { label: string; url: string }[] {
  const q = encodeURIComponent(queries[0] || "pioneer cdj");
  return [
    { label: "OfferUp — NYC", url: `https://offerup.com/search?q=${q}` },
    { label: "eBay — sold & completed", url: `https://www.ebay.com/sch/i.html?_nkw=${q}&LH_Sold=1` },
  ];
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const URL_ = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const admin = createClient(URL_, SRK);

  // auth: service-role bearer (cron) OR admin JWT (dashboard)
  const auth = req.headers.get("Authorization") || "";
  const bearer = auth.replace(/^Bearer\s+/i, "");
  let authed = bearer === SRK;
  if (!authed && bearer) {
    const uc = createClient(URL_, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
    const { data: { user } } = await uc.auth.getUser();
    if (user) {
      const { data: p } = await admin.from("profiles").select("role").eq("id", user.id).single();
      authed = !!p && ["master_admin", "sub_admin"].includes(p.role);
    }
  }
  if (!authed) return err(401, "unauthorized");

  const body = await req.json().catch(() => ({}));
  const dryRun = body.dryRun === true;

  const { data: cfg } = await admin.from("gear_watch_config").select("*").eq("id", true).single();
  if (!cfg) return err(500, "gear_watch_config missing — apply migration 146");
  if (!cfg.enabled) return new Response(JSON.stringify({ ok: true, skipped: "disabled" }), { headers: JH });

  const { data: targets } = await admin.from("gear_watch_targets").select("*").eq("active", true);
  if (!targets?.length) return err(500, "no active gear_watch_targets");

  // ── collect ───────────────────────────────────────────────────────────────
  const queries = (targets as Target[]).map((t) => `${t.make || ""} ${t.model_tokens[0]}`.trim());
  const listings: Listing[] = [];
  const counts: Record<string, number> = { reverb: 0, ebay: 0, craigslist: 0, facebook: 0 };
  const fails: Record<string, string> = {};

  for (const q of queries) {
    try { const r = await fromReverb(q); listings.push(...r); counts.reverb += r.length; }
    catch (e) { fails.reverb = e instanceof Error ? e.message : String(e); }
  }

  // Missing credentials is a CONFIGURATION state, not an outage. Reporting it as
  // FAILED sent a "1 source failed" digest three times a day saying nothing but
  // "eBay still isn't set up" — which is precisely how a reader learns to skim
  // past failure lines, and then misses the one that matters.
  const ebayConfigured = !!(Deno.env.get("EBAY_CLIENT_ID") && Deno.env.get("EBAY_CLIENT_SECRET"));
  if (ebayConfigured) {
    try {
      const tok = await ebayToken();
      for (const q of queries) {
        try { const r = await fromEbay(q, tok); listings.push(...r); counts.ebay += r.length; }
        catch (e) { fails.ebay = e instanceof Error ? e.message : String(e); }
      }
    } catch (e) { fails.ebay = e instanceof Error ? e.message : String(e); }
  }

  for (const q of queries) {
    try { const r = await fromCraigslist(q); listings.push(...r); counts.craigslist += r.length; }
    catch (e) { fails.craigslist = e instanceof Error ? e.message : String(e); }
  }

  // Facebook Marketplace is the only source that COSTS MONEY per result, so it
  // does not ride the 3x-daily cron by default — four queries of detail-mode
  // results, three times a day, would burn the free Apify credit in days. It
  // runs on a manual "Run scan now", and on cron only if FB_ON_CRON is set.
  // A source that is deliberately skipped says so; it never reports zero.
  const fbConfigured = !!Deno.env.get("APIFY_TOKEN");
  const fbCronOn = (Deno.env.get("FB_ON_CRON") || "").toLowerCase() === "true";
  const fbOn = fbConfigured && (body.trigger !== "cron" || fbCronOn);
  if (fbOn) {
    for (const q of queries) {
      try { const r = await fromFacebook(q); listings.push(...r); counts.facebook += r.length; }
      catch (e) { fails.facebook = e instanceof Error ? e.message : String(e); }
    }
  }

  // One line per source, and a source that threw says FAILED — never "0 found".
  const sourceStatus: Record<string, string> = {};
  for (const k of ["reverb", "craigslist"]) {
    sourceStatus[k] = fails[k] ? `FAILED: ${fails[k]}` : `ok — ${counts[k]} listing(s) fetched`;
  }
  sourceStatus.ebay = !ebayConfigured
    ? "NOT CONFIGURED — set EBAY_CLIENT_ID and EBAY_CLIENT_SECRET"
    : (fails.ebay ? `FAILED: ${fails.ebay}` : `ok — ${counts.ebay} listing(s) fetched`);
  // "Not configured" is its own answer. It is not a failure, and it is certainly
  // not zero results — nobody should read this line as "Facebook has nothing".
  sourceStatus.facebook = !fbConfigured
    ? "NOT CONFIGURED — set APIFY_TOKEN to include Facebook Marketplace"
    : !fbOn
      ? "skipped on the schedule to control cost — press Run scan now to include it"
      : (fails.facebook ? `FAILED: ${fails.facebook}` : `ok — ${counts.facebook} listing(s) fetched`);

  // ── score ─────────────────────────────────────────────────────────────────
  const scoreCfg = { theft_date: cfg.theft_date, geo_terms: cfg.geo_terms || [] };

  // Dedupe first, so a listing answering three queries is enriched once.
  const unique: Listing[] = [];
  const seen = new Set<string>();
  for (const l of listings) {
    const key = `${l.source}:${l.listing_id}`;
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(l);
  }

  // ── enrich Reverb candidates ──────────────────────────────────────────────
  // Reverb's SEARCH payload has no location and no seller feedback; its DETAIL
  // endpoint has both. Fetching detail for all ~50 results per query would be
  // hundreds of calls, so only listings that already clear the model gate are
  // enriched — typically a handful. The cap is reported, never silent (§18).
  const REVERB_ENRICH_CAP = 40;
  const reverbToken = Deno.env.get("REVERB_TOKEN");
  let enriched = 0, enrichSkipped = 0;
  if (reverbToken) {
    const candidates = unique.filter((l) => l.source === "reverb" && bestMatch(l, targets as Target[], scoreCfg));
    for (const l of candidates) {
      if (enriched >= REVERB_ENRICH_CAP) { enrichSkipped++; continue; }
      try {
        const res = await fetch(reverbDetailUrl(l.listing_id), { headers: reverbHeaders(reverbToken) });
        if (!res.ok) continue;
        Object.assign(l, mergeReverbDetail(l, await res.json()));
        enriched++;
      } catch { /* detail is an enhancement; a failure leaves the search data intact */ }
    }
    if (enrichSkipped) sourceStatus.reverb += ` (detail lookup capped at ${REVERB_ENRICH_CAP}; ${enrichSkipped} scored without location)`;
  }

  const scored: Array<Listing & { target_id: string; score: number; breakdown: Record<string, unknown> }> = [];
  for (const l of unique) {
    const best = bestMatch(l, targets as Target[], scoreCfg);
    if (best) scored.push({ ...l, target_id: best.target.id, score: best.score, breakdown: best.breakdown });
  }

  if (dryRun) {
    return new Response(JSON.stringify({
      ok: true, dryRun: true, sources: sourceStatus,
      fetched: listings.length, matched: scored.length,
      top: scored.sort((a, b) => b.score - a.score).slice(0, 10)
        .map((s) => ({ score: s.score, title: s.title, url: s.url, breakdown: s.breakdown })),
    }), { headers: JH });
  }

  // ── store (upsert on (source, listing_id): re-seeing a listing moves
  //     last_seen_at, it does not create a duplicate or re-alert) ────────────
  let inserted = 0;
  for (const s of scored) {
    const { data: existing } = await admin.from("gear_watch_hits")
      .select("id").eq("source", s.source).eq("listing_id", s.listing_id).maybeSingle();

    if (existing) {
      await admin.from("gear_watch_hits").update({
        last_seen_at: new Date().toISOString(), price: s.price, score: s.score, score_breakdown: s.breakdown,
      }).eq("id", existing.id);
      continue;
    }

    const { error } = await admin.from("gear_watch_hits").insert({
      source: s.source, listing_id: s.listing_id, url: s.url, title: s.title,
      price: s.price, currency: s.currency, location: s.location, seller: s.seller,
      seller_feedback: s.seller_feedback, posted_at: s.posted_at, image_url: s.image_url,
      target_id: s.target_id, score: s.score, score_breakdown: s.breakdown, raw: s.raw,
    });
    if (!error) inserted++;
  }

  // ── alert ─────────────────────────────────────────────────────────────────
  // What still needs announcing is a question for the DATABASE, not for this run.
  // Asking "what did I insert just now?" strands anything found while the digest
  // email was unset — which is exactly what happened: the Brooklyn CDJ-3000 was
  // stored on 2026-08-18, the email address arrived a day later, and that hit
  // would have sat unannounced forever. It also misses a hit whose score crosses
  // the threshold on a later scan, e.g. once a serial is pasted in.
  const minScore = cfg.min_score ?? 65;
  const { data: pending, error: pendErr } = await admin.from("gear_watch_hits")
    .select("id, score, title, url, price, location, source, score_breakdown")
    .is("alerted_at", null).gte("score", minScore).eq("status", "new")
    .order("score", { ascending: false }).limit(25);
  if (pendErr) console.error("pending-alert query failed:", pendErr.message);

  const fresh = (pending || []).map((p: Record<string, any>) => ({
    id: p.id as string, score: p.score as number, title: p.title as string, url: p.url as string,
    price: p.price as number | null, location: p.location as string | null,
    source: p.source as string, breakdown: (p.score_breakdown || {}) as Record<string, unknown>,
  }));
  const failed = Object.entries(sourceStatus).filter(([, v]) => v.startsWith("FAILED"));
  let emailed = false, pushed = 0;

  if ((fresh.length || failed.length) && cfg.email_to) {
    const rows = fresh.map((f) => `
      <tr>
        <td style="padding:8px;border-bottom:1px solid #eee;font-weight:700;font-size:18px">${f.score}</td>
        <td style="padding:8px;border-bottom:1px solid #eee">
          <a href="${f.url}" style="color:#1a1410">${f.title.replace(/</g, "&lt;")}</a><br>
          <span style="font-size:12px;color:#666">${f.source} · ${f.price != null ? "$" + f.price : "no price"} · ${f.location || "location unknown"}</span><br>
          <span style="font-size:11px;color:#999">${Object.entries(f.breakdown).map(([k, v]) => `${k}: ${v}`).join(" · ")}</span>
        </td>
      </tr>`).join("");

    const failHtml = failed.length
      ? `<p style="background:#fdecea;padding:10px;border-left:3px solid #c00"><b>${failed.length} source(s) failed this run:</b> ${failed.map(([k, v]) => `${k} — ${v}`).join("; ")}.<br>
         Treat this as "not checked", not as "nothing found".</p>` : "";

    const links = manualLinks(queries).map((l) => `<a href="${l.url}">${l.label}</a>`).join(" · ");

    const html = `
      <div style="font-family:Arial,sans-serif;max-width:640px">
        <h2 style="font-family:Georgia,serif">Gear Watch — ${fresh.length} new hit${fresh.length === 1 ? "" : "s"}</h2>
        ${failHtml}
        ${fresh.length ? `<table style="width:100%;border-collapse:collapse">${rows}</table>`
                       : "<p>No new listings above the score threshold this run.</p>"}
        <p style="margin-top:20px;font-size:13px">Sites with no API — check by hand:<br>${links}</p>
        <p style="margin-top:20px;padding:10px;background:#f5f2ed;font-size:13px">
          <b>If something looks right, send the link and the score breakdown to the detective.</b>
          Do not contact the seller and do not arrange a meet.
        </p>
      </div>`;

    const res = await fetch(`${URL_}/functions/v1/send-notice`, {
      method: "POST",
      headers: { "Authorization": `Bearer ${SRK}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        to: cfg.email_to,
        subject: fresh.length ? `Gear Watch — ${fresh.length} new hit${fresh.length === 1 ? "" : "s"} (top score ${fresh[0].score})`
                              : `Gear Watch — ${failed.length} source(s) failed`,
        html,
      }),
    });
    emailed = res.ok;
  }

  // High scores also buzz the phone: a serial match or a full-rig bundle is
  // worth interrupting someone for. Anything below push_score waits for the digest.
  const urgent = fresh.filter((f) => f.score >= (cfg.push_score ?? 85));
  if (urgent.length && cfg.push_user_id) {
    for (const u of urgent.slice(0, 5)) {
      await admin.from("notifications").insert({
        user_id: cfg.push_user_id, kind: "fyi",
        title: `Gear Watch: ${u.score}/100 — ${u.title.slice(0, 80)}`,
        body: `${u.source} · ${u.price != null ? "$" + u.price : "no price"} · ${u.location || "location unknown"}`,
        subject_type: "gearwatch",
      });
      const r = await fetch(`${URL_}/functions/v1/send-push`, {
        method: "POST",
        headers: { "Authorization": `Bearer ${SRK}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          user_id: cfg.push_user_id,
          title: `Gear Watch — ${u.score}/100`,
          body: u.title.slice(0, 120),
          url: u.url,
        }),
      });
      if (r.ok) pushed++;
    }
  }

  // Stamp only what was actually announced, keyed by primary key — so the next
  // run in eight hours does not re-announce the same listing.
  if (emailed || pushed) {
    await admin.from("gear_watch_hits")
      .update({ alerted_at: new Date().toISOString() })
      .in("id", fresh.map((f) => f.id));
  }

  const status = failed.length ? `partial: ${failed.map(([k]) => k).join(",")} failed` : "ok";
  await admin.from("gear_watch_config")
    .update({ last_run_at: new Date().toISOString(), last_status: status }).eq("id", true);

  return new Response(JSON.stringify({
    ok: true, trigger: body.trigger || "manual", sources: sourceStatus,
    fetched: listings.length, matched: scored.length, inserted, alerted: fresh.length,
    emailed, pushed,
  }), { headers: JH });
});
