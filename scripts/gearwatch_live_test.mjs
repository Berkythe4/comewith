// Live check of the Gear Watch pipeline against the one source that needs no
// credentials: Craigslist, through the internal JSON endpoint its own search box
// calls. Runs the REAL decoder and the REAL scorer, so what you see here is what
// the deployed function will do.
//
//   node scripts/gearwatch_live_test.mjs                  # all four targets
//   node scripts/gearwatch_live_test.mjs "pioneer xdj"    # one query
//   THEFT_DATE=2026-08-16 node scripts/gearwatch_live_test.mjs
//
// Prints every listing the scorer keeps, with its breakdown, and every listing it
// dropped with the reason — the drops are the half that proves the gates work.

import { writeFileSync } from "node:fs";
import {
  parseCraigslistSapi,
  craigslistSapiUrl,
  CRAIGSLIST_HEADERS,
  parseReverbListings,
  reverbSearchUrl,
  reverbHeaders,
  reverbDetailUrl,
  mergeReverbDetail,
} from "../supabase/functions/scan-gear-market/parsers.ts";
import { bestMatch } from "../supabase/functions/scan-gear-market/scoring.ts";

// Mirrors the rows seeded by migration 146.
const TARGETS = [
  { id: "t-xdj", label: "Pioneer XDJ-AZ", make: "Pioneer DJ",
    model_tokens: ["xdj-az", "xdjaz", "xdj az"],
    exclude_tokens: ["case", "cover", "decal", "skin", "stand", "bag", "lid", "sticker", "manual", "parts", "broken"],
    serial: null, qty: 1, typical_resale: 2400 },
  { id: "t-cdj", label: "Pioneer CDJ-3000", make: "Pioneer DJ",
    model_tokens: ["cdj-3000", "cdj3000", "cdj 3000"],
    exclude_tokens: ["case", "cover", "decal", "skin", "stand", "bag", "lid", "sticker", "manual", "parts", "broken"],
    serial: null, qty: 2, typical_resale: 2200 },
  { id: "t-wave", label: "AlphaTheta Wave 8", make: "AlphaTheta",
    model_tokens: ["wave 8", "wave-8", "wave8"],
    exclude_tokens: ["case", "cover", "decal", "skin", "bag", "sticker", "manual", "parts", "broken"],
    serial: null, qty: 2, typical_resale: 700 },
  { id: "t-krk", label: "KRK Rokit 5", make: "KRK",
    model_tokens: ["rokit 5", "rokit5", "rp5"],
    exclude_tokens: ["case", "cover", "stand", "bag", "sticker", "manual", "parts", "broken", "pair"],
    serial: null, qty: 1, typical_resale: 150 },
];

// Optional. With it, Reverb is scanned too — the fastest way to confirm a new
// token works before it goes anywhere near prod:
//   REVERB_TOKEN=xxxx node scripts/gearwatch_live_test.mjs
const REVERB_TOKEN = process.env.REVERB_TOKEN || null;

const CFG = {
  theft_date: process.env.THEFT_DATE || null,   // unset = no pre-theft gate
  geo_terms: ["new york", "ny", "nyc", "brooklyn", "queens", "bronx", "manhattan",
              "staten island", "jersey city", "newark", "hoboken", "yonkers", "long island"],
};

// Any number of queries can be passed. Craigslist matches on the literal words in
// the title, so the same model needs several spellings: a seller who wrote
// "XDJ AZ" is invisible to a search for "xdj-az". The scorer's gates do the real
// filtering, so it costs nothing to cast the search wide.
const argQueries = process.argv.slice(2).filter((a) => !a.startsWith("--"));
const queries = argQueries.length ? argQueries : [
  "pioneer cdj 3000", "cdj-3000", "pioneer xdj-az", "xdj az",
  "alphatheta wave 8", "wave 8 speaker", "krk rokit 5", "pioneer dj",
];

let totalFetched = 0, totalKept = 0;
const allKept = [];   // collected for the clickable HTML report

for (const q of queries) {
  console.log(`\n── "${q}" ${"─".repeat(Math.max(0, 52 - q.length))}`);

  const listings = [];

  try {
    const res = await fetch(craigslistSapiUrl(q), { headers: CRAIGSLIST_HEADERS });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const got = parseCraigslistSapi(await res.json());
    listings.push(...got);
    console.log(`  craigslist: ${got.length} listing(s)`);
  } catch (e) {
    console.log(`  craigslist FAILED: ${e.message}`);
  }

  // Reverb runs only when a token is in the environment. No token is reported as
  // NOT CHECKED, never as zero results — the whole point of §24.
  if (REVERB_TOKEN) {
    try {
      const res = await fetch(reverbSearchUrl(q), { headers: reverbHeaders(REVERB_TOKEN) });
      if (!res.ok) throw new Error(`HTTP ${res.status}${res.status === 401 ? " — token rejected" : ""}`);
      const got = parseReverbListings(await res.json());
      listings.push(...got);
      console.log(`  reverb:     ${got.length} listing(s)`);
    } catch (e) {
      console.log(`  reverb FAILED: ${e.message}`);
    }
  } else {
    console.log("  reverb:     NOT CHECKED (set REVERB_TOKEN to include it)");
  }

  totalFetched += listings.length;

  // Reverb's search payload carries no location or seller feedback; its detail
  // endpoint does. Only listings that already clear the model gate are enriched.
  if (REVERB_TOKEN) {
    let n = 0;
    for (const l of listings) {
      if (l.source !== "reverb" || !bestMatch(l, TARGETS, CFG)) continue;
      try {
        const res = await fetch(reverbDetailUrl(l.listing_id), { headers: reverbHeaders(REVERB_TOKEN) });
        if (res.ok) { Object.assign(l, mergeReverbDetail(l, await res.json())); n++; }
      } catch { /* enhancement only */ }
    }
    if (n) console.log(`  reverb:     ${n} candidate(s) enriched with location + seller feedback`);
  }

  const kept = [], dropped = [];
  for (const l of listings) {
    const m = bestMatch(l, TARGETS, CFG);
    if (m) { kept.push({ l, m }); continue; }
    // Why was it dropped? A silent drop is indistinguishable from a broken gate.
    const squash = (s) => (s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const named = TARGETS.find((t) => t.model_tokens.some((tok) => squash(l.title).includes(squash(tok))));
    if (named) {
      const why = named.exclude_tokens.find((tok) => squash(l.title).includes(squash(tok)))
        ? "accessory/excluded word in title"
        : "posted before the theft date";
      dropped.push({ l, why: `${why} (vs ${named.label})` });
    }
  }

  totalKept += kept.length;
  allKept.push(...kept.map(({ l, m }) => ({ ...l, score: m.score, breakdown: m.breakdown })));
  kept.sort((a, b) => b.m.score - a.m.score);
  for (const { l, m } of kept) {
    console.log(`\n  ${String(m.score).padStart(3)}/100  ${l.title}`);
    console.log(`          ${l.price != null ? "$" + l.price : "no price"} · ${l.location || "location unknown"} · ${l.posted_at ? l.posted_at.slice(0, 10) : "no date"}`);
    console.log(`          ${l.url}`);
    console.log(`          ${Object.entries(m.breakdown).map(([k, v]) => `${k}=${v}`).join("  ")}`);
  }
  for (const { l, why } of dropped) console.log(`\n   drop   ${l.title}\n          ${why}`);
  if (!kept.length && !dropped.length) console.log("  nothing matched the rig (the normal, expected result)");
}

console.log(`\n═══ ${totalFetched} listing(s) fetched, ${totalKept} scored as candidates ═══`);

// ── clickable report ────────────────────────────────────────────────────────
// Until the dashboard panel is deployed there is nowhere to click a result, and
// a URL in a terminal is not a link you can open on a phone. Written every run.
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
// One listing can answer several queries ("cdj-3000" and "pioneer cdj 3000" are
// the same posting). The edge function dedupes on (source, listing_id) in the
// database; here it is deduped by URL so the report shows each listing once.
const byUrl = new Map();
for (const h of allKept) if (!byUrl.has(h.url) || byUrl.get(h.url).score < h.score) byUrl.set(h.url, h);
allKept.length = 0;
allKept.push(...byUrl.values());
allKept.sort((a, b) => b.score - a.score);
const band = (s) => s >= 85 ? "#c0392b" : s >= 65 ? "#b58900" : "#666";
const rows = allKept.map((h) => `
  <tr>
    <td class="score" style="color:${band(h.score)}">${h.score}</td>
    <td>
      <a href="${esc(h.url)}" target="_blank" rel="noopener">${esc(h.title)}</a>
      <div class="meta">${esc(h.source)} · ${h.price != null ? "$" + h.price : "no price"} · ${esc(h.location || "location unknown")} · ${h.posted_at ? esc(h.posted_at.slice(0, 10)) : "no date"}</div>
      <div class="bd">${Object.entries(h.breakdown).map(([k, v]) => `${esc(k.replace(/_/g, " "))}: <b>${esc(v)}</b>`).join(" · ")}</div>
    </td>
  </tr>`).join("");

const html = `<!doctype html><meta charset="utf-8"><title>Gear Watch — scan results</title>
<style>
 body{font:15px/1.5 -apple-system,Segoe UI,Arial,sans-serif;max-width:900px;margin:32px auto;padding:0 16px;color:#1a1410}
 h1{font-size:22px;margin:0 0 4px} .sub{color:#666;font-size:13px;margin-bottom:20px}
 table{width:100%;border-collapse:collapse} td{padding:12px 8px;border-bottom:1px solid #eee;vertical-align:top}
 .score{font-size:22px;font-weight:700;width:60px;text-align:center}
 .meta{color:#666;font-size:13px;margin-top:2px} .bd{color:#999;font-size:12px;margin-top:4px}
 a{color:#1a1410} .note{margin-top:24px;padding:12px;background:#f5f2ed;font-size:13px;border-left:3px solid #c0392b}
 .empty{padding:24px;background:#f5f2ed;color:#666}
</style>
<h1>Gear Watch — scan results</h1>
<div class="sub">Craigslist NYC${REVERB_TOKEN ? " + Reverb" : " only — <b>Reverb not checked</b> (no token), eBay not checked"} · ${totalFetched} listing(s) fetched · ${totalKept} scored${CFG.theft_date ? ` · theft date ${esc(CFG.theft_date)}` : " · no theft date set, recency not scored"}</div>
${allKept.length ? `<table>${rows}</table>` : `<div class="empty">Nothing matched the rig this run. That is the expected result most days.</div>`}
<div class="note"><b>If one of these looks right, send the link and its score breakdown to the detective.</b>
Do not contact the seller and do not arrange a meet.</div>`;

const out = "gearwatch_results.html";
writeFileSync(out, html);
console.log(`\nClickable report written to ${out} — open it in a browser.`);
