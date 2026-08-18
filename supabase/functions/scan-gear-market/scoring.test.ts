// Tests for the Gear Watch confidence model.
//
//   node --test supabase/functions/scan-gear-market/
//
// No API keys, no network, no database. Every case below is a listing shape
// that actually shows up on Reverb / eBay / Craigslist when you search for this
// gear — the accessory listings and the wrong-model near-misses are the whole
// reason the gates exist.

import { test } from "node:test";
import assert from "node:assert/strict";
import { bestMatch, hasToken, scoreListing, type Listing, type Target } from "./scoring.ts";

const CFG = { theft_date: "2026-08-16", geo_terms: ["new york", "ny", "brooklyn", "queens"] };

const XDJ: Target = {
  id: "t-xdj", label: "Pioneer XDJ-AZ", make: "Pioneer DJ",
  model_tokens: ["xdj-az", "xdjaz", "xdj az"],
  exclude_tokens: ["case", "cover", "decal", "skin", "stand", "bag"],
  serial: null, qty: 1, typical_resale: 2400,
};
const CDJ: Target = {
  id: "t-cdj", label: "Pioneer CDJ-3000", make: "Pioneer DJ",
  model_tokens: ["cdj-3000", "cdj3000", "cdj 3000"],
  exclude_tokens: ["case", "cover", "decal", "skin", "stand", "bag"],
  serial: null, qty: 2, typical_resale: 2200,
};
const KRK: Target = {
  id: "t-krk", label: "KRK Rokit 5", make: "KRK",
  model_tokens: ["rokit 5", "rokit5", "rp5"],
  exclude_tokens: ["case", "cover", "stand", "pair"],
  serial: null, qty: 1, typical_resale: 150,
};
const ALL = [XDJ, CDJ, KRK];

function listing(over: Partial<Listing> = {}): Listing {
  return {
    source: "reverb", listing_id: "1", url: "https://reverb.com/item/1",
    title: "Pioneer CDJ-3000", price: null, currency: "USD",
    location: null, seller: null, seller_feedback: null,
    posted_at: "2026-08-17T12:00:00Z", image_url: null, description: null,
    ...over,
  };
}

// ── token matching ──────────────────────────────────────────────────────────
test("model spellings all match: hyphen, space, squashed", () => {
  for (const t of ["Pioneer CDJ-3000", "pioneer cdj 3000", "CDJ3000 for sale"]) {
    assert.ok(scoreListing(listing({ title: t }), CDJ, ALL, CFG), `${t} should match`);
  }
});

test("an en dash in a pasted title still matches", () => {
  assert.ok(hasToken("Pioneer CDJ–3000 mint", "cdj-3000"));
});

test("a different model does not match", () => {
  assert.equal(scoreListing(listing({ title: "Pioneer CDJ-2000NXS2" }), CDJ, ALL, CFG), null);
});

// ── gates ───────────────────────────────────────────────────────────────────
test("GATE: accessory listings are dropped, not scored low", () => {
  for (const t of [
    "Hard case for Pioneer CDJ-3000",
    "Decal skin for CDJ-3000",
    "Odyssey bag CDJ 3000",
  ]) {
    assert.equal(scoreListing(listing({ title: t }), CDJ, ALL, CFG), null, `${t} should be gated`);
  }
});

test("GATE: a case mentioned in the DESCRIPTION does not kill a real listing", () => {
  const l = listing({ title: "Pioneer CDJ-3000", description: "Includes original case and box." });
  assert.ok(scoreListing(l, CDJ, ALL, CFG));
});

test("GATE: listings posted before the theft are dropped", () => {
  const l = listing({ posted_at: "2026-07-04T12:00:00Z" });
  assert.equal(scoreListing(l, CDJ, ALL, CFG), null);
});

test("GATE: a listing with no date survives (we do not guess)", () => {
  const l = listing({ posted_at: null });
  assert.ok(scoreListing(l, CDJ, ALL, CFG));
});

test("GATE: same-day-as-theft listing survives", () => {
  const l = listing({ posted_at: "2026-08-16T09:00:00Z" });
  assert.ok(scoreListing(l, CDJ, ALL, CFG));
});

// ── signals ─────────────────────────────────────────────────────────────────
test("a bare model match scores the base 35 and nothing else", () => {
  const r = scoreListing(listing({ posted_at: null }), CDJ, ALL, CFG)!;
  assert.equal(r.score, 35);
  assert.deepEqual(Object.keys(r.breakdown), ["model_match"]);
});

// The calibration case, from the first live run: a CDJ-3000 in Brooklyn listed
// the day after the theft. This is the exact shape of a real hit and it has to
// land in "look now" territory, not in "probably nothing".
test("model + local + listed next day lands at 85", () => {
  const l = listing({ title: "Pioneer CDJ-3000 Professional DJ multi player",
    price: 2200, location: "brooklyn, newyork", posted_at: "2026-08-17T12:00:00Z" });
  const r = scoreListing(l, CDJ, ALL, CFG)!;
  assert.equal(r.score, 85);
  assert.equal(r.breakdown.listed_soon_after_theft, 25);
});

test("recency is graduated — a day after beats three weeks after", () => {
  const at = (d: string) => scoreListing(listing({ posted_at: d, location: "Brooklyn, NY" }), CDJ, ALL, CFG)!;
  const day1 = at("2026-08-17T00:00:00Z"), day10 = at("2026-08-26T00:00:00Z"), day40 = at("2026-09-25T00:00:00Z");
  assert.equal(day1.breakdown.listed_soon_after_theft, 25);
  assert.equal(day10.breakdown.listed_soon_after_theft, 15);
  assert.equal(day40.breakdown.listed_soon_after_theft, 5);
  assert.ok(day1.score > day10.score && day10.score > day40.score);
});

test("serial match dominates the score", () => {
  const withSerial = { ...CDJ, serial: "ABX1234567" };
  const l = listing({ title: "Pioneer CDJ-3000 s/n ABX1234567", posted_at: null });
  const r = scoreListing(l, withSerial, [withSerial], CFG)!;
  assert.equal(r.breakdown.serial_match, 60);
  assert.ok(r.score >= 85);
});

test("local listings score the geo bump, distant ones do not", () => {
  const near = scoreListing(listing({ location: "Brooklyn, NY", posted_at: null }), CDJ, ALL, CFG)!;
  const far = scoreListing(listing({ location: "Los Angeles, CA", posted_at: null }), CDJ, ALL, CFG)!;
  assert.equal(near.breakdown.local, 25);
  assert.equal(far.breakdown.local, undefined);
});

test("two stolen models in one listing is the strongest non-serial signal", () => {
  const l = listing({ title: "Pioneer XDJ-AZ and 2x CDJ-3000 DJ setup", posted_at: null });
  const r = bestMatch(l, ALL, CFG)!;
  assert.equal(r.breakdown.bundle_multi_model, 30);
  assert.ok(r.score >= 65);
});

test("the full-rig bundle, local and fresh, pins at 100", () => {
  const l = listing({ title: "Pioneer XDJ-AZ and 2x CDJ-3000 DJ setup",
    location: "Queens, NY", posted_at: "2026-08-17T00:00:00Z" });
  assert.equal(bestMatch(l, ALL, CFG)!.score, 100);
});

test("a pair of a model we lost a pair of scores the weaker pair bump", () => {
  const l = listing({ title: "Pair of Pioneer CDJ-3000", posted_at: null });
  const r = scoreListing(l, CDJ, ALL, CFG)!;
  assert.equal(r.breakdown.bundle_pair, 18);
  assert.equal(r.breakdown.bundle_multi_model, undefined);
});

// Calibration guard: model + recency alone, with no local signal, must stay
// BELOW the 65 email threshold. An upstate listing of a common monitor should
// sit in the panel for triage, not arrive in the inbox.
test("model + fresh but not local stays under the email threshold", () => {
  const l = listing({ title: "KRK Rokit 5 Studio Monitors", price: 200,
    location: "SCHENECTADY, albany", posted_at: "2026-08-17T00:00:00Z" });
  const r = scoreListing(l, KRK, ALL, CFG)!;
  assert.equal(r.score, 60);
  assert.ok(r.score < 65, "must not reach the digest on model + recency alone");
});

test("a single-unit target gets no pair bump", () => {
  const l = listing({ title: "Pair of KRK Rokit 5", posted_at: null });
  // 'pair' is an exclude token for the KRK: we lost one, a pair is not ours
  assert.equal(scoreListing(l, KRK, ALL, CFG), null);
});

test("price far below market scores more than price near market", () => {
  const cheap = scoreListing(listing({ price: 900, posted_at: null }), CDJ, ALL, CFG)!;
  const fair = scoreListing(listing({ price: 2100, posted_at: null }), CDJ, ALL, CFG)!;
  assert.equal(cheap.breakdown.price_far_below_market, 15);
  assert.equal(fair.breakdown.price_far_below_market, undefined);
  assert.equal(fair.breakdown.price_below_market, undefined);
});

test("a brand-new seller account scores the thin-history bump", () => {
  const r = scoreListing(listing({ seller_feedback: 2, posted_at: null }), CDJ, ALL, CFG)!;
  assert.equal(r.breakdown.new_seller, 10);
});

test("listing language is capped so it can never carry a hit alone", () => {
  const l = listing({
    title: "CDJ-3000 no box no receipt cash only must go moving firm on price",
    posted_at: null,
  });
  const r = scoreListing(l, CDJ, ALL, CFG)!;
  assert.equal(r.breakdown.listing_language, 10);
});

test("the score is capped at 100", () => {
  const withSerial = { ...CDJ, serial: "ABX1234567" };
  const l = listing({
    title: "XDJ-AZ + pair CDJ-3000 ABX1234567 no box cash only",
    price: 500, location: "Queens, NY", seller_feedback: 0,
    posted_at: "2026-08-17T00:00:00Z",
  });
  const r = bestMatch(l, [withSerial, XDJ, KRK], CFG)!;
  assert.equal(r.score, 100);
});

test("breakdown is always populated for a stored hit", () => {
  const r = bestMatch(listing({ location: "Brooklyn, NY", price: 800 }), ALL, CFG)!;
  assert.ok(Object.keys(r.breakdown).length >= 2);
  assert.equal(typeof r.score, "number");
});

test("a listing matching nothing returns null from bestMatch", () => {
  assert.equal(bestMatch(listing({ title: "Numark Mixtrack Pro" }), ALL, CFG), null);
});

// ── craigslist sapi decoder ─────────────────────────────────────────────────
// Shape captured from a real response on 2026-08-18. The payload is delta-encoded:
// ids accumulate from decode.minPostingId, dates are seconds after minPostedDate,
// and price -1 means "no price stated" rather than free.
// Structure copied from a real response on 2026-08-18 (query "krk rokit 5").
// Note the geo strings: "<locationIdx>:<descriptionIdx>", two DIFFERENT indexes.
const SAPI = {
  data: {
    decode: {
      minPostingId: 7944653815,
      minPostedDate: 1784118455,
      locations: [0, [349, "cnj"], [561, "jerseyshore"], [3, "newyork", "brk"], [59, "albany"]],
      locationDescriptions: [0, "Wycombe", "Lawrence Township, NJ", "Toms River", "brooklyn", "SCHENECTADY"],
    },
    items: [
      [0, 243008, 98, 2200, "3:4~40.6742~-73.7057", "0kl0qC",
        [13, "j9KYtjcUkmV7aB5NzkJ1Xk"], [6, "brooklyn-pioneer-cdj-3000"], [10, "$2,200"],
        "Pioneer CDJ-3000 Professional DJ multi player"],
      [8750644, 0, 98, 1100, "4:5~42.7823~-73.9448", 0,
        [13, "dzpAyVESDpcsBa9WW8iC3m"], [6, "schenectady-pioneer-cdj-3000"], [10, "$1,100"],
        "Pioneer CDJ-3000 club player"],
      [10, 500, 98, -1, "3:4~40.67~-73.70", 0,
        [13, "tokenNoPrice00000000000"], [6, "cdj-3000-no-price"], "Pioneer CDJ-3000 make offer"],
    ],
  },
};

test("sapi decoder: ids accumulate, dates decode, -1 price becomes null", async () => {
  const { parseCraigslistSapi } = await import("./parsers.ts");
  const out = parseCraigslistSapi(SAPI);
  assert.equal(out.length, 3);
  assert.equal(out[0].listing_id, "7944653815");
  assert.equal(out[1].listing_id, "7953404459", "ids are cumulative deltas, not absolute");
  assert.equal(out[0].price, 2200);
  assert.equal(out[2].price, null, "-1 means no price stated, never free");
  assert.equal(out[0].posted_at, new Date((1784118455 + 243008) * 1000).toISOString());
});

test("sapi decoder: canonical URL is /view/d/<slug>/<token>", async () => {
  const { parseCraigslistSapi } = await import("./parsers.ts");
  const out = parseCraigslistSapi(SAPI);
  assert.equal(out[0].url, "https://www.craigslist.org/view/d/brooklyn-pioneer-cdj-3000/j9KYtjcUkmV7aB5NzkJ1Xk");
});

// The geo string carries TWO indexes into TWO arrays. Using the first for both
// put a Schenectady listing in Bushwick on the first live run.
test("sapi decoder: location and neighbourhood use their own separate indexes", async () => {
  const { parseCraigslistSapi } = await import("./parsers.ts");
  const out = parseCraigslistSapi(SAPI);
  assert.equal(out[0].location, "brooklyn, newyork");
  assert.equal(out[1].location, "SCHENECTADY, albany");
});

// The NYC board returns nearby-area results — upstate, New Jersey, Connecticut.
test("an upstate listing is not scored as local", async () => {
  const { parseCraigslistSapi } = await import("./parsers.ts");
  const out = parseCraigslistSapi(SAPI);
  const scored = bestMatch(out[1], ALL, { theft_date: null, geo_terms: CFG.geo_terms });
  assert.equal(scored?.breakdown.local, undefined, "Schenectady is not New York City");
});

// "albany" contains "ny". A substring match on short geo terms scored every
// upstate listing as local on the first live run.
test("geo terms match on word boundaries, not substrings", async () => {
  const { isLocalTerm } = await import("./scoring.ts");
  assert.equal(isLocalTerm("SCHENECTADY, albany", "ny"), false, "'albany' must not match 'ny'");
  assert.equal(isLocalTerm("Brooklyn, NY", "ny"), true);
  assert.equal(isLocalTerm("brooklyn, newyork", "brooklyn"), true);
  assert.equal(isLocalTerm("Bethany Beach, DE", "ny"), false);
  assert.equal(isLocalTerm("Queens, New York", "new york"), true);
});

test("an out-of-state craigslist listing gets no local bump", () => {
  const phl = { source: "craigslist" as const, listing_id: "7", url: "u", title: "Pioneer CDJ-3000 single",
    price: 2150, currency: "USD", location: "Philadelphia", seller: null, seller_feedback: null,
    posted_at: null, image_url: null, description: null };
  const r = scoreListing(phl, CDJ, ALL, CFG)!;
  assert.equal(r.breakdown.local, undefined);
});

// The whole point of §24: a block page and an empty result set must not look
// alike. Craigslist answers a blocked request with HTML, and an error with a
// JSON envelope — both have to throw, not return [].
test("a block page throws rather than returning zero listings", async () => {
  const { parseCraigslistSapi } = await import("./parsers.ts");
  assert.throws(() => parseCraigslistSapi("<html><body>Your request has been blocked.</body></html>"));
  assert.throws(() => parseCraigslistSapi({ errors: [{ message: "Sorry, we don't recognize GET requests" }] }),
    /unexpected payload/);
  assert.throws(() => parseCraigslistSapi({ data: {} }), /unexpected payload/);
});

// Craigslist returns `decode: 0` — the NUMBER — when a query has no results.
// Treating that as a bad payload reported "nothing listed" as "we were blocked",
// which is §24's mistake pointed the other way. Seen live on 2026-08-18.
test("a genuinely empty result set returns [] and does not throw", async () => {
  const { parseCraigslistSapi } = await import("./parsers.ts");
  assert.deepEqual(parseCraigslistSapi({ data: { items: [], decode: 0 } }), []);
  assert.deepEqual(parseCraigslistSapi({ data: { items: [], decode: 0 }, errors: [] }), []);
});

test("items present but no decode table still throws", async () => {
  const { parseCraigslistSapi } = await import("./parsers.ts");
  assert.throws(() => parseCraigslistSapi({ data: { items: [[0, 0, 98, 100, "1:1~0~0", 0, [13, "t"], [6, "s"], "title"]], decode: 0 } }),
    /decode table missing/);
});

// ── reverb payload mapping ──────────────────────────────────────────────────
test("reverb listings map to the common shape", async () => {
  const { parseReverbListings } = await import("./parsers.ts");
  const out = parseReverbListings({
    listings: [{
      id: 12345, title: "Pioneer CDJ-3000", price: { amount: "2100.00", currency: "USD" },
      shop_name: "Some Shop", shop: { slug: "some-shop", preferred_seller: false },
      published_at: "2026-08-17T10:00:00Z",
      _links: { web: { href: "https://reverb.com/item/12345" } },
    }],
  });
  assert.equal(out.length, 1);
  assert.equal(out[0].source, "reverb");
  assert.equal(out[0].price, 2100);
  assert.equal(out[0].seller, "Some Shop");
  assert.equal(out[0].url, "https://reverb.com/item/12345");
});

// A 200 carrying an error body must not read as "no results" — the Bandcamp
// lesson in CLAUDE.md, applied to an authenticated endpoint.
test("a reverb error body throws instead of returning zero listings", async () => {
  const { parseReverbListings } = await import("./parsers.ts");
  assert.throws(() => parseReverbListings({ message: "Invalid token" }), /unexpected payload/);
  assert.deepEqual(parseReverbListings({ listings: [] }), []);
});

// ── evidence against ────────────────────────────────────────────────────────
// Both added 2026-08-18 after the first Reverb run: a $6,800 CDJ-3000 (3x market)
// scored the same as a plausible one, and dealer inventory is the bulk of Reverb.
test("priced well above market counts against a listing", () => {
  const r = scoreListing(listing({ price: 6800, posted_at: null }), CDJ, ALL, CFG)!;
  assert.equal(r.breakdown.price_above_market, -10);
  assert.equal(r.score, 25, "35 base minus 10");
});

test("an established dealer is de-emphasised, not gated out", () => {
  const l = listing({ posted_at: null, established_dealer: true });
  const r = scoreListing(l, CDJ, ALL, CFG)!;
  assert.equal(r.breakdown.established_dealer, -10);
  assert.equal(r.score, 25);
  assert.ok(r.score > 0, "still stored for triage, never dropped");
});

test("a score can never go negative", () => {
  const l = listing({ price: 99999, posted_at: null, established_dealer: true });
  const r = scoreListing(l, CDJ, ALL, CFG)!;
  assert.ok(r.score >= 0);
});

test("reverb listings carry no location and never score local", async () => {
  const { parseReverbListings } = await import("./parsers.ts");
  const out = parseReverbListings({ listings: [{
    id: 1, title: "Pioneer CDJ-3000", price: { amount: "2100.00", currency: "USD" },
    shop_name: "Robyn's Gear Garage", shop: { preferred_seller: false },
    shipping: { local: true }, published_at: "2026-08-17T10:00:00Z",
  }] });
  assert.equal(out[0].location, null, "Reverb exposes no seller location — never invent one");
  assert.equal(out[0].seller, "Robyn's Gear Garage", "seller comes from shop_name, not shop.name");
  assert.equal(out[0].local_pickup, true);
  const r = scoreListing(out[0], CDJ, ALL, CFG)!;
  assert.equal(r.breakdown.local, undefined);
  assert.equal(r.breakdown.note_local_pickup_offered, "yes", "context only, worth no points");
});
