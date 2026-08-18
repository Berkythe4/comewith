// parsers.ts — pure parsing for the sources that don't hand back friendly JSON.
//
// Split out of index.ts for the same reason scoring.ts was: a pure function over
// a payload can be run against the LIVE endpoint from a plain node script, with
// no Supabase, no credentials and no deploy (scripts/gearwatch_live_test.mjs).
//
// ── Craigslist ──────────────────────────────────────────────────────────────
// Craigslist has no public API and its search **RSS** is dead: every
// `format=rss` path answers HTTP 403 "Your request has been blocked" (verified
// 2026-08-18), and the HTML search page is a JS shell with no listings in it.
//
// But the site's own search box calls an internal JSON endpoint —
// `sapi.craigslist.org/web/v8/postings/search/full` — and that answers 200.
// Same approach `track-sources` already uses for Bandcamp: call the endpoint
// their own front-end calls. Verified 2026-08-18 returning 47 live NYC results
// for "cdj", with reconstructed posting URLs resolving 200.
//
// The payload is delta-encoded to keep it small, which is the only hard part:
//
//   item[0]  posting-id delta, CUMULATIVE across items from decode.minPostingId
//   item[1]  posted-at, in seconds AFTER decode.minPostedDate
//   item[3]  price as a number — **-1 means "no price given"**, not free
//   item[n]  a "geo" string "<locationIdx>:<descriptionIdx>~<lat>~<lon>"
//            TWO different indexes into TWO different arrays. Using the first
//            for both puts a Schenectady listing in Bushwick (seen 2026-08-18).
//   [6, s]   the URL slug
//   [13, t]  the posting token — the modern canonical URL is
//            https://www.craigslist.org/view/d/<slug>/<token>
//   last     the title
//
// decode.locations[locationIdx]        -> [regionId, "newyork", "brk"]
// decode.locationDescriptions[descIdx] -> "brooklyn"
//
// On a query with NO results, `decode` comes back as the NUMBER 0 rather than an
// object. An empty result set is a legitimate answer and must return [], not
// throw — reporting "nothing listed" as "we were blocked" is as wrong as the
// reverse, and this endpoint does both if you let it.

import type { Listing } from "./scoring.ts";

// ── Reverb ──────────────────────────────────────────────────────────────────
// Official API. Personal access tokens do NOT expire (unlike Beatport's
// 600-second ones), so this is a set-once secret. Scopes needed: public,
// read_listings. Docs: reverb-api.com/docs/authentication

export const reverbSearchUrl = (query: string, perPage = 50) =>
  `https://api.reverb.com/api/listings?query=${encodeURIComponent(query)}&per_page=${perPage}&item_region=US`;

export const reverbHeaders = (token: string) => ({
  "Authorization": `Bearer ${token}`,
  "Accept": "application/hal+json",
  "Accept-Version": "3.0",
  "Content-Type": "application/hal+json",
});

export const reverbDetailUrl = (id: string) => `https://api.reverb.com/api/listings/${id}`;

/**
 * Merges a Reverb LISTING DETAIL response into a search-result listing.
 *
 * The search endpoint deliberately omits seller location and feedback; the
 * detail endpoint carries both:
 *   location: {region, locality, country_code, display_location}
 *   shop:     {feedback_count, rating_percentage, preferred_seller}
 *
 * Without this, every Reverb hit was location-unknown and could never score the
 * local bonus — the largest source in the scan, structurally blind to the
 * strongest non-serial signal we have. Verified against the live API 2026-08-18.
 */
export function mergeReverbDetail(l: Listing, detail: unknown): Listing {
  const d = detail as Record<string, any>;
  if (!d || typeof d !== "object") return l;
  const fb = Number(d.shop?.feedback_count);
  return {
    ...l,
    location: d.location?.display_location || l.location,
    seller_feedback: Number.isFinite(fb) ? fb : l.seller_feedback,
    local_pickup: d.local_pickup_only === true || l.local_pickup,
    established_dealer: d.shop?.preferred_seller === true || l.established_dealer,
  };
}

/**
 * Maps a Reverb listings payload. Throws on anything that isn't a listings
 * response — a 200 carrying an error body must not read as "no results".
 */
export function parseReverbListings(payload: unknown): Listing[] {
  const j = payload as Record<string, any>;
  if (!Array.isArray(j?.listings)) {
    const m = j?.message || j?.error;
    throw new Error("Reverb: unexpected payload shape" + (m ? ` — ${String(m).slice(0, 120)}` : ""));
  }
  return j.listings.map((x: Record<string, any>) => ({
    source: "reverb" as const,
    listing_id: String(x.id),
    url: x._links?.web?.href || `https://reverb.com/item/${x.id}`,
    title: x.title || "",
    price: x.price?.amount != null ? Number(x.price.amount) : null,
    currency: x.price?.currency || "USD",
    // Reverb's search payload carries NO seller location — verified against the
    // live API 2026-08-18: no `location`, and `shop` holds only {slug,
    // preferred_seller}. The nearest thing is `shipping.local` (pickup offered),
    // which says nothing about WHERE. So location is honestly null and Reverb
    // hits never score the local bonus. Do not synthesise one (§26).
    location: null,
    seller: x.shop_name || null,
    seller_feedback: null,
    posted_at: x.published_at || x.created_at || null,
    image_url: x.photos?.[0]?._links?.large_crop?.href || null,
    description: x.description || null,
    // Reverb search is mostly dealer inventory. `preferred_seller` is Reverb's
    // badge for established high-volume shops — used to de-emphasise them, since
    // a fence is a private seller, not a shop with a return policy.
    established_dealer: x.shop?.preferred_seller === true,
    local_pickup: x.shipping?.local === true,
    raw: x,
  }));
}

/** NYC. The leading number in `batch` is the Craigslist area id. */
export const CRAIGSLIST_NYC_AREA = 17;

export const craigslistSapiUrl = (query: string, areaId = CRAIGSLIST_NYC_AREA) =>
  `https://sapi.craigslist.org/web/v8/postings/search/full` +
  `?batch=${areaId}-0-360-0-0&cc=US&lang=en&searchPath=sss&query=${encodeURIComponent(query)}`;

export const CRAIGSLIST_HEADERS = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
  "Accept": "*/*",
  "Referer": "https://newyork.craigslist.org/",
};

const tagged = (item: unknown[], tag: number): unknown[] | undefined =>
  item.find((x) => Array.isArray(x) && (x as unknown[])[0] === tag) as unknown[] | undefined;

/**
 * Decodes a Craigslist sapi search payload into listings.
 *
 * Throws when the body isn't a recognisable sapi response — a block page, an
 * error envelope or a shape change must surface as a FAILED source, never as
 * zero results. "Nothing found" and "we were blocked" are different answers.
 */
export function parseCraigslistSapi(payload: unknown): Listing[] {
  const j = payload as Record<string, any>;
  const d = j?.data;
  if (!d || !Array.isArray(d.items)) {
    const e = j?.errors?.[0]?.message;
    throw new Error("Craigslist: unexpected payload shape" + (e ? ` — ${String(e).slice(0, 120)}` : ""));
  }
  // No results: `decode` is 0, and that is a real answer, not a failure.
  if (d.items.length === 0) return [];
  if (!d.decode || typeof d.decode !== "object") {
    throw new Error("Craigslist: items present but decode table missing");
  }

  const minId = Number(d.decode.minPostingId) || 0;
  const minPosted = Number(d.decode.minPostedDate) || 0;
  const locations: any[] = Array.isArray(d.decode.locations) ? d.decode.locations : [];
  const hoods: any[] = Array.isArray(d.decode.locationDescriptions) ? d.decode.locationDescriptions : [];

  const out: Listing[] = [];
  let runningId = minId;

  for (const raw of d.items) {
    if (!Array.isArray(raw)) continue;
    const item = raw as unknown[];

    runningId += Number(item[0]) || 0;

    const slug = tagged(item, 6)?.[1] as string | undefined;
    const token = tagged(item, 13)?.[1] as string | undefined;
    const title = typeof item[item.length - 1] === "string" ? String(item[item.length - 1]) : "";
    if (!title || !slug || !token) continue;

    // -1 is Craigslist's "no price stated". Storing it as a number would make
    // every unpriced listing look like the cheapest thing on the board and hand
    // it the price-anomaly bonus.
    const rawPrice = Number(item[3]);
    const price = Number.isFinite(rawPrice) && rawPrice > 0 ? rawPrice : null;

    // "<locationIdx>:<descriptionIdx>~lat~lon" — two indexes, two arrays.
    const geo = item.find((x) => typeof x === "string" && (x as string).includes("~")) as string | undefined;
    const head = geo ? geo.split("~")[0] : "";
    const locIdx = Number(head.split(":")[0]);
    const descIdx = Number(head.split(":")[1]);
    const loc = Number.isFinite(locIdx) ? locations[locIdx] : undefined;
    const hood = Number.isFinite(descIdx) ? hoods[descIdx] : undefined;
    // Never fall back to "New York": the NYC board carries nearby-area results
    // from New Jersey, Connecticut and beyond, and a defaulted location would
    // hand every one of them the local bonus. Unknown scores nothing.
    const location = [typeof hood === "string" ? hood : null, Array.isArray(loc) ? loc[1] : null]
      .filter(Boolean).join(", ") || null;

    const postedDelta = Number(item[1]);
    const posted_at = minPosted && Number.isFinite(postedDelta)
      ? new Date((minPosted + postedDelta) * 1000).toISOString()
      : null;

    out.push({
      source: "craigslist",
      listing_id: String(runningId),
      url: `https://www.craigslist.org/view/d/${slug}/${token}`,
      title,
      price,
      currency: "USD",
      location,
      seller: null,
      seller_feedback: null,
      posted_at,
      image_url: null,
      description: null,
      raw: { id: runningId, slug, token },
    });
  }
  return out;
}
