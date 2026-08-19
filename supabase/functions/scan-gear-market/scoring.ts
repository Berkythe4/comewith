// scoring.ts — the Gear Watch confidence model.
//
// Split out of index.ts on purpose: this is the part that decides what reaches
// Keith's phone, and it is the only part testable without live API credentials.
// See scoring.test.ts (`node --test supabase/functions/scan-gear-market/`).
//
// Shape of the model:
//   GATES   — "this is not our gear". The listing is dropped, not stored.
//   SIGNALS — weighted points, capped at 100, with every award recorded in the
//             breakdown so a hit can be explained to a detective rather than
//             asserted. "87/100" means nothing on its own; "serial match +60,
//             local +20" is something a person can act on.

export type Target = {
  id: string;
  label: string;
  make: string | null;
  model_tokens: string[];
  exclude_tokens: string[];
  serial: string | null;
  qty: number;
  typical_resale: number | null;
};

export type Listing = {
  source: "reverb" | "ebay" | "craigslist" | "facebook";
  listing_id: string;
  url: string;
  title: string;
  price: number | null;
  currency: string;
  location: string | null;
  seller: string | null;
  seller_feedback: number | null;
  posted_at: string | null;
  image_url: string | null;
  description?: string | null;
  /** Reverb only: an established shop with Reverb's preferred-seller badge. */
  established_dealer?: boolean;
  /** Reverb only: seller offers local pickup. Context for triage, not scored. */
  local_pickup?: boolean;
  raw?: unknown;
};

export type ScoreResult = { score: number; breakdown: Record<string, number | string> };

// "CDJ-3000", "cdj 3000" and "CDJ3000" are one model to a seller and three
// different strings to a computer. Compare the spaced form AND a squashed form
// so no spelling slips past. Unicode dashes are folded to ASCII first — a title
// pasted from a store page often carries an en dash.
export const norm = (s: string) =>
  (s || "").toLowerCase().replace(/[‐-―]/g, "-").replace(/\s+/g, " ").trim();

export const squash = (s: string) => norm(s).replace(/[^a-z0-9]/g, "");

export function hasToken(hay: string, token: string): boolean {
  const h = norm(hay), t = norm(token);
  if (!t) return false;
  if (h.includes(t)) return true;
  return squash(hay).includes(squash(token));
}

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Word-boundary match, so "ny" hits "Brooklyn, NY" and misses "albany". */
export function isLocalTerm(location: string, term: string): boolean {
  const t = norm(term);
  if (!t) return false;
  return new RegExp(`(^|[^a-z0-9])${escapeRe(t)}([^a-z0-9]|$)`, "i").test(norm(location));
}

export function scoreListing(
  l: Listing,
  t: Target,
  allTargets: Target[],
  cfg: { theft_date: string | null; geo_terms: string[] },
): ScoreResult | null {
  const text = `${l.title} ${l.description || ""}`;
  const b: Record<string, number | string> = {};

  // GATE 1 — the model has to be named somewhere in the listing.
  if (!t.model_tokens.some((tok) => hasToken(text, tok))) return null;

  // GATE 2 — accessories. "Hard case for CDJ-3000" names the model and is not
  // the model. Checked against the TITLE only: half the legitimate listings on
  // Reverb mention a case in the description ("comes with case"), and gating on
  // the description would throw away real hits.
  if (t.exclude_tokens.some((tok) => hasToken(l.title, tok))) return null;

  // GATE 3 — a listing that went up before the theft is not the stolen unit.
  // Only applies when the source gave us a date; absent a date we keep the
  // listing rather than guess.
  if (cfg.theft_date && l.posted_at) {
    if (new Date(l.posted_at) < new Date(cfg.theft_date + "T00:00:00Z")) return null;
  }

  // Matching one of a handful of specific stolen models is already meaningful —
  // this is not a keyword search over all DJ gear. Recalibrated 2026-08-18 from
  // 25: the old base made a CDJ-3000 in Brooklyn listed the day after the theft
  // come out at 55/100, which read as "probably nothing" when it is the exact
  // shape of a real hit.
  let score = 35;
  b.model_match = 35;

  // Serial number. Dormant until serials are recovered — decisive the moment
  // they are, which is why recovering them is worth more than any other input.
  if (t.serial && t.serial.length >= 5 && hasToken(text, t.serial)) {
    b.serial_match = 60;
    score += 60;
  }

  // Local. Stolen gear moves locally far more often than it ships.
  // Matched on WORD BOUNDARIES, not substrings: a plain `includes("ny")` scores
  // "albany" as New York, which is how a Schenectady listing collected the local
  // bonus on 2026-08-18. Short terms are exactly the dangerous ones.
  const loc = norm(l.location || "");
  if (loc && cfg.geo_terms.some((g) => g && isLocalTerm(loc, g))) {
    b.local = 25;
    score += 25;
  }

  // Bundle — the strongest non-serial signal available. The rig is distinctive:
  // XDJ-AZ + two CDJ-3000s + two Wave 8s. Two different stolen models in one
  // listing is close to a fingerprint; a "pair" of something we lost a pair of
  // is the weaker cousin of the same idea.
  const alsoMatches = allTargets
    .filter((o) => o.id !== t.id)
    .some((o) => o.model_tokens.some((tok) => hasToken(text, tok)));
  if (alsoMatches) {
    b.bundle_multi_model = 30;
    score += 30;
  } else if (t.qty > 1 && /(\bpair\b|\btwo\b|\b2x\b|\bx2\b|\(2\)|\bboth\b)/i.test(text)) {
    b.bundle_pair = 18;
    score += 18;
  }

  // Priced to move — and priced NOT to move.
  if (l.price != null && t.typical_resale) {
    const ratio = l.price / Number(t.typical_resale);
    if (ratio < 0.6) { b.price_far_below_market = 15; score += 15; }
    else if (ratio < 0.8) { b.price_below_market = 8; score += 8; }
    // Well above market is evidence AGAINST: a fence prices to move quickly, and
    // a 3x-market listing is a dealer's retail price or a multi-unit bundle. Seen
    // live on Reverb 2026-08-18 — a $6,800 CDJ-3000 scoring the same as a
    // plausible one purely on model and recency.
    else if (ratio > 1.5) { b.price_above_market = -10; score -= 10; }
    b.price_ratio = Math.round(ratio * 100) / 100;
  }

  // Established dealers are the bulk of Reverb's inventory and the least likely
  // route for stolen goods — a shop with a storefront, a return policy and a
  // reputation to lose. A de-noiser, not a verdict: it lowers a score, it never
  // gates a listing out.
  if (l.established_dealer) {
    b.established_dealer = -10;
    score -= 10;
  }

  // Listed soon after the theft — GRADUATED, not a flat bonus. Stolen gear is
  // flipped fast: something appearing the day after the break-in is a far
  // stronger signal than the same listing three weeks later, and a flat +10 for
  // anything inside 30 days threw that away.
  if (cfg.theft_date && l.posted_at) {
    const days = (new Date(l.posted_at).getTime() - new Date(cfg.theft_date + "T00:00:00Z").getTime()) / 86400000;
    const pts = days < 0 ? 0
      : days <= 3 ? 25
      : days <= 7 ? 20
      : days <= 14 ? 15
      : days <= 30 ? 10
      : days <= 60 ? 5
      : 0;
    if (pts) { b.listed_soon_after_theft = pts; score += pts; }
    b.days_after_theft = Math.round(days);
  }

  // Thin seller history.
  if (l.seller_feedback != null && l.seller_feedback < 10) {
    b.new_seller = 10;
    score += 10;
  }

  // Phrasing common on gear the seller can't document. Weak individually,
  // which is why it is capped at 10 — it should never carry a hit on its own.
  const phrases = [/no box/i, /no receipt/i, /cash only/i, /must go/i, /moving/i, /firm on price/i, /no returns/i];
  const n = phrases.filter((re) => re.test(text)).length;
  if (n) { const pts = Math.min(n * 5, 10); b.listing_language = pts; score += pts; }

  // Context for triage — carried in the breakdown, never scored, so it can't
  // become invented evidence.
  if (l.local_pickup) b.note_local_pickup_offered = "yes";

  return { score: Math.max(0, Math.min(score, 100)), breakdown: b };
}

// Picks the best-scoring target for one listing. Returns null when the listing
// matches nothing — the common case, and the reason the hits table stays small.
export function bestMatch(
  l: Listing,
  targets: Target[],
  cfg: { theft_date: string | null; geo_terms: string[] },
): { target: Target; score: number; breakdown: Record<string, number | string> } | null {
  let best: { target: Target; score: number; breakdown: Record<string, number | string> } | null = null;
  for (const t of targets) {
    const s = scoreListing(l, t, targets, cfg);
    if (s && (!best || s.score > best.score)) best = { target: t, score: s.score, breakdown: s.breakdown };
  }
  return best;
}
