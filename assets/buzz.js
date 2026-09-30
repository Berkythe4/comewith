// Buzz — the ONE copy of the formula. Imported by dashboard.html (Come With Radio /
// NYC Scene) and dj.html (the guest-DJ crate), so a DJ ranks artists exactly the way
// the dashboard does. Moved here verbatim on 2026-09-30; before that it lived inline
// in dashboard.html and the DJ page had no Buzz at all. Change it here only.
//
// Input (all on `c`): topPlays, topTrack, reach, demand, demandKnown, songCount,
// scanOk, raFollowers, editorial, recent. Output: c.buzz (null = nothing measurable),
// c.coverage, and the per-input s*/k* fields.

// ── Buzz ────────────────────────────────────────────────────────────────────
// Four inputs, each mapped 0–100 against a FIXED anchor, weighted, and — the
// part that matters — dropped entirely when the input is unknown, with the
// remaining weights renormalised over what was actually measurable.
//
// WHY FIXED ANCHORS. The old score scaled every input against the current pool's
// maximum, so an artist's buzz moved when somebody else was scanned or a bigger
// name entered the list. "72 last week, 68 now" said nothing about the artist.
// Anchors are absolute, so the number means the same thing every week — the same
// reason v_kpi_prior refuses to compare a metric against the latest reading.
//
// WHY UNKNOWN IS NOT ZERO. Scoring a missing input as 0 is invented evidence
// (§26): it says "we measured this and it was nothing" when nothing was measured.
// It penalised two whole groups invisibly — artists whose shows are only on DICE
// or Ticketmaster (neither publishes RSVPs), and artists whose SoundCloud scan
// failed. Both now shrink the DENOMINATOR instead of the numerator, and the
// coverage they were scored on is shown next to the number.
export const BUZZ_W = { plays: 0.30, reach: 0.30, demand: 0.25, catalog: 0.15 };
// lo scores 0, hi scores 100, log-spaced between. Chosen off the real prod
// spread: top-track plays run p50 18k / p90 900k / p99 13.5M, followers
// p50 1.5k / p90 44k / p99 445k, RA attending has a median of 5.
export const BUZZ_ANCHOR = {
  plays:   [1000, 10000000],
  reach:   [100, 1000000],
  demand:  [2, 1000],
  catalog: [1, 200],
};
export function buzzAnchor(kind, x) {
  const [lo, hi] = BUZZ_ANCHOR[kind];
  const v = Number(x) || 0;
  if (v <= lo) return 0;
  return Math.max(0, Math.min(100, Math.round(100 * Math.log(v / lo) / Math.log(hi / lo))));
}
export function scScoreArtist(c) {
  c.sPlays = buzzAnchor('plays', c.topPlays);
  c.sReach = buzzAnchor('reach', c.reach);
  c.sDemand = buzzAnchor('demand', c.demand);
  c.sCatalog = buzzAnchor('catalog', c.songCount);
  // A successful scan makes the CATALOGUE known even at zero — "we looked, they
  // upload nothing" is a real answer, and for a radio show that plays records it
  // is the answer that matters.
  //
  // PLAYS is different: the plays of a catalogue that does not exist is not zero,
  // it is undefined, and scoring it 0 punishes an artist twice for one fact. It
  // buries precisely the people buzz should surface — the selector DJs who play
  // other artists' records and upload none of their own. Measured on prod: Ben UFO
  // (109k followers, 2,010 RSVPs, no uploads) scores 56 if the missing plays count
  // as zero and 76 if they are left out. Craig Richards, Joseph Capriati and Jyoty
  // all move the same way. 393 scanned artists have no tracks at all.
  c.kPlays = c.scanOk && c.topTrack != null;
  c.kCatalog = c.scanOk;
  c.kReach = c.scanOk || c.raFollowers > 0;
  c.kDemand = !!c.demandKnown;
  const parts = [
    ['plays', c.kPlays, c.sPlays], ['reach', c.kReach, c.sReach],
    ['demand', c.kDemand, c.sDemand], ['catalog', c.kCatalog, c.sCatalog],
  ].filter(p => p[1]);
  const wsum = parts.reduce((a, p) => a + BUZZ_W[p[0]], 0);
  c.coverage = Math.round(wsum * 100);
  if (!wsum) {
    // Nothing measurable at all. A 0 here would rank this artist below one we
    // know to be small, which is a claim the data does not support.
    c.buzz = null;
    return;
  }
  let s = parts.reduce((a, p) => a + BUZZ_W[p[0]] * p[2], 0) / wsum;
  if (c.editorial) s += 8;   // RA editorial pick
  if (c.recent) s += 5;      // released something in the last 6 months
  c.buzz = Math.max(0, Math.min(100, Math.round(s)));
}
export const buzzSort = b => (b == null ? -1 : b);   // unknown ranks below a known zero
