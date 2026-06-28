// pricing-engine.js — pure pricing logic + industry-standard defaults for the
// Come With pricing tool. Imported by dashboard.html AND by scripts/test_pricing.mjs
// (single source of truth — no duplicated math).
//
// Defaults reflect 2025–26 market rates for a SMALL DJ / production company:
//  • DJ fees by tier (club/bar set → experienced private → premium/wedding); club
//    hourly ~$50–200/hr, national wedding-DJ avg ~$1.7k, premium $3k+ (sources in
//    the build notes). We sit a small Brooklyn house/disco company in that band.
//  • Equipment = the company's own daily_rate from equipment_inventory; unpriced
//    gear is suggested at ~10% of purchase price (industry "1/10th" rule of thumb).
//    Multi-day at 50% per extra day (rule-of-thirds style discounting).
//  • Labor: AV tech ~$65/hr, 10-hr "day" = $600, half-day $350, OT 1.5× after 10.
//  • Delivery: base + $0.50/mile beyond a free radius; flat setup/strike.
//  • Surcharges: weekend / peak-season / rush; deposit 50%.
// Every number here is a DEFAULT — the dashboard lets you edit any of them and
// override per-DJ rates; saved overrides are merged over these at load time.

export const PRICING_DEFAULTS = {
  dj: {
    tiers: [
      { key: 'resident', label: 'Resident / club & bar set', rate: 400, hours: 3 },
      { key: 'standard', label: 'Standard / experienced (private)', rate: 750, hours: 4 },
      { key: 'premium', label: 'Premium / headliner & weddings', rate: 1500, hours: 5 },
    ],
    hourly_rate: 150,
    min_hours: 2,
    extra_hour_rate: 150,
    addons: [
      { key: 'mc', label: 'MC / hosting', amount: 150 },
      { key: 'ceremony', label: 'Ceremony / 2nd setup', amount: 300 },
      { key: 'planning', label: 'Planning & consult', amount: 150 },
    ],
  },
  dj_overrides: {}, // actorId -> custom per-event rate
  rental: { multi_day_factor: 0.5, suggest_rule: 0.10, damage_waiver_pct: 8 },
  labor: { tech_hourly: 65, day_rate: 600, day_hours: 10, half_day: 350, ot_multiplier: 1.5 },
  delivery: { base: 75, per_mile: 0.5, free_radius_miles: 15, setup_strike: 150 },
  lighting: {
    tiers: [
      { key: 'none', label: 'None', amount: 0 },
      { key: 'basic', label: 'Basic wash + uplights', amount: 500 },
      { key: 'standard', label: 'Standard (movers + control)', amount: 999 },
      { key: 'large', label: 'Large (truss + FX)', amount: 2000 },
    ],
  },
  modifiers: { weekend_pct: 15, peak_season_pct: 20, rush_pct: 15, deposit_pct: 50 },
};

const num = (v, d = 0) => { const n = Number(v); return Number.isFinite(n) ? n : d; };
const round2 = n => Math.round(n * 100) / 100;

// Deep-merge saved config over the defaults (objects merge; arrays/scalars replace).
export function mergeConfig(saved) {
  const out = JSON.parse(JSON.stringify(PRICING_DEFAULTS));
  const merge = (a, b) => {
    if (!b || typeof b !== 'object' || Array.isArray(b)) return b;
    for (const k of Object.keys(b)) {
      a[k] = (a[k] && typeof a[k] === 'object' && !Array.isArray(a[k]) && b[k] && typeof b[k] === 'object' && !Array.isArray(b[k]))
        ? merge(a[k], b[k]) : b[k];
    }
    return a;
  };
  return merge(out, saved || {});
}

// Suggested daily rental rate for gear with no rate set (10% of purchase price).
export function suggestDailyRate(purchasePrice, cfg = PRICING_DEFAULTS) {
  const p = num(purchasePrice);
  if (p <= 0) return 0;
  return Math.round(p * num(cfg.rental.suggest_rule, 0.10));
}

// Rental line total: daily × qty, with extra days discounted by multi_day_factor.
export function rentalLineTotal(daily, qty, days, cfg = PRICING_DEFAULTS) {
  const d = Math.max(1, num(days, 1));
  const mult = 1 + (d - 1) * num(cfg.rental.multi_day_factor, 0.5);
  return round2(num(daily) * Math.max(1, num(qty, 1)) * mult);
}

// Labor line total by mode: day | half | hourly (hourly applies OT after day_hours).
export function laborLineTotal(line, cfg = PRICING_DEFAULTS) {
  const L = cfg.labor, count = Math.max(1, num(line.count, 1));
  if (line.mode === 'day') return round2(num(L.day_rate) * count);
  if (line.mode === 'half') return round2(num(L.half_day) * count);
  const hrs = num(line.hours), dh = num(L.day_hours, 10);
  const base = Math.min(hrs, dh) * num(L.tech_hourly);
  const ot = Math.max(0, hrs - dh) * num(L.tech_hourly) * num(L.ot_multiplier, 1.5);
  return round2((base + ot) * count);
}

// Master quote calculator. Returns line items + subtotal, surcharges, discount,
// total, deposit. Pure — same inputs always give the same numbers.
export function computeQuote(cfg, q) {
  cfg = cfg || PRICING_DEFAULTS;
  q = q || {};
  const lines = [];

  // DJ services — each line carries a precomputed fee + optional add-ons.
  let djSub = 0;
  for (const dj of (q.djs || [])) {
    const fee = num(dj.fee);
    const addons = (dj.addons || []).reduce((s, a) => s + num(a.amount), 0);
    const amt = round2(fee + addons);
    djSub += amt;
    lines.push({ group: 'DJ', label: dj.label || 'DJ', amount: amt });
  }

  // Equipment rental.
  let rentSub = 0;
  for (const r of (q.equipment || [])) {
    const amt = rentalLineTotal(r.daily, r.qty, r.days, cfg);
    rentSub += amt;
    lines.push({ group: 'Rental', label: r.label || 'Equipment', amount: amt });
  }
  if (q.damageWaiver && rentSub > 0) {
    const w = round2(rentSub * num(cfg.rental.damage_waiver_pct, 8) / 100);
    rentSub += w;
    lines.push({ group: 'Rental', label: `Damage waiver (${num(cfg.rental.damage_waiver_pct, 8)}%)`, amount: w });
  }

  // Labor / techs.
  for (const t of (q.techs || [])) {
    const amt = laborLineTotal(t, cfg);
    lines.push({ group: 'Labor', label: t.label || 'Technician', amount: amt });
  }

  // Lighting (a single chosen tier amount).
  if (num(q.lighting) > 0) lines.push({ group: 'Production', label: q.lightingLabel || 'Lighting', amount: round2(num(q.lighting)) });

  // Delivery + setup/strike.
  if (q.delivery && q.delivery.include) {
    const D = cfg.delivery, miles = num(q.delivery.miles);
    const dAmt = round2(num(D.base) + Math.max(0, miles - num(D.free_radius_miles)) * num(D.per_mile));
    lines.push({ group: 'Production', label: `Delivery${miles ? ` (${miles} mi)` : ''}`, amount: dAmt });
  }
  if (q.setupStrike) lines.push({ group: 'Production', label: 'Setup & strike', amount: round2(num(q.setupStrike)) });

  const subtotal = round2(lines.reduce((s, l) => s + l.amount, 0));

  // Surcharges (% of subtotal), shown as their own lines.
  const M = cfg.modifiers, surLines = [];
  const addSur = (on, pct, label) => { if (on && pct) { const a = round2(subtotal * num(pct) / 100); surLines.push({ label: `${label} (+${num(pct)}%)`, amount: a }); } };
  addSur(q.weekend, M.weekend_pct, 'Weekend');
  addSur(q.peak, M.peak_season_pct, 'Peak season');
  addSur(q.rush, M.rush_pct, 'Rush');
  const surcharge = round2(surLines.reduce((s, l) => s + l.amount, 0));

  // Discount: flat amount + percent of subtotal.
  const discount = round2(num(q.discountAmt) + subtotal * num(q.discountPct) / 100);

  const total = round2(subtotal + surcharge - discount);
  const depositPct = num(q.depositPct, num(M.deposit_pct, 50));
  const deposit = round2(total * depositPct / 100);

  return { lines, subtotal, surLines, surcharge, discount, total, depositPct, deposit, djSub: round2(djSub), rentSub: round2(rentSub) };
}
