// Scenario tests for the pricing engine. Run: node scripts/test_pricing.mjs
import { PRICING_DEFAULTS, computeQuote, mergeConfig, suggestDailyRate, rentalLineTotal, laborLineTotal } from '../assets/pricing-engine.js';

let pass = 0, fail = 0;
const approx = (a, b) => Math.abs(a - b) < 0.011;
function check(name, got, want) {
  const ok = approx(got, want);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}: got ${got}${ok ? '' : `  want ${want}`}`);
  ok ? pass++ : fail++;
}
const cfg = mergeConfig({});

// --- unit checks ---
check('suggestDailyRate(2922) = 292', suggestDailyRate(2922, cfg), 292);
check('rental 2 days @440 = 660 (1.5x)', rentalLineTotal(440, 1, 2, cfg), 660);
check('labor hourly 12h x1 (10@65 + 2@97.5)', laborLineTotal({ mode: 'hourly', hours: 12, count: 1 }, cfg), 650 + 195);
check('labor day x2 = 1200', laborLineTotal({ mode: 'day', count: 2 }, cfg), 1200);

// --- Scenario 1: DJ-only (resident $400 + MC $150), weekend +15%, 50% deposit ---
const s1 = computeQuote(cfg, {
  djs: [{ label: 'Berky — resident', fee: 400, addons: [{ label: 'MC', amount: 150 }] }],
  weekend: true, depositPct: 50,
});
check('S1 subtotal', s1.subtotal, 550);
check('S1 weekend surcharge', s1.surcharge, 82.5);
check('S1 total', s1.total, 632.5);
check('S1 deposit', s1.deposit, 316.25);

// --- Scenario 2: Rental (CDJ pair $440×2d, Wave8 pair $132×2d) + 8% waiver ---
const s2 = computeQuote(cfg, {
  equipment: [
    { label: 'CDJ-3000 pair', daily: 440, qty: 1, days: 2 },
    { label: 'Wave 8 pair', daily: 132, qty: 1, days: 2 },
  ],
  damageWaiver: true, depositPct: 50,
});
check('S2 rental subtotal (660+198+8% waiver)', s2.subtotal, 926.64);
check('S2 total', s2.total, 926.64);

// --- Scenario 3: Full production ---
const s3 = computeQuote(cfg, {
  djs: [{ label: 'Premium DJ', fee: 1500, addons: [{ label: 'Planning', amount: 150 }] }],
  equipment: [{ label: 'CDJ-3000 pair', daily: 440, qty: 1, days: 1 }],
  techs: [{ label: 'AV tech', mode: 'day', count: 2 }],
  lighting: 999, lightingLabel: 'Standard lighting',
  delivery: { include: true, miles: 25 },
  setupStrike: 150,
  weekend: true, peak: true, depositPct: 50,
});
check('S3 subtotal', s3.subtotal, 1650 + 440 + 1200 + 999 + 80 + 150);
check('S3 surcharge (35%)', s3.surcharge, (1650 + 440 + 1200 + 999 + 80 + 150) * 0.35);
check('S3 deposit 50%', s3.deposit, s3.total / 2);
console.log('\nS3 line items:'); s3.lines.forEach(l => console.log(`   [${l.group}] ${l.label}: $${l.amount}`));
console.log(`   subtotal $${s3.subtotal} + surcharge $${s3.surcharge} = total $${s3.total}, deposit $${s3.deposit}`);

console.log(`\n${fail ? 'FAILED' : 'ALL PASS'} — ${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
