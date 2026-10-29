// Tests for assets/control-center.js. Pure logic, plus the real DI #2 RA and
// MS Society DIY exports where this machine has them (they hold guest data and
// are not in git; those blocks skip elsewhere).
// Run: node scripts/test_control_center.mjs
import { readFileSync, existsSync } from 'node:fs';
import {
  parseCsv, parseCsvRaw, controlFor, tierSplit, matchTier, ticketStats, donationStats, sponsorStats,
  missionMath, drinkLedger, doorStats, audience, winbackList, scenario, parseZeffy, parseRa, parseDiy,
  planImport, execSnapshot, execDiff, daysUntil, toCsv, personIndex,
} from '../assets/control-center.js';

let pass = 0, fail = 0;
const eq = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (ok) pass++; else fail++;
  console.log(`${ok ? '  ok ' : 'FAIL '} ${label}${ok ? '' : `\n       got  ${JSON.stringify(got)}\n       want ${JSON.stringify(want)}`}`);
};

// ---- DI #3 seed shapes -----------------------------------------------------
const DI3 = { id: 'ev3', slug: 'dance-infusion-3', name: 'Dance Infusion #3', series: 'Dance Infusion', event_date: '2026-11-14', bar_minimum: 6000 };
const T25 = { id: 't25', label: '$25 ticket', price: 25, to_mission: 25, drink_tickets: null, aliases: [], active: true };
const T40 = { id: 't40', label: '$40 ticket', price: 40, to_mission: 25, production_label: 'Production support', drink_tickets: null, aliases: [], active: true };
const TIERS = [T25, T40];
const SP = [
  { id: 's1', tier: 'Champion', cash_amount: 500, entry_tickets: 4, drink_tickets: 0, drink_tickets_tbd: true, status: 'pending', actor: { display_name: 'Pulse Devices' } },
  { id: 's2', tier: 'Champion', cash_amount: 500, entry_tickets: 3, drink_tickets: 0, drink_tickets_tbd: true, status: 'pending', actor: { display_name: 'Theresa Berkman' } },
  { id: 's3', tier: 'Champion', cash_amount: 500, entry_tickets: 4, drink_tickets: 0, drink_tickets_tbd: true, status: 'pending', actor: { display_name: 'Patrick Savery' } },
];
const GEA = ['Patrick Savery', 'Arnaldo Martinez', 'Alice Zhang'].map((n, i) => ({ guest_id: 'g' + i, name: n, sponsorship_id: 's3', entry_type: 'sponsor_comp', quantity: 1, drink_tickets: 0 }));

console.log('config');
eq('no row: DI defaults fundraising on, 60% target', [controlFor(DI3, null).fundraising, controlFor(DI3, null).mission_target_pct], [true, 60]);
eq('no row: a party defaults off', controlFor({ series: 'Come With Parties' }, null).fundraising, false);
eq('$40 split: 25 mission / 15 production', tierSplit(T40), { price: 40, toMission: 25, production: 15 });
eq('countdown is NY calendar days (23:30 Oct 29 in NY)', daysUntil('2026-11-14', new Date('2026-10-30T03:30:00Z')), 16);

console.log('tier matching');
eq('"$40 Ticket" -> $40 tier by dollar text', matchTier('$40 Ticket', 40, TIERS).id, 't40');
eq('RA net-of-fee $21.90 -> $25 tier', matchTier('General admission', 21.9, TIERS).id, 't25');
eq('alias wins', matchTier('Guest List', 25, [{ ...T25, aliases: ['guest list'] }, T40]).id, 't25');
eq('$28 is in neither fee band -> null', matchTier('Mystery', 28, TIERS), null);

console.log('money');
const tkRows = [
  { id: 'a', source: 'zeffy', tier_id: 't25', amount_paid: 25, quantity: 1, ticket_type: 'Guest List' },
  { id: 'b', source: 'zeffy', tier_id: 't40', amount_paid: 80, quantity: 2, ticket_type: '$40 Ticket' },
  { id: 'c', source: 'resident_advisor', tier_id: 't25', amount_paid: 21.9, quantity: 1, ticket_type: 'GA' },
  { id: 'd', source: 'ms_diy', tier_id: null, amount_paid: 50, quantity: 1, ticket_type: 'Early Bird' },
  { id: 'e', source: 'resident_advisor', tier_id: null, amount_paid: 30, quantity: 1, ticket_type: 'Weird' },
  { id: 'f', source: 'zeffy', tier_id: 't25', amount_paid: 0, quantity: 1, ticket_type: 'Comp' },
];
const tk = ticketStats(tkRows, TIERS);
eq('qty / revenue', [tk.qty, tk.revenue], [7, 206.9]);
eq('mission: 25 + 50 + 25 (RA face) + 50 (DIY direct) ; unassigned unknown, comp 0', tk.mission, 150);
eq('unassigned ticket type surfaced, not zeroed', tk.unassigned.map(u => [u.ticket_type, u.qty]), [['Weird', 1]]);
const don = donationStats([{ amount: 100, payment_processor: 'Zeffy' }, { amount: 50, payment_processor: 'MS Society DIY' }, { amount: 162.44, payment_processor: 'Founder out-of-pocket' }]);
eq('founder money is raised but not mission', [don.mission, don.other, don.total], [150, 162.44, 312.44]);
const sp = sponsorStats(SP, GEA);
eq('sponsors: pledged / paid / comps', [sp.pledged, sp.paid, sp.compsUsed, sp.compsTotal], [1500, 0, 3, 11]);
eq('Savery: 3 named, 1 open', [sp.cards[2].compsUsed, sp.cards[2].compsOpen], [3, 1]);
const ctl = controlFor(DI3, null);
const mm = missionMath({ tk, don, sp, control: ctl });
eq('to mission / raised', [mm.toMission, mm.raised], [300, 2019.34]);
eq('comp credit only when the event sets it', missionMath({ tk, don, sp, control: { ...ctl, comp_mission_value: 25 } }).compCredit, 275);

console.log('drinks');
let dl = drinkLedger({ sp, tiers: TIERS, tk, gea: GEA, allocations: [], control: ctl, barMinimum: 6000 });
eq('everything TBD: 0 allocated, 5 TBD lines, no bar % (no $ value)', [dl.allocated, dl.tbdLines, dl.barPct, dl.locked], [0, 5, null, null]);
const sp2 = sponsorStats(SP.map(s => ({ ...s, drink_tickets: 20, drink_tickets_tbd: false })), GEA.map((g, i) => ({ ...g, drink_tickets: i === 0 ? 3 : 0 })));
dl = drinkLedger({
  sp: sp2, tiers: [T25, { ...T40, drink_tickets: 2 }], tk,
  gea: [...GEA.map((g, i) => ({ ...g, drink_tickets: i === 0 ? 3 : 0 })), { guest_id: 'dj', entry_type: 'dj', drink_tickets: 5, sponsorship_id: null }],
  allocations: [{ category: 'staff', tickets: 10, label: 'Staff pool' }],
  control: { ...ctl, drink_tickets_locked: 100, drink_ticket_value: 12 }, barMinimum: 6000,
});
eq('sponsor tickets handed to a guest are NOT added on top of the package', dl.lines.filter(l => l.cat === 'sponsor').map(l => l.tickets), [20, 20, 20]);
eq('allocated = 60 sponsor + 4 tier + 5 dj + 10 staff ($25 tier TBD)', dl.allocated, 79);
eq('remaining / overflow', [dl.remaining, dl.overflow], [21, 0]);
eq('bar progress: 79 x $12 of $6,000', Math.round(dl.barPct * 10) / 10, 15.8);
const over = sponsorStats([{ ...SP[0], drink_tickets: 2, drink_tickets_tbd: false }], [{ sponsorship_id: 's1', drink_tickets: 5 }]);
eq('handing out more than the package counts what exists', drinkLedger({ sp: over, tiers: [], tk: ticketStats([], []), gea: [], allocations: [], control: ctl, barMinimum: 6000 }).allocated, 5);

console.log('door');
eq('door stats', doorStats([{ quantity: 2, checked_in_at: 'x', scan_count: 2 }, { quantity: 1 }, { entry_type: 'sponsor_comp', quantity: 1 }]), { rows: 3, people: 4, inside: 2, comps: 1 });
eq('csv quoting', toCsv(['a', 'b'], [['x,y', 'q"z']]), 'a,b\r\n"x,y","q""z"\r\n');

console.log('audience');
const evs = [
  { id: 'di1', name: 'DI #1', event_date: '2025-09-08', series: 'Dance Infusion' },
  { id: 'cw', name: '7-11', event_date: '2026-07-11', series: 'Come With Parties' },
  { id: 'ev3', name: 'DI #3', event_date: '2026-11-14', series: 'Dance Infusion' },
];
const guests = [
  { id: 'A1', full_name: 'Ana Ruiz', email: 'ana@x.com' }, { id: 'A2', full_name: 'ana  ruiz', email: null },
  { id: 'B', full_name: 'Bo Li', email: 'bo@x.com' }, { id: 'C', full_name: 'Cy', email: 'cy@x.com' },
  { id: 'D', full_name: 'Dee', email: 'dee@x.com', opted_in_mailing: false },
];
eq('a name-only row merges into the email row with that name', personIndex(guests).keyOf.get('A2'), 'e:ana@x.com');
const aud = audience({ gea: [{ guest_id: 'A2', event_id: 'di1' }, { guest_id: 'A1', event_id: 'ev3' }, { guest_id: 'B', event_id: 'cw' }, { guest_id: 'C', event_id: 'ev3' }, { guest_id: 'D', event_id: 'di1' }], guests, events: evs, currentEventId: 'ev3' });
eq('returning / new / lapsed', [aud.returning.map(p => p.name), aud.fresh.map(p => p.name), aud.lapsed.map(p => p.name)], [['Ana Ruiz'], ['Cy'], ['Bo Li', 'Dee']]);
eq('series scope drops the party-only person', audience({ gea: [{ guest_id: 'B', event_id: 'cw' }], guests, events: evs, currentEventId: 'ev3', scope: 'series' }).lapsed.length, 0);
const wb = winbackList(aud.lapsed, [{ id: 'sb', email: 'bo@x.com', status: 'unsubscribed' }], guests);
eq('win-back never mails unsubscribed or opted-out', wb.map(p => [p.name, p.status, p.send]), [['Bo Li', 'unsubscribed', false], ['Dee', 'opted out', false]]);

console.log('scenario');
const sc = scenario({ tiers: TIERS, counts: { t25: 100, t40: 50 }, extraDonations: 500, expenses: 6000, sponsorCash: 1500, targetPct: 60 });
eq('to MS = 2500 + 1250 + 500', sc.toMission, 4250);
eq('raised = 2500 + 2000 + 500 + 1500', sc.raised, 6500);
eq('production = 2250, short 3750 vs $6,000', [sc.production, sc.covered], [2250, -3750]);
eq('65.4% meets 60%', [Math.round(sc.pct * 10) / 10, sc.meetsTarget], [65.4, true]);

console.log('parsers');
const z = parseZeffy([
  { 'Payment Date (America/New_York)': '5/9/2026', 'Payment Time (America/New_York)': '9:32 PM', 'Total Amount': 25, 'Payment Status': 'Succeeded', 'First Name': 'Reshad', 'Last Name': 'Alam', Email: 'R@X.com', Details: '1x Guest List', 'Tax Receipt #': 'SI2-1' },
  { 'Payment Date (America/New_York)': '5/9/2026', 'Payment Time (America/New_York)': '8:58 PM', 'Total Amount': 30, 'Payment Status': 'Succeeded', 'First Name': 'Tarik', 'Last Name': 'H', Email: 't@x.com', Details: '1x Donations' },
  { 'Payment Date (America/New_York)': '5/8/2026', 'Payment Time (America/New_York)': '1:00 PM', 'Total Amount': 90, 'Extra Donation': 10, 'Payment Status': 'Succeeded', 'First Name': 'Pat', 'Last Name': 'Q', Email: 'p@x.com', Details: '2x $40 Ticket' },
  { 'Payment Date (America/New_York)': '5/8/2026', 'Payment Time (America/New_York)': '1:00 PM', 'Total Amount': 40, 'Payment Status': 'Incomplete', 'First Name': 'No', 'Last Name': 'Pay', Details: '1x $40 Ticket' },
]);
eq('zeffy kinds', z.sales.map(s => [s.kind, s.qty || null, s.amount]), [['ticket', 1, 25], ['donation', null, 30], ['ticket', 2, 80], ['donation', null, 10]]);
eq('zeffy: incomplete skipped, email lowercased, receipt # is the key', [z.skipped.length, z.sales[0].email, z.sales[0].external_id], [1, 'r@x.com', 'zeffy:SI2-1']);
const d = parseDiy([['', 'Victoria Treadwell$25February 9, 2026 2:59 PM to the incentive "Early Bird Guest List"'], ['', 'Victoria Treadwell$25February 9, 2026 2:59 PM to the incentive "Early Bird Guest List"'], ['', 'Anonymous$10January 17, 2026 7:40 PM'], ['', 'Oheyitsamanda$50February 5, 2026 6:20 PMWish I could be there!']]);
eq('diy: incentive = ticket, else donation; identical repeats keep distinct keys', [d.sales.map(s => s.kind), d.sales[0].external_id !== d.sales[1].external_id, d.sales[3].note], [['ticket', 'ticket', 'donation', 'donation'], true, 'Wish I could be there!']);

console.log('import plan');
const db = {
  guests: [{ id: 'gR', full_name: 'Reshad Alam', email: 'r@x.com' }, { id: 'gP', full_name: 'Pat Q', email: null }, { id: 'gU', full_name: 'Unsub', email: 'u@x.com' }],
  eventGea: [{ guest_id: 'gP', entry_type: 'sponsor_comp' }],
  ticketKeys: new Set(['zeffy|zeffy:SI2-1']), donationKeys: new Set(),
  subscribers: [{ id: 'su', email: 'u@x.com', status: 'unsubscribed' }],
};
const plan = planImport({ sales: [...z.sales, { kind: 'ticket', source: 'resident_advisor', external_id: '999', name: 'Unsub', email: 'u@x.com', ticket_type: 'GA', qty: 1, amount: 21.9 }], tiers: TIERS, event: DI3, db });
eq('dupe ticket skipped, others planned with tiers', [plan.report.ticketsSkipped, plan.ticketing.map(t => [t.source, t.tier_id])], [1, [['zeffy', 't40'], ['resident_advisor', 't25']]]);
eq('Pat (name match, already a sponsor comp) is recomputed, not re-inserted', [plan.geaRecompute.has('gP'), plan.geaInsert.some(r => r.guestId === 'gP')], [true, false]);
eq('unsubscribed address never re-subscribed (others are, per the import rule)', plan.subscribers.map(s => s.email), ['r@x.com', 'p@x.com']);
eq('donations planned', plan.donations.map(x => [x.payment_processor, x.amount]), [['Zeffy', 30], ['Zeffy', 10]]);
const past = planImport({ sales: z.sales, tiers: [], event: { id: 'di1', slug: 'dance-infusion-1', series: 'Dance Infusion' }, db: { ...db, ticketKeys: new Set(), eventGea: [] }, mode: 'past' });
eq('past-event import writes NO money', [past.ticketing.length, past.donations.length, past.geaInsert.length > 0], [0, 0, true]);

console.log('exec update');
const tasks = [{ id: 'x', title: 'Blocker', status: 'blocked' }, { id: 'y', title: 'Book DJ', status: 'done', updated_at: '2026-10-30T00:00:00Z' }];
const snap = execSnapshot({ event: DI3, control: ctl, tk, don, sp, mm, door: doorStats(GEA), drinks: dl, tasks, at: '2026-10-31T00:00:00Z' });
eq('snapshot lists blockers', snap.blockers, ['Blocker']);
const diff = execDiff(snap, { at: '2026-10-29T00:00:00Z', kpis: { ...snap.kpis, tickets: 2 }, openTaskIds: ['x', 'y'] }, tasks);
eq('closed since + deltas', [diff.closed.map(t => t.title), diff.deltas.tickets], [['Book DJ'], 5]);

console.log('zeffy sponsorships + held tickets');
const zs = parseZeffy([
  { 'Payment Date': '3/1/2026', 'Payment Time': '1:00 PM', 'Total Amount': 500, 'Payment Status': 'Succeeded', 'First Name': 'Theresa', 'Last Name': 'Berkman', Email: 'tb@x.com', Details: '1x Champion Sponsorship' },
  { 'Payment Date': '3/1/2026', 'Payment Time': '2:00 PM', 'Total Amount': 60, 'Payment Status': 'Succeeded', 'First Name': 'Bottle', 'Last Name': 'Buyer', Email: 'bb@x.com', Details: '1x Bottle of Tequila' },
]);
eq('a Zeffy sponsorship is a sponsor payment, never a ticket', zs.sales.map(s => s.kind), ['sponsor_payment', 'ticket']);
const pz = planImport({ sales: zs.sales, tiers: TIERS, event: DI3, db: { guests: [], eventGea: [], ticketKeys: new Set(), donationKeys: new Set(), subscribers: [], sponsorships: SP } });
eq('sponsor payment matched to the card and marked to apply', [pz.sponsorPayments[0].sponsorship.id, pz.sponsorPayments[0].apply], ['s2', true]);
eq('unmapped $60 item held: no ticket, no guest, no door row', [pz.ticketing.length, pz.guestsToCreate.length, pz.geaInsert.length, pz.held.length], [0, 0, 0, 1]);

// ---- real DI #2 exports, where present ------------------------------------
const SRC = new URL('../events/dance-infusion/di-02-2026-05/source-files/', import.meta.url);
const raFile = new URL('20260509-DanceInfusion%232-list%20(4).csv', SRC), diyFile = new URL('DIY%20Page%20transactions.csv', SRC);
if (existsSync(raFile) && existsSync(diyFile)) {
  console.log('real DI #2 exports');
  const ra = parseRa(parseCsv(readFileSync(raFile, 'utf8')).rows);
  eq('RA: 20 barcodes', ra.sales.length, 20);
  eq('RA: every GA ticket maps to the $25 tier', ra.sales.every(s => matchTier(s.ticket_type, s.amount, TIERS)?.id === 't25'), true);
  const diy = parseDiy(parseCsvRaw(readFileSync(diyFile, 'utf8')));
  eq('DIY: all 12 lines parse', [diy.sales.length, diy.skipped.length], [12, 0]);
  eq('DIY: total $445 (the DI #2 "DIY paid" figure)', diy.sales.reduce((a, s) => a + s.amount, 0), 445);
} else console.log('  (skipped: DI #2 source exports not on this machine)');

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
