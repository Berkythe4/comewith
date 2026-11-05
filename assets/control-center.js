// Event Control Center — pure logic for control.html.
// No DOM, no Supabase: control.html loads the rows, hands them here, renders
// what comes back, and executes import plans. scripts/test_control_center.mjs
// runs all of it against the real DI #2 exports and DI #3 seed shapes.
//
// Rules encoded here (DI #2 carry-overs, non-negotiable):
//  - every drink ticket counts toward the venue bar minimum, whatever its source
//  - a sponsor's drink tickets live on the guest rows they are handed to and are
//    PART of the sponsor's package — counted once, as the larger of the package
//    and what has been handed out
//  - imports and audience matching dedupe on email first, name as fallback
//  - the public metric is "% to the mission" = to mission / total raised
//  - nothing here invents a number: TBD stays null and renders as TBD

import { parseCsv, parseCsvRaw, normName } from './import-engine.js';

export { parseCsv, parseCsvRaw, normName };

const num = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };
const sum = (xs, f) => xs.reduce((a, x) => a + num(f ? f(x) : x), 0);
const round2 = (n) => Math.round(n * 100) / 100;

// ---------------------------------------------------------------------------
// Event config
// ---------------------------------------------------------------------------

export const DEFAULT_TARGET_PCT = 60;

// No event_control row = defaults. Fundraising defaults ON for Dance Infusion,
// because that is the series that exists to raise money.
export function controlFor(event, row) {
  const isDI = event && event.series === 'Dance Infusion';
  return {
    event_id: event && event.id,
    exists: !!row,
    fundraising: row ? !!row.fundraising : isDI,
    beneficiary: row ? row.beneficiary : (isDI ? 'National MS Society' : null),
    mission_target_pct: row && row.mission_target_pct != null ? num(row.mission_target_pct) : (isDI ? DEFAULT_TARGET_PCT : null),
    drink_tickets_locked: row && row.drink_tickets_locked != null ? num(row.drink_tickets_locked) : null,
    drink_ticket_value: row && row.drink_ticket_value != null ? num(row.drink_ticket_value) : null,
    comp_mission_value: row && row.comp_mission_value != null ? num(row.comp_mission_value) : null,
    notes: row ? row.notes : null,
  };
}

export function tierSplit(t) {
  const price = num(t.price), toMission = Math.min(num(t.to_mission), price);
  return { price, toMission, production: round2(price - toMission) };
}

// Calendar-day countdown in America/New_York, so "today" flips at local
// midnight, not UTC. Negative = days since.
export function daysUntil(eventDate, now = new Date()) {
  if (!eventDate) return null;
  const today = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/New_York' }).format(now);
  const a = Date.UTC(...today.split('-').map((x, i) => i === 1 ? Number(x) - 1 : Number(x)));
  const [y, m, d] = String(eventDate).slice(0, 10).split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - a) / 86400000);
}

// ---------------------------------------------------------------------------
// Tier matching (imports + the "unassigned ticket types" fixer)
// ---------------------------------------------------------------------------

// 1) an alias somebody already confirmed  2) the tier's own label  3) a "$NN"
// in the text equal to one tier's price  4) the unit price: exact, else within
// the band marketplace fees can take (RA paid out $21.90 on a $25 ticket).
// Ambiguity returns null — a wrong tier misstates the mission split.
export function matchTier(ticketType, unitPrice, tiers) {
  const live = (tiers || []).filter(t => t.active !== false);
  const label = normName(ticketType);
  if (label) {
    const byAlias = live.filter(t => (t.aliases || []).map(normName).includes(label));
    if (byAlias.length === 1) return byAlias[0];
    const byLabel = live.filter(t => normName(t.label) === label);
    if (byLabel.length === 1) return byLabel[0];
    const dollars = [...String(ticketType).matchAll(/\$\s?(\d+(?:\.\d\d)?)/g)].map(m => Number(m[1]));
    const byDollar = live.filter(t => dollars.includes(num(t.price)));
    if (byDollar.length === 1) return byDollar[0];
  }
  const u = Number(unitPrice);
  if (Number.isFinite(u) && u > 0) {
    const exact = live.filter(t => num(t.price) === u);
    if (exact.length === 1) return exact[0];
    if (exact.length > 1) return null;
    const band = live.filter(t => u <= num(t.price) * 1.02 && u >= num(t.price) * 0.8);
    if (band.length === 1) return band[0];
  }
  return null;
}

// ---------------------------------------------------------------------------
// Money: tickets, donations, sponsors -> to mission / raised
// ---------------------------------------------------------------------------

// The MS Society DIY page pays the Society directly, so the whole amount of a
// DIY ticket is mission money. Everything else: the tier's to_mission per
// ticket, face value (DI #2 counted $25 per non-comp ticket regardless of
// marketplace fees). A $0 ticket is a comp and credits nothing.
export const DIRECT_TO_MISSION_SOURCES = ['ms_diy'];
export const NON_MISSION_DONATION = /founder|out-of-pocket/i;

export function ticketMission(row, tierById) {
  const amt = num(row.amount_paid);
  if (amt <= 0) return 0;
  if (DIRECT_TO_MISSION_SOURCES.includes(row.source)) return amt;
  const t = row.tier_id && tierById[row.tier_id];
  if (!t) return null;                                   // unassigned: unknown, not 0
  // face value even when a marketplace paid out net of fees (DI #2 rule)
  return tierSplit(t).toMission * Math.max(1, num(row.quantity) || 1);
}

export function ticketStats(ticketing, tiers) {
  const tierById = Object.fromEntries((tiers || []).map(t => [t.id, t]));
  const byTier = {}, bySource = {}, unassigned = {};
  for (const t of tiers || []) byTier[t.id] = { tier: t, qty: 0, revenue: 0, mission: 0, bySource: {} };
  let qty = 0, revenue = 0, mission = 0, missionUnknownQty = 0, comps = 0;
  for (const r of ticketing || []) {
    const q = Math.max(1, num(r.quantity) || 1), amt = num(r.amount_paid);
    const m = ticketMission(r, tierById);
    qty += q; revenue += amt;
    if (amt <= 0) comps += q;
    const src = r.source || 'manual';
    const s = bySource[src] || (bySource[src] = { qty: 0, revenue: 0, last: null });
    s.qty += q; s.revenue += amt;
    if (r.created_at && (!s.last || r.created_at > s.last)) s.last = r.created_at;
    if (m == null) {
      missionUnknownQty += q;
      const k = r.ticket_type || '(no ticket type)';
      const u = unassigned[k] || (unassigned[k] = { ticket_type: k, qty: 0, revenue: 0, ids: [] });
      u.qty += q; u.revenue += amt; u.ids.push(r.id);
    } else mission += m;
    const bt = r.tier_id && byTier[r.tier_id];
    if (bt) {
      bt.qty += q; bt.revenue += amt; bt.mission += m || 0;
      bt.bySource[src] = (bt.bySource[src] || 0) + q;
    }
  }
  return { qty, revenue: round2(revenue), mission: round2(mission), missionUnknownQty, comps,
           byTier, bySource, unassigned: Object.values(unassigned) };
}

export function donationStats(donations) {
  let mission = 0, other = 0;
  const byProcessor = {};
  for (const d of donations || []) {
    const a = num(d.amount);
    if (NON_MISSION_DONATION.test(d.payment_processor || '')) other += a; else mission += a;
    const p = d.payment_processor || 'Direct';
    const b = byProcessor[p] || (byProcessor[p] = { n: 0, amount: 0, last: null });
    b.n++; b.amount += a;
    if (d.created_at && (!b.last || d.created_at > b.last)) b.last = d.created_at;
  }
  return { mission: round2(mission), other: round2(other), total: round2(mission + other), byProcessor };
}

export function sponsorStats(sponsorships, gea) {
  const live = (sponsorships || []).filter(s => s.status !== 'cancelled');
  const cards = live.map(s => {
    const slots = (gea || []).filter(g => g.sponsorship_id === s.id);
    const handed = sum(slots, g => g.drink_tickets);
    const pkg = s.drink_tickets_tbd ? null : num(s.drink_tickets);
    return {
      // comps count ENTRIES, not rows: a comp party of 2 uses 2 (228, JayJay)
      s, slots, compsTotal: num(s.entry_tickets), compsUsed: sum(slots, g => Math.max(1, num(g.quantity) || 1)),
      compsOpen: Math.max(0, num(s.entry_tickets) - sum(slots, g => Math.max(1, num(g.quantity) || 1))),
      drinkPackage: pkg, drinksHanded: handed,
      drinksLeftover: pkg == null ? null : pkg - handed,
      paid: s.status === 'paid',
    };
  });
  return {
    cards,
    pledged: round2(sum(live, s => s.cash_amount)),
    paid: round2(sum(live.filter(s => s.status === 'paid'), s => s.cash_amount)),
    compsTotal: sum(cards, c => c.compsTotal),
    compsUsed: sum(cards, c => c.compsUsed),
    unpaidCount: cards.filter(c => !c.paid && num(c.s.cash_amount) > 0).length,
  };
}

// "% to the mission" = to mission / total raised. Sponsor comps move money from
// the production share of sponsor cash to the mission only when the event sets
// comp_mission_value (DI #2 credited $25 per sponsor entry); otherwise 0.
export function missionMath({ tk, don, sp, control }) {
  const compCredit = control.comp_mission_value != null ? round2(control.comp_mission_value * sp.compsTotal) : 0;
  const toMission = round2(tk.mission + don.mission + Math.min(compCredit, sp.pledged));
  const raised = round2(tk.revenue + don.total + sp.pledged);
  const pct = raised > 0 ? toMission / raised * 100 : null;
  return { toMission, raised, production: round2(raised - toMission), pct, compCredit,
           target: control.mission_target_pct, missionUnknownQty: tk.missionUnknownQty };
}

// ---------------------------------------------------------------------------
// Drink tickets
// ---------------------------------------------------------------------------

export const DRINK_CATS = {
  sponsor: 'Sponsors', tier: 'Ticket tiers', dj: 'DJs / artists', staff: 'Staff',
  volunteer: 'Volunteers', comp: 'Comps', purchased: 'Purchased', other: 'Other',
};
const ENTRY_TO_DRINK_CAT = { dj: 'dj', artist: 'dj', staff: 'staff', volunteer: 'volunteer', vendor: 'staff',
  comp: 'comp', guest_list: 'comp', ticket: 'purchased', walk_in: 'purchased' };

export function drinkLedger({ sp, tiers, tk, gea, allocations, control, barMinimum }) {
  const lines = [];
  let tbd = 0;
  for (const c of sp.cards) {
    const name = (c.s.actor && c.s.actor.display_name) || c.s.notes || 'Sponsor';
    // A package not agreed yet still has the tickets already handed out; those
    // are real and count. Once agreed, the package counts (or what was handed
    // out, if somebody went over it — those tickets exist either way).
    const n = c.drinkPackage == null ? c.drinksHanded : Math.max(c.drinkPackage, c.drinksHanded);
    if (c.drinkPackage == null) tbd++;
    lines.push({ cat: 'sponsor', label: name, tickets: n, tbd: c.drinkPackage == null,
                 note: c.drinkPackage == null ? (c.drinksHanded ? `${c.drinksHanded} handed out; package TBD` : 'package TBD')
                       : `${c.drinksHanded} of ${c.drinkPackage} handed out` });
  }
  for (const t of tiers || []) {
    const sold = (tk.byTier[t.id] || {}).qty || 0;
    if (t.drink_tickets == null) { tbd++; lines.push({ cat: 'tier', label: t.label, tickets: 0, tbd: true, note: `${sold} sold × TBD` }); continue; }
    lines.push({ cat: 'tier', label: t.label, tickets: sold * num(t.drink_tickets), tbd: false, note: `${sold} sold × ${t.drink_tickets}` });
  }
  const people = {};
  for (const g of gea || []) {
    if (g.sponsorship_id || !num(g.drink_tickets)) continue;
    const cat = ENTRY_TO_DRINK_CAT[g.entry_type] || 'other';
    const p = people[cat] || (people[cat] = { n: 0, people: 0 });
    p.n += num(g.drink_tickets); p.people++;
  }
  for (const [cat, p] of Object.entries(people))
    lines.push({ cat, label: `${DRINK_CATS[cat]} on the door list`, tickets: p.n, tbd: false, note: `${p.people} ${p.people === 1 ? 'person' : 'people'}` });
  for (const a of allocations || [])
    lines.push({ cat: a.category, label: a.label || DRINK_CATS[a.category], tickets: num(a.tickets), tbd: false, note: a.notes || 'block allocation', id: a.id });

  const allocated = sum(lines, l => l.tickets);
  const locked = control.drink_tickets_locked;
  const value = control.drink_ticket_value;
  const bar = barMinimum != null ? num(barMinimum) : null;
  return {
    lines, allocated, locked, tbdLines: tbd,
    remaining: locked == null ? null : locked - allocated,
    overflow: locked == null ? null : Math.max(0, allocated - locked),
    barValue: value == null ? null : round2(allocated * value),
    lockedValue: value == null || locked == null ? null : round2(locked * value),
    barMinimum: bar,
    barPct: value == null || !bar ? null : allocated * value / bar * 100,
    lockedBarPct: value == null || !bar || locked == null ? null : locked * value / bar * 100,
  };
}

// ---------------------------------------------------------------------------
// Door
// ---------------------------------------------------------------------------

export const ENTRY_TYPES = {
  ticket: 'Ticket', comp: 'Comp', sponsor_comp: 'Sponsor comp', guest_list: 'Guest list',
  dj: 'DJ', artist: 'Artist', staff: 'Staff', volunteer: 'Volunteer', vendor: 'Vendor', walk_in: 'Walk-in',
};

export function doorStats(gea) {
  const rows = gea || [];
  const people = sum(rows, g => Math.max(1, num(g.quantity) || 1));
  const inside = sum(rows.filter(g => g.checked_in_at || g.attended), g => num(g.scan_count) || Math.max(1, num(g.quantity) || 1));
  const comps = rows.filter(g => ['comp', 'sponsor_comp', 'guest_list'].includes(g.entry_type));
  return { rows: rows.length, people, inside, comps: sum(comps, g => Math.max(1, num(g.quantity) || 1)) };
}

export function searchDoor(rows, q) {
  const s = normName(q);
  if (!s) return rows;
  return rows.filter(r => normName(r.name).includes(s) || (r.email || '').toLowerCase().includes(s));
}

const csvCell = (v) => { const s = v == null ? '' : String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
export function toCsv(headers, rows) {
  return [headers.map(csvCell).join(','), ...rows.map(r => r.map(csvCell).join(','))].join('\r\n') + '\r\n';
}

// ---------------------------------------------------------------------------
// Audience: Returning / New / Lapsed
// ---------------------------------------------------------------------------

// One PERSON can be several guest rows (an email row from Zeffy, a name-only row
// from a door list). Merge: same email = same person; else same normalised name
// as an email-bearing row = that person; else name alone.
export function personIndex(guests) {
  const byEmail = new Map(), byName = new Map(), keyOf = new Map(), info = new Map();
  for (const g of guests || []) {
    const e = (g.email || '').trim().toLowerCase();
    const n = normName(g.full_name);
    if (e) { if (!byEmail.has(e)) byEmail.set(e, 'e:' + e); if (n && !byName.has(n)) byName.set(n, 'e:' + e); }
  }
  for (const g of guests || []) {
    const e = (g.email || '').trim().toLowerCase();
    const n = normName(g.full_name);
    const k = e ? byEmail.get(e) : (n && byName.get(n)) || (n ? 'n:' + n : 'g:' + g.id);
    keyOf.set(g.id, k);
    const cur = info.get(k) || { key: k, name: g.full_name || '', email: e || null, guestIds: [] };
    cur.guestIds.push(g.id);
    if (!cur.email && e) cur.email = e;
    if (!cur.name && g.full_name) cur.name = g.full_name;
    info.set(k, cur);
  }
  return { keyOf, info };
}

// gea: [{guest_id, event_id}] for every event; events: [{id, name, event_date, series}]
export function audience({ gea, guests, events, currentEventId, scope = 'all' }) {
  const evById = Object.fromEntries((events || []).map(e => [e.id, e]));
  const cur = evById[currentEventId];
  const { keyOf, info } = personIndex(guests);
  const past = new Map(), now = new Set();
  for (const r of gea || []) {
    const k = keyOf.get(r.guest_id); if (!k) continue;
    if (r.event_id === currentEventId) { now.add(k); continue; }
    const e = evById[r.event_id]; if (!e || !cur) continue;
    if (!(e.event_date < cur.event_date)) continue;
    if (scope === 'series' && e.series !== cur.series) continue;
    if (!past.has(k)) past.set(k, []);
    if (!past.get(k).some(x => x.id === e.id)) past.get(k).push(e);
  }
  const person = (k) => {
    const p = info.get(k), evs = (past.get(k) || []).slice().sort((a, b) => b.event_date.localeCompare(a.event_date));
    return { ...p, events: evs, last: evs[0] || null };
  };
  const returning = [], fresh = [], lapsed = [];
  for (const k of now) (past.has(k) ? returning : fresh).push(person(k));
  for (const k of past.keys()) if (!now.has(k)) lapsed.push(person(k));
  const byLast = (a, b) => ((b.last && b.last.event_date) || '').localeCompare((a.last && a.last.event_date) || '') || a.name.localeCompare(b.name);
  return { returning: returning.sort(byLast), fresh: fresh.sort((a, b) => a.name.localeCompare(b.name)), lapsed: lapsed.sort(byLast) };
}

// Who a win-back email may go to. Never an unsubscribed address, never a guest
// who opted out — the mailing rule from every import since 2026-07-13.
export function winbackList(lapsed, subscribers, guests) {
  const subBy = new Map((subscribers || []).map(s => [(s.email || '').toLowerCase(), s]));
  const optedOut = new Set((guests || []).filter(g => g.opted_in_mailing === false && g.email).map(g => g.email.toLowerCase()));
  const out = [];
  for (const p of lapsed) {
    if (!p.email) { out.push({ ...p, status: 'no email', send: false }); continue; }
    const s = subBy.get(p.email);
    if (s && s.status === 'unsubscribed') out.push({ ...p, status: 'unsubscribed', send: false, subscriberId: s.id });
    else if (optedOut.has(p.email)) out.push({ ...p, status: 'opted out', send: false });
    else if (s && s.status === 'subscribed') out.push({ ...p, status: 'subscribed', send: true, subscriberId: s.id });
    else out.push({ ...p, status: s ? s.status : 'not on list', send: false, subscriberId: s && s.id });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Scenario planner
// ---------------------------------------------------------------------------

// counts: {tierId: n}. Face value, no fees — it is a plan, not a settlement.
export function scenario({ tiers, counts, extraDonations, expenses, sponsorCash, compCredit = 0, targetPct }) {
  let tickets = 0, mission = 0, gross = 0;
  for (const t of tiers || []) {
    const n = Math.max(0, num(counts[t.id]));
    const sp = tierSplit(t);
    tickets += n; gross += n * sp.price; mission += n * sp.toMission;
  }
  const donations = Math.max(0, num(extraDonations));
  const cash = Math.max(0, num(sponsorCash));
  const credit = Math.min(Math.max(0, num(compCredit)), cash);
  const toMission = mission + donations + credit;
  const raised = gross + donations + cash;
  const production = raised - toMission;
  const exp = Math.max(0, num(expenses));
  const pct = raised > 0 ? toMission / raised * 100 : null;
  return {
    tickets, gross: round2(gross), toMission: round2(toMission), raised: round2(raised),
    production: round2(production), expenses: round2(exp),
    covered: round2(production - exp), pct, target: targetPct,
    meetsTarget: pct != null && targetPct != null ? pct >= targetPct : null,
  };
}

// ---------------------------------------------------------------------------
// Import parsers -> one normalised "sale" shape
// ---------------------------------------------------------------------------
// sale: { kind: 'ticket'|'donation', source, external_id, name, email,
//         ticket_type, qty, amount, purchased_at, note }

const occ = () => { const seen = new Map(); return (k) => { const n = (seen.get(k) || 0) + 1; seen.set(k, n); return n === 1 ? k : k + '#' + n; }; };

export function looksLikeZeffy(headers) {
  const h = (headers || []).map(x => String(x).trim().toLowerCase());
  return h.includes('total amount') && h.includes('payment status');
}
export function looksLikeRa(headers) {
  const h = (headers || []).map(x => String(x).trim().toLowerCase());
  return h.includes('barcode') && h.includes('billing name');
}

// Zeffy "Export" sheet. Details reads like "1x Guest List" / "2x $40 Ticket, 1x
// Donations". Extra Donation is its own column. Only Succeeded rows count;
// refunds come off. There is no order id, so the key is the tax receipt number
// when present, else date+time+buyer+amount (+#n for an identical repeat).
export function parseZeffy(rows) {
  const key = occ(), out = [], skipped = [];
  for (const r of rows || []) {
    const g = (k) => r[k] == null ? '' : String(r[k]).trim();
    const status = g('Payment Status');
    const name = [g('First Name'), g('Last Name')].filter(Boolean).join(' ') || g('Company Name');
    if (status && status !== 'Succeeded') { skipped.push(`${name || '(no name)'}: payment ${status}`); continue; }
    const total = num(g('Total Amount')), extra = num(g('Extra Donation')), refund = num(g('Refund Amount'));
    const email = g('Email').toLowerCase() || null;
    const date = g('Payment Date (America/New_York)') || g('Payment Date');
    const time = g('Payment Time (America/New_York)') || g('Payment Time');
    const base = g('Tax Receipt #') ? 'zeffy:' + g('Tax Receipt #') : key(`zeffy:${date} ${time}|${email || normName(name)}|${total}`);
    const items = [...g('Details').matchAll(/(\d+)\s*x\s*([^,]+)/gi)].map(m => ({ qty: Number(m[1]), label: m[2].trim() }));
    // Zeffy sells sponsorships through the same form. Those are sponsor CASH
    // (sponsorships.cash_amount), never tickets - importing one as a ticket
    // counts the same $500 twice.
    // ...but a TICKET may carry the word too: DI #3's $40 tier sells as "Guest
    // List + Sponsorship". Anything naming a guest list / ticket is a ticket.
    const isSponsor = (l) => /sponsor/i.test(l) && !/guest\s*list|ticket/i.test(l);
    const sponsorItems = items.filter(i => isSponsor(i.label));
    const ticketItems = items.filter(i => !/donation/i.test(i.label) && !isSponsor(i.label));
    const net = Math.max(0, total - refund);
    if (net <= 0 && total > 0) { skipped.push(`${name}: fully refunded`); continue; }
    const purchased_at = toIso(date, time);
    if (sponsorItems.length && !ticketItems.length) {
      out.push({ kind: 'sponsor_payment', source: 'zeffy', external_id: base, name: name || '(unknown)', email,
                 ticket_type: sponsorItems.map(i => i.label).join(', '), amount: Math.max(0, net - extra), purchased_at, note: g('Details') || null });
      if (extra > 0) out.push({ kind: 'donation', source: 'zeffy', external_id: base + '/donation', name: name || 'Anonymous', email, amount: extra, purchased_at, note: 'Extra donation on a sponsorship' });
      continue;
    }
    if (sponsorItems.length) skipped.push(`${name}: mixes a sponsorship with tickets ("${g('Details')}") - enter it by hand`);
    if (!ticketItems.length) {
      if (net > 0) out.push({ kind: 'donation', source: 'zeffy', external_id: base, name: name || 'Anonymous', email, amount: net, purchased_at, note: g('Details') || null });
      continue;
    }
    const ticketAmount = Math.max(0, net - extra);
    const qtyAll = sum(ticketItems, i => i.qty) || 1;
    ticketItems.forEach((it, i) => out.push({
      kind: 'ticket', source: 'zeffy', external_id: ticketItems.length > 1 ? base + '/' + i : base,
      name: name || email || '(unknown)', email, ticket_type: it.label, qty: it.qty,
      amount: round2(ticketAmount * it.qty / qtyAll), purchased_at, note: null,
    }));
    if (extra > 0) out.push({ kind: 'donation', source: 'zeffy', external_id: base + '/donation', name: name || 'Anonymous', email, amount: extra, purchased_at, note: 'Extra donation on a ticket order' });
  }
  return { sales: out, skipped };
}

// RA ticket list. Same source/external_id as the event hub importer
// ('resident_advisor' + barcode) so the two can never double-count each other.
export function parseRa(rows) {
  const out = [], seen = new Map();
  for (const r of rows || []) {
    const bc = (r['Barcode'] || '').trim(); if (!bc) continue;
    const prev = seen.get(bc);
    if (prev && (prev.email || !(r['Email'] || '').trim())) continue;
    seen.set(bc, {
      kind: 'ticket', source: 'resident_advisor', external_id: bc,
      name: (r['Billing name'] || '').trim() || (r['Email'] || '').trim() || '(unknown)',
      email: (r['Email'] || '').trim().toLowerCase() || null,
      ticket_type: (r['Ticket type'] || 'General admission').trim(),
      qty: Number(r['Quantity'] || 1) || 1, amount: num(r['Price']),
      purchased_at: (r['Date purchased'] || '').trim() || null, note: null,
    });
  }
  out.push(...seen.values());
  return { sales: out, skipped: [] };
}

// MS Society DIY "transactions" export: no header, one text blob per row, e.g.
//   Victoria Treadwell$25February 9, 2026 2:59 PM to the incentive "Early Bird Guest List"
//   Anonymous$10January 17, 2026 7:40 PM
//   Oheyitsamanda$50February 5, 2026 6:20 PMWish I could be there!
// An incentive = a ticket; none = a donation. The DIY page pays the Society
// directly, so all of it is mission money (DIRECT_TO_MISSION_SOURCES).
const DIY_RE = /^(.*?)\$([\d,]+(?:\.\d{1,2})?)\s*([A-Z][a-z]+ \d{1,2}, \d{4}\s+\d{1,2}:\d{2}\s*[AP]M)\s*(?:to the incentive\s+"([^"]+)")?\s*(.*)$/s;
export function parseDiy(rawRows) {
  const key = occ(), out = [], skipped = [];
  for (const row of rawRows || []) {
    const cell = row.map(c => String(c || '').replace(/ /g, ' ').trim()).sort((a, b) => b.length - a.length)[0] || '';
    if (!cell) continue;
    const m = cell.match(DIY_RE);
    if (!m) { skipped.push(cell.slice(0, 60)); continue; }
    const name = m[1].trim() || 'Anonymous', amount = num(m[2].replace(/,/g, '')), when = m[3];
    const ext = key(`diy:${when}|${normName(name)}|${amount}`);
    const purchased_at = toIso(when);
    if (m[4]) out.push({ kind: 'ticket', source: 'ms_diy', external_id: ext, name, email: null, ticket_type: m[4].trim(), qty: 1, amount, purchased_at, note: m[5].trim() || null });
    else out.push({ kind: 'donation', source: 'ms_diy', external_id: ext, name, email: null, amount, purchased_at, note: m[5].trim() || null });
  }
  return { sales: out, skipped };
}

function toIso(date, time) {
  if (!date) return null;
  const d = new Date(time ? `${date} ${time}` : date);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

export const SOURCE_LABELS = { zeffy: 'Zeffy', resident_advisor: 'Resident Advisor', ms_diy: 'MS Society DIY', partiful: 'Partiful', manual: 'Manual', sponsor: 'Sponsor' };
export const DONATION_PROCESSOR = { zeffy: 'Zeffy', ms_diy: 'MS Society DIY' };

// ---------------------------------------------------------------------------
// Import plan
// ---------------------------------------------------------------------------
// db: { guests:[{id,full_name,email,opted_in_mailing}], eventGea:[{guest_id,entry_type}],
//       ticketKeys:Set('source|external_id'), donationKeys:Set('processor|external_ref'),
//       subscribers:[{id,email,status}] }
// mode 'live'  -> tickets + donations + door rows + mailing (the event being run)
// mode 'past'  -> door rows + mailing ONLY. A past event's money is already
//                 reconciled (DI #1 holds aggregate RA rows); writing tickets
//                 for it again would double its revenue.
export function planImport({ sales, tiers, event, db, mode = 'live' }) {
  const plan = {
    mode, guestsToCreate: [], ticketing: [], donations: [], sponsorPayments: [], held: [], geaInsert: [], geaRecompute: new Set(),
    subscribers: [], segmentEmails: new Set(),
    eventSeg: event.slug || String(event.id),
    brandSeg: event.series === 'Dance Infusion' ? 'dance_infusion' : 'come_with',
    report: { tickets: 0, ticketsSkipped: 0, donations: 0, donationsSkipped: 0, newGuests: 0, matchedByEmail: 0, matchedByName: 0, unmatchedTiers: {}, flags: [] },
  };
  const byEmail = new Map(), byName = new Map();
  for (const g of db.guests || []) {
    if (g.email) byEmail.set(g.email.toLowerCase(), g);
    const n = normName(g.full_name); if (n && !byName.has(n)) byName.set(n, g);
  }
  const onEvent = new Map((db.eventGea || []).map(r => [r.guest_id, r]));
  const subs = new Map((db.subscribers || []).map(s => [(s.email || '').toLowerCase(), s]));
  const optedOut = new Set((db.guests || []).filter(g => g.opted_in_mailing === false && g.email).map(g => g.email.toLowerCase()));
  const newByKey = new Map();
  const counted = new Set();

  const resolve = (s) => {
    if (s.email && byEmail.has(s.email)) { plan.report.matchedByEmail++; return { guestId: byEmail.get(s.email).id }; }
    const n = normName(s.name);
    if (n && byName.has(n) && !(s.email && byName.get(n).email && byName.get(n).email.toLowerCase() !== s.email)) {
      plan.report.matchedByName++;
      const g = byName.get(n);
      if (s.email && !g.email) plan.report.flags.push(`${s.name}: matched by name to a record with no email; the email in this file was not copied onto it`);
      return { guestId: g.id };
    }
    const k = s.email ? 'e:' + s.email : 'n:' + n;
    if (!newByKey.has(k)) {
      newByKey.set(k, true);
      plan.guestsToCreate.push({ key: k, full_name: s.name, email: s.email || null, source: 'import:' + s.source + ':' + plan.eventSeg });
      plan.report.newGuests++;
    }
    return { guestKey: k };
  };

  const sponsorByName = new Map();
  for (const sp of db.sponsorships || []) {
    for (const n of [sp.actor && sp.actor.display_name, sp.notes]) if (n) sponsorByName.set(normName(n), sp);
  }
  // Source keys of payments already applied to a sponsorship (224). A payment
  // can sit in the processor's export as something else - Savery's $250 is a
  // "Donations" line in Zeffy - so without this a re-import books it twice.
  const appliedRefs = new Map();
  for (const sp of db.sponsorships || []) for (const r of sp.payment_refs || []) appliedRefs.set(r, sp);
  for (const s of sales || []) {
    if ((s.kind === 'sponsor_payment' || s.kind === 'donation') && appliedRefs.has(s.external_id)) {
      const sp = appliedRefs.get(s.external_id);
      plan.report.paymentsAlreadyApplied = (plan.report.paymentsAlreadyApplied || 0) + 1;
      plan.report.flags.push(`${s.name}: ${s.amount} already applied to the ${(sp.actor && sp.actor.display_name) || 'sponsor'} sponsorship - skipped`);
      continue;
    }
    if (s.kind === 'sponsor_payment') {
      if (mode !== 'live') continue;
      const n = normName(s.name);
      const sp = sponsorByName.get(n) || [...sponsorByName.entries()].find(([k]) => n && k.includes(n))?.[1] || null;
      plan.sponsorPayments.push({ ...s, sponsorship: sp,
        apply: !!(sp && sp.status !== 'paid' && s.amount >= Number(sp.cash_amount || 0) && Number(sp.cash_amount || 0) > 0) });
      continue;
    }
    if (s.kind === 'donation') {
      if (mode !== 'live') continue;
      const proc = DONATION_PROCESSOR[s.source] || s.source;
      if ((db.donationKeys || new Set()).has(proc + '|' + s.external_id)) { plan.report.donationsSkipped++; continue; }
      plan.donations.push({ donor_name: s.name, amount: s.amount, payment_processor: proc, external_ref: s.external_id,
                            date: (s.purchased_at || '').slice(0, 10) || null, notes: s.note || null });
      plan.report.donations++;
      continue;
    }
    const dupe = (db.ticketKeys || new Set()).has(s.source + '|' + s.external_id);
    let tier = null;
    if (mode === 'live' && !dupe) {
      tier = matchTier(s.ticket_type, s.qty ? s.amount / s.qty : s.amount, tiers);
      // A ticket with no tier has no mission split, so it is HELD, not written:
      // the preview asks for a mapping (saved as a tier alias) and the next run
      // picks it up. Nothing about this person is written until then.
      if (!tier) { plan.report.unmatchedTiers[s.ticket_type] = (plan.report.unmatchedTiers[s.ticket_type] || 0) + s.qty; plan.held.push(s); continue; }
    }
    const ref = resolve(s);
    if (mode === 'live') {
      if (dupe) plan.report.ticketsSkipped++;
      else {
        plan.ticketing.push({ ...ref, external_id: s.external_id, source: s.source, ticket_type: s.ticket_type,
                              amount_paid: s.amount, quantity: s.qty, purchased_at: s.purchased_at, tier_id: tier.id,
                              notes: s.note });
        plan.report.tickets++;
      }
    }
    // the door row: one per person per event (unique guest_id, event_id)
    const pk = ref.guestId || ref.guestKey;
    if (!counted.has(pk)) {
      counted.add(pk);
      if (ref.guestId && onEvent.has(ref.guestId)) plan.geaRecompute.add(ref.guestId);
      else plan.geaInsert.push({ ...ref, entry_type: 'ticket', source: s.source, ticket_type: s.ticket_type, quantity: s.qty, amount_spent: s.amount });
    } else {
      const row = plan.geaInsert.find(r => (r.guestId || r.guestKey) === pk);
      if (row) { row.quantity += s.qty; row.amount_spent = round2(row.amount_spent + s.amount); }
    }
    if (ref.guestId) plan.geaRecompute.add(ref.guestId);
    // mailing: subscribe unless unsubscribed / opted out; tag event + brand
    if (s.email && !optedOut.has(s.email)) {
      const sub = subs.get(s.email);
      if (!sub) { if (!plan.subscribers.some(x => x.email === s.email)) plan.subscribers.push({ ...ref, email: s.email, full_name: s.name }); plan.segmentEmails.add(s.email); }
      else if (sub.status === 'subscribed') plan.segmentEmails.add(s.email);
    }
  }
  return plan;
}

// ---------------------------------------------------------------------------
// Exec update
// ---------------------------------------------------------------------------

export function execSnapshot({ event, control, tk, don, sp, mm, door, drinks, tasks, at = new Date().toISOString() }) {
  return {
    at, event: { id: event.id, name: event.name, date: event.event_date },
    kpis: {
      tickets: tk.qty, ticketRevenue: tk.revenue,
      toMission: control.fundraising ? mm.toMission : null, raised: mm.raised,
      pct: control.fundraising ? mm.pct : null, target: control.fundraising ? control.mission_target_pct : null,
      sponsorPledged: sp.pledged, sponsorPaid: sp.paid,
      comps: `${sp.compsUsed}/${sp.compsTotal}`, door: door.people, inside: door.inside,
      drinksAllocated: drinks.allocated, drinksLocked: drinks.locked, drinksTbd: drinks.tbdLines,
    },
    blockers: (tasks || []).filter(t => t.status === 'blocked').map(t => t.title),
    openTaskIds: (tasks || []).filter(t => t.status !== 'done').map(t => t.id),
  };
}

// What changed since the last saved update. "Closed" = a task that was open in
// the last snapshot and is done now (or done after it, if it was not listed).
export function execDiff(snap, prev, tasks) {
  if (!prev) return { closed: [], deltas: {}, since: null };
  const wasOpen = new Set(prev.openTaskIds || []);
  const closed = (tasks || []).filter(t => t.status === 'done' && (wasOpen.has(t.id) || (t.updated_at && t.updated_at > prev.at)));
  const deltas = {};
  for (const [k, v] of Object.entries(snap.kpis)) {
    const p = prev.kpis && prev.kpis[k];
    if (typeof v === 'number' && typeof p === 'number' && v !== p) deltas[k] = round2(v - p);
  }
  return { closed, deltas, since: prev.at };
}

export function nextMilestones(tasks, eventDate, todayIso, n = 6) {
  const open = (tasks || []).filter(t => t.status !== 'done' && t.due_date && t.due_date >= todayIso)
    .sort((a, b) => a.due_date.localeCompare(b.due_date) || (b.milestone ? 1 : 0) - (a.milestone ? 1 : 0));
  return open.slice(0, n);
}

// ---------------------------------------------------------------------------
// Roster: ONE answer to "who is tied to this event".
// ---------------------------------------------------------------------------
// Used by control.html (Command, Guest list) AND the event hub's Customers tab
// in dashboard.html, so the two can never disagree (Keith, 2026-11-05).
//
// A person is keyed by GUEST record. Lineup and sponsor ACTORS join a person
// through guests.actor_id, then by name. Role comes from entry_type (222).
// ticket_type is free text ("Sponsor comp", "DJ"), and reading it is how the
// hub used to show every comp as a paying attendee. Pre-222 rows have no
// entry_type, so they fall back to the old reading: 'comp' -> guest,
// 'donation' -> donor, anything else -> ticket.
export const ROSTER_GROUPS = [
  ['sponsor', 'Sponsors & their comps'], ['lineup', 'Lineup & crew'], ['comp', 'Comps'],
  ['ticket', 'Paid tickets'], ['donor', 'Donors'], ['other', 'Also tied to the event'],
];
export const ENTRY_LABEL = {
  ticket: 'ticket', walk_in: 'walk-in', comp: 'comp', guest_list: 'guest', sponsor_comp: 'sponsor comp',
  dj: 'dj', artist: 'artist', staff: 'staff', volunteer: 'volunteer', vendor: 'vendor', donation: 'donor',
};
const LINEUP_ROLES = new Set(['dj', 'artist', 'staff', 'volunteer', 'vendor', 'performer', 'dancer', 'host', 'crew', 'photographer']);

export function entryOf(r) {
  if (r.entry_type) return r.entry_type;
  const tt = String(r.ticket_type || '').trim().toLowerCase();
  if (tt === 'donation') return 'donation';
  if (tt === 'comp') return 'guest_list';
  return 'ticket';
}

export function eventRoster({ gea = [], sponsorships = [], participants = [], donations = [] } = {}) {
  const people = new Map(), byGuest = new Map(), byActor = new Map(), byName = new Map();
  const make = (key, name, email) => {
    const p = { key, name: (name || '').trim() || '(no name)', email: email || null, guestIds: [], actorIds: [], rows: [],
                roles: new Set(), sources: new Set(), heads: 0, spent: 0, donated: 0, sponsored: 0, fee: 0,
                attended: null, scans: 0, sponsorshipId: null, sponsorOf: [], set: null, stage: null, group: 'other' };
    people.set(key, p);
    const n = normName(name); if (n && !byName.has(n)) byName.set(n, p);
    return p;
  };
  for (const r of gea || []) {
    const g = r.guest || {};
    const gid = r.guest_id || g.id;
    const name = g.full_name ?? r.name, email = g.email ?? r.email ?? null, actorId = g.actor_id ?? r.actor_id ?? null;
    let p = gid && byGuest.get(gid);
    if (!p) {
      p = make('g:' + gid, name, email);
      if (gid) { byGuest.set(gid, p); p.guestIds.push(gid); }
    }
    if (actorId && !byActor.has(actorId)) { byActor.set(actorId, p); p.actorIds.push(actorId); }
    if (email && !p.email) p.email = email;
    p.rows.push(r);
    const e = entryOf(r);
    if (e === 'donation') { p.roles.add('donor'); p.donated += num(r.amount_spent); }
    else {
      p.roles.add(ENTRY_LABEL[e] || e);
      p.heads += Math.max(1, num(r.quantity) || 1);
      if (e === 'ticket' || e === 'walk_in') p.spent += num(r.amount_spent);
      if (e === 'sponsor_comp' && r.sponsorship_id) p.sponsorshipId = r.sponsorship_id;
    }
    if (r.source) p.sources.add(r.source);
    p.scans += num(r.scan_count);
    if (r.attended === true || r.checked_in_at) p.attended = true;
    else if (r.attended === false && p.attended !== true) p.attended = false;
  }
  const findByActor = (a) => (a && a.id && byActor.get(a.id)) || (a && byName.get(normName(a.display_name))) || null;
  for (const pp of participants || []) {
    const a = pp.actor || {};
    let p = findByActor(a);
    if (!p) p = make('a:' + (a.id || normName(a.display_name)), a.display_name, a.email);
    if (a.id && !p.actorIds.includes(a.id)) { p.actorIds.push(a.id); if (!byActor.has(a.id)) byActor.set(a.id, p); }
    (pp.roles && pp.roles.length ? pp.roles : [pp.role]).filter(Boolean).forEach(x => p.roles.add(x));
    p.fee += num(pp.fee);
    if (a.display_name && normName(a.display_name) !== normName(p.name)) p.stage = a.display_name;
    if (pp.set_start || pp.bill_order != null) p.set = { start: pp.set_start || null, end: pp.set_end || null, order: pp.bill_order ?? null };
  }
  const sponsors = [];
  for (const s of (sponsorships || []).filter(x => x.status !== 'cancelled')) {
    const a = s.actor || {};
    const name = a.display_name || s.notes || 'Sponsor';
    const p = findByActor(a);
    if (p) { p.roles.add('sponsor'); p.sponsored += num(s.cash_amount) + num(s.in_kind_value); p.sponsorOf.push(s.id); }
    sponsors.push({ id: s.id, name, tier: s.tier || null, cash: num(s.cash_amount), inKind: num(s.in_kind_value),
                    status: s.status, paid: s.status === 'paid', compsTotal: num(s.entry_tickets), person: p || null, comps: [], open: 0, s });
  }
  for (const d of donations || []) {
    let p = byName.get(normName(d.donor_name));
    if (!p) p = make('d:' + normName(d.donor_name || 'anonymous'), d.donor_name || 'Anonymous', null);
    p.roles.add('donor'); p.donated += num(d.amount);
  }
  // groups: a sponsor comp sits under its sponsor; then lineup, comps, tickets, donors
  const spById = new Map(sponsors.map(s => [s.id, s]));
  const groups = Object.fromEntries(ROSTER_GROUPS.map(([k]) => [k, []]));
  for (const p of people.values()) {
    const has = (...r) => r.some(x => p.roles.has(x));
    if (p.sponsorshipId && spById.has(p.sponsorshipId)) { p.group = 'sponsor'; spById.get(p.sponsorshipId).comps.push(p); continue; }
    if ([...p.roles].some(r => LINEUP_ROLES.has(r))) p.group = 'lineup';
    else if (has('comp', 'guest', 'sponsor comp')) p.group = 'comp';
    else if (has('ticket', 'walk-in')) p.group = 'ticket';
    else if (has('donor')) p.group = 'donor';
    else if (p.roles.has('sponsor')) { p.group = 'sponsor'; continue; }   // drawn as its sponsor header
    else p.group = 'other';
    groups[p.group].push(p);
  }
  const byNameSort = (a, b) => a.name.localeCompare(b.name);
  for (const s of sponsors) { s.comps.sort(byNameSort); s.compHeads = sum(s.comps, p => p.heads); s.open = Math.max(0, s.compsTotal - s.compHeads); }
  sponsors.sort((a, b) => (b.cash - a.cash) || a.name.localeCompare(b.name));
  // timed sets in running order, then anyone with no set (crew) by name.
  // Plain comparison, not localeCompare: collation sorts symbols before digits.
  const setKey = (p) => p.set && (p.set.start || p.set.order != null) ? '0' + (p.set.start || '') + String(p.set.order ?? 999).padStart(4, '0') : '1' + normName(p.name);
  groups.lineup.sort((a, b) => { const x = setKey(a), y = setKey(b); return x < y ? -1 : x > y ? 1 : 0; });
  for (const k of ['comp', 'ticket', 'donor', 'other']) groups[k].sort(byNameSort);
  const all = [...people.values()];
  const onList = all.filter(p => p.heads > 0);
  return {
    sponsors, groups, people: all, onList,
    counts: {
      people: onList.length, heads: sum(onList, p => p.heads),
      byGroup: Object.fromEntries(ROSTER_GROUPS.map(([k]) => [k, k === 'sponsor' ? sum(sponsors, s => s.compHeads) : groups[k].length])),
    },
  };
}

// ---------------------------------------------------------------------------
// Where we stand: one answer to "are we going to make it"
// ---------------------------------------------------------------------------
// Raised = tickets + donations + sponsor cash (pledged). What is promised to the
// mission comes off first; what is left is what can pay the costs. Expected
// cost per category is the larger of planned and booked, so a booked expense
// never double-counts the plan line it realises. TBD lines are COUNTED, never $0.
export function whereWeStand({ tk, don, sp, mm, fund, tiers = [], budget = [], expenses = [], advances = [] }) {
  const raised = round2(tk.revenue + don.total + sp.pledged);
  const received = round2(tk.revenue + don.total + sp.paid);
  const promised = fund ? round2(mm.toMission) : 0;
  const forCosts = round2(raised - promised);
  const cats = {};
  const cat = (k) => cats[k] || (cats[k] = { planned: 0, actual: 0, paid: 0, tbd: 0 });
  for (const b of budget.filter(x => x.direction === 'expense' && !x.deleted_at)) {
    const x = cat(b.category || 'Uncategorised'); if (b.amount_tbd) x.tbd++; else x.planned += num(b.planned_amount);
  }
  const live = expenses.filter(x => !x.deleted_at);
  for (const e of live) {
    const x = cat(e.category || 'Uncategorised'); x.actual += num(e.amount); if (e.status === 'paid') x.paid += num(e.amount);
  }
  const vals = Object.values(cats);
  const expected = round2(sum(vals, x => Math.max(x.planned, x.actual)));
  const booked = round2(sum(vals, x => x.actual));
  const paid = round2(sum(vals, x => x.paid));
  const ownerPaid = round2(sum(live.filter(e => e.funded_by === 'owner' && e.status === 'paid' && !e.reimbursed_at), e => e.amount));
  const owedBack = round2(sum((advances || []).filter(a => a.status === 'contingent' || a.status === 'owed'), a => a.amount));
  const gap = round2(expected - forCosts);
  // what one more of each thing adds toward costs (the lever)
  const levers = tiers.filter(t => t.active !== false).map(t => ({ label: t.label, toCosts: fund ? round2(num(t.price) - num(t.to_mission)) : num(t.price) }));
  return {
    raised, received, toCollect: round2(raised - received), promised, forCosts,
    costs: { expected, booked, paid, toPay: round2(booked - paid), planOnly: round2(expected - booked), tbd: sum(vals, x => x.tbd), cats },
    gap, covered: expected > 0 ? Math.max(0, Math.min(100, forCosts / expected * 100)) : null,
    ownerPaid, owedBack, levers,
  };
}

// ---------------------------------------------------------------------------
// RSVPs (event_rsvps, 224): a layer, never tickets or attendance
// ---------------------------------------------------------------------------
export const RSVP_STATUS = { going: 'Going', maybe: 'Maybe', cant_go: "Can't go", interested: 'Interested' };

export function rsvpStats(rsvps, roster) {
  const onList = new Set((roster ? roster.onList : []).flatMap(p => p.guestIds));
  const byStatus = { going: 0, maybe: 0, cant_go: 0, interested: 0 };
  const promoters = {};
  const goingNoTicket = [], toConfirm = [], goingOnList = [];
  for (const r of rsvps || []) {
    byStatus[r.status] = (byStatus[r.status] || 0) + 1;
    if (r.invited_by) {
      const p = promoters[r.invited_by] || (promoters[r.invited_by] = { name: r.invited_by, invited: 0, going: 0, onList: 0 });
      p.invited++; if (r.status === 'going') p.going++; if (r.guest_id && onList.has(r.guest_id)) p.onList++;
    }
    const pending = !r.guest_id && !!(r.suggested_guest_id || r.match_note);
    if (pending) toConfirm.push(r);
    if (r.status !== 'going') continue;
    if (r.guest_id && onList.has(r.guest_id)) goingOnList.push(r);
    else if (!pending) goingNoTicket.push(r);
  }
  const host = (r) => normName(r.plus_one_of || r.name);
  goingNoTicket.sort((a, b) => host(a).localeCompare(host(b)) || (a.plus_one_of ? 1 : 0) - (b.plus_one_of ? 1 : 0));
  return { total: (rsvps || []).length, byStatus, goingNoTicket, goingOnList, toConfirm,
           promoters: Object.values(promoters).sort((a, b) => b.invited - a.invited || a.name.localeCompare(b.name)) };
}

// Partiful guest export: Name, Status, RSVP date, Invited By, Explore Source, Is Plus One Of.
export function looksLikePartiful(headers) {
  const h = (headers || []).map(x => String(x).trim().toLowerCase());
  return h.includes('name') && h.includes('status') && h.includes('invited by');
}
const RSVP_FROM = { going: 'going', maybe: 'maybe', "can't go": 'cant_go', 'cant go': 'cant_go', interested: 'interested' };

// Re-importing an export UPDATES status/date/inviter on rows we have (links are
// kept). New rows auto-link only on an exact full-name match to ONE person;
// a single first name (or first + initial) that fits exactly one person on the
// guest list becomes a suggestion for a human to confirm, never a link.
// Keys match 224: 'partiful:' + normName(name), '#n' for a repeated name.
export function planRsvps({ rows, existing = [], guests = [], roster = null }) {
  const seen = new Map(), have = new Map(existing.map(r => [r.external_key, r]));
  const byFull = new Map();
  for (const g of guests) { const n = normName(g.full_name); if (!n) continue; byFull.set(n, byFull.has(n) ? null : g); }
  const listPeople = (roster ? roster.onList : []).filter(p => p.guestIds[0]);
  for (const p of listPeople) byFull.set(normName(p.name), { id: p.guestIds[0], full_name: p.name });   // the guest list wins a tie
  const inserts = [], updates = [], keys = new Set();
  const report = { rows: 0, created: 0, updated: 0, unchanged: 0, linked: 0, suggested: 0, skipped: [] };
  for (const r of rows || []) {
    const g = (k) => { const hit = Object.keys(r).find(x => x.trim().toLowerCase() === k); return hit == null ? '' : String(r[hit] ?? '').replace(/ /g, ' ').trim(); };
    const name = g('name'); if (!name) continue;
    const status = RSVP_FROM[normName(g('status'))];
    if (!status) { report.skipped.push(`${name}: unknown status "${g('status')}"`); continue; }
    report.rows++;
    const n = normName(name), c = (seen.get(n) || 0) + 1; seen.set(n, c);
    const key = 'partiful:' + n + (c === 1 ? '' : '#' + c);
    keys.add(key);
    const fields = { status, rsvp_at: g('rsvp date') || null, invited_by: g('invited by') || null, plus_one_of: g('is plus one of') || null, explore_source: g('explore source') || null };
    const prev = have.get(key);
    if (prev) {
      const t = (v) => v ? new Date(v).getTime() : null;
      const diff = Object.keys(fields).some(k => k === 'rsvp_at' ? t(prev[k]) !== t(fields[k]) : (prev[k] || null) !== fields[k]);
      if (diff) { updates.push({ id: prev.id, ...fields }); report.updated++; } else report.unchanged++;
      continue;
    }
    const row = { external_key: key, name, ...fields, guest_id: null, suggested_guest_id: null, match_note: null };
    const exact = c === 1 ? byFull.get(n) : null;
    if (exact) { row.guest_id = exact.id; report.linked++; }
    else if (c === 1 && !/'s \+1$|^friend of /i.test(name)) {
      const parts = n.replace(/\(.*?\)/g, ' ').replace(/[^a-z0-9 ]/g, ' ').trim().split(/\s+/);
      const short = parts.length === 1 || (parts.length === 2 && parts[1].length === 1);
      if (short && parts[0].length >= 2) {
        const hits = listPeople.filter(p => { const q = normName(p.name).split(' '); return q[0] === parts[0] && (parts.length === 1 || (q[1] || '').startsWith(parts[1])); });
        if (hits.length === 1) { row.suggested_guest_id = hits[0].guestIds[0]; row.match_note = `${hits[0].name}? (on the guest list)`; report.suggested++; }
        else if (hits.length > 1) { row.match_note = hits.map(h => h.name).join(' / ') + '?'; report.suggested++; }
      }
    }
    inserts.push(row); report.created++;
  }
  // rows we hold that this export no longer has (un-RSVP'd, or renamed): kept, and reported
  report.notInFile = existing.filter(e => (e.source || 'partiful') === 'partiful' && !keys.has(e.external_key)).length;
  return { inserts, updates, report };
}

// ---------------------------------------------------------------------------
// Chase list: everything that needs a person before the night, in one place
// ---------------------------------------------------------------------------
const money0 = (v) => '$' + Math.round(num(v)).toLocaleString('en-US');
export function chaseList({ sp, rsvp, expenses = [], tasks = [], todayIso, eventDate }) {
  const out = [];
  for (const k of sp.cards) {
    const nm = (k.s.actor && k.s.actor.display_name) || 'Sponsor';
    if (!k.paid && num(k.s.cash_amount) > 0) out.push({ rank: 2, sec: 'sponsors', kind: 'money', title: `Collect ${money0(k.s.cash_amount)} from ${nm}`, detail: [k.s.tier, k.s.status].filter(Boolean).join(' · ') });
    if (k.compsOpen > 0) out.push({ rank: 3, sec: 'sponsors', kind: 'comp', title: `${nm}: ${k.compsOpen} comp${k.compsOpen === 1 ? '' : 's'} not named yet`, detail: `${k.compsUsed} of ${k.compsTotal} named` });
  }
  if (rsvp && rsvp.goingNoTicket.length) out.push({ rank: 2, sec: 'rsvps', kind: 'people', title: `${rsvp.goingNoTicket.length} going on Partiful with no ticket`, detail: 'nudge them to buy' });
  if (rsvp && rsvp.toConfirm.length) out.push({ rank: 3, sec: 'rsvps', kind: 'people', title: `${rsvp.toConfirm.length} RSVP match${rsvp.toConfirm.length === 1 ? '' : 'es'} to confirm`, detail: rsvp.toConfirm.map(r => r.name).join(', ') });
  for (const e of expenses.filter(x => !x.deleted_at && x.status !== 'paid')) {
    const late = !!(e.due_date && todayIso && e.due_date < todayIso);
    out.push({ rank: late ? 0 : 2, sec: 'finance', kind: 'money', late, title: `Pay ${e.vendor || e.category || 'expense'} ${money0(e.amount)}`,
               detail: [e.category, e.due_date ? `due ${e.due_date}` : 'no due date'].filter(Boolean).join(' · ') });
  }
  for (const t of tasks.filter(x => !['done', 'cancelled'].includes(x.status) && !x.deleted_at && (!x.due_date || !eventDate || x.due_date <= eventDate))) {
    const late = !!(t.due_date && todayIso && t.due_date < todayIso);
    out.push({ rank: t.status === 'blocked' ? 1 : late ? 0 : 3, sec: 'command', kind: t.status === 'blocked' ? 'blocked' : 'task', late, taskId: t.id,
               blocked: t.status === 'blocked', title: t.title, detail: [t.status === 'blocked' ? 'blocked' : '', t.due_date ? `due ${t.due_date}` : ''].filter(Boolean).join(' · ') });
  }
  return out.sort((a, b) => a.rank - b.rank);
}
